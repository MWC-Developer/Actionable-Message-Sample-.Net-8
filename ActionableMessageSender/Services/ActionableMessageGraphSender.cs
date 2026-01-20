/*
 * By David Barrett, Microsoft Ltd. Use at your own risk.  No warranties are given.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 * */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Threading;
using ActionableMessageSender.Options;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace ActionableMessageSender.Services;

/// <summary>
/// Provides Graph sendMail helpers for delivering actionable message cards.
/// </summary>
public sealed class ActionableMessageGraphSender
{
    private static readonly HttpClient HttpClient = new();
    private static readonly SemaphoreSlim TraceFileSemaphore = new(1, 1);
    private const string ActionControllerRoute = "api/ActionableMessage";
    private const string AdaptiveCardSchema = "https://adaptivecards.io/schemas/adaptive-card.json";
    private const string CardVersion = "1.0";
    private static readonly JsonSerializerOptions ActionRequestSerializerOptions = new()
    {
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    private readonly GraphOptions _options;
    private readonly GraphTokenProvider _tokenProvider;
    private readonly ILogger<ActionableMessageGraphSender> _logger;

    /// <summary>
    /// Initializes a new instance of the <see cref="ActionableMessageGraphSender"/> class.
    /// </summary>
    /// <param name="tokenProvider">Provider used to fetch Microsoft Graph tokens.</param>
    /// <param name="options">Graph configuration bound from settings.</param>
    /// <param name="logger">Sink for diagnostic logging.</param>
    public ActionableMessageGraphSender(
        GraphTokenProvider tokenProvider,
        IOptions<GraphOptions> options,
        ILogger<ActionableMessageGraphSender> logger)
    {
        _tokenProvider = tokenProvider;
        _logger = logger;
        _options = options.Value;
    }

    /// <summary>
    /// Sends the actionable message email using Microsoft Graph.
    /// </summary>
    /// <param name="cancellationToken">Token used to cancel the send operation.</param>
    public async Task SendAsync(CancellationToken cancellationToken)
    {
        ValidateOptions();

        var token = await _tokenProvider.GetAccessTokenAsync(cancellationToken);
        var endpoint = BuildSendMailEndpoint();
        var payload = BuildSendMailPayload();

        using var request = new HttpRequestMessage(HttpMethod.Post, endpoint)
        {
            Content = new StringContent(payload, Encoding.UTF8, "application/json")
        };
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

        _logger.LogInformation("Sending actionable message via {Endpoint} using {Flow} flow.", endpoint, _options.AuthFlow);

		await TraceGraphRequestAsync(request, payload, cancellationToken);

        using var response = await HttpClient.SendAsync(request, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            var details = await response.Content.ReadAsStringAsync(cancellationToken);
            throw new InvalidOperationException($"Graph sendMail request failed ({(int)response.StatusCode}). {details}");
        }

        _logger.LogInformation("Actionable message submitted successfully.");
    }

    /// <summary>
    /// Validates that the configured options contain the required values.
    /// </summary>
    private void ValidateOptions()
    {
        if (_options.Mail.RecipientAddresses is null || _options.Mail.RecipientAddresses.Length == 0)
        {
            throw new InvalidOperationException("At least one recipient address is required.");
        }

        if (string.IsNullOrWhiteSpace(_options.ActionableMessage.ServerBaseUrl))
        {
            throw new InvalidOperationException("Server base URL is required to build action targets.");
        }

        if (string.IsNullOrWhiteSpace(_options.ActionableMessage.OriginatorId))
        {
            throw new InvalidOperationException("OriginatorId is required to use the actionable message endpoint.");
        }

        if (_options.AuthFlow == GraphAuthFlow.Application && string.IsNullOrWhiteSpace(_options.Application.SenderUserId))
        {
            throw new InvalidOperationException("Sender user id must be provided for the application flow.");
        }
    }

    /// <summary>
    /// Builds the Graph sendMail endpoint for the configured authentication flow.
    /// </summary>
    /// <returns>Absolute sendMail endpoint.</returns>
    private string BuildSendMailEndpoint()
    {
        if (_options.AuthFlow == GraphAuthFlow.Delegated)
        {
            var userId = string.IsNullOrWhiteSpace(_options.Delegated.SenderUserId) ||
                         string.Equals(_options.Delegated.SenderUserId, "me", StringComparison.OrdinalIgnoreCase)
                ? "me"
                : Uri.EscapeDataString(_options.Delegated.SenderUserId);

            return userId == "me"
                ? "https://graph.microsoft.com/v1.0/me/sendMail"
                : $"https://graph.microsoft.com/v1.0/users/{userId}/sendMail";
        }

        var sender = Uri.EscapeDataString(_options.Application.SenderUserId);
        return $"https://graph.microsoft.com/v1.0/users/{sender}/sendMail";
    }

    /// <summary>
    /// Constructs the JSON payload that contains the actionable message card and recipients.
    /// </summary>
    /// <returns>Serialized JSON payload.</returns>
    private string BuildSendMailPayload()
    {
        var htmlBody = BuildHtmlBody();
        var payload = new
        {
            message = new
            {
                subject = _options.Mail.Subject,
                body = new
                {
                    contentType = "HTML",
                    content = htmlBody
                },
                toRecipients = _options.Mail.RecipientAddresses
                    .Select(address => new
                    {
                        emailAddress = new { address }
                    })
                    .ToArray()
            },
            saveToSentItems = true
        };

        return JsonSerializer.Serialize(
            payload,
            new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            });
    }

    /// <summary>
    /// Builds the HTML body with an embedded actionable message card.
    /// </summary>
    /// <returns>HTML string containing the message card.</returns>
    private string BuildHtmlBody()
    {
        var stateKey = Guid.NewGuid().ToString("N");

        var card = BuildAdaptiveCard(stateKey);

        var cardJson = JsonSerializer.Serialize(
            card,
            new JsonSerializerOptions
            {
                PropertyNamingPolicy = null,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            });

        return $@"<html><body><p>This email contains an actionable message.</p><script type=""application/adaptivecard+json"">{cardJson}</script></body></html>";
    }

    private IDictionary<string, object> BuildAdaptiveCard(string stateKey)
    {
        var bodyElements = BuildInitialBodyElements();
        var actionDefinitions = BuildInitialActions(stateKey);

        return CreateCardEnvelope(stateKey, bodyElements, actionDefinitions, CardVersion);
    }

    private static List<object> BuildInitialBodyElements()
        => new()
        {
            new { type = "TextBlock", size = "Medium", weight = "Bolder", text = "Actionable message sample" },
            new { type = "TextBlock", text = "Choose one of the actions below to call back into the API.", wrap = true }
        };

    private List<object> BuildInitialActions(string stateKey)
        => new()
        {
            CreateHttpAction("action1", "Action 1", stateKey),
            CreateHttpAction("action2", "Action 2", stateKey)
        };

    private IDictionary<string, object> CreateCardEnvelope(
        string cardId,
        IEnumerable<object> bodyElements,
        IEnumerable<object> actions,
        string cardVersion)
    {
        var envelope = new Dictionary<string, object>
        {
            ["type"] = "AdaptiveCard",
            ["$schema"] = AdaptiveCardSchema,
            ["version"] = cardVersion,
            ["originator"] = _options.ActionableMessage.OriginatorId,
            ["cardId"] = cardId,
            ["body"] = bodyElements.ToArray()
        };

        var actionArray = actions.ToArray();
        if (actionArray.Length > 0)
        {
            envelope["actions"] = actionArray;
        }

        return envelope;
    }

    /// <summary>
    /// Creates an Action.Http block for the actionable message card.
    /// </summary>
    private IDictionary<string, object> CreateHttpAction(string actionName, string friendlyName, string stateKey)
    {
        var actionData = new Dictionary<string, string?>
        {
            ["stateKey"] = stateKey,
            ["originatorId"] = _options.ActionableMessage.OriginatorId,
            ["completedActions"] = string.Empty,
            ["action"] = actionName,
            ["verb"] = actionName
        };

        var requestBody = SerializeActionRequest(actionName, stateKey, actionData);

        return new Dictionary<string, object>
        {
            ["type"] = "Action.Http",
            ["title"] = friendlyName,
            ["method"] = "POST",
            ["url"] = BuildActionUrl(actionName),
            ["body"] = requestBody,
			["headers"] = new object[]
			{
				new { name = "Content-Type", value = "application/json" }
			}
        };
    }

    private string BuildActionUrl(string routeName)
    {
        if (string.IsNullOrWhiteSpace(routeName))
        {
            throw new ArgumentException("Route name is required.", nameof(routeName));
        }

        var baseUrl = _options.ActionableMessage.ServerBaseUrl.TrimEnd('/');
        var normalizedRoute = routeName.Trim('/');
        return $"{baseUrl}/{ActionControllerRoute}/{normalizedRoute}";
    }

    private static string SerializeActionRequest(string actionId, string cardId, IDictionary<string, string?> payload)
    {
        var request = new Dictionary<string, object?>
        {
            ["actionId"] = actionId,
            ["cardId"] = cardId,
            ["data"] = payload
        };

        return JsonSerializer.Serialize(request, ActionRequestSerializerOptions);
    }

	private async Task TraceGraphRequestAsync(HttpRequestMessage request, string requestBody, CancellationToken cancellationToken)
	{
		var traceFilePath = _options.TraceFilePath;
		if (string.IsNullOrWhiteSpace(traceFilePath))
		{
			return;
		}

		var logBuilder = new StringBuilder();
		logBuilder.AppendLine($"[{DateTimeOffset.UtcNow:O}] {request.Method} {request.RequestUri}");
		AppendHeaders(logBuilder, request.Headers);
		AppendHeaders(logBuilder, request.Content?.Headers);
		logBuilder.AppendLine();
		logBuilder.AppendLine(requestBody);
		logBuilder.AppendLine(new string('-', 80));

		var directory = Path.GetDirectoryName(traceFilePath);
		if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
		{
			Directory.CreateDirectory(directory);
		}

		await TraceFileSemaphore.WaitAsync(cancellationToken);
		try
		{
			await File.AppendAllTextAsync(traceFilePath, logBuilder.ToString(), cancellationToken);
		}
		finally
		{
			TraceFileSemaphore.Release();
		}
	}

	private static void AppendHeaders(StringBuilder builder, HttpHeaders? headers)
	{
		if (headers is null)
		{
			return;
		}

		foreach (var header in headers)
		{
			var value = string.Join(", ", header.Value);
			if (string.Equals(header.Key, "Authorization", StringComparison.OrdinalIgnoreCase))
			{
				value = MaskAuthorizationHeader(value);
			}

			builder.AppendLine($"{header.Key}: {value}");
		}
	}

	private static string MaskAuthorizationHeader(string headerValue)
	{
		if (string.IsNullOrWhiteSpace(headerValue))
		{
			return string.Empty;
		}

		const string bearerPrefix = "Bearer ";
		return headerValue.StartsWith(bearerPrefix, StringComparison.OrdinalIgnoreCase)
			? $"{bearerPrefix}***"
			: "***";
	}
}
