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

using System.Net.Http.Headers;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
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
        var card = new Dictionary<string, object>
        {
            ["@type"] = "MessageCard",
            ["@context"] = "http://schema.org/extensions",
            ["summary"] = "Actionable message sample",
            ["themeColor"] = "0072C6",
            ["title"] = "Actionable message sample",
            ["text"] = "Choose one of the actions below to call back into the API.",
            ["potentialAction"] = new object[]
            {
                CreateAction("action1", "Action 1"),
                CreateAction("action2", "Action 2")
            }
        };

        var cardJson = JsonSerializer.Serialize(
            card,
            new JsonSerializerOptions
            {
                PropertyNamingPolicy = null,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            });

        return $@"<html><body><p>This email contains an actionable message.</p><script type=""application/ld+json"">{cardJson}</script></body></html>";
    }

    /// <summary>
    /// Creates an HttpPOST action block for the actionable message card.
    /// </summary>
    /// <param name="actionName">Internal action identifier.</param>
    /// <param name="friendlyName">Display name shown to recipients.</param>
    /// <returns>Dictionary representing the action JSON.</returns>
    private IDictionary<string, object> CreateAction(string actionName, string friendlyName)
    {
        var target = BuildActionTarget(actionName);
        var body = JsonSerializer.Serialize(new { action = actionName, completedActions = string.Empty });

        return new Dictionary<string, object>
        {
            ["@type"] = "HttpPOST",
            ["name"] = friendlyName,
            ["target"] = target,
            ["headers"] = new[]
            {
                new { name = "Content-Type", value = "application/json" }
            },
            ["body"] = body
        };
    }

    /// <summary>
    /// Builds the callback URL used by Microsoft Graph when an action is triggered.
    /// </summary>
    /// <param name="actionName">Name of the triggered action.</param>
    /// <returns>Absolute callback URL.</returns>
    private string BuildActionTarget(string actionName)
    {
        var trimmedBase = _options.ActionableMessage.ServerBaseUrl.TrimEnd('/');
        var encodedOriginator = Uri.EscapeDataString(_options.ActionableMessage.OriginatorId);
        return $"{trimmedBase}/api/ActionableMessage/{actionName}?originatorId={encodedOriginator}";
    }
}
