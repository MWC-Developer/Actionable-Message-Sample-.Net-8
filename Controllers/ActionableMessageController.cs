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

using System.Text;
using System.Text.Json;
using ActionableMessageSample.Models;
using ActionableMessageSample.Options;
using ActionableMessageSample.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Tokens;

namespace ActionableMessageSample.Controllers;

[ApiController]
[Route("api/[controller]")]
/// <summary>
/// Handles actionable message callbacks and returns updated cards to Outlook.
/// </summary>
public sealed class ActionableMessageController : ControllerBase
{
    private const string CompletedActionsFieldName = "completedActions";
    private static readonly (string RouteName, string FriendlyName)[] SupportedActions =
    {
        ("action1", "Action 1"),
        ("action2", "Action 2")
    };

    private readonly IActionableMessageTokenValidator _tokenValidator;
    private readonly ActionableMessageOptions _options;
    private readonly ILogger<ActionableMessageController> _logger;

    /// <summary>
    /// Initializes a new instance of the <see cref="ActionableMessageController"/> class.
    /// </summary>
    /// <param name="tokenValidator">Service used to validate actionable message tokens.</param>
    /// <param name="options">Bound actionable message configuration.</param>
    /// <param name="logger">Logger for diagnostic events.</param>
    public ActionableMessageController(
        IActionableMessageTokenValidator tokenValidator,
        IOptions<ActionableMessageOptions> options,
        ILogger<ActionableMessageController> logger)
    {
        _tokenValidator = tokenValidator;
        _options = options.Value;
        _logger = logger;
    }

    /// <summary>
    /// Handles the Action 1 callback from Outlook.
    /// </summary>
    /// <param name="request">Payload posted by Outlook.</param>
    /// <param name="originatorId">Originator identifier from the query string.</param>
    /// <param name="cancellationToken">Token used to cancel the request.</param>
    /// <returns>Action result containing the updated card.</returns>
    [HttpPost("action1")]
    public Task<IActionResult> Action1(
        [FromBody] ActionableMessageRequest? request,
        [FromQuery(Name = "originatorId")] string? originatorId,
        CancellationToken cancellationToken)
        => HandleActionAsync("Action1", request, originatorId, cancellationToken);

    /// <summary>
    /// Handles the Action 2 callback from Outlook.
    /// </summary>
    /// <param name="request">Payload posted by Outlook.</param>
    /// <param name="originatorId">Originator identifier from the query string.</param>
    /// <param name="cancellationToken">Token used to cancel the request.</param>
    /// <returns>Action result containing the updated card.</returns>
    [HttpPost("action2")]
    public Task<IActionResult> Action2(
        [FromBody] ActionableMessageRequest? request,
        [FromQuery(Name = "originatorId")] string? originatorId,
        CancellationToken cancellationToken)
        => HandleActionAsync("Action2", request, originatorId, cancellationToken);

    /// <summary>
    /// Validates the request and builds the updated actionable card for the provided action.
    /// </summary>
    /// <param name="actionName">Name of the controller action invoked.</param>
    /// <param name="request">Payload posted by Outlook.</param>
    /// <param name="originatorId">Originator identifier from the query string.</param>
    /// <param name="cancellationToken">Token used to cancel the operation.</param>
    /// <returns>Action result indicating success or validation failure.</returns>
    private async Task<IActionResult> HandleActionAsync(
        string actionName,
        ActionableMessageRequest? request,
        string? originatorId,
        CancellationToken cancellationToken)
    {
        var validationError = await EnsureRequestIsTrustedAsync(originatorId, cancellationToken);
        if (validationError is not null)
        {
            return validationError;
        }

        if (request is null)
        {
            return BadRequest("Request payload is required.");
        }

        _logger.LogInformation("Received {Action} callback for message {MessageId}.", actionName, request.MessageId ?? "n/a");

        var payload = request.Data is { Count: > 0 }
            ? new Dictionary<string, string>(request.Data, StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        var completedActions = ExtractCompletedActions(payload);
        var routeActionName = NormalizeActionRoute(actionName);
        completedActions.Add(routeActionName);

        var card = BuildUpdatedMessageCard(routeActionName, request, payload, completedActions);
        Response.Headers["CARD-ACTION-STATUS"] = $"{GetFriendlyActionName(routeActionName)} completed.";
        Response.Headers["CARD-UPDATE-IN-BODY"] = "true";

        var serializedResponse = JsonSerializer.Serialize(card);
        _logger.LogInformation("Responding to {Action} with payload: {Response}", routeActionName, serializedResponse);

        return Ok(card);
    }

    /// <summary>
    /// Ensures the request comes from Outlook by validating the originator id and bearer token.
    /// </summary>
    /// <param name="originatorId">Originator id provided in the query string.</param>
    /// <param name="cancellationToken">Token used to cancel the validation.</param>
    /// <returns><c>null</c> when validation succeeds; otherwise an action result describing the error.</returns>
    private async Task<IActionResult?> EnsureRequestIsTrustedAsync(
        string? originatorId,
        CancellationToken cancellationToken)
    {
        _logger.LogInformation("Validating actionable message request.");

        if (string.IsNullOrWhiteSpace(_options.OriginatorId))
        {
            _logger.LogWarning("OriginatorId option is not configured.");
            return StatusCode(StatusCodes.Status500InternalServerError, "OriginatorId is not configured.");
        }

        _logger.LogInformation("OriginatorId option detected.");

        if (!string.Equals(_options.OriginatorId, originatorId, StringComparison.OrdinalIgnoreCase))
        {
            _logger.LogWarning("OriginatorId mismatch. Expected {Expected} but received {Received}.", _options.OriginatorId, originatorId);
            return BadRequest("OriginatorId mismatch.");
        }

        _logger.LogInformation("OriginatorId verified.");

        if (!Request.Headers.TryGetValue("Authorization", out var authorizationHeader))
        {
            _logger.LogWarning("Authorization header missing from request.");
            return Unauthorized("Authorization header missing.");
        }

        _logger.LogInformation("Authorization header present in request.");

        var token = ExtractBearerToken(authorizationHeader.ToString());
        if (string.IsNullOrWhiteSpace(token))
        {
            _logger.LogWarning("Bearer token missing from Authorization header.");
            return Unauthorized("Bearer token missing.");
        }

        _logger.LogInformation("Bearer token extracted successfully.");

        try
        {
            var principal = await _tokenValidator.ValidateAsync(token, cancellationToken);
            HttpContext.User = principal;
            _logger.LogInformation("Bearer token validated successfully.");
            return null;
        }
        catch (SecurityTokenException ex)
        {
            _logger.LogWarning(ex, "Actionable message token validation failed.");
            return Unauthorized("Invalid bearer token.");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unexpected error validating actionable message token.");
            return StatusCode(StatusCodes.Status500InternalServerError, "Unable to validate token at this time.");
        }
    }

    /// <summary>
    /// Extracts the completed action list from the payload dictionary.
    /// </summary>
    /// <param name="payload">Posted payload dictionary.</param>
    /// <returns>Set of completed action route names.</returns>
    private static HashSet<string> ExtractCompletedActions(IDictionary<string, string> payload)
    {
        var completed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (payload.TryGetValue(CompletedActionsFieldName, out var completedValue) &&
            !string.IsNullOrWhiteSpace(completedValue))
        {
            var entries = completedValue.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            foreach (var entry in entries)
            {
                completed.Add(entry);
            }
        }

        return completed;
    }

    /// <summary>
    /// Builds the updated actionable message card with the latest action results.
    /// </summary>
    /// <param name="completedAction">Route name of the completed action.</param>
    /// <param name="request">Original actionable request.</param>
    /// <param name="payload">Payload data submitted with the action.</param>
    /// <param name="completedActions">Set of actions already completed.</param>
    /// <returns>Dictionary representing the updated card.</returns>
    private IDictionary<string, object> BuildUpdatedMessageCard(
        string completedAction,
        ActionableMessageRequest request,
        IDictionary<string, string> payload,
        IReadOnlyCollection<string> completedActions)
    {
        var friendlyAction = GetFriendlyActionName(completedAction);
        var card = new Dictionary<string, object>
        {
            ["@type"] = "MessageCard",
            ["@context"] = "http://schema.org/extensions",
            ["summary"] = "Actionable message sample",
            ["themeColor"] = "0072C6",
            ["title"] = "Actionable message sample",
            ["text"] = $"Processed {friendlyAction} for message {request.MessageId ?? "n/a"}."
        };

        card["sections"] = new object[]
        {
            new
            {
                activityTitle = "Action details",
                facts = BuildFactEntries(completedAction, request, payload)
            }
        };

        var actions = BuildRemainingActions(completedActions);
        if (actions.Count > 0)
        {
            card["potentialAction"] = actions;
        }

        return card;
    }

    /// <summary>
    /// Builds the list of remaining actions that should stay on the card.
    /// </summary>
    /// <param name="completedActions">Actions already completed by the user.</param>
    /// <returns>Remaining action definitions.</returns>
    private List<object> BuildRemainingActions(IReadOnlyCollection<string> completedActions)
    {
        var completed = new HashSet<string>(completedActions, StringComparer.OrdinalIgnoreCase);
        var remainingActions = new List<object>();
        var completedValue = BuildCompletedActionsValue(completed);

        foreach (var (routeName, friendlyName) in SupportedActions)
        {
            if (completed.Contains(routeName))
            {
                continue;
            }

            remainingActions.Add(CreateHttpPostAction(routeName, friendlyName, completedValue));
        }

        return remainingActions;
    }

    /// <summary>
    /// Creates the serialized representation of the completed actions list.
    /// </summary>
    /// <param name="completedActions">Completed action identifiers.</param>
    /// <returns>Comma-delimited action list.</returns>
    private static string BuildCompletedActionsValue(IReadOnlyCollection<string> completedActions)
        => completedActions.Count == 0 ? string.Empty : string.Join(',', completedActions);

    /// <summary>
    /// Builds an HttpPOST potential action block for the card.
    /// </summary>
    /// <param name="routeName">Route name the action should call.</param>
    /// <param name="friendlyName">Display name for the button.</param>
    /// <param name="completedActionsValue">Serialized completed actions string.</param>
    /// <returns>Dictionary describing the action.</returns>
    private object CreateHttpPostAction(string routeName, string friendlyName, string completedActionsValue)
    {
        var target = BuildActionTarget(routeName);
        var actionBody = JsonSerializer.Serialize(new
        {
            action = routeName,
            completedActions = completedActionsValue
        });

        return new Dictionary<string, object>
        {
            ["@type"] = "HttpPOST",
            ["name"] = friendlyName,
            ["target"] = target,
            ["headers"] = new[]
            {
                new { name = "Content-Type", value = "application/json" }
            },
            ["body"] = actionBody
        };
    }

    /// <summary>
    /// Builds the callback target URL for the given route.
    /// </summary>
    /// <param name="routeName">Route segment representing the action.</param>
    /// <returns>Absolute callback URL.</returns>
    private string BuildActionTarget(string routeName)
    {
        var scheme = string.IsNullOrWhiteSpace(Request.Scheme) ? "https" : Request.Scheme;
        var host = Request.Host.HasValue ? Request.Host.Value : string.Empty;
        var pathBase = Request.PathBase.HasValue ? Request.PathBase.Value!.TrimEnd('/') : string.Empty;
        var baseUriBuilder = new StringBuilder();
        baseUriBuilder.Append(scheme);
        baseUriBuilder.Append("://");
        baseUriBuilder.Append(host);
        if (!string.IsNullOrEmpty(pathBase))
        {
            baseUriBuilder.Append(pathBase);
        }

        var baseUri = baseUriBuilder.Length > 0 ? baseUriBuilder.ToString().TrimEnd('/') : string.Empty;
        var encodedOriginator = Uri.EscapeDataString(_options.OriginatorId);
        return $"{baseUri}/api/ActionableMessage/{routeName}?originatorId={encodedOriginator}";
    }

    /// <summary>
    /// Builds the fact entries displayed in the card details section.
    /// </summary>
    /// <param name="completedAction">Completed action identifier.</param>
    /// <param name="request">Incoming actionable message request.</param>
    /// <param name="payload">Submitted payload dictionary.</param>
    /// <returns>Enumerable list of fact objects.</returns>
    private static IEnumerable<object> BuildFactEntries(
        string completedAction,
        ActionableMessageRequest request,
        IDictionary<string, string> payload)
    {
        var facts = new List<object>
        {
            new { name = "Completed action", value = GetFriendlyActionName(completedAction) }
        };

        if (!string.IsNullOrWhiteSpace(request.MessageId))
        {
            facts.Add(new { name = "Message Id", value = request.MessageId });
        }

        if (!string.IsNullOrWhiteSpace(request.UserId))
        {
            facts.Add(new { name = "User Id", value = request.UserId });
        }

        foreach (var pair in payload)
        {
            if (string.Equals(pair.Key, CompletedActionsFieldName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            facts.Add(new { name = pair.Key, value = pair.Value });
        }

        return facts;
    }

    /// <summary>
    /// Normalizes an action name so it can be compared to route names.
    /// </summary>
    /// <param name="actionName">Incoming action name.</param>
    /// <returns>Normalized action route name.</returns>
    private static string NormalizeActionRoute(string actionName)
        => string.IsNullOrWhiteSpace(actionName)
            ? string.Empty
            : actionName.Trim().ToLowerInvariant();

    /// <summary>
    /// Resolves the friendly name for a given action route.
    /// </summary>
    /// <param name="routeName">Route identifier.</param>
    /// <returns>User-friendly action name.</returns>
    private static string GetFriendlyActionName(string routeName)
        => routeName switch
        {
            "action1" => "Action 1",
            "action2" => "Action 2",
            _ => string.IsNullOrWhiteSpace(routeName) ? "Unknown action" : routeName
        };

    /// <summary>
    /// Extracts the bearer token from the Authorization header.
    /// </summary>
    /// <param name="headerValue">Authorization header value.</param>
    /// <returns>Bearer token string.</returns>
    private static string ExtractBearerToken(string headerValue)
    {
        if (string.IsNullOrWhiteSpace(headerValue))
        {
            return string.Empty;
        }

        const string prefix = "Bearer ";
        if (headerValue.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            return headerValue[prefix.Length..].Trim();
        }

        return headerValue.Trim();
    }
}
