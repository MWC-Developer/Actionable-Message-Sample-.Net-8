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
using System.Linq;
using System.Security.Claims;
using System.Text.Encodings.Web;
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
    private const string StateKeyFieldName = "stateKey";
    private const string ActionControllerRoute = "api/ActionableMessage";
    private const string AdaptiveCardSchema = "https://adaptivecards.io/schemas/adaptive-card.json";
    private const string CardVersion = "1.0";
    private static readonly JsonSerializerOptions ActionRequestSerializerOptions = new()
    {
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };
    private static readonly (string RouteName, string FriendlyName)[] SupportedActions =
    {
        ("action1", "Action 1"),
        ("action2", "Action 2")
    };

    private readonly IActionableMessageTokenValidator _tokenValidator;
    private readonly ActionableMessageOptions _options;
    private readonly ILogger<ActionableMessageController> _logger;
    private readonly IActionableMessageStateStore _stateStore;

    /// <summary>
    /// Initializes a new instance of the <see cref="ActionableMessageController"/> class.
    /// </summary>
    /// <param name="tokenValidator">Service used to validate actionable message tokens.</param>
    /// <param name="options">Bound actionable message configuration.</param>
    /// <param name="logger">Logger for diagnostic events.</param>
    /// <param name="stateStore">In-memory store for per-message action state.</param>
    public ActionableMessageController(
        IActionableMessageTokenValidator tokenValidator,
        IOptions<ActionableMessageOptions> options,
        ILogger<ActionableMessageController> logger,
        IActionableMessageStateStore stateStore)
    {
        _tokenValidator = tokenValidator;
        _options = options.Value;
        _logger = logger;
        _stateStore = stateStore;
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
	/// Handles Action.Execute callbacks emitted by Adaptive Cards.
	/// </summary>
	[HttpPost("execute")]
	public Task<IActionResult> Execute(
		[FromBody] ActionableMessageRequest? request,
		[FromQuery(Name = "originatorId")] string? originatorId,
		CancellationToken cancellationToken)
	{
		var resolvedAction = ResolveActionName(request);
		if (string.IsNullOrWhiteSpace(resolvedAction))
		{
			return Task.FromResult<IActionResult>(BadRequest("Action identifier is required."));
		}

		var shouldRecord = !string.Equals(resolvedAction, "refresh", StringComparison.OrdinalIgnoreCase);
		return HandleActionAsync(resolvedAction, request, originatorId, cancellationToken, shouldRecord);
	}

    /// <summary>
    /// Handles refresh callbacks so Outlook can display the latest card state when reopening the message.
    /// </summary>
    /// <param name="request">Payload posted by Outlook.</param>
    /// <param name="originatorId">Originator identifier from the query string.</param>
    /// <param name="cancellationToken">Token used to cancel the request.</param>
    /// <returns>Action result containing the refreshed card.</returns>

	[HttpPost("refresh")]
	public Task<IActionResult> Refresh(
		[FromBody] ActionableMessageRequest? request,
		[FromQuery(Name = "originatorId")] string? originatorId,
		CancellationToken cancellationToken)
		=> HandleActionAsync("Refresh", request, originatorId, cancellationToken, recordAction: false);

	/// <summary>
	/// Handles refresh callbacks issued via HTTP GET (used by some Outlook clients during auto-refresh).
	/// </summary>
	/// <param name="originatorId">Originator identifier from the query string.</param>
	/// <param name="cancellationToken">Token used to cancel the request.</param>
	/// <returns>Action result containing the refreshed card.</returns>
	[HttpGet("refresh")]
	public Task<IActionResult> RefreshGet(
		[FromQuery(Name = "originatorId")] string? originatorId,
		CancellationToken cancellationToken)
	{
		var syntheticRequest = new ActionableMessageRequest
		{
			Data = new Dictionary<string, string>()
		};

		return HandleActionAsync("Refresh", syntheticRequest, originatorId, cancellationToken, recordAction: false);
	}

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
		CancellationToken cancellationToken,
		bool recordAction = true)
	{
		if (request is null)
		{
			return BadRequest("Request payload is required.");
		}

		var payload = request.Data is { Count: > 0 }
			? new Dictionary<string, string>(request.Data, StringComparer.OrdinalIgnoreCase)
			: new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

		var effectiveOriginatorId = ResolveOriginatorId(originatorId, payload);
		var (validationError, validatedOriginatorId) = await EnsureRequestIsTrustedAsync(effectiveOriginatorId, cancellationToken);
		if (validationError is not null)
		{
			return validationError;
		}

		if (validatedOriginatorId is null)
		{
			_logger.LogError("Validated originator id missing after successful trust check.");
			return StatusCode(StatusCodes.Status500InternalServerError, "Unable to process originator id.");
		}

		var effectiveUserId = ResolveUserId(request);

		_logger.LogInformation("Received {Action} callback for message {MessageId}.", actionName, request.MessageId ?? "n/a");

		var payloadStateKeyPresent = payload.TryGetValue(StateKeyFieldName, out var payloadStateKey);
		_logger.LogInformation(
			"Payload inspection for {Action}: keys={Keys}; containsCompletedActions={HasCompletedActions}; containsStateKey={HasStateKey}; stateKeyValue={StateKeyValue}",
			actionName,
			payload.Count == 0 ? "<none>" : string.Join(",", payload.Keys),
			payload.ContainsKey(CompletedActionsFieldName),
			payloadStateKeyPresent,
			payloadStateKey ?? "<null>");

		var queryStateKey = ExtractStateKeyFromQuery();
		var stateKey = ResolveStateKey(request, payload, queryStateKey);
		if (!string.IsNullOrWhiteSpace(stateKey))
		{
			payload[StateKeyFieldName] = stateKey!;
		}
		else
		{
			_logger.LogWarning("State key missing for action {Action} and message {MessageId}.", actionName, request.MessageId ?? "n/a");
		}

		var messageKey = BuildMessageStoreKey(stateKey, effectiveUserId);
		_logger.LogInformation("Message store key resolution for {Action}: stateKey={StateKey}; userId={UserId}; storeKey={StoreKey}", actionName, stateKey ?? "<null>", request.UserId ?? "<null>", messageKey ?? "<null>");
		HashSet<string> completedActions;

		if (messageKey is not null)
		{
			completedActions = new HashSet<string>(_stateStore.GetCompletedActions(messageKey), StringComparer.OrdinalIgnoreCase);
			_logger.LogInformation("Retrieved {Count} completed actions for store key {StoreKey}: {Actions}", completedActions.Count, messageKey, completedActions.Count == 0 ? "<none>" : string.Join(",", completedActions));
		}
		else
		{
			_logger.LogWarning("Unable to resolve message key for message {MessageId}. Falling back to payload state only.", request.MessageId ?? "n/a");
			completedActions = ExtractCompletedActions(payload);
		}

        var routeActionName = NormalizeActionRoute(actionName);
        var friendlyActionName = GetFriendlyActionName(routeActionName);
        var actionRecorded = true;

		if (recordAction && !string.IsNullOrEmpty(routeActionName))
        {
			if (messageKey is not null)
			{
				actionRecorded = _stateStore.TryMarkActionCompleted(messageKey, routeActionName);
				completedActions = new HashSet<string>(_stateStore.GetCompletedActions(messageKey), StringComparer.OrdinalIgnoreCase);
				_logger.LogInformation(
					"Attempted to mark action {Action} for store key {StoreKey}. Recorded={Recorded}. Completed actions now: {Actions}",
					routeActionName,
					messageKey,
					actionRecorded,
					completedActions.Count == 0 ? "<none>" : string.Join(",", completedActions));

				if (!actionRecorded)
				{
					_logger.LogInformation("Action {Action} already completed for message key {MessageKey}.", routeActionName, messageKey);
				}
			}
			else
			{
				_logger.LogWarning("Recording action {Action} using transient payload state because store key is unavailable.", routeActionName);
				completedActions.Add(routeActionName);
			}
        }

		var statusMessage = recordAction
			? (actionRecorded ? $"{friendlyActionName} completed." : $"{friendlyActionName} was already completed.")
			: "Card refreshed.";

		var card = BuildUpdatedMessageCard(
			recordAction ? routeActionName : null,
			request,
			payload,
			completedActions,
			validatedOriginatorId,
			effectiveUserId,
			statusMessage,
			stateKey);

        Response.Headers["CARD-ACTION-STATUS"] = statusMessage;
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
	/// <returns>Tuple containing either the validation error or the verified originator id.</returns>
	private async Task<(IActionResult? Error, string? ValidatedOriginatorId)> EnsureRequestIsTrustedAsync(
		string? originatorId,
		CancellationToken cancellationToken)
	{
		_logger.LogInformation("Validating actionable message request.");

		var configuredOriginators = _options.GetConfiguredOriginatorIds();
		if (configuredOriginators.Count == 0)
		{
			_logger.LogWarning("No originator ids are configured.");
			return (StatusCode(StatusCodes.Status500InternalServerError, "OriginatorId is not configured."), null);
		}

		if (string.IsNullOrWhiteSpace(originatorId))
		{
			_logger.LogWarning("OriginatorId missing from request.");
			return (BadRequest("OriginatorId is required."), null);
		}

		var matchedOriginator = configuredOriginators.FirstOrDefault(
			configured => string.Equals(configured, originatorId, StringComparison.OrdinalIgnoreCase));
		if (matchedOriginator is null)
		{
			_logger.LogWarning("OriginatorId mismatch. Received {Received} but not in configured list.", originatorId);
			return (BadRequest("OriginatorId mismatch."), null);
		}

		_logger.LogInformation("OriginatorId verified.");

		if (!Request.Headers.TryGetValue("Authorization", out var authorizationHeader))
		{
			_logger.LogWarning("Authorization header missing from request.");
			return (Unauthorized("Authorization header missing."), null);
		}

		_logger.LogInformation("Authorization header present in request.");

		var token = ExtractBearerToken(authorizationHeader.ToString());
		if (string.IsNullOrWhiteSpace(token))
		{
			_logger.LogWarning("Bearer token missing from Authorization header.");
			return (Unauthorized("Bearer token missing."), null);
		}

		_logger.LogInformation("Bearer token extracted successfully.");

		try
		{
			var principal = await _tokenValidator.ValidateAsync(token, cancellationToken);
			HttpContext.User = principal;
			_logger.LogInformation("Bearer token validated successfully.");
			return (null, matchedOriginator);
		}
		catch (SecurityTokenException ex)
		{
			_logger.LogWarning(ex, "Actionable message token validation failed.");
			return (Unauthorized("Invalid bearer token."), null);
		}
		catch (Exception ex)
		{
			_logger.LogError(ex, "Unexpected error validating actionable message token.");
			return (StatusCode(StatusCodes.Status500InternalServerError, "Unable to validate token at this time."), null);
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
		string? completedAction,
		ActionableMessageRequest request,
		IDictionary<string, string> payload,
		IReadOnlyCollection<string> completedActions,
		string originatorId,
		string? effectiveUserId,
		string statusMessage,
		string? stateKey)
	{
		var resolvedCardId = !string.IsNullOrWhiteSpace(stateKey)
			? stateKey!
			: (!string.IsNullOrWhiteSpace(request.CardId)
				? request.CardId!
				: (!string.IsNullOrWhiteSpace(request.MessageId)
					? request.MessageId!
					: Guid.NewGuid().ToString("N")));
		var actionStateKey = string.IsNullOrWhiteSpace(stateKey) ? resolvedCardId : stateKey;

		var bodyElements = new List<object>
		{
			new { type = "TextBlock", size = "Medium", weight = "Bolder", text = "Actionable message sample" },
			new { type = "TextBlock", text = statusMessage, wrap = true }
		};

		var facts = BuildFactEntries(completedAction, request, payload, effectiveUserId).ToArray();
		if (facts.Length > 0)
		{
			bodyElements.Add(new { type = "FactSet", facts });
		}

		var actions = BuildRemainingActions(completedActions, originatorId, actionStateKey);
		var card = CreateCardEnvelope(
			resolvedCardId,
			originatorId,
			bodyElements,
			actions,
			CardVersion);

        return card;
    }

    /// <summary>
    /// Builds the list of remaining actions that should stay on the card.
    /// </summary>
    /// <param name="completedActions">Actions already completed by the user.</param>
    /// <returns>Remaining action definitions.</returns>
	private List<object> BuildRemainingActions(IReadOnlyCollection<string> completedActions, string originatorId, string? stateKey)
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

		remainingActions.Add(CreateHttpAction(routeName, friendlyName, completedValue, originatorId, stateKey));
        }

        return remainingActions;
    }

	private static IDictionary<string, object> CreateCardEnvelope(
		string cardId,
		string originatorId,
		IEnumerable<object> bodyElements,
		IEnumerable<object> actions,
		string version)
	{
		var envelope = new Dictionary<string, object>
		{
			["type"] = "AdaptiveCard",
			["$schema"] = AdaptiveCardSchema,
			["version"] = version,
			["originator"] = originatorId,
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
    /// Creates the serialized representation of the completed actions list.
    /// </summary>
    /// <param name="completedActions">Completed action identifiers.</param>
    /// <returns>Comma-delimited action list.</returns>
    private static string BuildCompletedActionsValue(IReadOnlyCollection<string> completedActions)
        => completedActions.Count == 0 ? string.Empty : string.Join(',', completedActions);

    /// <summary>
    /// Builds an Action.Http block for the card.
    /// </summary>
    /// <param name="routeName">Route name the action should call.</param>
    /// <param name="friendlyName">Display name for the button.</param>
    /// <param name="completedActionsValue">Serialized completed actions string.</param>
    /// <returns>Dictionary describing the action.</returns>
	private object CreateHttpAction(string routeName, string friendlyName, string completedActionsValue, string originatorId, string? stateKey)
    {
		var actionPayload = new Dictionary<string, string?>
		{
			["action"] = routeName,
			[CompletedActionsFieldName] = completedActionsValue,
			["originatorId"] = originatorId,
			["verb"] = routeName
		};

		if (!string.IsNullOrWhiteSpace(stateKey))
		{
			actionPayload[StateKeyFieldName] = stateKey;
		}

		var body = SerializeActionRequest(routeName, stateKey, actionPayload);

		return new Dictionary<string, object>
		{
			["type"] = "Action.Http",
			["title"] = friendlyName,
			["method"] = "POST",
			["url"] = BuildActionUrl(routeName),
			["body"] = body,
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

		var baseUrl = !string.IsNullOrWhiteSpace(_options.CallbackBaseUrl)
			? _options.CallbackBaseUrl!.TrimEnd('/')
			: $"{Request.Scheme}://{Request.Host}".TrimEnd('/');

		var normalizedRoute = routeName.Trim('/');
		return $"{baseUrl}/{ActionControllerRoute}/{normalizedRoute}";
	}

	private static string SerializeActionRequest(string actionId, string? cardId, IDictionary<string, string?> payload)
	{
		var request = new Dictionary<string, object?>
		{
			["actionId"] = actionId,
			["cardId"] = cardId,
			["data"] = payload
		};

		return JsonSerializer.Serialize(request, ActionRequestSerializerOptions);
	}


    /// <summary>
    /// Builds the callback target URL for the given route.
    /// </summary>
    /// <param name="routeName">Route segment representing the action.</param>
    /// <returns>Absolute callback URL.</returns>

    /// <summary>
    /// Builds the fact entries displayed in the card details section.
    /// </summary>
    /// <param name="completedAction">Completed action identifier.</param>
    /// <param name="request">Incoming actionable message request.</param>
    /// <param name="payload">Submitted payload dictionary.</param>
    /// <returns>Enumerable list of fact objects.</returns>
	private static IEnumerable<object> BuildFactEntries(
		string? completedAction,
		ActionableMessageRequest request,
		IDictionary<string, string> payload,
		string? effectiveUserId)
    {
		var facts = new List<object>();

		if (!string.IsNullOrWhiteSpace(completedAction))
		{
			facts.Add(new { title = "Completed action", value = GetFriendlyActionName(completedAction) });
		}

        if (!string.IsNullOrWhiteSpace(request.MessageId))
        {
			facts.Add(new { title = "Message Id", value = request.MessageId });
        }

		var userIdToDisplay = !string.IsNullOrWhiteSpace(request.UserId) ? request.UserId : effectiveUserId;
		if (!string.IsNullOrWhiteSpace(userIdToDisplay))
		{
			facts.Add(new { title = "User Id", value = userIdToDisplay });
		}

		foreach (var pair in payload)
        {
			if (string.Equals(pair.Key, CompletedActionsFieldName, StringComparison.OrdinalIgnoreCase) ||
			    string.Equals(pair.Key, StateKeyFieldName, StringComparison.OrdinalIgnoreCase) ||
			    string.Equals(pair.Key, "originatorId", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

			facts.Add(new { title = pair.Key, value = pair.Value });
        }

        return facts;
    }

	private static string ResolveActionName(ActionableMessageRequest? request)
	{
		if (request is null)
		{
			return string.Empty;
		}

		if (!string.IsNullOrWhiteSpace(request.ActionId))
		{
			return request.ActionId;
		}

		if (request.Data is { Count: > 0 })
		{
			if (request.Data.TryGetValue("action", out var actionValue) && !string.IsNullOrWhiteSpace(actionValue))
			{
				return actionValue;
			}

			if (request.Data.TryGetValue("verb", out var verbValue) && !string.IsNullOrWhiteSpace(verbValue))
			{
				return verbValue;
			}
		}

		return string.Empty;
	}

	private static string? ResolveOriginatorId(string? originatorId, IDictionary<string, string> payload)
	{
		if (!string.IsNullOrWhiteSpace(originatorId))
		{
			return originatorId;
		}

		if (payload.TryGetValue("originatorId", out var payloadOriginator) && !string.IsNullOrWhiteSpace(payloadOriginator))
		{
			return payloadOriginator;
		}

		return null;
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

	/// <summary>
	/// Resolves a stable state key for the actionable message instance.
	/// </summary>
	/// <param name="request">Incoming actionable message request.</param>
	/// <param name="payload">Payload sent alongside the action.</param>
	/// <returns>Underlying state identifier or <c>null</c> when missing.</returns>
	private static string? ResolveStateKey(ActionableMessageRequest request, IDictionary<string, string> payload, string? queryStateKey)
	{
		if (request is null)
		{
			return null;
		}

		if (!string.IsNullOrWhiteSpace(queryStateKey))
		{
			return queryStateKey;
		}

		if (payload.TryGetValue(StateKeyFieldName, out var explicitStateKey) && !string.IsNullOrWhiteSpace(explicitStateKey))
		{
			return explicitStateKey;
		}

		if (!string.IsNullOrWhiteSpace(request.CardId))
		{
			return request.CardId;
		}

		if (!string.IsNullOrWhiteSpace(request.MessageId))
		{
			return request.MessageId;
		}

		return null;
	}

	private string? ExtractStateKeyFromQuery()
	{
		if (Request.Query.TryGetValue(StateKeyFieldName, out var stateKeyValues) && stateKeyValues.Count > 0)
		{
			var stateKey = stateKeyValues[0];
			if (!string.IsNullOrWhiteSpace(stateKey))
			{
				return stateKey;
			}
		}

		return null;
	}

	/// <summary>
	/// Builds the per-user message state key used to track completed actions in memory.
	/// </summary>
	/// <param name="stateKey">Underlying message instance key.</param>
	/// <param name="userId">User executing the action.</param>
	/// <returns>Composite key or <c>null</c> if insufficient data.</returns>
	private static string? BuildMessageStoreKey(string? stateKey, string? userId)
	{
		if (string.IsNullOrWhiteSpace(stateKey))
		{
			return null;
		}

		return string.IsNullOrWhiteSpace(userId)
			? stateKey
			: $"{stateKey}:{userId}";
	}

	private string? ResolveUserId(ActionableMessageRequest request)
	{
		if (request is not null && !string.IsNullOrWhiteSpace(request.UserId))
		{
			return request.UserId;
		}

		var principal = HttpContext.User;
		if (principal?.Identity is not { IsAuthenticated: true })
		{
			return null;
		}

		var preferredUsername = principal.FindFirst("preferred_username")?.Value;
		if (!string.IsNullOrWhiteSpace(preferredUsername))
		{
			return preferredUsername;
		}

		var upn = principal.FindFirst("upn")?.Value;
		if (!string.IsNullOrWhiteSpace(upn))
		{
			return upn;
		}

		var email = principal.FindFirst(ClaimTypes.Email)?.Value;
		if (!string.IsNullOrWhiteSpace(email))
		{
			return email;
		}

		var name = principal.FindFirst(ClaimTypes.Name)?.Value;
		return string.IsNullOrWhiteSpace(name) ? null : name;
	}
}
