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

using System.Text.Json;
using ActionableMessageSender.Options;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace ActionableMessageSender.Services;

/// <summary>
/// Handles acquiring Microsoft Graph access tokens for both application and delegated flows.
/// </summary>
public sealed class GraphTokenProvider
{
    private static readonly HttpClient HttpClient = new();

    private readonly GraphOptions _options;
    private readonly ILogger<GraphTokenProvider> _logger;

    /// <summary>
    /// Initializes a new instance of the <see cref="GraphTokenProvider"/> class.
    /// </summary>
    /// <param name="options">Graph configuration containing auth settings.</param>
    /// <param name="logger">Logger used for diagnostics.</param>
    public GraphTokenProvider(IOptions<GraphOptions> options, ILogger<GraphTokenProvider> logger)
    {
        _options = options.Value;
        _logger = logger;
    }

    /// <summary>
    /// Gets a Microsoft Graph access token for the configured authentication flow.
    /// </summary>
    /// <param name="cancellationToken">Token used to cancel the request.</param>
    /// <returns>Access token string.</returns>
    public Task<string> GetAccessTokenAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("Acquiring Graph access token using {Flow} flow.", _options.AuthFlow);

        return _options.AuthFlow switch
        {
            GraphAuthFlow.Application => AcquireApplicationTokenAsync(cancellationToken),
            GraphAuthFlow.Delegated => AcquireDelegatedTokenAsync(cancellationToken),
            _ => throw new InvalidOperationException($"Unsupported auth flow '{_options.AuthFlow}'.")
        };
    }

    /// <summary>
    /// Acquires an application token using the client credentials flow.
    /// </summary>
    /// <param name="cancellationToken">Token used to cancel the request.</param>
    /// <returns>Access token string.</returns>
    private async Task<string> AcquireApplicationTokenAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_options.Application.TenantId))
        {
            throw new InvalidOperationException("Application tenant id is required.");
        }

        if (string.IsNullOrWhiteSpace(_options.Application.ClientId))
        {
            throw new InvalidOperationException("Application client id is required.");
        }

        if (string.IsNullOrWhiteSpace(_options.Application.ClientSecret))
        {
            throw new InvalidOperationException("Application client secret is required.");
        }

        var form = new Dictionary<string, string>
        {
            ["client_id"] = _options.Application.ClientId,
            ["client_secret"] = _options.Application.ClientSecret,
            ["scope"] = "https://graph.microsoft.com/.default",
            ["grant_type"] = "client_credentials"
        };

        var tokenEndpoint = BuildTokenEndpoint(_options.Application.TenantId);
        return await SendTokenRequestAsync(tokenEndpoint, form, cancellationToken);
    }

    /// <summary>
    /// Runs the device code flow to obtain a delegated access token.
    /// </summary>
    /// <param name="cancellationToken">Token used to cancel the request.</param>
    /// <returns>Access token string.</returns>
    private async Task<string> AcquireDelegatedTokenAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_options.Delegated.TenantId))
        {
            throw new InvalidOperationException("Delegated tenant id is required.");
        }

        if (string.IsNullOrWhiteSpace(_options.Delegated.ClientId))
        {
            throw new InvalidOperationException("Delegated client id is required.");
        }

        var scopes = _options.Delegated.Scopes is { Length: > 0 }
            ? string.Join(' ', NormalizeScopes(_options.Delegated.Scopes))
            : "https://graph.microsoft.com/Mail.Send offline_access";

        var deviceEndpoint = BuildDeviceCodeEndpoint(_options.Delegated.TenantId);
        var deviceContent = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["client_id"] = _options.Delegated.ClientId,
            ["scope"] = scopes
        });

        using var deviceResponse = await HttpClient.PostAsync(deviceEndpoint, deviceContent, cancellationToken);
        var devicePayload = await ReadJsonAsync(deviceResponse, cancellationToken);
        if (!deviceResponse.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Device code request failed: {devicePayload}");
        }

        var deviceCode = devicePayload.GetProperty("device_code").GetString()!;
        var interval = devicePayload.TryGetProperty("interval", out var intervalElement)
            ? intervalElement.GetInt32()
            : 5;
        var message = devicePayload.TryGetProperty("message", out var messageElement)
            ? messageElement.GetString()
            : "Complete authentication to continue.";

        Console.WriteLine(message);

        var tokenEndpoint = BuildTokenEndpoint(_options.Delegated.TenantId);
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await Task.Delay(TimeSpan.FromSeconds(interval), cancellationToken);

            var tokenContent = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["grant_type"] = "urn:ietf:params:oauth:grant-type:device_code",
                ["client_id"] = _options.Delegated.ClientId,
                ["device_code"] = deviceCode
            });

            using var tokenResponse = await HttpClient.PostAsync(tokenEndpoint, tokenContent, cancellationToken);
            var tokenPayload = await ReadJsonAsync(tokenResponse, cancellationToken);
            if (tokenResponse.IsSuccessStatusCode && tokenPayload.TryGetProperty("access_token", out var tokenElement))
            {
                return tokenElement.GetString()!;
            }

            if (tokenPayload.TryGetProperty("error", out var errorElement))
            {
                var error = errorElement.GetString();
                switch (error)
                {
                    case "authorization_pending":
                        continue;
                    case "slow_down":
                        interval += 5;
                        continue;
                    default:
                        throw new InvalidOperationException($"Delegated token request failed: {error}");
                }
            }

            throw new InvalidOperationException("Delegated token request failed.");
        }
    }

    /// <summary>
    /// Sends a token request to the provided endpoint and extracts the access token from the response.
    /// </summary>
    /// <param name="endpoint">Token endpoint URL.</param>
    /// <param name="formValues">Form values required by the token request.</param>
    /// <param name="cancellationToken">Token used to cancel the request.</param>
    /// <returns>Access token string.</returns>
    private static async Task<string> SendTokenRequestAsync(
        string endpoint,
        IDictionary<string, string> formValues,
        CancellationToken cancellationToken)
    {
        using var content = new FormUrlEncodedContent(formValues);
        using var response = await HttpClient.PostAsync(endpoint, content, cancellationToken);
        var payload = await ReadJsonAsync(response, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Token request failed: {payload}");
        }

        if (!payload.TryGetProperty("access_token", out var tokenElement))
        {
            throw new InvalidOperationException("Token response did not contain an access token.");
        }

        return tokenElement.GetString()!;
    }

    /// <summary>
    /// Normalizes scope strings so they are accepted by Microsoft identity endpoints.
    /// </summary>
    /// <param name="scopes">Original scopes from configuration.</param>
    /// <returns>Enumeration of normalized scopes.</returns>
    private static IEnumerable<string> NormalizeScopes(IEnumerable<string> scopes)
    {
        foreach (var scope in scopes)
        {
            if (string.IsNullOrWhiteSpace(scope))
            {
                continue;
            }

            var trimmed = scope.Trim();
            if (trimmed.Contains("://", StringComparison.OrdinalIgnoreCase) || trimmed.StartsWith("api://", StringComparison.OrdinalIgnoreCase))
            {
                yield return trimmed;
            }
            else if (string.Equals(trimmed, "offline_access", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(trimmed, "openid", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(trimmed, "profile", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(trimmed, ".default", StringComparison.OrdinalIgnoreCase))
            {
                yield return trimmed;
            }
            else
            {
                yield return $"https://graph.microsoft.com/{trimmed}";
            }
        }
    }

    /// <summary>
    /// Builds the OAuth token endpoint for the provided tenant.
    /// </summary>
    /// <param name="tenant">Tenant identifier or domain.</param>
    /// <returns>Token endpoint URL.</returns>
    private static string BuildTokenEndpoint(string tenant)
        => $"https://login.microsoftonline.com/{tenant.Trim()}/oauth2/v2.0/token";

    /// <summary>
    /// Builds the OAuth device code endpoint for the provided tenant.
    /// </summary>
    /// <param name="tenant">Tenant identifier or domain.</param>
    /// <returns>Device code endpoint URL.</returns>
    private static string BuildDeviceCodeEndpoint(string tenant)
        => $"https://login.microsoftonline.com/{tenant.Trim()}/oauth2/v2.0/devicecode";

    /// <summary>
    /// Reads the HTTP response content as JSON and returns its root element.
    /// </summary>
    /// <param name="response">Response received from the identity endpoint.</param>
    /// <param name="cancellationToken">Token used to cancel the operation.</param>
    /// <returns>Parsed <see cref="JsonElement"/>.</returns>
    private static async Task<JsonElement> ReadJsonAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var document = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        return document.RootElement.Clone();
    }
}
