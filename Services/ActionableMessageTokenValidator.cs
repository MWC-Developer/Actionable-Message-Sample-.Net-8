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

using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using ActionableMessageSample.Options;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;

namespace ActionableMessageSample.Services;

/// <summary>
/// Validates incoming actionable message bearer tokens from Outlook.
/// </summary>
public sealed class ActionableMessageTokenValidator : IActionableMessageTokenValidator
{
    private readonly ActionableMessageOptions _options;
    private readonly ILogger<ActionableMessageTokenValidator> _logger;
    private readonly IConfigurationManager<OpenIdConnectConfiguration> _authorityConfigurationManager;
    private readonly IConfigurationManager<OpenIdConnectConfiguration> _substrateConfigurationManager;
    private readonly JwtSecurityTokenHandler _tokenHandler = new();

    /// <summary>
    /// Initializes a new instance of the <see cref="ActionableMessageTokenValidator"/> class.
    /// </summary>
    /// <param name="options">Actionable message configuration options.</param>
    /// <param name="logger">Application logger used for diagnostics.</param>
    public ActionableMessageTokenValidator(IOptions<ActionableMessageOptions> options, ILogger<ActionableMessageTokenValidator> logger)
    {
        _options = options.Value ?? throw new ArgumentNullException(nameof(options));
        _logger = logger;

        if (string.IsNullOrWhiteSpace(_options.EntraTenantId))
        {
            throw new InvalidOperationException("Actionable message options must include EntraTenantId.");
        }

        var authority = BuildAuthority();
        var documentRetriever = new HttpDocumentRetriever { RequireHttps = true };
        _authorityConfigurationManager = new ConfigurationManager<OpenIdConnectConfiguration>(
            $"{authority}/v2.0/.well-known/openid-configuration",
            new OpenIdConnectConfigurationRetriever(),
            documentRetriever);

        _substrateConfigurationManager = new ConfigurationManager<OpenIdConnectConfiguration>(
            "https://substrate.office.com/sts/common/.well-known/openid-configuration",
            new OpenIdConnectConfigurationRetriever(),
            documentRetriever);
    }

    /// <inheritdoc />
    public async Task<ClaimsPrincipal> ValidateAsync(string bearerToken, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(bearerToken))
        {
            throw new SecurityTokenException("Bearer token missing.");
        }

        var unsignedToken = _tokenHandler.ReadJwtToken(bearerToken);
        var configuration = await GetConfigurationForIssuerAsync(unsignedToken.Issuer, cancellationToken);
        var validationParameters = BuildValidationParameters(configuration);

        try
        {
            var principal = _tokenHandler.ValidateToken(bearerToken, validationParameters, out _);
            return principal;
        }
        catch (SecurityTokenSignatureKeyNotFoundException)
        {
            _logger.LogWarning("Token validation failed due to missing signing key. Refreshing metadata for issuer {Issuer}.", unsignedToken.Issuer);
            RequestRefreshForIssuer(unsignedToken.Issuer);
            throw;
        }
    }

    /// <summary>
    /// Builds validation parameters tailored for actionable message tokens.
    /// </summary>
    /// <param name="configuration">Issuer configuration containing signing keys.</param>
    /// <returns>Populated <see cref="TokenValidationParameters"/>.</returns>
    private TokenValidationParameters BuildValidationParameters(OpenIdConnectConfiguration configuration)
    {
        var audiences = GetAudiences().ToArray();
        if (audiences.Length == 0)
        {
            throw new InvalidOperationException("No actionable message audiences configured.");
        }

        return new TokenValidationParameters
        {
            ValidAudiences = audiences,
            ValidateAudience = true,
            ValidIssuers = GetValidIssuers().ToArray(),
            ValidateIssuer = true,
            IssuerSigningKeys = configuration.SigningKeys,
            RequireExpirationTime = true,
            RequireSignedTokens = true,
            ValidateIssuerSigningKey = true,
            ValidateLifetime = true,
            ClockSkew = TimeSpan.FromMinutes(2)
        };
    }

    /// <summary>
    /// Enumerates all audiences accepted for actionable message tokens.
    /// </summary>
    /// <returns>Collection of audience values.</returns>
    private IEnumerable<string> GetAudiences()
    {
        if (!string.IsNullOrWhiteSpace(_options.EntraAudience))
        {
            yield return _options.EntraAudience!;
        }

        if (!string.IsNullOrWhiteSpace(_options.EntraClientId))
        {
            yield return $"api://{_options.EntraClientId}";
            yield return _options.EntraClientId;
        }

        if (_options.AdditionalAudiences is { Length: > 0 })
        {
            foreach (var audience in _options.AdditionalAudiences)
            {
                if (!string.IsNullOrWhiteSpace(audience))
                {
                    yield return audience;
                }
            }
        }
    }

    /// <summary>
    /// Enumerates the issuer values considered valid for actionable message tokens.
    /// </summary>
    /// <returns>Collection of issuer URIs.</returns>
    private IEnumerable<string> GetValidIssuers()
    {
        var authority = BuildAuthority();
        if (!string.IsNullOrWhiteSpace(authority))
        {
            // Tokens issued from Entra may present the issuer with, without, or with a trailing slash plus /v2.0 suffix.
            var trimmedAuthority = authority.TrimEnd('/');
            yield return trimmedAuthority;
            yield return $"{trimmedAuthority}/";
            yield return $"{trimmedAuthority}/v2.0";
        }

        // Outlook actionable message tokens are issued by the substrate service
        yield return "https://substrate.office.com/sts/";
        //yield return "https://substrate.office.com/sts/v1";
    }

    /// <summary>
    /// Builds the Entra authority URL used when fetching OpenID metadata.
    /// </summary>
    /// <returns>Authority base URL.</returns>
    private string BuildAuthority()
    {
        var host = string.IsNullOrWhiteSpace(_options.EntraAuthorityHost)
            ? "https://login.microsoftonline.com"
            : _options.EntraAuthorityHost.TrimEnd('/');

        var tenantSegment = $"/{_options.EntraTenantId}";

        if (host.EndsWith(tenantSegment, StringComparison.OrdinalIgnoreCase))
        {
            // Host already contains the tenant identifier, so reuse as-is to avoid duplicating it.
            return host;
        }

        return $"{host}{tenantSegment}";
    }

    /// <summary>
    /// Retrieves the appropriate OpenID configuration for the provided issuer.
    /// </summary>
    /// <param name="issuer">Issuer extracted from the JWT.</param>
    /// <param name="cancellationToken">Cancellation token for the metadata request.</param>
    /// <returns>OpenID configuration document.</returns>
    private async Task<OpenIdConnectConfiguration> GetConfigurationForIssuerAsync(string? issuer, CancellationToken cancellationToken)
    {
        if (IsSubstrateIssuer(issuer))
        {
            return await _substrateConfigurationManager.GetConfigurationAsync(cancellationToken);
        }

        return await _authorityConfigurationManager.GetConfigurationAsync(cancellationToken);
    }

    /// <summary>
    /// Requests a metadata refresh on the configuration manager that matches the issuer.
    /// </summary>
    /// <param name="issuer">JWT issuer value.</param>
    private void RequestRefreshForIssuer(string? issuer)
    {
        if (IsSubstrateIssuer(issuer))
        {
            _substrateConfigurationManager.RequestRefresh();
        }
        else
        {
            _authorityConfigurationManager.RequestRefresh();
        }
    }

    /// <summary>
    /// Determines whether the issuer points to the Outlook substrate service.
    /// </summary>
    /// <param name="issuer">JWT issuer value.</param>
    /// <returns><c>true</c> if the issuer matches the substrate host; otherwise <c>false</c>.</returns>
    private static bool IsSubstrateIssuer(string? issuer)
        => !string.IsNullOrWhiteSpace(issuer) && issuer.StartsWith("https://substrate.office.com/sts", StringComparison.OrdinalIgnoreCase);
}
