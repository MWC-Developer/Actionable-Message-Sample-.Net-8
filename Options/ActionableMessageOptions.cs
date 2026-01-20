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

namespace ActionableMessageSample.Options;

/// <summary>
/// Configuration required to validate actionable message callbacks from Outlook.
/// </summary>
public sealed class ActionableMessageOptions
{
    /// <summary>
    /// Legacy single originator identifier maintained for backward compatibility.
    /// Use <see cref="OriginatorIds"/> to configure multiple registrations.
    /// </summary>
    public string OriginatorId { get; set; } = string.Empty;

    /// <summary>
    /// Collection of originator identifiers accepted by the server.
    /// </summary>
    public string[] OriginatorIds { get; set; } = Array.Empty<string>();

		/// <summary>
		/// Optional absolute base URL used when building callback targets.
		/// </summary>
		public string? CallbackBaseUrl { get; set; }

    /// <summary>
    /// Entra ID tenant identifier (GUID or domain).
    /// </summary>
    public string EntraTenantId { get; set; } = string.Empty;

    /// <summary>
    /// Entra ID application (client) identifier registered for the actionable message target.
    /// </summary>
    public string EntraClientId { get; set; } = string.Empty;

    /// <summary>
    /// Optional explicit audience the incoming tokens must target. If omitted, api://&lt;clientId&gt; is used.
    /// </summary>
    public string? EntraAudience { get; set; }

    /// <summary>
    /// Optional additional audiences (e.g., HTTPS endpoints) accepted for tokens.
    /// </summary>
    public string[]? AdditionalAudiences { get; set; }

    /// <summary>
    /// Authority host to use when building the metadata endpoint. Defaults to the public cloud authority.
    /// </summary>
    public string EntraAuthorityHost { get; set; } = "https://login.microsoftonline.com";

	/// <summary>
	/// Returns the configured originator identifiers, falling back to the legacy single value when necessary.
	/// </summary>
	public IReadOnlyCollection<string> GetConfiguredOriginatorIds()
	{
		var configured = OriginatorIds?
			.Where(id => !string.IsNullOrWhiteSpace(id))
			.Distinct(StringComparer.OrdinalIgnoreCase)
			.ToArray();

		if (configured is { Length: > 0 })
		{
			return configured;
		}

		if (!string.IsNullOrWhiteSpace(OriginatorId))
		{
			return new[] { OriginatorId };
		}

		return Array.Empty<string>();
	}
}
