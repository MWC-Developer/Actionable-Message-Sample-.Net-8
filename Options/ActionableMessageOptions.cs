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

namespace ActionableMessageSample.Options;

/// <summary>
/// Configuration required to validate actionable message callbacks from Outlook.
/// </summary>
public sealed class ActionableMessageOptions
{
    /// <summary>
    /// Originator identifier that must match the actionable message registration.
    /// </summary>
    public string OriginatorId { get; set; } = string.Empty;

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
}
