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

namespace ActionableMessageSender.Options;

/// <summary>
/// Determines which Microsoft Graph authentication flow to use.
/// </summary>
public enum GraphAuthFlow
{
    /// <summary>
    /// Use application permissions with client credentials.
    /// </summary>
    Application,

    /// <summary>
    /// Use delegated permissions for a specific user context.
    /// </summary>
    Delegated
}

/// <summary>
/// Strongly typed configuration used to send Graph actionable messages.
/// </summary>
public sealed class GraphOptions
{
    /// <summary>
    /// Gets or sets the authentication flow used to acquire tokens.
    /// </summary>
    public GraphAuthFlow AuthFlow { get; set; } = GraphAuthFlow.Application;

    /// <summary>
    /// Gets or sets app-only authentication settings.
    /// </summary>
    public GraphApplicationOptions Application { get; set; } = new();

    /// <summary>
    /// Gets or sets delegated authentication settings.
    /// </summary>
    public GraphDelegatedOptions Delegated { get; set; } = new();

    /// <summary>
    /// Gets or sets mail-specific values such as recipients and subject.
    /// </summary>
    public GraphMailOptions Mail { get; set; } = new();

    /// <summary>
    /// Gets or sets actionable message specific configuration.
    /// </summary>
    public GraphActionableMessageOptions ActionableMessage { get; set; } = new();

    /// <summary>
    /// Gets or sets the optional file path used to trace outbound Graph requests.
    /// </summary>
    public string? TraceFilePath { get; set; }
}

/// <summary>
/// Holds configuration for app-only Microsoft Graph access.
/// </summary>
public sealed class GraphApplicationOptions
{
    /// <summary>
    /// Gets or sets the Entra tenant identifier.
    /// </summary>
    public string TenantId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the application (client) identifier.
    /// </summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the client secret used for token acquisition.
    /// </summary>
    public string ClientSecret { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the sender user id that owns the mailbox.
    /// </summary>
    public string SenderUserId { get; set; } = string.Empty;
}

/// <summary>
/// Holds configuration for delegated Microsoft Graph access.
/// </summary>
public sealed class GraphDelegatedOptions
{
    /// <summary>
    /// Gets or sets the Entra tenant identifier.
    /// </summary>
    public string TenantId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the application (client) identifier.
    /// </summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the user id (or "me") whose mailbox will be used.
    /// </summary>
    public string SenderUserId { get; set; } = "me";

    /// <summary>
    /// Gets or sets OAuth scopes required for delegated Graph access.
    /// </summary>
    public string[] Scopes { get; set; } = Array.Empty<string>();
}

/// <summary>
/// Holds configuration related to the Graph mail message.
/// </summary>
public sealed class GraphMailOptions
{
    /// <summary>
    /// Gets or sets the subject line for the actionable email.
    /// </summary>
    public string Subject { get; set; } = "Actionable message";

    /// <summary>
    /// Gets or sets the list of recipient SMTP addresses.
    /// </summary>
    public string[] RecipientAddresses { get; set; } = Array.Empty<string>();
}

/// <summary>
/// Holds values required to build actionable message callback URLs.
/// </summary>
public sealed class GraphActionableMessageOptions
{
    /// <summary>
    /// Gets or sets the public base URL that hosts the callback API.
    /// </summary>
    public string ServerBaseUrl { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the registered originator id needed for actionable cards.
    /// </summary>
    public string OriginatorId { get; set; } = string.Empty;
}
