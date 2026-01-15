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

using System.Text.Json.Serialization;

namespace ActionableMessageSample.Models;

/// <summary>
/// Represents the payload sent by Outlook when an actionable card callback fires.
/// </summary>
public sealed class ActionableMessageRequest
{
    /// <summary>
    /// Gets the identifier of the action invoked by the user.
    /// </summary>
    [JsonPropertyName("actionId")]
    public string? ActionId { get; init; }

    /// <summary>
    /// Gets the identifier of the actionable card instance.
    /// </summary>
    [JsonPropertyName("cardId")]
    public string? CardId { get; init; }

    /// <summary>
    /// Gets the Exchange message identifier tied to the card.
    /// </summary>
    [JsonPropertyName("messageId")]
    public string? MessageId { get; init; }

    /// <summary>
    /// Gets the user identifier performing the action.
    /// </summary>
    [JsonPropertyName("userId")]
    public string? UserId { get; init; }

    /// <summary>
    /// Gets the custom data payload submitted with the action.
    /// </summary>
    [JsonPropertyName("data")]
    public Dictionary<string, string>? Data { get; init; }
}
