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
/// Represents the payload returned to Outlook to update the actionable card.
/// </summary>
public sealed class ActionableMessageResponse
{
    /// <summary>
    /// Gets the title displayed in the updated card.
    /// </summary>
    [JsonPropertyName("title")]
    public string Title { get; init; } = string.Empty;

    /// <summary>
    /// Gets the message or status shown in the card body.
    /// </summary>
    [JsonPropertyName("message")]
    public string Message { get; init; } = string.Empty;

    /// <summary>
    /// Gets the data echoed back to the user for confirmation.
    /// </summary>
    [JsonPropertyName("echo")]
    public IReadOnlyDictionary<string, string>? Echo { get; init; }

    /// <summary>
    /// Creates a response instance with optional payload echoing.
    /// </summary>
    /// <param name="title">Title shown in the card.</param>
    /// <param name="message">Body message shown in the card.</param>
    /// <param name="payload">Optional payload to echo back.</param>
    /// <returns>Configured <see cref="ActionableMessageResponse"/>.</returns>
    public static ActionableMessageResponse Create(string title, string message, IDictionary<string, string>? payload)
    {
        IReadOnlyDictionary<string, string>? echo = null;
        if (payload is { Count: > 0 })
        {
            echo = new Dictionary<string, string>(payload, StringComparer.OrdinalIgnoreCase);
        }

        return new ActionableMessageResponse
        {
            Title = title,
            Message = message,
            Echo = echo
        };
    }
}
