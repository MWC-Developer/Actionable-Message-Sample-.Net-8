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

using System.Security.Claims;

namespace ActionableMessageSample.Services;

/// <summary>
/// Provides a contract for validating actionable message bearer tokens.
/// </summary>
public interface IActionableMessageTokenValidator
{
    /// <summary>
    /// Validates the supplied token and returns the associated principal when successful.
    /// </summary>
    /// <param name="bearerToken">JWT bearer token from the actionable message callback.</param>
    /// <param name="cancellationToken">Token used to cancel the operation.</param>
    /// <returns>Claims principal extracted from the token.</returns>
    Task<ClaimsPrincipal> ValidateAsync(string bearerToken, CancellationToken cancellationToken);
}
