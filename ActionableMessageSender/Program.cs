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

using ActionableMessageSender.Options;
using ActionableMessageSender.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var builder = Host.CreateApplicationBuilder(new HostApplicationBuilderSettings
{
    Args = args,
    ContentRootPath = AppContext.BaseDirectory
});

builder.Configuration
    .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
    .AddJsonFile($"appsettings.{builder.Environment.EnvironmentName}.json", optional: true, reloadOnChange: false)
    .AddJsonFile("appsettings.Local.json", optional: true, reloadOnChange: false)
    .AddEnvironmentVariables();

builder.Services.Configure<GraphOptions>(builder.Configuration.GetSection("Graph"));
builder.Services.AddSingleton<GraphTokenProvider>();
builder.Services.AddSingleton<ActionableMessageGraphSender>();

using var host = builder.Build();
using var cts = new CancellationTokenSource();
Console.CancelKeyPress += (_, eventArgs) =>
{
    eventArgs.Cancel = true;
    cts.Cancel();
};

var sender = host.Services.GetRequiredService<ActionableMessageGraphSender>();

try
{
    await sender.SendAsync(cts.Token);
    Console.WriteLine("Actionable message submitted via Microsoft Graph.");
    return 0;
}
catch (OperationCanceledException)
{
    Console.WriteLine("Action canceled.");
    return 1;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Failed to send actionable message: {ex.Message}");
    return 1;
}
