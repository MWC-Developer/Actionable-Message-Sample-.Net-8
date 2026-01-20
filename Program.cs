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

using ActionableMessageSample.Options;
using ActionableMessageSample.Services;
using Microsoft.Extensions.Logging;

namespace ActionableMessageSample
{
    /// <summary>
    /// Entry point for the sample ASP.NET Core application.
    /// </summary>
    public class Program
    {
        /// <summary>
        /// Configures and runs the web host.
        /// </summary>
        /// <param name="args">Command-line arguments supplied at startup.</param>
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            builder.Configuration.AddJsonFile("appsettings.Local.json", optional: true, reloadOnChange: true);

            builder.Logging.ClearProviders();
            builder.Logging.AddSimpleConsole(options =>
            {
                options.TimestampFormat = "yyyy-MM-dd HH:mm:ss ";
                options.SingleLine = true;
                options.UseUtcTimestamp = true;
            });

            // Add services to the container.

            builder.Services.AddControllers();
            // Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();
            builder.Services.Configure<ActionableMessageOptions>(builder.Configuration.GetSection("ActionableMessages"));
            builder.Services.AddSingleton<IActionableMessageTokenValidator, ActionableMessageTokenValidator>();
            builder.Services.AddSingleton<IActionableMessageStateStore, InMemoryActionableMessageStateStore>();

            var app = builder.Build();

            // Configure the HTTP request pipeline.
            if (app.Environment.IsDevelopment())
            {
                app.UseSwagger();
                app.UseSwaggerUI();
            }

            app.UseHttpsRedirection();

            app.UseAuthorization();


            app.MapControllers();

            app.Run();
        }
    }
}
