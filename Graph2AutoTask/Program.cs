using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Serilog;

namespace Graph2AutoTask
{
    public class Program
    {
        public static void Main(string[] args)
        {
            System.Console.WriteLine("Graph2AutoTask v2.0");
            System.Threading.Thread.CurrentThread.Name = "main";
            try
            {
                System.Console.WriteLine("[LOGGING] Initializing");
                IConfigurationRoot _logConfig = new ConfigurationBuilder()
                    .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                    .AddJsonFile("serilog.json")
                    .Build();
                Log.Logger = new LoggerConfiguration()
                    .ReadFrom.Configuration(_logConfig)
                    .CreateLogger();
                System.Console.WriteLine("[LOGGING] Initialized");
            }
            catch(System.Exception _ex)
            {
                System.Console.WriteLine("[LOGGING] Initialization Failed! Contact developer.");
                return;
            }
            try { 
                Log.Information("System starting");
                CreateHostBuilder(args).Build().Run();
                Log.Information("System shutting down");
            }
            catch(System.Exception _ex)
            {
                Log.Error($"System exception during startup or shutdown. Provide details to developer.",_ex);
            }
            finally
            {
                Log.Information("System shutdown complete");
                Log.CloseAndFlush();
            }
        }
        public static IHostBuilder CreateHostBuilder(string[] args)
        {
            AutomationConfig _configuration = new AutomationConfig();

            IHostBuilder _builder = Host.CreateDefaultBuilder(args)
                .ConfigureServices((hostContext, services) =>
                {
                    hostContext.Configuration.Bind(_configuration);

                    foreach (MailboxConfig _mailbox in _configuration.MailBoxes)
                    {
                        if (_mailbox.Processing.Enabled)
                        {
                            services.AddSingleton<IHostedService>(sp => new MailMonitorWorker(sp.GetService<ILogger<MailMonitorWorker>>(), _mailbox));
                        }
                    }
                }).UseSerilog();
            IHostBuilder _result = null;
            if (OperatingSystem.IsWindows())
            {
                Log.Information("Detected starting on Windows");
                _result = _builder.UseWindowsService();
            }
            else if (OperatingSystem.IsLinux())
            {
                Log.Information("Detected starting on Linux");
                _result = _builder.UseSystemd();
            }
            else
            { 
                Log.Information("Unable to start as service. Running in console session.");
                _result = _builder;
            }
            Log.Information("System started");
            return _result;
        }

    }
}
