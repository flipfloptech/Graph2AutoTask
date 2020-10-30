using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Serilog;
using System.Diagnostics;

namespace Graph2AutoTask
{
    public class Program
    {
        private static string _configuration_path = null;
        public static void Main(string[] args)
        {
            _configuration_path = $"{System.IO.Path.GetDirectoryName(new System.Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath)}/configs/";
            System.Console.WriteLine("Graph2AutoTask v2.0");
            System.Threading.Thread.CurrentThread.Name = "main";
            try
            {
                System.Console.WriteLine("[LOGGING] Initializing");
                IConfigurationRoot _logConfig = new ConfigurationBuilder()
                    .SetBasePath(_configuration_path)
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
                System.Console.WriteLine($"Current Directory: {System.IO.Directory.GetCurrentDirectory()}");
                System.Console.WriteLine($"Exception: {_ex.Message}");
                return;
            }
            try { 
                Log.Information("System starting");
                CreateHostBuilder(args).Build().Run();
                Log.Information("System shutting down");
            }
            catch(System.Exception _ex)
            {
                Log.Error(_ex, $"System exception during startup or shutdown. Provide details to developer.");
                Log.Error($"Exception: {_ex.Message}");
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
                 .ConfigureAppConfiguration(_configBuilder =>
                    {
                        PhysicalFileProvider _provider = new PhysicalFileProvider(_configuration_path);
                        if (!System.Diagnostics.Debugger.IsAttached)
                            _configBuilder.AddJsonFile(_provider, "appsettings.json", false, false);
                        else
                            _configBuilder.AddJsonFile(_provider, "appsettings.Development.json", false, false);
                    })
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
