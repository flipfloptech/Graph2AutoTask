using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Serilog;
using System.Diagnostics;
using System.IO;
using System.CommandLine.DragonFruit;

namespace Graph2AutoTask
{
    public class Program
    {
        private static string appsettings_configuration_path = null;
        private static string appsettings_configuration_name = (System.Diagnostics.Debugger.IsAttached ? "appsettings.Development.json" : "appsettings.json");
        private static string serilog_configuration_path = null;
        private static string serilog_configuration_name = "serilog.json";
        /// <summary>Processes emails from Graph and creates tickets in Autotask.</summary>
        /// <param name="appSettingsConfigFile">The path to the appsettings.config file that is to be used.</param>
        /// <param name="serilogConfigFile">The path to the serilog.config file that is to be used.</param>
        public static void Main(FileInfo appSettingsConfigFile=null, FileInfo serilogConfigFile=null)
        {
            appsettings_configuration_path = $"{System.IO.Path.GetDirectoryName(new System.Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath)}/configs/";
            serilog_configuration_path = $"{System.IO.Path.GetDirectoryName(new System.Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath)}/configs/";
            System.Console.WriteLine("Graph2AutoTask v2.0");

            if (appSettingsConfigFile?.DirectoryName != null && appSettingsConfigFile?.Name != null) {
                appsettings_configuration_path = $"{appSettingsConfigFile?.DirectoryName}/";
                appsettings_configuration_name = appSettingsConfigFile?.Name;
            }
            if (serilogConfigFile?.DirectoryName != null && serilogConfigFile?.Name != null) {
                serilog_configuration_path = $"{serilogConfigFile?.DirectoryName}/";
                serilog_configuration_name = serilogConfigFile?.Name;
            }

            System.Threading.Thread.CurrentThread.Name = "main";
            try
            {
                System.Console.WriteLine("[LOGGING] Initializing");
                IConfigurationRoot _logConfig = new ConfigurationBuilder()
                    .SetBasePath(serilog_configuration_path)
                    .AddJsonFile(serilog_configuration_name)
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
                CreateHostBuilder().Build().Run();
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
        private static IHostBuilder CreateHostBuilder()
        {
            AutomationConfig _configuration = new AutomationConfig();

            IHostBuilder _builder = Host.CreateDefaultBuilder()
                 .ConfigureAppConfiguration(_configBuilder =>
                    {
                        PhysicalFileProvider _provider = new PhysicalFileProvider(appsettings_configuration_path);
                        _configBuilder.AddJsonFile(_provider, appsettings_configuration_name, false, false);
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
