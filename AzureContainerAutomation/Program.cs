using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace AzureContainerAutomation
{
    public class Program
    {
        public static void Main(string[] args)
        {
            CreateHostBuilder(args).Build().Run();
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
                });
            if (OperatingSystem.IsWindows())
                return _builder.UseWindowsService();
            else if (OperatingSystem.IsLinux())
                return _builder.UseSystemd();
            else
                return _builder;
        }

    }
}
