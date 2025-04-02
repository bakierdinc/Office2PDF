using System.Diagnostics;
using System.Runtime.InteropServices;
using System.ServiceProcess;

namespace Office2PDF
{
    internal enum SetupResult
    {
        Setup,
        NotSetup
    }

    internal static class WindowsServiceInstaller
    {
        private const string InstallCommand = "--install";
        private const string UninstallCommand = "--uninstall";

        private static string[] _commands;

        private static bool IsSetup()
        {
            return _commands.Contains(InstallCommand) || _commands.Contains(UninstallCommand);
        }

        private static bool IsInstallCommand()
        {
            return _commands.Contains(InstallCommand);
        }

        private static bool IsUninstallCommand()
        {
            return _commands.Contains(UninstallCommand);
        }

        private static bool IsServiceInstalled(string serviceName)
        {
            return ServiceController.GetServices().Any(s => s.ServiceName.Equals(serviceName, StringComparison.OrdinalIgnoreCase));
        }

        private static void InstallService(string serviceName)
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                Console.WriteLine("Invalid OS Platform.");
                return;
            }

            var exePath = Environment.ProcessPath;

            var command = $"create {serviceName} binPath= \"{exePath}\" start= auto";

            var psi = new ProcessStartInfo("sc", command)
            {
                Verb = "runas",
                CreateNoWindow = true,
                UseShellExecute = true
            };

            try
            {
                Process.Start(psi);
                Console.WriteLine("Service installed successfully.");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Installation failed: {e.Message}");
            }
        }

        private static void UninstallService(string serviceName)
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                Console.WriteLine("Invalid OS Platform.");
                return;
            }

            var command = $"delete {serviceName}";

            var psi = new ProcessStartInfo("sc", command)
            {
                Verb = "runas",
                CreateNoWindow = true,
                UseShellExecute = true
            };

            try
            {
                Process.Start(psi);
                Console.WriteLine("Service uninstalled successfully.");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Uninstallation failed: {e.Message}");
            }

        }

        private static void StartService(string serviceName)
        {
            try
            {
                var startCommand = $"start {serviceName}";

                var psi = new ProcessStartInfo("sc", startCommand)
                {
                    Verb = "runas",
                    UseShellExecute = true,
                    CreateNoWindow = true
                };

                Process.Start(psi)?.WaitForExit();
            }
            catch (Exception e)
            {
                Console.WriteLine($"Failed to start service: {e.Message}");
            }
        }

        private static void StopService(string serviceName)
        {
            try
            {
                var stopCommand = $"stop {serviceName}";

                var psi = new ProcessStartInfo("sc", stopCommand)
                {
                    UseShellExecute = true,
                    Verb = "runas",
                    CreateNoWindow = true
                };

                Process.Start(psi)?.WaitForExit();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to stop service: {ex.Message}");
            }
        }

        public static SetupResult SetupIfRequired(string[] commands, string serviceName)
        {
            _commands = commands;

            if (!IsSetup())
            {
                return SetupResult.NotSetup;
            }

            if (IsInstallCommand())
            {
                if (IsServiceInstalled(serviceName))
                {
                    Console.WriteLine($"Service \"{serviceName}\" is already installed.");
                    Console.WriteLine("Please run with '--uninstall' to remove the existing service first.");
                    Environment.Exit(1);
                }

                InstallService(serviceName);
                StartService(serviceName);
            }

            if (IsUninstallCommand())
            {
                if (!IsServiceInstalled(serviceName))
                {
                    Console.WriteLine($"Service \"{serviceName}\" is not installed.");
                    Console.WriteLine("You can run with '--install' to install the service.");
                    Console.Read();
                    Environment.Exit(1);
                }

                StopService(serviceName);
                UninstallService(serviceName);
            }

            return SetupResult.Setup;
        }
    }
}