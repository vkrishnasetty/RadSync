using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace DeviceProfileManager.Services
{
    public class LogitechService : IDeviceService
    {
        public string DeviceName => "Logitech";
        public string DisplayName => "Logitech G Hub";

        // Primary location for G Hub settings (SQLite databases)
        private static readonly string LghubPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LGHUB");

        // Secondary location for additional G Hub data
        private static readonly string LghubAppDataPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "lghub");

        public static readonly string ExecutablePath = @"C:\Program Files\LGHUB\lghub.exe";
        public static readonly string DownloadUrl = "https://download01.logi.com/web/ftp/pub/techsupport/gaming/lghub_installer.exe";

        // Only copy essential files - the SQLite databases contain all settings
        private static readonly string[] EssentialFilePatterns = { "*.db", "*.db-shm", "*.db-wal" };

        public bool IsInstalled()
        {
            return File.Exists(ExecutablePath) || Directory.Exists(LghubPath);
        }

        public static bool IsRunning()
        {
            return Process.GetProcessesByName("lghub").Length > 0 ||
                   Process.GetProcessesByName("lghub_agent").Length > 0 ||
                   Process.GetProcessesByName("lghub_updater").Length > 0 ||
                   Process.GetProcessesByName("lghub_system_tray").Length > 0;
        }

        public static void CloseApplication()
        {
            // Use PowerShell Stop-Process which works reliably
            var processNames = new[] { "lghub", "lghub_agent", "lghub_updater", "lghub_system_tray" };

            foreach (var name in processNames)
            {
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = "powershell",
                        Arguments = $"-Command \"Stop-Process -Name {name} -Force -ErrorAction SilentlyContinue\"",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };
                    using (var proc = Process.Start(psi))
                    {
                        proc?.WaitForExit(3000);
                    }
                }
                catch { }
            }

            // Wait for processes to terminate (up to 3 seconds)
            for (int i = 0; i < 6; i++)
            {
                if (!IsRunning())
                    break;
                Thread.Sleep(500);
            }
        }

        public static void StartApplication()
        {
            if (File.Exists(ExecutablePath))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = ExecutablePath,
                        UseShellExecute = true
                    });
                }
                catch { }
            }
        }

        public async Task<bool> ExportAsync(string profilePath)
        {
            try
            {
                if (!Directory.Exists(LghubPath))
                    return true;

                bool wasRunning = IsRunning();
                if (wasRunning)
                {
                    await Task.Run(() => CloseApplication()).ConfigureAwait(false);
                }

                var exportDir = Path.Combine(profilePath, "logitech");

                // Clear existing export directory
                if (Directory.Exists(exportDir))
                {
                    Directory.Delete(exportDir, true);
                }
                Directory.CreateDirectory(exportDir);

                bool copiedSomething = false;

                // Only copy essential database files from LGHUB
                var lghubExportDir = Path.Combine(exportDir, "LGHUB");
                Directory.CreateDirectory(lghubExportDir);

                foreach (var pattern in EssentialFilePatterns)
                {
                    foreach (var file in Directory.GetFiles(LghubPath, pattern))
                    {
                        try
                        {
                            var destFile = Path.Combine(lghubExportDir, Path.GetFileName(file));
                            File.Copy(file, destFile, true);
                            copiedSomething = true;
                        }
                        catch { }
                    }
                }

                // Copy AppData Roaming lghub files (small JSON files)
                if (Directory.Exists(LghubAppDataPath))
                {
                    var appDataExportDir = Path.Combine(exportDir, "lghub_appdata");
                    Directory.CreateDirectory(appDataExportDir);

                    foreach (var file in Directory.GetFiles(LghubAppDataPath))
                    {
                        try
                        {
                            var destFile = Path.Combine(appDataExportDir, Path.GetFileName(file));
                            File.Copy(file, destFile, true);
                        }
                        catch { }
                    }
                }

                // Restart G Hub if it was running
                if (wasRunning)
                {
                    StartApplication();
                }

                return copiedSomething;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public async Task<bool> ImportAsync(string profilePath)
        {
            try
            {
                var importDir = Path.Combine(profilePath, "logitech");
                if (!Directory.Exists(importDir))
                    return false;

                bool wasRunning = IsRunning();
                if (wasRunning)
                {
                    await Task.Run(() => CloseApplication()).ConfigureAwait(false);
                }

                // Restore database files to LGHUB
                var lghubImportDir = Path.Combine(importDir, "LGHUB");
                if (Directory.Exists(lghubImportDir))
                {
                    // Only delete and replace the database files, keep other files intact
                    foreach (var pattern in EssentialFilePatterns)
                    {
                        // Delete existing db files
                        foreach (var file in Directory.GetFiles(LghubPath, pattern))
                        {
                            try { File.Delete(file); } catch { }
                        }
                    }

                    // Copy new db files
                    foreach (var file in Directory.GetFiles(lghubImportDir))
                    {
                        try
                        {
                            var destFile = Path.Combine(LghubPath, Path.GetFileName(file));
                            File.Copy(file, destFile, true);
                        }
                        catch { }
                    }
                }

                // Restore AppData Roaming lghub files
                var appDataImportDir = Path.Combine(importDir, "lghub_appdata");
                if (Directory.Exists(appDataImportDir))
                {
                    if (!Directory.Exists(LghubAppDataPath))
                    {
                        Directory.CreateDirectory(LghubAppDataPath);
                    }

                    foreach (var file in Directory.GetFiles(appDataImportDir))
                    {
                        try
                        {
                            var destFile = Path.Combine(LghubAppDataPath, Path.GetFileName(file));
                            File.Copy(file, destFile, true);
                        }
                        catch { }
                    }
                }

                // Restart G Hub
                if (wasRunning || File.Exists(ExecutablePath))
                {
                    StartApplication();
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool HasConfigData(string profilePath)
        {
            try
            {
                var importDir = Path.Combine(profilePath, "logitech");
                if (!Directory.Exists(importDir))
                    return false;

                return Directory.GetFileSystemEntries(importDir).Length > 0;
            }
            catch
            {
                return false;
            }
        }
    }
}
