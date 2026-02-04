using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;

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

        // Database files to copy (main db files only, WAL will be checkpointed first)
        private static readonly string[] DatabaseFiles = { "settings.db", "privacy_settings.db" };

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

            // Wait for processes to terminate (up to 5 seconds)
            for (int i = 0; i < 10; i++)
            {
                if (!IsRunning())
                    break;
                Thread.Sleep(500);
            }

            // Extra wait for file handles to be released
            Thread.Sleep(1000);
        }

        /// <summary>
        /// Checkpoints the SQLite WAL (Write-Ahead Log) into the main database file.
        /// This ensures all pending changes are written to the .db file before copying.
        /// </summary>
        private static void CheckpointDatabase(string dbPath)
        {
            if (!File.Exists(dbPath))
                return;

            try
            {
                var connectionString = new SqliteConnectionStringBuilder
                {
                    DataSource = dbPath,
                    Mode = SqliteOpenMode.ReadWrite
                }.ToString();

                using (var connection = new SqliteConnection(connectionString))
                {
                    connection.Open();
                    using (var cmd = connection.CreateCommand())
                    {
                        // TRUNCATE mode: checkpoint and then truncate the WAL file
                        cmd.CommandText = "PRAGMA wal_checkpoint(TRUNCATE);";
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch
            {
                // If checkpoint fails, we'll still copy the files as-is
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

                // Checkpoint WAL files to ensure all changes are in the main .db files
                foreach (var dbFile in DatabaseFiles)
                {
                    var dbPath = Path.Combine(LghubPath, dbFile);
                    CheckpointDatabase(dbPath);
                }

                var exportDir = Path.Combine(profilePath, "logitech");

                // Clear existing export directory
                if (Directory.Exists(exportDir))
                {
                    Directory.Delete(exportDir, true);
                }
                Directory.CreateDirectory(exportDir);

                bool copiedSomething = false;

                // Copy database files from LGHUB (only main .db files after checkpoint)
                var lghubExportDir = Path.Combine(exportDir, "LGHUB");
                Directory.CreateDirectory(lghubExportDir);

                foreach (var dbFile in DatabaseFiles)
                {
                    var sourcePath = Path.Combine(LghubPath, dbFile);
                    if (File.Exists(sourcePath))
                    {
                        try
                        {
                            var destFile = Path.Combine(lghubExportDir, dbFile);
                            File.Copy(sourcePath, destFile, true);
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

                // Ensure LGHUB directory exists
                if (!Directory.Exists(LghubPath))
                {
                    Directory.CreateDirectory(LghubPath);
                }

                // Restore database files to LGHUB
                var lghubImportDir = Path.Combine(importDir, "LGHUB");
                if (Directory.Exists(lghubImportDir))
                {
                    // Delete existing db, shm, and wal files to ensure clean import
                    foreach (var dbFile in DatabaseFiles)
                    {
                        var basePath = Path.Combine(LghubPath, dbFile);
                        try { if (File.Exists(basePath)) File.Delete(basePath); } catch { }
                        try { if (File.Exists(basePath + "-shm")) File.Delete(basePath + "-shm"); } catch { }
                        try { if (File.Exists(basePath + "-wal")) File.Delete(basePath + "-wal"); } catch { }
                    }

                    // Copy the checkpointed database files
                    foreach (var dbFile in DatabaseFiles)
                    {
                        var sourcePath = Path.Combine(lghubImportDir, dbFile);
                        if (File.Exists(sourcePath))
                        {
                            try
                            {
                                var destFile = Path.Combine(LghubPath, dbFile);
                                File.Copy(sourcePath, destFile, true);
                            }
                            catch { }
                        }
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
