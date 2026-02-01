using System;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace DeviceProfileManager.Services
{
    public class StreamDeckService : IDeviceService
    {
        public string DeviceName => "StreamDeck";
        public string DisplayName => "Elgato Stream Deck";

        private static readonly string AppDataPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Elgato", "StreamDeck");

        // Cache file to store local device info (persists across profile changes)
        private static readonly string DeviceCacheFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Elgato", "StreamDeck", "radsync_device_cache.json");

        public static readonly string ExecutablePath = @"C:\Program Files\Elgato\StreamDeck\StreamDeck.exe";
        public static readonly string DownloadUrl = "https://edge.elgato.com/egc/windows/sd/Stream_Deck_6.7.1.21877.msi";

        // Only copy these essential folders - skip plugins, logs, cache, etc.
        private static readonly string[] EssentialFolders = { "ProfilesV3", "BackupV3" };

        public bool IsInstalled()
        {
            return File.Exists(ExecutablePath) || Directory.Exists(AppDataPath);
        }

        public static bool IsRunning()
        {
            return Process.GetProcessesByName("StreamDeck").Length > 0;
        }

        public static void CloseApplication()
        {
            // Use PowerShell Stop-Process which works reliably
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell",
                    Arguments = "-Command \"Stop-Process -Name StreamDeck -Force -ErrorAction SilentlyContinue\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using (var proc = Process.Start(psi))
                {
                    proc?.WaitForExit(5000);
                }
            }
            catch { }

            // Wait for process to terminate (up to 3 seconds)
            for (int i = 0; i < 6; i++)
            {
                if (Process.GetProcessesByName("StreamDeck").Length == 0)
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
                if (!Directory.Exists(AppDataPath))
                    return true;

                bool wasRunning = IsRunning();
                if (wasRunning)
                {
                    await Task.Run(() => CloseApplication()).ConfigureAwait(false);
                }

                // Cache local device info before saving (for future imports)
                CacheLocalDeviceInfo();

                var exportDir = Path.Combine(profilePath, "streamdeck");

                if (Directory.Exists(exportDir))
                {
                    Directory.Delete(exportDir, true);
                }
                Directory.CreateDirectory(exportDir);

                bool copiedSomething = false;

                // Only copy essential folders (ProfilesV3 contains button configs)
                foreach (var folderName in EssentialFolders)
                {
                    var sourceFolder = Path.Combine(AppDataPath, folderName);
                    if (Directory.Exists(sourceFolder))
                    {
                        var destFolder = Path.Combine(exportDir, folderName);
                        await Task.Run(() => CopyDirectory(sourceFolder, destFolder)).ConfigureAwait(false);
                        copiedSomething = true;
                    }
                }

                // Restart Stream Deck if it was running
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
                var importDir = Path.Combine(profilePath, "streamdeck");

                if (!Directory.Exists(importDir))
                    return false;

                bool wasRunning = IsRunning();
                if (wasRunning)
                {
                    await Task.Run(() => CloseApplication()).ConfigureAwait(false);
                }

                // Get local device info from cache file (saved during previous save/export)
                // This persists even after we delete profiles
                var localDeviceInfo = GetCachedDeviceInfo() ?? GetDeviceInfoFromProfiles();

                // Cache it now if we found it from profiles (before deleting them)
                if (localDeviceInfo != null)
                {
                    SaveDeviceCache(localDeviceInfo);
                }

                // Only replace the essential folders, keep plugins/resources intact
                foreach (var folderName in EssentialFolders)
                {
                    var sourceFolder = Path.Combine(importDir, folderName);
                    var destFolder = Path.Combine(AppDataPath, folderName);

                    if (Directory.Exists(sourceFolder))
                    {
                        // Remove existing folder
                        if (Directory.Exists(destFolder))
                        {
                            try { Directory.Delete(destFolder, true); } catch { }
                        }

                        // Copy new folder
                        await Task.Run(() => CopyDirectory(sourceFolder, destFolder)).ConfigureAwait(false);
                    }
                }

                // Fix device bindings in profile manifests to use local device
                if (localDeviceInfo != null)
                {
                    FixDeviceBindings(localDeviceInfo);
                }

                // Restart Stream Deck
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
                var importDir = Path.Combine(profilePath, "streamdeck");
                if (!Directory.Exists(importDir))
                    return false;

                return Directory.GetFileSystemEntries(importDir).Length > 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Cache local device info to a file that persists across profile changes
        /// </summary>
        private static void CacheLocalDeviceInfo()
        {
            var deviceInfo = GetDeviceInfoFromProfiles();
            if (deviceInfo != null)
            {
                SaveDeviceCache(deviceInfo);
            }
        }

        private static void SaveDeviceCache(DeviceInfo deviceInfo)
        {
            try
            {
                var json = JsonSerializer.Serialize(deviceInfo);
                File.WriteAllText(DeviceCacheFile, json);
            }
            catch { }
        }

        private static DeviceInfo GetCachedDeviceInfo()
        {
            try
            {
                if (File.Exists(DeviceCacheFile))
                {
                    var json = File.ReadAllText(DeviceCacheFile);
                    return JsonSerializer.Deserialize<DeviceInfo>(json);
                }
            }
            catch { }
            return null;
        }

        private static DeviceInfo GetDeviceInfoFromProfiles()
        {
            try
            {
                var profilesV3 = Path.Combine(AppDataPath, "ProfilesV3");
                if (!Directory.Exists(profilesV3))
                    return null;

                foreach (var profileDir in Directory.GetDirectories(profilesV3, "*.sdProfile"))
                {
                    var manifestPath = Path.Combine(profileDir, "manifest.json");
                    if (File.Exists(manifestPath))
                    {
                        var json = File.ReadAllText(manifestPath);
                        var doc = JsonDocument.Parse(json);

                        if (doc.RootElement.TryGetProperty("Device", out var device))
                        {
                            var model = device.TryGetProperty("Model", out var m) ? m.GetString() : null;
                            var uuid = device.TryGetProperty("UUID", out var u) ? u.GetString() : null;

                            if (!string.IsNullOrEmpty(model) && !string.IsNullOrEmpty(uuid))
                            {
                                return new DeviceInfo { Model = model, UUID = uuid };
                            }
                        }
                    }
                }
            }
            catch { }

            return null;
        }

        private static void FixDeviceBindings(DeviceInfo deviceInfo)
        {
            try
            {
                var profilesV3 = Path.Combine(AppDataPath, "ProfilesV3");
                if (!Directory.Exists(profilesV3))
                    return;

                foreach (var profileDir in Directory.GetDirectories(profilesV3, "*.sdProfile"))
                {
                    var manifestPath = Path.Combine(profileDir, "manifest.json");
                    if (File.Exists(manifestPath))
                    {
                        try
                        {
                            var json = File.ReadAllText(manifestPath);
                            var doc = JsonDocument.Parse(json);

                            using (var ms = new MemoryStream())
                            {
                                using (var writer = new Utf8JsonWriter(ms))
                                {
                                    writer.WriteStartObject();

                                    foreach (var prop in doc.RootElement.EnumerateObject())
                                    {
                                        if (prop.Name == "Device")
                                        {
                                            writer.WritePropertyName("Device");
                                            writer.WriteStartObject();
                                            writer.WriteString("Model", deviceInfo.Model);
                                            writer.WriteString("UUID", deviceInfo.UUID);
                                            writer.WriteEndObject();
                                        }
                                        else
                                        {
                                            prop.WriteTo(writer);
                                        }
                                    }

                                    writer.WriteEndObject();
                                }

                                var updatedJson = System.Text.Encoding.UTF8.GetString(ms.ToArray());
                                File.WriteAllText(manifestPath, updatedJson);
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }

        private class DeviceInfo
        {
            public string Model { get; set; }
            public string UUID { get; set; }
        }

        private static void CopyDirectory(string sourceDir, string destDir)
        {
            Directory.CreateDirectory(destDir);

            foreach (var file in Directory.GetFiles(sourceDir))
            {
                try
                {
                    var destFile = Path.Combine(destDir, Path.GetFileName(file));
                    File.Copy(file, destFile, true);
                }
                catch { }
            }

            foreach (var dir in Directory.GetDirectories(sourceDir))
            {
                var destSubDir = Path.Combine(destDir, Path.GetFileName(dir));
                CopyDirectory(dir, destSubDir);
            }
        }
    }
}
