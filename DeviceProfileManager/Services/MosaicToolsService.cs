using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace DeviceProfileManager.Services
{
    public class MosaicToolsService : IDeviceService
    {
        public string DeviceName => "MosaicTools";
        public string DisplayName => "Mosaic Tools Settings";

        private const string ProcessName = "MosaicTools";

        private static readonly string ConfigBasePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MosaicTools");

        private static readonly string SettingsFile = "MosaicToolsSettings.json";

        public static bool IsRunning()
        {
            return Process.GetProcessesByName(ProcessName).Length > 0;
        }

        public static string GetExecutablePath()
        {
            var procs = Process.GetProcessesByName(ProcessName);
            if (procs.Length > 0)
            {
                try
                {
                    return procs[0].MainModule?.FileName;
                }
                catch { }
            }
            return null;
        }

        public static void CloseApplication()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell",
                    Arguments = $"-Command \"Stop-Process -Name '{ProcessName}' -Force -ErrorAction SilentlyContinue\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using (var proc = Process.Start(psi))
                {
                    proc?.WaitForExit(3000);
                }
            }
            catch { }

            Thread.Sleep(300);
        }

        public static void StartApplication(string exePath)
        {
            if (!string.IsNullOrEmpty(exePath) && File.Exists(exePath))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = exePath,
                        UseShellExecute = true
                    });
                }
                catch { }
            }
        }

        // Machine-specific keys to exclude (window positions, paths, etc.)
        // These settings are workstation-specific and should not be synced
        private static readonly HashSet<string> MachineSpecificKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // Window positions (all use snake_case in JSON)
            "window_x",
            "window_y",
            "clinical_history_x",
            "clinical_history_y",
            "impression_x",
            "impression_y",
            "floating_toolbar_x",
            "floating_toolbar_y",
            "indicator_x",
            "indicator_y",
            "report_popup_x",
            "report_popup_y",
            "settings_x",
            "settings_y",
            "pick_list_popup_x",
            "pick_list_popup_y",

            // Window sizes (optional - may want to sync these)
            // "pick_list_editor_width",
            // "pick_list_editor_height",
            // "report_popup_width",

            // Paths - machine-specific
            "rvucounter_path",

            // Legacy keys (in case of older versions)
            "ExePath",
            "Mic Indicator EXE",
            "PosX",
            "PosY",
            "WindowX",
            "WindowY",
            "LastDirectory",
            "MachineName"
        };

        public bool IsInstalled()
        {
            return File.Exists(Path.Combine(ConfigBasePath, SettingsFile));
        }

        public async Task<bool> ExportAsync(string profilePath)
        {
            try
            {
                var settingsPath = Path.Combine(ConfigBasePath, SettingsFile);
                if (!File.Exists(settingsPath))
                    return true; // No config to save

                var exportDir = Path.Combine(profilePath, "mosaictools");
                Directory.CreateDirectory(exportDir);

                // Copy the main settings file, filtering machine-specific keys
                var content = await File.ReadAllTextAsync(settingsPath);
                var filtered = FilterJsonContent(content);
                await File.WriteAllTextAsync(Path.Combine(exportDir, SettingsFile), filtered);

                return true;
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
                var importDir = Path.Combine(profilePath, "mosaictools");
                var srcPath = Path.Combine(importDir, SettingsFile);

                if (!File.Exists(srcPath))
                    return false;

                // Check if running and get path before closing
                bool wasRunning = IsRunning();
                string exePath = wasRunning ? GetExecutablePath() : null;

                if (wasRunning)
                {
                    await Task.Run(() => CloseApplication()).ConfigureAwait(false);
                }

                Directory.CreateDirectory(ConfigBasePath);

                var destPath = Path.Combine(ConfigBasePath, SettingsFile);
                var profileContent = await File.ReadAllTextAsync(srcPath);

                // Merge with existing local settings to preserve machine-specific values
                if (File.Exists(destPath))
                {
                    var localContent = await File.ReadAllTextAsync(destPath);
                    var merged = MergeJsonContent(profileContent, localContent);
                    await File.WriteAllTextAsync(destPath, merged);
                }
                else
                {
                    await File.WriteAllTextAsync(destPath, profileContent);
                }

                // Restart if it was running
                if (wasRunning && !string.IsNullOrEmpty(exePath))
                {
                    StartApplication(exePath);
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
            var importDir = Path.Combine(profilePath, "mosaictools");
            return File.Exists(Path.Combine(importDir, SettingsFile));
        }

        private string FilterJsonContent(string json)
        {
            try
            {
                using (var doc = JsonDocument.Parse(json))
                {
                    var options = new JsonWriterOptions { Indented = true };
                    using (var stream = new MemoryStream())
                    {
                        using (var writer = new Utf8JsonWriter(stream, options))
                        {
                            FilterJsonElement(doc.RootElement, writer);
                        }
                        return Encoding.UTF8.GetString(stream.ToArray());
                    }
                }
            }
            catch
            {
                return json; // Return original if parsing fails
            }
        }

        private void FilterJsonElement(JsonElement element, Utf8JsonWriter writer)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.Object:
                    writer.WriteStartObject();
                    foreach (var property in element.EnumerateObject())
                    {
                        if (!MachineSpecificKeys.Contains(property.Name))
                        {
                            writer.WritePropertyName(property.Name);
                            FilterJsonElement(property.Value, writer);
                        }
                    }
                    writer.WriteEndObject();
                    break;
                case JsonValueKind.Array:
                    writer.WriteStartArray();
                    foreach (var item in element.EnumerateArray())
                    {
                        FilterJsonElement(item, writer);
                    }
                    writer.WriteEndArray();
                    break;
                default:
                    element.WriteTo(writer);
                    break;
            }
        }

        private string MergeJsonContent(string profileJson, string localJson)
        {
            try
            {
                using (var profileDoc = JsonDocument.Parse(profileJson))
                using (var localDoc = JsonDocument.Parse(localJson))
                {
                    var merged = new Dictionary<string, JsonElement>();

                    // Add all profile settings
                    if (profileDoc.RootElement.ValueKind == JsonValueKind.Object)
                    {
                        foreach (var property in profileDoc.RootElement.EnumerateObject())
                        {
                            merged[property.Name] = property.Value;
                        }
                    }

                    // Restore machine-specific settings from local
                    if (localDoc.RootElement.ValueKind == JsonValueKind.Object)
                    {
                        foreach (var property in localDoc.RootElement.EnumerateObject())
                        {
                            if (MachineSpecificKeys.Contains(property.Name))
                            {
                                merged[property.Name] = property.Value;
                            }
                        }
                    }

                    // Serialize merged result
                    var options = new JsonSerializerOptions { WriteIndented = true };
                    return JsonSerializer.Serialize(merged, options);
                }
            }
            catch
            {
                return profileJson; // Return profile content if merging fails
            }
        }

        private static async Task CopyFileAsync(string source, string destination)
        {
            using (var sourceStream = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var destStream = new FileStream(destination, FileMode.Create, FileAccess.Write))
            {
                await sourceStream.CopyToAsync(destStream);
            }
        }
    }
}
