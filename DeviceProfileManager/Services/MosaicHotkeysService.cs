using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DeviceProfileManager.Services
{
    public class MosaicHotkeysService : IDeviceService
    {
        public string DeviceName => "MosaicHotkeys";
        public string DisplayName => "Mosaic Combined Hotkeys";

        private const string ProcessName = "Mosaic Combined Hotkeys";

        private static readonly string ConfigPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MosaicCombinedTools", "HotkeyConfig.ini");

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

        // Machine-specific keys to exclude when saving
        // These settings are workstation-specific and should not be synced
        private static readonly HashSet<string> MachineSpecificKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // Paths - machine-specific executable locations
            "Mic Indicator EXE",
            "ExePath",

            // Window positions - screen layout specific
            "PosX",
            "PosY"

            // Note: User Name is intentionally synced as it's the same person across workstations
            // Note: AutoStart is intentionally synced as user preference
        };

        public bool IsInstalled()
        {
            return File.Exists(ConfigPath);
        }

        public async Task<bool> ExportAsync(string profilePath)
        {
            try
            {
                if (!File.Exists(ConfigPath))
                    return true; // No config to save

                var exportDir = Path.Combine(profilePath, "mosaichotkeys");
                Directory.CreateDirectory(exportDir);

                // Read and filter the INI file
                var filteredContent = await FilterMachineSpecificSettingsAsync(ConfigPath);
                var destPath = Path.Combine(exportDir, "HotkeyConfig.ini");

                await File.WriteAllTextAsync(destPath, filteredContent, Encoding.Unicode);
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
                var importDir = Path.Combine(profilePath, "mosaichotkeys");
                var srcPath = Path.Combine(importDir, "HotkeyConfig.ini");

                if (!File.Exists(srcPath))
                    return false;

                // Check if running and get path before closing
                bool wasRunning = IsRunning();
                string exePath = wasRunning ? GetExecutablePath() : null;

                if (wasRunning)
                {
                    await Task.Run(() => CloseApplication()).ConfigureAwait(false);
                }

                // Read the saved profile config
                var profileSettings = await ParseIniFileAsync(srcPath);

                // Read current local config to preserve machine-specific settings
                var localSettings = new Dictionary<string, Dictionary<string, string>>();
                if (File.Exists(ConfigPath))
                {
                    localSettings = await ParseIniFileAsync(ConfigPath);
                }

                // Merge: use profile settings but preserve local machine-specific values
                var mergedSettings = MergeSettings(profileSettings, localSettings);

                // Write merged config
                Directory.CreateDirectory(Path.GetDirectoryName(ConfigPath));
                var content = BuildIniContent(mergedSettings);
                await File.WriteAllTextAsync(ConfigPath, content, Encoding.Unicode);

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
            var importDir = Path.Combine(profilePath, "mosaichotkeys");
            return File.Exists(Path.Combine(importDir, "HotkeyConfig.ini"));
        }

        private async Task<string> FilterMachineSpecificSettingsAsync(string filePath)
        {
            var settings = await ParseIniFileAsync(filePath);

            // Remove machine-specific keys
            foreach (var section in settings.Values)
            {
                var keysToRemove = new List<string>();
                foreach (var key in section.Keys)
                {
                    if (MachineSpecificKeys.Contains(key))
                    {
                        keysToRemove.Add(key);
                    }
                }
                foreach (var key in keysToRemove)
                {
                    section.Remove(key);
                }
            }

            return BuildIniContent(settings);
        }

        private async Task<Dictionary<string, Dictionary<string, string>>> ParseIniFileAsync(string filePath)
        {
            var result = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
            var currentSection = "";

            var content = await File.ReadAllTextAsync(filePath);
            // Handle potential BOM and different encodings
            content = content.TrimStart('\uFEFF', '\uFFFE');

            using (var reader = new StringReader(content))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    line = line.Trim();

                    // Skip empty lines and comments
                    if (string.IsNullOrEmpty(line) || line.StartsWith(";") || line.StartsWith("#"))
                        continue;

                    // Section header
                    if (line.StartsWith("[") && line.EndsWith("]"))
                    {
                        currentSection = line.Substring(1, line.Length - 2).Trim();
                        if (!result.ContainsKey(currentSection))
                        {
                            result[currentSection] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        }
                    }
                    // Key=Value pair
                    else if (line.Contains("="))
                    {
                        var idx = line.IndexOf('=');
                        var key = line.Substring(0, idx).Trim();
                        var value = line.Substring(idx + 1).Trim();

                        if (!string.IsNullOrEmpty(currentSection) && !result.ContainsKey(currentSection))
                        {
                            result[currentSection] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        }

                        if (!string.IsNullOrEmpty(currentSection))
                        {
                            result[currentSection][key] = value;
                        }
                    }
                }
            }

            return result;
        }

        private Dictionary<string, Dictionary<string, string>> MergeSettings(
            Dictionary<string, Dictionary<string, string>> profileSettings,
            Dictionary<string, Dictionary<string, string>> localSettings)
        {
            var merged = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

            // Start with profile settings
            foreach (var section in profileSettings)
            {
                merged[section.Key] = new Dictionary<string, string>(section.Value, StringComparer.OrdinalIgnoreCase);
            }

            // Restore machine-specific keys from local settings
            foreach (var section in localSettings)
            {
                if (!merged.ContainsKey(section.Key))
                {
                    merged[section.Key] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                }

                foreach (var kvp in section.Value)
                {
                    if (MachineSpecificKeys.Contains(kvp.Key))
                    {
                        merged[section.Key][kvp.Key] = kvp.Value;
                    }
                }
            }

            return merged;
        }

        private string BuildIniContent(Dictionary<string, Dictionary<string, string>> settings)
        {
            var sb = new StringBuilder();

            foreach (var section in settings)
            {
                sb.AppendLine($"[{section.Key}]");
                foreach (var kvp in section.Value)
                {
                    sb.AppendLine($"{kvp.Key}={kvp.Value}");
                }
                sb.AppendLine();
            }

            return sb.ToString();
        }
    }
}
