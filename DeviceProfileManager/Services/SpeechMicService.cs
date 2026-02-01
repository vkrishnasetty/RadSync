using System;
using System.IO;
using System.Threading.Tasks;

namespace DeviceProfileManager.Services
{
    public class SpeechMicService : IDeviceService
    {
        public string DeviceName => "SpeechMic";
        public string DisplayName => "Philips SpeechMic";

        private static readonly string ConfigPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Philips Device Control Center");

        public static readonly string ExecutablePath = @"C:\Program Files (x86)\Philips Speech\Device Control Center\PDCC.exe";
        public static readonly string DownloadPageUrl = "https://www.dictation.philips.com/us/products/desktop-dictation/speechmike-premium-dictation-microphone-lfh3500/#productsupport";

        public bool IsInstalled()
        {
            return File.Exists(ExecutablePath);
        }

        public async Task<bool> ExportAsync(string profilePath)
        {
            try
            {
                var exportDir = Path.Combine(profilePath, "speechmic");
                Directory.CreateDirectory(exportDir);

                if (!Directory.Exists(ConfigPath))
                    return true; // No config to export

                // Copy AppControlConfig.*.xml files
                foreach (var file in Directory.GetFiles(ConfigPath, "AppControlConfig.*.xml"))
                {
                    var destFile = Path.Combine(exportDir, Path.GetFileName(file));
                    await CopyFileAsync(file, destFile);
                }

                // Also copy any other XML config files
                foreach (var file in Directory.GetFiles(ConfigPath, "*.xml"))
                {
                    if (!Path.GetFileName(file).StartsWith("AppControlConfig."))
                    {
                        var destFile = Path.Combine(exportDir, Path.GetFileName(file));
                        await CopyFileAsync(file, destFile);
                    }
                }

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
                var importDir = Path.Combine(profilePath, "speechmic");

                if (!Directory.Exists(importDir))
                    return false;

                Directory.CreateDirectory(ConfigPath);

                foreach (var file in Directory.GetFiles(importDir, "*.xml"))
                {
                    var destFile = Path.Combine(ConfigPath, Path.GetFileName(file));
                    await CopyFileAsync(file, destFile);
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
            var importDir = Path.Combine(profilePath, "speechmic");
            return Directory.Exists(importDir) && Directory.GetFiles(importDir, "*.xml").Length > 0;
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
