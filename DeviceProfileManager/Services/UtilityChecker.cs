using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace DeviceProfileManager.Services
{
    public class UtilityChecker
    {
        // Download page URLs - users download and install manually
        private const string LogitechDownloadPage = "https://www.logitechg.com/en-us/software/ghub";
        private const string StreamDeckDownloadPage = "https://www.elgato.com/us/en/s/downloads";
        private const string SpeechMicDownloadPage = "https://www.dictation.philips.com/us/products/desktop-dictation/speechmike-premium-dictation-microphone-lfh3500/#productsupport";

        // GitHub API URLs for Mosaic tools
        private const string MosaicHotkeysGitHubApi = "https://api.github.com/repos/vkrishnasetty/Mosaic-Combined-Tools/releases/latest";
        private const string MosaicToolsGitHubApi = "https://api.github.com/repos/erichter2018/MosaicTools/releases/latest";

        // Installation paths
        public static readonly string MosaicCombinedToolsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MosaicCombinedTools");

        public static readonly string MosaicHotkeysExePath = Path.Combine(MosaicCombinedToolsPath, "Mosaic Combined Hotkeys.exe");
        public static readonly string MosaicToolsExePath = Path.Combine(MosaicCombinedToolsPath, "MosaicTools.exe");

        public UtilityChecker(string basePath)
        {
            // No longer need installers path since we just open browser pages
        }

        public bool IsLogitechInstalled() => File.Exists(LogitechService.ExecutablePath);
        public bool IsStreamDeckInstalled() => File.Exists(StreamDeckService.ExecutablePath);
        public bool IsSpeechMicInstalled() => File.Exists(SpeechMicService.ExecutablePath);
        public bool IsMosaicHotkeysInstalled() => File.Exists(MosaicHotkeysExePath);
        public bool IsMosaicToolsInstalled() => File.Exists(MosaicToolsExePath);

        public (bool success, string message) OpenLogitechDownloadPage()
        {
            try
            {
                OpenUrl(LogitechDownloadPage);
                return (true, "Opened Logitech G Hub download page. Download and follow the installation instructions.");
            }
            catch (Exception ex)
            {
                return (false, $"Failed to open download page: {ex.Message}");
            }
        }

        public (bool success, string message) OpenStreamDeckDownloadPage()
        {
            try
            {
                OpenUrl(StreamDeckDownloadPage);
                return (true, "Opened Stream Deck download page. Download and follow the installation instructions.");
            }
            catch (Exception ex)
            {
                return (false, $"Failed to open download page: {ex.Message}");
            }
        }

        public (bool success, string message) OpenSpeechMicDownloadPage()
        {
            try
            {
                OpenUrl(SpeechMicDownloadPage);
                return (true, "Opened Philips download page. Download and follow the installation instructions.");
            }
            catch (Exception ex)
            {
                return (false, $"Failed to open download page: {ex.Message}");
            }
        }

        public async Task<(bool success, string message)> DownloadAndInstallMosaicHotkeysAsync(IProgress<string> progress = null)
        {
            return await DownloadAndInstallFromGitHubAsync(
                MosaicHotkeysGitHubApi,
                "Mosaic Combined Hotkeys",
                "Mosaic Combined Hotkeys.exe",
                true, // Create desktop shortcut
                progress);
        }

        public async Task<(bool success, string message)> DownloadAndInstallMosaicToolsAsync(IProgress<string> progress = null)
        {
            return await DownloadAndInstallFromGitHubAsync(
                MosaicToolsGitHubApi,
                "MosaicTools",
                "MosaicTools.exe",
                false, // No desktop shortcut
                progress);
        }

        private async Task<(bool success, string message)> DownloadAndInstallFromGitHubAsync(
            string apiUrl, string appName, string exeName, bool createShortcut, IProgress<string> progress = null)
        {
            try
            {
                progress?.Report($"Checking for latest {appName}...");

                // Get latest release info from GitHub
                string downloadUrl = null;
                string version = null;

                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "RadSync");
                    client.Timeout = TimeSpan.FromSeconds(30);

                    var response = await client.GetStringAsync(apiUrl);
                    var json = JsonDocument.Parse(response);

                    version = json.RootElement.GetProperty("tag_name").GetString()?.TrimStart('v', 'V');

                    // Find the .zip asset
                    if (json.RootElement.TryGetProperty("assets", out var assets))
                    {
                        foreach (var asset in assets.EnumerateArray())
                        {
                            var name = asset.GetProperty("name").GetString();
                            if (name != null && name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                            {
                                downloadUrl = asset.GetProperty("browser_download_url").GetString();
                                break;
                            }
                        }
                    }
                }

                if (string.IsNullOrEmpty(downloadUrl))
                {
                    return (false, $"Could not find download for {appName}.");
                }

                progress?.Report($"Downloading {appName} v{version}...");

                // Create temp directory for download
                var tempDir = Path.Combine(Path.GetTempPath(), $"{appName.Replace(" ", "_")}_Install");
                var zipPath = Path.Combine(tempDir, "download.zip");
                var extractPath = Path.Combine(tempDir, "extracted");

                // Clean up previous attempt
                if (Directory.Exists(tempDir))
                {
                    try { Directory.Delete(tempDir, true); } catch { }
                }
                Directory.CreateDirectory(tempDir);

                // Download the zip
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "RadSync");
                    client.Timeout = TimeSpan.FromMinutes(5);

                    using (var response = await client.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead))
                    {
                        response.EnsureSuccessStatusCode();

                        var totalBytes = response.Content.Headers.ContentLength ?? -1;
                        var downloadedBytes = 0L;

                        using (var contentStream = await response.Content.ReadAsStreamAsync())
                        using (var fileStream = new FileStream(zipPath, FileMode.Create, FileAccess.Write))
                        {
                            var buffer = new byte[81920];
                            int bytesRead;
                            while ((bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                            {
                                await fileStream.WriteAsync(buffer, 0, bytesRead);
                                downloadedBytes += bytesRead;

                                if (totalBytes > 0)
                                {
                                    var percent = (int)((downloadedBytes * 100) / totalBytes);
                                    progress?.Report($"Downloading {appName}... {percent}%");
                                }
                            }
                        }
                    }
                }

                progress?.Report($"Extracting {appName}...");

                // Extract zip
                ZipFile.ExtractToDirectory(zipPath, extractPath);

                // Find the exe
                var sourceExe = FindFile(extractPath, exeName);
                if (sourceExe == null)
                {
                    return (false, $"Could not find {exeName} in download package.");
                }

                progress?.Report($"Installing {appName}...");

                // Create installation directory
                Directory.CreateDirectory(MosaicCombinedToolsPath);

                // Copy all files from the extracted folder to installation path
                var sourceDir = Path.GetDirectoryName(sourceExe);
                foreach (var file in Directory.GetFiles(sourceDir))
                {
                    var destFile = Path.Combine(MosaicCombinedToolsPath, Path.GetFileName(file));
                    File.Copy(file, destFile, true);
                }

                // Create desktop shortcut if requested
                if (createShortcut)
                {
                    var destExe = Path.Combine(MosaicCombinedToolsPath, exeName);
                    CreateDesktopShortcut(destExe, appName);
                }

                // Cleanup temp files
                try { Directory.Delete(tempDir, true); } catch { }

                var shortcutMsg = createShortcut ? " Desktop shortcut created." : "";
                return (true, $"{appName} v{version} installed successfully.{shortcutMsg}");
            }
            catch (Exception ex)
            {
                return (false, $"Failed to install {appName}: {ex.Message}");
            }
        }

        private string FindFile(string directory, string fileName)
        {
            // Check root
            var rootPath = Path.Combine(directory, fileName);
            if (File.Exists(rootPath))
                return rootPath;

            // Check subdirectories
            foreach (var subDir in Directory.GetDirectories(directory))
            {
                var result = FindFile(subDir, fileName);
                if (result != null)
                    return result;
            }

            return null;
        }

        private void CreateDesktopShortcut(string targetPath, string shortcutName)
        {
            try
            {
                var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                var shortcutPath = Path.Combine(desktopPath, $"{shortcutName}.lnk");
                var workingDir = Path.GetDirectoryName(targetPath);

                // Use PowerShell to create the shortcut
                var psScript = $@"
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut('{shortcutPath.Replace("'", "''")}')
$Shortcut.TargetPath = '{targetPath.Replace("'", "''")}'
$Shortcut.WorkingDirectory = '{workingDir.Replace("'", "''")}'
$Shortcut.Description = '{shortcutName.Replace("'", "''")}'
$Shortcut.Save()
";
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-ExecutionPolicy Bypass -NoProfile -Command \"{psScript.Replace("\"", "\\\"")}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using (var proc = Process.Start(psi))
                {
                    proc?.WaitForExit(5000);
                }
            }
            catch
            {
                // Shortcut creation is optional, don't fail the install
            }
        }

        private void OpenUrl(string url)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = url,
                UseShellExecute = true
            });
        }
    }
}
