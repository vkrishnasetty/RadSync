using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;

namespace DeviceProfileManager.Services
{
    public class UpdateService
    {
        private const string GitHubApiUrl = "https://api.github.com/repos/vkrishnasetty/RadSync/releases/latest";
        private const string ReleasesUrl = "https://github.com/vkrishnasetty/RadSync/releases/latest";

        public static string CurrentVersion
        {
            get
            {
                var version = Assembly.GetExecutingAssembly().GetName().Version;
                return version != null ? $"{version.Major}.{version.Minor}.{version.Build}" : "1.0.0";
            }
        }

        public static async Task<(bool hasUpdate, string latestVersion, string downloadUrl)> CheckForUpdatesAsync()
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "RadSync");
                    client.Timeout = TimeSpan.FromSeconds(10);

                    var response = await client.GetStringAsync(GitHubApiUrl);
                    var json = JsonDocument.Parse(response);

                    var tagName = json.RootElement.GetProperty("tag_name").GetString();
                    var latestVersion = tagName?.TrimStart('v', 'V') ?? "0.0.0";

                    // Find the .zip asset download URL
                    string zipDownloadUrl = null;
                    if (json.RootElement.TryGetProperty("assets", out var assets))
                    {
                        foreach (var asset in assets.EnumerateArray())
                        {
                            var name = asset.GetProperty("name").GetString();
                            if (name != null && name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                            {
                                zipDownloadUrl = asset.GetProperty("browser_download_url").GetString();
                                break;
                            }
                        }
                    }

                    var hasUpdate = CompareVersions(latestVersion, CurrentVersion) > 0;

                    return (hasUpdate, latestVersion, zipDownloadUrl ?? ReleasesUrl);
                }
            }
            catch
            {
                return (false, null, null);
            }
        }

        public static void OpenReleasesPage()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = ReleasesUrl,
                    UseShellExecute = true
                });
            }
            catch { }
        }

        public static async Task<(bool success, string message)> DownloadAndInstallUpdateAsync(string downloadUrl, IProgress<string> statusProgress = null)
        {
            try
            {
                // If no direct download URL, open releases page
                if (string.IsNullOrEmpty(downloadUrl) || !downloadUrl.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                {
                    OpenReleasesPage();
                    return (true, "Opened releases page. Please download manually.");
                }

                statusProgress?.Report("Downloading update...");

                var tempDir = Path.Combine(Path.GetTempPath(), "RadSync_Update");
                var zipPath = Path.Combine(tempDir, "RadSync_Update.zip");
                var extractPath = Path.Combine(tempDir, "extracted");

                // Clean up any previous update attempt
                if (Directory.Exists(tempDir))
                {
                    Directory.Delete(tempDir, true);
                }
                Directory.CreateDirectory(tempDir);

                // Download the zip file
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "RadSync");
                    client.Timeout = TimeSpan.FromMinutes(5);

                    using (var response = await client.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead))
                    {
                        response.EnsureSuccessStatusCode();

                        using (var contentStream = await response.Content.ReadAsStreamAsync())
                        using (var fileStream = new FileStream(zipPath, FileMode.Create, FileAccess.Write))
                        {
                            await contentStream.CopyToAsync(fileStream);
                        }
                    }
                }

                statusProgress?.Report("Extracting update...");

                // Extract the zip
                ZipFile.ExtractToDirectory(zipPath, extractPath);

                // Find the new exe - could be in root or a subfolder
                var newExe = FindExecutable(extractPath, "RadSync.exe") ??
                             FindExecutable(extractPath, "DeviceProfileManager.exe");

                if (newExe == null)
                {
                    return (false, "Could not find executable in update package.");
                }

                statusProgress?.Report("Installing update...");

                // Get current exe path
                var currentExe = Process.GetCurrentProcess().MainModule?.FileName;
                if (string.IsNullOrEmpty(currentExe))
                {
                    return (false, "Could not determine current executable path.");
                }

                var currentDir = Path.GetDirectoryName(currentExe);
                var backupExe = currentExe + ".old";

                // Create a batch script to:
                // 1. Wait for current process to exit
                // 2. Replace the exe
                // 3. Start the new exe
                // 4. Clean up
                var batchPath = Path.Combine(tempDir, "update.bat");
                var batchContent = $@"@echo off
timeout /t 2 /nobreak >nul
del ""{backupExe}"" 2>nul
move ""{currentExe}"" ""{backupExe}""
xcopy ""{Path.GetDirectoryName(newExe)}\*.*"" ""{currentDir}\"" /E /Y /Q
start """" ""{currentExe}""
timeout /t 3 /nobreak >nul
rd /s /q ""{tempDir}""
del ""%~f0""
";
                await File.WriteAllTextAsync(batchPath, batchContent);

                // Start the batch script hidden
                var psi = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c \"{batchPath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                Process.Start(psi);

                return (true, "Update downloaded. Restarting application...");
            }
            catch (Exception ex)
            {
                // Fall back to opening releases page
                OpenReleasesPage();
                return (false, $"Auto-update failed: {ex.Message}. Opened releases page for manual download.");
            }
        }

        private static string FindExecutable(string directory, string fileName)
        {
            // Check root
            var rootPath = Path.Combine(directory, fileName);
            if (File.Exists(rootPath))
                return rootPath;

            // Check subdirectories
            foreach (var subDir in Directory.GetDirectories(directory))
            {
                var result = FindExecutable(subDir, fileName);
                if (result != null)
                    return result;
            }

            return null;
        }

        private static int CompareVersions(string v1, string v2)
        {
            try
            {
                var parts1 = v1.Split('.');
                var parts2 = v2.Split('.');

                for (int i = 0; i < Math.Max(parts1.Length, parts2.Length); i++)
                {
                    int p1 = i < parts1.Length && int.TryParse(parts1[i], out var n1) ? n1 : 0;
                    int p2 = i < parts2.Length && int.TryParse(parts2[i], out var n2) ? n2 : 0;

                    if (p1 > p2) return 1;
                    if (p1 < p2) return -1;
                }

                return 0;
            }
            catch
            {
                return 0;
            }
        }
    }
}
