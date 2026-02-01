using System;
using System.Diagnostics;
using System.IO;

namespace DeviceProfileManager.Services
{
    public class UtilityChecker
    {
        // Download page URLs - users download and install manually
        private const string LogitechDownloadPage = "https://www.logitechg.com/en-us/software/ghub";
        private const string StreamDeckDownloadPage = "https://www.elgato.com/us/en/s/downloads";
        private const string SpeechMicDownloadPage = "https://www.dictation.philips.com/us/products/desktop-dictation/speechmike-premium-dictation-microphone-lfh3500/#productsupport";

        public UtilityChecker(string basePath)
        {
            // No longer need installers path since we just open browser pages
        }

        public bool IsLogitechInstalled() => File.Exists(LogitechService.ExecutablePath);
        public bool IsStreamDeckInstalled() => File.Exists(StreamDeckService.ExecutablePath);
        public bool IsSpeechMicInstalled() => File.Exists(SpeechMicService.ExecutablePath);

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
