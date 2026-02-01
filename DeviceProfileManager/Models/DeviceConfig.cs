namespace DeviceProfileManager.Models
{
    public class DeviceConfig
    {
        public string Name { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public bool IsInstalled { get; set; }
        public bool IsEnabled { get; set; } = true;
        public string ExecutablePath { get; set; } = "";
        public string ConfigPath { get; set; } = "";
        public string DownloadUrl { get; set; } = "";
        public string InstallerFileName { get; set; } = "";
    }

    public enum DeviceType
    {
        Logitech,
        StreamDeck,
        SpeechMic
    }
}
