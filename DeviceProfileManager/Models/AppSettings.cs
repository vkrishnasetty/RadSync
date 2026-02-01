using System.Collections.Generic;

namespace DeviceProfileManager.Models
{
    public class AppSettings
    {
        public string ProfilesPath { get; set; } = @"H:\DeviceProfiles";
        public string LastSelectedProfile { get; set; } = "Default";
        public bool RunOnStartup { get; set; } = false;
        public bool AutoApplyOnStartup { get; set; } = false;
        public Dictionary<string, bool> DeviceStates { get; set; } = new Dictionary<string, bool>
        {
            { "Logitech", true },
            { "StreamDeck", true },
            { "SpeechMic", true },
            { "MosaicHotkeys", true },
            { "MosaicTools", true }
        };
    }
}
