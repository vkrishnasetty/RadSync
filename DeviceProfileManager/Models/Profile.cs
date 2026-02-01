using System;
using System.Collections.Generic;

namespace DeviceProfileManager.Models
{
    public class Profile
    {
        public string Name { get; set; } = "Default";
        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public DateTime LastModified { get; set; } = DateTime.Now;
        public Dictionary<string, bool> EnabledDevices { get; set; } = new Dictionary<string, bool>
        {
            { "Logitech", true },
            { "StreamDeck", true },
            { "SpeechMic", true },
            { "MosaicHotkeys", true },
            { "MosaicTools", true }
        };
        public string Notes { get; set; } = "";
    }
}
