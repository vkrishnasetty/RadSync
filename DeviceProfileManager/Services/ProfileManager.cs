using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using DeviceProfileManager.Models;

namespace DeviceProfileManager.Services
{
    public class ProfileManager
    {
        private readonly string _basePath;
        private readonly string _profilesPath;
        private readonly string _backupPath;
        private readonly string _settingsPath;
        private readonly List<IDeviceService> _deviceServices;

        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        public ProfileManager(string basePath)
        {
            _basePath = basePath;
            _profilesPath = Path.Combine(basePath, "profiles");
            _backupPath = Path.Combine(basePath, "backup");
            _settingsPath = Path.Combine(basePath, "settings.json");

            _deviceServices = new List<IDeviceService>
            {
                new LogitechService(),
                new StreamDeckService(),
                new SpeechMicService(),
                new MosaicHotkeysService(),
                new MosaicToolsService()
            };

            EnsureDirectoryStructure();
        }

        public IReadOnlyList<IDeviceService> DeviceServices => _deviceServices;
        public string BackupPath => _backupPath;

        private void EnsureDirectoryStructure()
        {
            Directory.CreateDirectory(_basePath);
            Directory.CreateDirectory(_profilesPath);
            Directory.CreateDirectory(_backupPath);
            Directory.CreateDirectory(Path.Combine(_basePath, "installers"));

            // Create default profile if none exist
            if (GetProfiles().Count == 0)
            {
                CreateProfile("Default");
            }
        }

        public List<string> GetProfiles()
        {
            var profiles = new List<string>();
            if (!Directory.Exists(_profilesPath))
                return profiles;

            foreach (var dir in Directory.GetDirectories(_profilesPath))
            {
                profiles.Add(Path.GetFileName(dir));
            }
            return profiles;
        }

        public Profile GetProfile(string name)
        {
            var profileJsonPath = Path.Combine(_profilesPath, name, "profile.json");
            if (!File.Exists(profileJsonPath))
                return null;

            var json = File.ReadAllText(profileJsonPath);
            return JsonSerializer.Deserialize<Profile>(json, JsonOptions);
        }

        public bool CreateProfile(string name)
        {
            try
            {
                var profileDir = Path.Combine(_profilesPath, name);
                if (Directory.Exists(profileDir))
                    return false;

                Directory.CreateDirectory(profileDir);

                var profile = new Profile { Name = name };
                SaveProfile(profile);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool DeleteProfile(string name)
        {
            try
            {
                var profileDir = Path.Combine(_profilesPath, name);
                if (!Directory.Exists(profileDir))
                    return false;

                Directory.Delete(profileDir, true);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool RenameProfile(string oldName, string newName)
        {
            try
            {
                var oldDir = Path.Combine(_profilesPath, oldName);
                var newDir = Path.Combine(_profilesPath, newName);

                if (!Directory.Exists(oldDir) || Directory.Exists(newDir))
                    return false;

                Directory.Move(oldDir, newDir);

                // Update profile.json with new name
                var profile = GetProfile(newName);
                if (profile != null)
                {
                    profile.Name = newName;
                    SaveProfile(profile);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void SaveProfile(Profile profile)
        {
            var profileDir = Path.Combine(_profilesPath, profile.Name);
            Directory.CreateDirectory(profileDir);

            profile.LastModified = DateTime.Now;
            var json = JsonSerializer.Serialize(profile, JsonOptions);
            File.WriteAllText(Path.Combine(profileDir, "profile.json"), json);
        }

        public string GetProfilePath(string profileName)
        {
            return Path.Combine(_profilesPath, profileName);
        }

        public async Task<Dictionary<string, bool>> SaveAllAsync(string profileName, IEnumerable<string> enabledDevices)
        {
            var results = new Dictionary<string, bool>();
            var profilePath = GetProfilePath(profileName);
            var enabledSet = new HashSet<string>(enabledDevices);

            foreach (var service in _deviceServices)
            {
                if (!enabledSet.Contains(service.DeviceName))
                {
                    results[service.DeviceName] = true; // Skipped
                    continue;
                }

                results[service.DeviceName] = await service.ExportAsync(profilePath);
            }

            // Update profile metadata with enabled devices
            var profile = GetProfile(profileName) ?? new Profile { Name = profileName };
            profile.LastModified = DateTime.Now;
            profile.EnabledDevices = new Dictionary<string, bool>();
            foreach (var service in _deviceServices)
            {
                profile.EnabledDevices[service.DeviceName] = enabledSet.Contains(service.DeviceName);
            }
            SaveProfile(profile);

            return results;
        }

        public async Task<Dictionary<string, bool>> ApplyAllAsync(string profileName, IEnumerable<string> enabledDevices)
        {
            var results = new Dictionary<string, bool>();
            var profilePath = GetProfilePath(profileName);
            var enabledSet = new HashSet<string>(enabledDevices);

            foreach (var service in _deviceServices)
            {
                if (!enabledSet.Contains(service.DeviceName))
                {
                    results[service.DeviceName] = true; // Skipped
                    continue;
                }

                if (service.HasConfigData(profilePath))
                {
                    // Backup current state before applying
                    await service.ExportAsync(_backupPath);
                    results[service.DeviceName] = await service.ImportAsync(profilePath);
                }
                else
                {
                    results[service.DeviceName] = true; // No data to apply
                }
            }

            return results;
        }

        /// <summary>
        /// Backup a single device's current state before applying
        /// </summary>
        public async Task<bool> BackupDeviceAsync(IDeviceService service)
        {
            return await service.ExportAsync(_backupPath);
        }

        /// <summary>
        /// Revert a device to its backed-up state
        /// </summary>
        public async Task<bool> RevertDeviceAsync(IDeviceService service)
        {
            if (!service.HasConfigData(_backupPath))
                return false;

            return await service.ImportAsync(_backupPath);
        }

        /// <summary>
        /// Check if a device has backup data available
        /// </summary>
        public bool HasBackup(IDeviceService service)
        {
            return service.HasConfigData(_backupPath);
        }

        public AppSettings LoadSettings()
        {
            try
            {
                if (File.Exists(_settingsPath))
                {
                    var json = File.ReadAllText(_settingsPath);
                    return JsonSerializer.Deserialize<AppSettings>(json, JsonOptions) ?? new AppSettings();
                }
            }
            catch { }
            return new AppSettings();
        }

        public void SaveSettings(AppSettings settings)
        {
            var json = JsonSerializer.Serialize(settings, JsonOptions);
            File.WriteAllText(_settingsPath, json);
        }
    }
}
