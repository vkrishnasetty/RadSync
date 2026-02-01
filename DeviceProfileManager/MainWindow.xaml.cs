using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using DeviceProfileManager.Models;
using DeviceProfileManager.Services;

namespace DeviceProfileManager
{
    public partial class MainWindow : Window
    {
        private readonly ProfileManager _profileManager;
        private readonly UtilityChecker _utilityChecker;
        private AppSettings _settings;
        private readonly LogitechService _logitechService;
        private readonly StreamDeckService _streamDeckService;
        private readonly SpeechMicService _speechMicService;
        private readonly MosaicHotkeysService _mosaicHotkeysService;
        private readonly MosaicToolsService _mosaicToolsService;

        private const string ProfilesBasePath = @"H:\DeviceProfiles";

        public MainWindow()
        {
            InitializeComponent();

            _profileManager = new ProfileManager(ProfilesBasePath);
            _utilityChecker = new UtilityChecker(ProfilesBasePath);
            _settings = _profileManager.LoadSettings();

            _logitechService = new LogitechService();
            _streamDeckService = new StreamDeckService();
            _speechMicService = new SpeechMicService();
            _mosaicHotkeysService = new MosaicHotkeysService();
            _mosaicToolsService = new MosaicToolsService();

            LoadProfiles();
            LoadSettings();
            UpdateDeviceStatuses();
            UpdateAllDeviceRowStates();
            VersionText.Text = $"v{UpdateService.CurrentVersion}";
        }

        private void LoadProfiles()
        {
            var profiles = _profileManager.GetProfiles();
            ProfileComboBox.ItemsSource = profiles;

            if (!string.IsNullOrEmpty(_settings.LastSelectedProfile) && profiles.Contains(_settings.LastSelectedProfile))
            {
                ProfileComboBox.SelectedItem = _settings.LastSelectedProfile;
            }
            else if (profiles.Count > 0)
            {
                ProfileComboBox.SelectedIndex = 0;
            }
        }

        private void LoadSettings()
        {
            RunOnStartupCheckBox.IsChecked = StartupManager.IsRunOnStartupEnabled();
            AutoApplyCheckBox.IsChecked = _settings.AutoApplyOnStartup;

            if (_settings.DeviceStates.TryGetValue("Logitech", out var logitech))
                LogitechCheckBox.IsChecked = logitech;
            if (_settings.DeviceStates.TryGetValue("StreamDeck", out var streamDeck))
                StreamDeckCheckBox.IsChecked = streamDeck;
            if (_settings.DeviceStates.TryGetValue("SpeechMic", out var speechMic))
                SpeechMicCheckBox.IsChecked = speechMic;
            if (_settings.DeviceStates.TryGetValue("MosaicHotkeys", out var mosaicHotkeys))
                MosaicHotkeysCheckBox.IsChecked = mosaicHotkeys;
            if (_settings.DeviceStates.TryGetValue("MosaicTools", out var mosaicTools))
                MosaicToolsCheckBox.IsChecked = mosaicTools;

            // Auto-apply profile on startup if enabled
            if (_settings.AutoApplyOnStartup && !string.IsNullOrEmpty(_settings.LastSelectedProfile))
            {
                _ = ApplyProfileAsync(_settings.LastSelectedProfile);
            }
        }

        private void UpdateDeviceStatuses()
        {
            UpdateDeviceStatus(LogitechStatus, _logitechService.IsInstalled());
            UpdateDeviceStatus(StreamDeckStatus, _streamDeckService.IsInstalled());
            UpdateDeviceStatus(SpeechMicStatus, _speechMicService.IsInstalled());
            UpdateDeviceStatus(MosaicHotkeysStatus, _mosaicHotkeysService.IsInstalled());
            UpdateDeviceStatus(MosaicToolsStatus, _mosaicToolsService.IsInstalled());
        }

        private void UpdateDeviceStatus(TextBlock statusText, bool hasConfig)
        {
            if (hasConfig)
            {
                statusText.Text = "Config found";
                statusText.Foreground = FindResource("SuccessBrush") as SolidColorBrush ?? Brushes.Green;
            }
            else
            {
                statusText.Text = "No config";
                statusText.Foreground = FindResource("WarningBrush") as SolidColorBrush ?? Brushes.Orange;
            }
        }

        private void SaveSettings()
        {
            _settings.DeviceStates["Logitech"] = LogitechCheckBox.IsChecked ?? true;
            _settings.DeviceStates["StreamDeck"] = StreamDeckCheckBox.IsChecked ?? true;
            _settings.DeviceStates["SpeechMic"] = SpeechMicCheckBox.IsChecked ?? true;
            _settings.DeviceStates["MosaicHotkeys"] = MosaicHotkeysCheckBox.IsChecked ?? true;
            _settings.DeviceStates["MosaicTools"] = MosaicToolsCheckBox.IsChecked ?? true;
            _settings.LastSelectedProfile = ProfileComboBox.SelectedItem?.ToString() ?? "";
            _profileManager.SaveSettings(_settings);
        }

        private void DeviceCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox cb && cb.Tag is string deviceName)
            {
                UpdateDeviceRowState(deviceName, cb.IsChecked ?? false);
                SaveSettings();
            }
        }

        private void UpdateDeviceRowState(string deviceName, bool isEnabled)
        {
            double opacity = isEnabled ? 1.0 : 0.4;

            switch (deviceName)
            {
                case "Logitech":
                    SetRowEnabled(LogitechApplyBtn, LogitechSaveLink, LogitechRevertLink, LogitechSep1, LogitechSep2, isEnabled, opacity);
                    break;
                case "StreamDeck":
                    SetRowEnabled(StreamDeckApplyBtn, StreamDeckSaveLink, StreamDeckRevertLink, StreamDeckSep1, StreamDeckSep2, isEnabled, opacity);
                    break;
                case "SpeechMic":
                    SetRowEnabled(SpeechMicApplyBtn, SpeechMicSaveLink, SpeechMicRevertLink, SpeechMicSep1, SpeechMicSep2, isEnabled, opacity);
                    break;
                case "MosaicHotkeys":
                    SetRowEnabled(MosaicHotkeysApplyBtn, MosaicHotkeysSaveLink, MosaicHotkeysRevertLink, MosaicHotkeysSep1, MosaicHotkeysSep2, isEnabled, opacity);
                    break;
                case "MosaicTools":
                    SetRowEnabled(MosaicToolsApplyBtn, MosaicToolsSaveLink, MosaicToolsRevertLink, MosaicToolsSep1, MosaicToolsSep2, isEnabled, opacity);
                    break;
            }
        }

        private void SetRowEnabled(Button applyBtn, TextBlock saveLink, TextBlock revertLink, TextBlock sep1, TextBlock sep2, bool isEnabled, double opacity)
        {
            applyBtn.IsEnabled = isEnabled;
            saveLink.IsEnabled = isEnabled;
            saveLink.Opacity = opacity;
            revertLink.IsEnabled = isEnabled;
            revertLink.Opacity = opacity;
            sep1.Opacity = opacity;
            sep2.Opacity = opacity;
        }

        private void UpdateAllDeviceRowStates()
        {
            UpdateDeviceRowState("Logitech", LogitechCheckBox.IsChecked ?? false);
            UpdateDeviceRowState("StreamDeck", StreamDeckCheckBox.IsChecked ?? false);
            UpdateDeviceRowState("SpeechMic", SpeechMicCheckBox.IsChecked ?? false);
            UpdateDeviceRowState("MosaicHotkeys", MosaicHotkeysCheckBox.IsChecked ?? false);
            UpdateDeviceRowState("MosaicTools", MosaicToolsCheckBox.IsChecked ?? false);
        }

        private enum StatusType { Normal, Success, Error }

        private void SetStatus(string message, bool isError = false)
        {
            SetStatusEx(message, isError ? StatusType.Error : StatusType.Normal);
        }

        private void SetStatusSuccess(string message)
        {
            SetStatusEx(message, StatusType.Success);
        }

        private void SetStatusEx(string message, StatusType type)
        {
            StatusText.Text = message;
            StatusText.Foreground = type switch
            {
                StatusType.Success => FindResource("SuccessBrush") as SolidColorBrush ?? Brushes.LimeGreen,
                StatusType.Error => FindResource("ErrorBrush") as SolidColorBrush ?? Brushes.Red,
                _ => FindResource("SecondaryTextBrush") as SolidColorBrush ?? Brushes.Gray
            };
        }

        private string GetSelectedProfile()
        {
            return ProfileComboBox.SelectedItem?.ToString();
        }

        private List<string> GetEnabledDevices()
        {
            var devices = new List<string>();
            if (LogitechCheckBox.IsChecked == true) devices.Add("Logitech");
            if (StreamDeckCheckBox.IsChecked == true) devices.Add("StreamDeck");
            if (SpeechMicCheckBox.IsChecked == true) devices.Add("SpeechMic");
            if (MosaicHotkeysCheckBox.IsChecked == true) devices.Add("MosaicHotkeys");
            if (MosaicToolsCheckBox.IsChecked == true) devices.Add("MosaicTools");
            return devices;
        }

        // Profile Selection
        private void ProfileComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ProfileComboBox.SelectedItem != null)
            {
                // Load profile's enabled devices
                var profile = _profileManager.GetProfile(ProfileComboBox.SelectedItem.ToString());
                if (profile != null && profile.EnabledDevices != null)
                {
                    if (profile.EnabledDevices.TryGetValue("Logitech", out var logitech))
                        LogitechCheckBox.IsChecked = logitech;
                    if (profile.EnabledDevices.TryGetValue("StreamDeck", out var streamDeck))
                        StreamDeckCheckBox.IsChecked = streamDeck;
                    if (profile.EnabledDevices.TryGetValue("SpeechMic", out var speechMic))
                        SpeechMicCheckBox.IsChecked = speechMic;
                    if (profile.EnabledDevices.TryGetValue("MosaicHotkeys", out var mosaicHotkeys))
                        MosaicHotkeysCheckBox.IsChecked = mosaicHotkeys;
                    if (profile.EnabledDevices.TryGetValue("MosaicTools", out var mosaicTools))
                        MosaicToolsCheckBox.IsChecked = mosaicTools;
                }
                UpdateAllDeviceRowStates();
                SaveSettings();
            }
        }

        // Profile Management
        private void AddProfile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new InputDialog("New Profile", "Enter profile name:");
            if (dialog.ShowDialog() == true && !string.IsNullOrWhiteSpace(dialog.ResponseText))
            {
                if (_profileManager.CreateProfile(dialog.ResponseText))
                {
                    LoadProfiles();
                    ProfileComboBox.SelectedItem = dialog.ResponseText;
                    SetStatusSuccess($"Profile '{dialog.ResponseText}' created");
                }
                else
                {
                    SetStatus("Failed to create profile (may already exist)", true);
                }
            }
        }

        private void DeleteProfile_Click(object sender, RoutedEventArgs e)
        {
            var profile = GetSelectedProfile();
            if (profile == null)
            {
                SetStatus("No profile selected", true);
                return;
            }

            if (MessageBox.Show($"Delete profile '{profile}'?", "Confirm Delete",
                MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                if (_profileManager.DeleteProfile(profile))
                {
                    LoadProfiles();
                    SetStatusSuccess($"Profile '{profile}' deleted");
                }
                else
                {
                    SetStatus("Failed to delete profile", true);
                }
            }
        }

        private void RenameProfile_Click(object sender, RoutedEventArgs e)
        {
            var profile = GetSelectedProfile();
            if (profile == null)
            {
                SetStatus("No profile selected", true);
                return;
            }

            var dialog = new InputDialog("Rename Profile", "Enter new name:", profile);
            if (dialog.ShowDialog() == true && !string.IsNullOrWhiteSpace(dialog.ResponseText))
            {
                if (_profileManager.RenameProfile(profile, dialog.ResponseText))
                {
                    LoadProfiles();
                    ProfileComboBox.SelectedItem = dialog.ResponseText;
                    SetStatusSuccess($"Profile renamed to '{dialog.ResponseText}'");
                }
                else
                {
                    SetStatus("Failed to rename profile", true);
                }
            }
        }

        // Apply Profile
        private async void ApplyProfile_Click(object sender, RoutedEventArgs e)
        {
            var profile = GetSelectedProfile();
            if (profile == null)
            {
                SetStatus("No profile selected", true);
                return;
            }

            await ApplyProfileAsync(profile);
        }

        // Save to Profile
        private async void SaveProfile_Click(object sender, RoutedEventArgs e)
        {
            var profile = GetSelectedProfile();
            if (profile == null)
            {
                SetStatus("No profile selected", true);
                return;
            }

            await SaveToProfileAsync(profile);
        }

        private async Task ApplyProfileAsync(string profileName)
        {
            SetStatus($"Applying profile '{profileName}' (closing apps, copying files, restarting)...");
            var enabledDevices = GetEnabledDevices();
            var results = await _profileManager.ApplyAllAsync(profileName, enabledDevices);

            var failures = results.Where(r => !r.Value).Select(r => r.Key).ToList();
            if (failures.Count > 0)
            {
                SetStatus($"Failed to apply: {string.Join(", ", failures)}.", true);
            }
            else
            {
                SetStatusSuccess($"Profile '{profileName}' applied. Apps restarted with new settings.");
            }
        }

        private async Task SaveToProfileAsync(string profileName)
        {
            SetStatus($"Saving to profile '{profileName}' (closing apps to capture settings)...");
            var enabledDevices = GetEnabledDevices();
            var results = await _profileManager.SaveAllAsync(profileName, enabledDevices);

            var failures = results.Where(r => !r.Value).Select(r => r.Key).ToList();
            if (failures.Count > 0)
            {
                SetStatus($"Failed to save: {string.Join(", ", failures)}", true);
            }
            else
            {
                SetStatusSuccess($"Saved to profile '{profileName}'");
            }
        }

        // Individual Device Save (MouseDown events for text links)
        private async void SaveLogitech_Click(object sender, MouseButtonEventArgs e)
        {
            await SaveDeviceAsync(_logitechService);
        }

        private async void SaveStreamDeck_Click(object sender, MouseButtonEventArgs e)
        {
            await SaveDeviceAsync(_streamDeckService);
        }

        private async void SaveSpeechMic_Click(object sender, MouseButtonEventArgs e)
        {
            await SaveDeviceAsync(_speechMicService);
        }

        private async void SaveMosaicHotkeys_Click(object sender, MouseButtonEventArgs e)
        {
            await SaveDeviceAsync(_mosaicHotkeysService);
        }

        private async void SaveMosaicTools_Click(object sender, MouseButtonEventArgs e)
        {
            await SaveDeviceAsync(_mosaicToolsService);
        }

        // Individual Device Apply
        private async void ApplyLogitech_Click(object sender, RoutedEventArgs e)
        {
            await ApplyDeviceAsync(_logitechService);
        }

        private async void ApplyStreamDeck_Click(object sender, RoutedEventArgs e)
        {
            await ApplyDeviceAsync(_streamDeckService);
        }

        private async void ApplySpeechMic_Click(object sender, RoutedEventArgs e)
        {
            await ApplyDeviceAsync(_speechMicService);
        }

        private async void ApplyMosaicHotkeys_Click(object sender, RoutedEventArgs e)
        {
            await ApplyDeviceAsync(_mosaicHotkeysService);
        }

        private async void ApplyMosaicTools_Click(object sender, RoutedEventArgs e)
        {
            await ApplyDeviceAsync(_mosaicToolsService);
        }

        // Individual Device Revert
        private async void RevertLogitech_Click(object sender, MouseButtonEventArgs e)
        {
            await RevertDeviceAsync(_logitechService);
        }

        private async void RevertStreamDeck_Click(object sender, MouseButtonEventArgs e)
        {
            await RevertDeviceAsync(_streamDeckService);
        }

        private async void RevertSpeechMic_Click(object sender, MouseButtonEventArgs e)
        {
            await RevertDeviceAsync(_speechMicService);
        }

        private async void RevertMosaicHotkeys_Click(object sender, MouseButtonEventArgs e)
        {
            await RevertDeviceAsync(_mosaicHotkeysService);
        }

        private async void RevertMosaicTools_Click(object sender, MouseButtonEventArgs e)
        {
            await RevertDeviceAsync(_mosaicToolsService);
        }

        private async Task SaveDeviceAsync(IDeviceService service)
        {
            var profile = GetSelectedProfile();
            if (profile == null)
            {
                SetStatus("No profile selected", true);
                return;
            }

            // Show status about closing apps for Logitech and StreamDeck
            if (service.DeviceName == "Logitech" && LogitechService.IsRunning())
            {
                SetStatus("Closing Logitech G Hub to save settings...");
            }
            else if (service.DeviceName == "StreamDeck" && StreamDeckService.IsRunning())
            {
                SetStatus("Closing Stream Deck to save settings...");
            }
            else
            {
                SetStatus($"Saving {service.DisplayName}...");
            }

            var success = await service.ExportAsync(_profileManager.GetProfilePath(profile));
            if (success)
            {
                SetStatusSuccess($"{service.DisplayName} saved to profile");
            }
            else
            {
                SetStatus($"Failed to save {service.DisplayName}", true);
            }
        }

        private async Task ApplyDeviceAsync(IDeviceService service)
        {
            var profile = GetSelectedProfile();
            if (profile == null)
            {
                SetStatus("No profile selected", true);
                return;
            }

            var profilePath = _profileManager.GetProfilePath(profile);

            if (!service.HasConfigData(profilePath))
            {
                SetStatus($"No {service.DisplayName} data in profile. Save it first on source machine.", true);
                return;
            }

            // Show status about closing and restarting apps
            if (service.DeviceName == "Logitech")
            {
                SetStatus("Backing up current settings, then applying...");
            }
            else if (service.DeviceName == "StreamDeck")
            {
                SetStatus("Backing up current settings, then applying...");
            }
            else
            {
                SetStatus($"Backing up and applying {service.DisplayName}...");
            }

            // Backup current state before applying
            await _profileManager.BackupDeviceAsync(service);

            var success = await service.ImportAsync(profilePath);

            if (success)
            {
                SetStatusSuccess($"{service.DisplayName} applied. Use Revert to undo.");
            }
            else
            {
                SetStatus($"Failed to apply {service.DisplayName}.", true);
            }
        }

        private async Task RevertDeviceAsync(IDeviceService service)
        {
            if (!_profileManager.HasBackup(service))
            {
                SetStatus($"No backup available for {service.DisplayName}. Apply first to create a backup.", true);
                return;
            }

            SetStatus($"Reverting {service.DisplayName} to previous state...");

            var success = await _profileManager.RevertDeviceAsync(service);

            if (success)
            {
                SetStatusSuccess($"{service.DisplayName} reverted to previous state.");
            }
            else
            {
                SetStatus($"Failed to revert {service.DisplayName}.", true);
            }
        }

        // Bulk Operations
        private async void SaveAll_Click(object sender, RoutedEventArgs e)
        {
            var profile = GetSelectedProfile();
            if (profile == null)
            {
                SetStatus("No profile selected", true);
                return;
            }

            await SaveToProfileAsync(profile);
        }

        private async void ApplyAll_Click(object sender, RoutedEventArgs e)
        {
            var profile = GetSelectedProfile();
            if (profile == null)
            {
                SetStatus("No profile selected", true);
                return;
            }

            await ApplyProfileAsync(profile);
        }

        // Utility Checker
        private void CheckUtilities_Click(object sender, RoutedEventArgs e)
        {
            UpdateDeviceStatuses();

            // Test mode: Hold Shift to simulate all utilities missing (for testing dialog UI)
            // bool testMode = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift);

            // Get actual install status
            var logitechInstalled = _utilityChecker.IsLogitechInstalled();
            var streamDeckInstalled = _utilityChecker.IsStreamDeckInstalled();
            var speechMicInstalled = _utilityChecker.IsSpeechMicInstalled();

            if (logitechInstalled && streamDeckInstalled && speechMicInstalled)
            {
                SetStatusSuccess("All device utilities are installed");
            }

            // Show install dialog with individual download buttons (always allow access for updates)
            var dialog = new UtilityInstallDialog(logitechInstalled, streamDeckInstalled, speechMicInstalled, _utilityChecker);
            dialog.Owner = this;
            dialog.ShowDialog();
        }

        // Startup Options
        private void RunOnStartup_Changed(object sender, RoutedEventArgs e)
        {
            var enable = RunOnStartupCheckBox.IsChecked == true;
            if (!StartupManager.SetRunOnStartup(enable))
            {
                SetStatus("Failed to update startup setting", true);
                RunOnStartupCheckBox.IsChecked = !enable;
            }
            else
            {
                _settings.RunOnStartup = enable;
                SaveSettings();
            }
        }

        private void AutoApply_Changed(object sender, RoutedEventArgs e)
        {
            _settings.AutoApplyOnStartup = AutoApplyCheckBox.IsChecked == true;
            SaveSettings();
        }

        private async void CheckUpdates_Click(object sender, MouseButtonEventArgs e)
        {
            SetStatus("Checking for updates...");

            var (hasUpdate, latestVersion, downloadUrl) = await UpdateService.CheckForUpdatesAsync();

            if (hasUpdate)
            {
                var result = MessageBox.Show(
                    $"A new version ({latestVersion}) is available. Current version: {UpdateService.CurrentVersion}\n\nDownload and install automatically?",
                    "Update Available",
                    MessageBoxButton.YesNoCancel,
                    MessageBoxImage.Information);

                if (result == MessageBoxResult.Yes)
                {
                    // Auto-download and install
                    var progress = new Progress<string>(msg => SetStatus(msg));
                    var (success, message) = await UpdateService.DownloadAndInstallUpdateAsync(downloadUrl, progress);

                    if (success && message.Contains("Restarting"))
                    {
                        // Close the application to allow update
                        Application.Current.Shutdown();
                    }
                    else
                    {
                        SetStatus(message, !success);
                    }
                }
                else if (result == MessageBoxResult.No)
                {
                    // Open releases page for manual download
                    UpdateService.OpenReleasesPage();
                    SetStatus($"Opened releases page for v{latestVersion}");
                }
                else
                {
                    SetStatus($"Update available: v{latestVersion}");
                }
            }
            else if (latestVersion != null)
            {
                SetStatusSuccess($"You have the latest version (v{UpdateService.CurrentVersion})");
            }
            else
            {
                SetStatus("Could not check for updates. Check your internet connection.", true);
            }
        }
    }

    // Simple input dialog for profile name entry
    public class InputDialog : Window
    {
        private readonly TextBox _textBox;
        public string ResponseText => _textBox.Text;

        public InputDialog(string title, string prompt, string defaultValue = "")
        {
            Title = title;
            Width = 350;
            Height = 150;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;

            var grid = new Grid { Margin = new Thickness(16) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var label = new TextBlock { Text = prompt, Margin = new Thickness(0, 0, 0, 8) };
            Grid.SetRow(label, 0);
            grid.Children.Add(label);

            _textBox = new TextBox { Text = defaultValue, Margin = new Thickness(0, 0, 0, 16) };
            Grid.SetRow(_textBox, 1);
            grid.Children.Add(_textBox);

            var buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            Grid.SetRow(buttonPanel, 2);

            var okButton = new Button { Content = "OK", Width = 75, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            okButton.Click += (s, e) => { DialogResult = true; Close(); };
            buttonPanel.Children.Add(okButton);

            var cancelButton = new Button { Content = "Cancel", Width = 75, IsCancel = true };
            cancelButton.Click += (s, e) => { DialogResult = false; Close(); };
            buttonPanel.Children.Add(cancelButton);

            grid.Children.Add(buttonPanel);
            Content = grid;
        }
    }

    // Dialog for downloading/updating utilities - dark mode themed
    public class UtilityInstallDialog : Window
    {
        private readonly UtilityChecker _utilityChecker;

        public UtilityInstallDialog(bool logitechInstalled, bool streamDeckInstalled, bool speechMicInstalled, UtilityChecker utilityChecker)
        {
            _utilityChecker = utilityChecker;

            Title = "Download Utilities";
            Width = 450;
            Height = 380;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(Color.FromRgb(0x0F, 0x17, 0x2A)); // BackgroundColor

            var border = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x29, 0x3B)), // CardBackground
                BorderBrush = new SolidColorBrush(Color.FromRgb(0x33, 0x41, 0x55)), // BorderColor
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(12),
                Margin = new Thickness(16),
                Padding = new Thickness(24)
            };

            var mainStack = new StackPanel();

            var headerText = new TextBlock
            {
                Text = "Download Utilities",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(0xF1, 0xF5, 0xF9)), // TextColor
                Margin = new Thickness(0, 0, 0, 12)
            };
            mainStack.Children.Add(headerText);

            var noteText = new TextBlock
            {
                Text = "Click Download to open the vendor's download page. You can download updates even if already installed.",
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(0x94, 0xA3, 0xB8)), // SecondaryTextColor
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 20)
            };
            mainStack.Children.Add(noteText);

            // Logitech row
            mainStack.Children.Add(CreateUtilityRow("Logitech G Hub", logitechInstalled, () =>
            {
                _utilityChecker.OpenLogitechDownloadPage();
            }));

            // Stream Deck row
            mainStack.Children.Add(CreateUtilityRow("Elgato Stream Deck", streamDeckInstalled, () =>
            {
                _utilityChecker.OpenStreamDeckDownloadPage();
            }));

            // SpeechMic row
            mainStack.Children.Add(CreateUtilityRow("Philips Device Control Center", speechMicInstalled, () =>
            {
                _utilityChecker.OpenSpeechMicDownloadPage();
            }));

            // Close button
            var closeButton = new Button
            {
                Content = "Close",
                Width = 80,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 20, 0, 0),
                Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x29, 0x3B)),
                Foreground = new SolidColorBrush(Color.FromRgb(0xF1, 0xF5, 0xF9)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0x33, 0x41, 0x55)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(12, 8, 12, 8),
                Cursor = Cursors.Hand
            };
            closeButton.Click += (s, e) => Close();
            mainStack.Children.Add(closeButton);

            border.Child = mainStack;
            Content = border;
        }

        private Grid CreateUtilityRow(string name, bool isInstalled, Action downloadAction)
        {
            var grid = new Grid { Margin = new Thickness(0, 0, 0, 14) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var nameStack = new StackPanel { VerticalAlignment = VerticalAlignment.Center };

            var nameText = new TextBlock
            {
                Text = name,
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(0xF1, 0xF5, 0xF9))
            };
            nameStack.Children.Add(nameText);

            var statusText = new TextBlock
            {
                Text = isInstalled ? "Installed" : "Not installed",
                FontSize = 12,
                Foreground = isInstalled
                    ? new SolidColorBrush(Color.FromRgb(0x22, 0xC5, 0x5E)) // SuccessColor
                    : new SolidColorBrush(Color.FromRgb(0xF5, 0x9E, 0x0B)), // WarningColor
                Margin = new Thickness(0, 2, 0, 0)
            };
            nameStack.Children.Add(statusText);

            Grid.SetColumn(nameStack, 0);
            grid.Children.Add(nameStack);

            var downloadButton = new Button
            {
                Content = "Download",
                Width = 90,
                Background = new SolidColorBrush(Color.FromRgb(0x3B, 0x82, 0xF6)), // PrimaryColor
                Foreground = new SolidColorBrush(Colors.White),
                BorderThickness = new Thickness(0),
                Padding = new Thickness(12, 8, 12, 8),
                Cursor = Cursors.Hand
            };
            downloadButton.Click += (s, e) => downloadAction();

            Grid.SetColumn(downloadButton, 1);
            grid.Children.Add(downloadButton);

            return grid;
        }
    }
}
