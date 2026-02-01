using System;
using System.Diagnostics;
using System.Reflection;
using Microsoft.Win32;

namespace DeviceProfileManager.Services
{
    public static class StartupManager
    {
        private const string RegistryKeyPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private const string AppName = "DeviceProfileManager";

        public static bool IsRunOnStartupEnabled()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, false))
                {
                    return key?.GetValue(AppName) != null;
                }
            }
            catch
            {
                return false;
            }
        }

        public static bool SetRunOnStartup(bool enable)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, true))
                {
                    if (key == null)
                        return false;

                    if (enable)
                    {
                        var exePath = GetExecutablePath();
                        key.SetValue(AppName, $"\"{exePath}\"");
                    }
                    else
                    {
                        key.DeleteValue(AppName, false);
                    }
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private static string GetExecutablePath()
        {
            // For single-file apps, Environment.ProcessPath is the reliable way
            return Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName ?? "";
        }
    }
}
