using System.Threading.Tasks;

namespace DeviceProfileManager.Services
{
    public interface IDeviceService
    {
        string DeviceName { get; }
        string DisplayName { get; }
        bool IsInstalled();
        Task<bool> ExportAsync(string profilePath);
        Task<bool> ImportAsync(string profilePath);
        bool HasConfigData(string profilePath);
    }
}
