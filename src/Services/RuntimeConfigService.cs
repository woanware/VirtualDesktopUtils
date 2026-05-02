using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.IO;
using System.Security;
using System.Security.Principal;
using Microsoft.Win32;

namespace VirtualDesktopUtils.Services;

internal sealed class RuntimeConfigService
{
    private const string ConfigDirectoryName = "VirtualDesktopUtils";
    private const string ConfigFileName = "config.json";
    private const string StartupRunRegistryPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string StartupRunValueName = "VirtualDesktopUtils";
    private const string StartupTaskName = "VirtualDesktopUtils";
    private const string UpstreamGuidSourceUrl = "https://raw.githubusercontent.com/MScholtes/VirtualDesktop/master/VirtualDesktop11-24H2.cs";
    private const string UpstreamGuidSourceName = "MScholtes/VirtualDesktop11-24H2.cs";
    private const int TaskActionExec = 0;
    private const int TaskCreateOrUpdate = 6;
    private const int TaskLogonInteractiveToken = 3;
    private const int TaskRunLevelHighest = 1;
    private const int TaskTriggerLogon = 9;

    private static readonly Guid DefaultImmersiveShellClsid = new("C2F03A33-21F5-47FA-B4BB-156362A2F239");
    private static readonly Guid DefaultVirtualDesktopManagerInternalServiceClsid = new("C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B");
    private static readonly Guid DefaultVirtualDesktopManagerClsid = new("AA509086-5CA9-4C25-8F95-589D3C07B48A");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    private static readonly Regex GuidDeclarationPattern = new(
        "CLSID_(?<name>[A-Za-z0-9_]+)\\s*=\\s*new\\s+Guid\\(\"(?<value>[0-9A-Fa-f-]{36})\"\\)",
        RegexOptions.Compiled);

    private readonly string _configDirectoryPath;
    private readonly string _configFilePath;
    private readonly object _sync = new();

    public RuntimeConfigService()
    {
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        _configDirectoryPath = Path.Combine(localAppData, ConfigDirectoryName);
        _configFilePath = Path.Combine(_configDirectoryPath, ConfigFileName);
    }

    public GuidConfiguration LoadGuidConfiguration()
    {
        var config = LoadConfig();
        return new GuidConfiguration(
            ImmersiveShellClsid: ParseOrDefault(config.Guids.ImmersiveShellClsid, DefaultImmersiveShellClsid),
            VirtualDesktopManagerInternalServiceClsid: ParseOrDefault(
                config.Guids.VirtualDesktopManagerInternalServiceClsid,
                DefaultVirtualDesktopManagerInternalServiceClsid),
            VirtualDesktopManagerClsid: ParseOrDefault(config.Guids.VirtualDesktopManagerClsid, DefaultVirtualDesktopManagerClsid),
            Source: string.IsNullOrWhiteSpace(config.Guids.Source) ? "defaults" : config.Guids.Source,
            LastUpdatedUtc: config.Guids.LastUpdatedUtc ?? string.Empty);
    }

    public bool IsGuidAutoUpdateOnStartupEnabled()
    {
        return LoadConfig().EnableGuidAutoUpdateOnStartup;
    }

    public bool IsAppUpdateCheckOnStartupEnabled()
    {
        return LoadConfig().EnableAppUpdateCheckOnStartup;
    }

    public void SetAppUpdateCheckOnStartupEnabled(bool enabled)
    {
        var config = LoadConfig();
        config.EnableAppUpdateCheckOnStartup = enabled;
        SaveConfig(config);
    }

    public string GetLastAppUpdateCheckUtc()
    {
        return LoadConfig().LastAppUpdateCheckUtc ?? string.Empty;
    }

    public void SetLastAppUpdateCheckUtc(string value)
    {
        var config = LoadConfig();
        config.LastAppUpdateCheckUtc = value;
        SaveConfig(config);
    }

    public bool IsStartWithWindowsEnabled()
    {
        return LoadConfig().EnableStartWithWindows;
    }

    public bool HasCompletedFirstLaunch()
    {
        return LoadConfig().HasCompletedFirstLaunch == true;
    }

    public void MarkFirstLaunchCompleted()
    {
        var config = LoadConfig();
        config.HasCompletedFirstLaunch = true;
        SaveConfig(config);
    }

    public (bool Success, string Message) SetStartWithWindowsEnabled(bool enabled)
    {
        var config = LoadConfig();
        var (success, message) = ApplyStartWithWindowsRegistration(enabled);
        if (!success)
        {
            return (false, message);
        }

        if (config.EnableStartWithWindows != enabled)
        {
            config.EnableStartWithWindows = enabled;
            SaveConfig(config);
        }

        return (true, message);
    }

    public (bool Success, string Message) EnsureStartWithWindowsSettingApplied()
    {
        var enabled = LoadConfig().EnableStartWithWindows;
        return ApplyStartWithWindowsRegistration(enabled);
    }

    public void SetGuidAutoUpdateOnStartupEnabled(bool enabled)
    {
        var config = LoadConfig();
        config.EnableGuidAutoUpdateOnStartup = enabled;
        SaveConfig(config);
    }

    public async Task<(bool Success, string Message)> SyncGuidConfigFromUpstreamAsync()
    {
        string upstreamSource;
        try
        {
            using var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(12) };
            upstreamSource = await httpClient.GetStringAsync(UpstreamGuidSourceUrl);
        }
        catch (HttpRequestException ex)
        {
            return (false, $"GUID sync failed: {ex.Message}");
        }
        catch (TaskCanceledException)
        {
            return (false, "GUID sync failed: request timed out.");
        }

        var guidMap = ParseGuidMap(upstreamSource);
        if (!guidMap.TryGetValue("ImmersiveShell", out var immersiveShell) ||
            !guidMap.TryGetValue("VirtualDesktopManagerInternal", out var managerInternal) ||
            !guidMap.TryGetValue("VirtualDesktopManager", out var manager))
        {
            return (false, "GUID sync failed: required GUIDs were not found in upstream source.");
        }

        var config = LoadConfig();
        config.Guids.ImmersiveShellClsid = immersiveShell;
        config.Guids.VirtualDesktopManagerInternalServiceClsid = managerInternal;
        config.Guids.VirtualDesktopManagerClsid = manager;
        config.Guids.Source = UpstreamGuidSourceName;
        config.Guids.LastUpdatedUtc = DateTime.UtcNow.ToString("O");

        try
        {
            SaveConfig(config);
        }
        catch (IOException ex)
        {
            return (false, $"GUID sync failed while saving config: {ex.Message}");
        }
        catch (UnauthorizedAccessException ex)
        {
            return (false, $"GUID sync failed while saving config: {ex.Message}");
        }

        return (true, $"GUID config updated from {UpstreamGuidSourceName}.");
    }

    private RuntimeConfig LoadConfig()
    {
        lock (_sync)
        {
            if (!File.Exists(_configFilePath))
            {
                var defaults = CreateDefaultConfig();
                SaveConfigInternal(defaults);
                return defaults;
            }

            try
            {
                var json = File.ReadAllText(_configFilePath);
                var parsed = JsonSerializer.Deserialize<RuntimeConfig>(json, JsonOptions);
                return NormalizeConfig(parsed, treatMissingFirstLaunchAsCompleted: true);
            }
            catch (JsonException)
            {
                var defaults = CreateDefaultConfig();
                SaveConfigInternal(defaults);
                return defaults;
            }
            catch (IOException)
            {
                return CreateDefaultConfig();
            }
            catch (UnauthorizedAccessException)
            {
                return CreateDefaultConfig();
            }
        }
    }

    private void SaveConfig(RuntimeConfig config)
    {
        lock (_sync)
        {
            SaveConfigInternal(NormalizeConfig(config));
        }
    }

    private void SaveConfigInternal(RuntimeConfig config)
    {
        Directory.CreateDirectory(_configDirectoryPath);
        var json = JsonSerializer.Serialize(config, JsonOptions);
        File.WriteAllText(_configFilePath, json);
    }

    private static Dictionary<string, string> ParseGuidMap(string source)
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (Match match in GuidDeclarationPattern.Matches(source))
        {
            var name = match.Groups["name"].Value;
            var guid = match.Groups["value"].Value;
            values[name] = guid;
        }

        return values;
    }

    private static (bool Success, string Message) ApplyStartWithWindowsRegistration(bool enabled)
    {
        try
        {
            if (enabled)
            {
                var startupCommand = GetStartupCommand();
                RegisterStartupTask(startupCommand.ExecutablePath, startupCommand.Arguments, startupCommand.WorkingDirectory);
                RemoveLegacyStartupRegistryValue();
                return (true, "Start with Windows enabled.");
            }

            UnregisterStartupTask();
            RemoveLegacyStartupRegistryValue();
            return (true, "Start with Windows disabled.");
        }
        catch (IOException ex)
        {
            return (false, $"Windows startup registration failed: {ex.Message}");
        }
        catch (COMException ex)
        {
            return (false, $"Windows startup registration failed: {ex.Message}");
        }
        catch (InvalidOperationException ex)
        {
            return (false, $"Windows startup registration failed: {ex.Message}");
        }
        catch (UnauthorizedAccessException ex)
        {
            return (false, $"Windows startup registration failed: {ex.Message}");
        }
        catch (SecurityException ex)
        {
            return (false, $"Windows startup registration failed: {ex.Message}");
        }
    }

    private static (string ExecutablePath, string Arguments, string WorkingDirectory) GetStartupCommand()
    {
        var processPath = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(processPath))
        {
            throw new InvalidOperationException("app executable path not found.");
        }

        var workingDirectory = Path.GetDirectoryName(processPath);
        if (string.IsNullOrWhiteSpace(workingDirectory))
        {
            throw new InvalidOperationException("app working directory not found.");
        }

        var entryAssemblyPath = Assembly.GetEntryAssembly()?.Location;
        if (Path.GetFileName(processPath).Equals("dotnet.exe", StringComparison.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(entryAssemblyPath))
            {
                throw new InvalidOperationException("app entry assembly path not found.");
            }

            return (processPath, $"\"{entryAssemblyPath}\"", Path.GetDirectoryName(entryAssemblyPath) ?? workingDirectory);
        }

        return (processPath, string.Empty, workingDirectory);
    }

    private static void RegisterStartupTask(string executablePath, string arguments, string workingDirectory)
    {
        var serviceType = Type.GetTypeFromProgID("Schedule.Service");
        if (serviceType is null)
        {
            throw new InvalidOperationException("Task Scheduler service is unavailable.");
        }

        var currentUser = WindowsIdentity.GetCurrent().Name;
        if (string.IsNullOrWhiteSpace(currentUser))
        {
            throw new InvalidOperationException("current Windows user could not be determined.");
        }

        dynamic? service = null;
        dynamic? rootFolder = null;
        dynamic? taskDefinition = null;
        dynamic? logonTrigger = null;
        dynamic? execAction = null;

        try
        {
            service = Activator.CreateInstance(serviceType)
                ?? throw new InvalidOperationException("Task Scheduler service could not be created.");
            service.Connect();
            rootFolder = service.GetFolder("\\");
            taskDefinition = service.NewTask(0);

            taskDefinition.RegistrationInfo.Description = "Launches VirtualDesktopUtils at user logon.";
            taskDefinition.Settings.Enabled = true;
            taskDefinition.Settings.StartWhenAvailable = true;
            taskDefinition.Settings.DisallowStartIfOnBatteries = false;
            taskDefinition.Settings.StopIfGoingOnBatteries = false;
            taskDefinition.Settings.MultipleInstances = 0;
            taskDefinition.Principal.UserId = currentUser;
            taskDefinition.Principal.LogonType = TaskLogonInteractiveToken;
            taskDefinition.Principal.RunLevel = TaskRunLevelHighest;

            logonTrigger = taskDefinition.Triggers.Create(TaskTriggerLogon);
            logonTrigger.Enabled = true;
            logonTrigger.UserId = currentUser;

            execAction = taskDefinition.Actions.Create(TaskActionExec);
            execAction.Path = executablePath;
            execAction.Arguments = arguments;
            execAction.WorkingDirectory = workingDirectory;

            rootFolder.RegisterTaskDefinition(
                StartupTaskName,
                taskDefinition,
                TaskCreateOrUpdate,
                Type.Missing,
                Type.Missing,
                TaskLogonInteractiveToken,
                Type.Missing);
        }
        finally
        {
            ReleaseComObject(execAction);
            ReleaseComObject(logonTrigger);
            ReleaseComObject(taskDefinition);
            ReleaseComObject(rootFolder);
            ReleaseComObject(service);
        }
    }

    private static void UnregisterStartupTask()
    {
        var serviceType = Type.GetTypeFromProgID("Schedule.Service");
        if (serviceType is null)
        {
            throw new InvalidOperationException("Task Scheduler service is unavailable.");
        }

        dynamic? service = null;
        dynamic? rootFolder = null;

        try
        {
            service = Activator.CreateInstance(serviceType)
                ?? throw new InvalidOperationException("Task Scheduler service could not be created.");
            service.Connect();
            rootFolder = service.GetFolder("\\");
            rootFolder.DeleteTask(StartupTaskName, 0);
        }
        catch (COMException ex) when ((uint)ex.HResult == 0x80070002)
        {
        }
        finally
        {
            ReleaseComObject(rootFolder);
            ReleaseComObject(service);
        }
    }

    private static void RemoveLegacyStartupRegistryValue()
    {
        using var runKey = Registry.CurrentUser.CreateSubKey(StartupRunRegistryPath, writable: true);
        runKey?.DeleteValue(StartupRunValueName, throwOnMissingValue: false);
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            Marshal.ReleaseComObject(value);
        }
    }

    private static RuntimeConfig NormalizeConfig(RuntimeConfig? config)
    private static RuntimeConfig NormalizeConfig(
        RuntimeConfig? config,
        bool treatMissingFirstLaunchAsCompleted = false)
    {
        var normalized = config ?? CreateDefaultConfig();
        normalized.Guids ??= new GuidConfigSection();

        normalized.HasCompletedFirstLaunch ??= treatMissingFirstLaunchAsCompleted;

        if (!Guid.TryParse(normalized.Guids.ImmersiveShellClsid, out _))
        {
            normalized.Guids.ImmersiveShellClsid = DefaultImmersiveShellClsid.ToString("D");
        }

        if (!Guid.TryParse(normalized.Guids.VirtualDesktopManagerInternalServiceClsid, out _))
        {
            normalized.Guids.VirtualDesktopManagerInternalServiceClsid = DefaultVirtualDesktopManagerInternalServiceClsid.ToString("D");
        }

        if (!Guid.TryParse(normalized.Guids.VirtualDesktopManagerClsid, out _))
        {
            normalized.Guids.VirtualDesktopManagerClsid = DefaultVirtualDesktopManagerClsid.ToString("D");
        }

        normalized.Guids.Source = string.IsNullOrWhiteSpace(normalized.Guids.Source)
            ? "defaults"
            : normalized.Guids.Source;

        normalized.Guids.LastUpdatedUtc ??= string.Empty;
        normalized.LastAppUpdateCheckUtc ??= string.Empty;
        normalized.PickerHotkey ??= new HotkeyConfigSection();
        normalized.MoveHotkey ??= new HotkeyConfigSection();
        return normalized;
    }

    private static RuntimeConfig CreateDefaultConfig()
    {
        return new RuntimeConfig
        {
            EnableAppUpdateCheckOnStartup = true,
            LastAppUpdateCheckUtc = string.Empty,
            HasCompletedFirstLaunch = false,
            EnableGuidAutoUpdateOnStartup = false,
            Guids = new GuidConfigSection
            {
                ImmersiveShellClsid = DefaultImmersiveShellClsid.ToString("D"),
                VirtualDesktopManagerInternalServiceClsid = DefaultVirtualDesktopManagerInternalServiceClsid.ToString("D"),
                VirtualDesktopManagerClsid = DefaultVirtualDesktopManagerClsid.ToString("D"),
                Source = "defaults",
                LastUpdatedUtc = string.Empty
            }
        };
    }

    private static Guid ParseOrDefault(string? value, Guid fallback)
    {
        return Guid.TryParse(value, out var parsed) ? parsed : fallback;
    }

    internal readonly record struct GuidConfiguration(
        Guid ImmersiveShellClsid,
        Guid VirtualDesktopManagerInternalServiceClsid,
        Guid VirtualDesktopManagerClsid,
        string Source,
        string LastUpdatedUtc);

    public HotkeyConfiguration LoadPickerHotkeyConfiguration()
    {
        var config = LoadConfig();
        return new HotkeyConfiguration(
            Modifiers: config.PickerHotkey.Modifiers,
            Vk: config.PickerHotkey.Vk,
            DisplayText: string.IsNullOrWhiteSpace(config.PickerHotkey.DisplayText) ? "Ctrl+Alt+Space" : config.PickerHotkey.DisplayText);
    }

    public void SavePickerHotkeyConfiguration(uint modifiers, uint vk, string displayText)
    {
        var config = LoadConfig();
        config.PickerHotkey.Modifiers = modifiers;
        config.PickerHotkey.Vk = vk;
        config.PickerHotkey.DisplayText = displayText;
        SaveConfig(config);
    }

    public HotkeyConfiguration LoadMoveHotkeyConfiguration()
    {
        var config = LoadConfig();
        return new HotkeyConfiguration(
            Modifiers: config.MoveHotkey.Modifiers,
            Vk: 0,
            DisplayText: string.IsNullOrWhiteSpace(config.MoveHotkey.DisplayText) ? "Ctrl+Alt" : config.MoveHotkey.DisplayText);
    }

    public void SaveMoveHotkeyConfiguration(uint modifiers, string displayText)
    {
        var config = LoadConfig();
        config.MoveHotkey.Modifiers = modifiers;
        config.MoveHotkey.DisplayText = displayText;
        SaveConfig(config);
    }

    internal readonly record struct HotkeyConfiguration(uint Modifiers, uint Vk, string DisplayText);

    private sealed class RuntimeConfig
    {
        public bool EnableAppUpdateCheckOnStartup { get; set; } = true;
        public string LastAppUpdateCheckUtc { get; set; } = string.Empty;
        public bool EnableGuidAutoUpdateOnStartup { get; set; }
        public bool EnableStartWithWindows { get; set; }
        public bool? HasCompletedFirstLaunch { get; set; }
        public GuidConfigSection Guids { get; set; } = new();
        public HotkeyConfigSection PickerHotkey { get; set; } = new();
        public HotkeyConfigSection MoveHotkey { get; set; } = new();
    }

    private sealed class GuidConfigSection
    {
        public string ImmersiveShellClsid { get; set; } = DefaultImmersiveShellClsid.ToString("D");
        public string VirtualDesktopManagerInternalServiceClsid { get; set; } =
            DefaultVirtualDesktopManagerInternalServiceClsid.ToString("D");
        public string VirtualDesktopManagerClsid { get; set; } = DefaultVirtualDesktopManagerClsid.ToString("D");
        public string Source { get; set; } = "defaults";
        public string LastUpdatedUtc { get; set; } = string.Empty;
    }

    private sealed class HotkeyConfigSection
    {
        public uint Modifiers { get; set; } = 0x0003;
        public uint Vk { get; set; } = 0x20; // VK_SPACE
        public string DisplayText { get; set; } = "Ctrl+Alt+Space";
    }
}
