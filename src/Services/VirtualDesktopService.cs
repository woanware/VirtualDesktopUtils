using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using VirtualDesktopUtils.Models;
using Microsoft.Win32;

namespace VirtualDesktopUtils.Services;

internal sealed class VirtualDesktopService
{
    private const string VirtualDesktopsRegistryPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops";
    private const string SessionInfoRegistryPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SessionInfo";

    private readonly RuntimeConfigService.GuidConfiguration _guidConfiguration;
    private readonly IVirtualDesktopManager? _virtualDesktopManager;
    private readonly IVirtualDesktopManagerInternal? _virtualDesktopManagerInternal;
    private readonly IApplicationViewCollection? _applicationViewCollection;
    private readonly IVirtualDesktopPinnedApps? _virtualDesktopPinnedApps;

    private string? _initError;

    public VirtualDesktopService(RuntimeConfigService runtimeConfigService)
    {
        _guidConfiguration = runtimeConfigService.LoadGuidConfiguration();

        try
        {
            _virtualDesktopManager = (IVirtualDesktopManager)Activator.CreateInstance(
                Type.GetTypeFromCLSID(_guidConfiguration.VirtualDesktopManagerClsid)!)!;
        }
        catch (Exception ex)
        {
            _virtualDesktopManager = null;
            _initError = $"IVirtualDesktopManager init failed: {ex.Message}";
        }

        try
        {
            var shell = (IServiceProvider10)Activator.CreateInstance(
                Type.GetTypeFromCLSID(_guidConfiguration.ImmersiveShellClsid)!)!;
            _virtualDesktopManagerInternal = (IVirtualDesktopManagerInternal)shell.QueryService(
                _guidConfiguration.VirtualDesktopManagerInternalServiceClsid,
                typeof(IVirtualDesktopManagerInternal).GUID);
            _applicationViewCollection = (IApplicationViewCollection)shell.QueryService(
                typeof(IApplicationViewCollection).GUID,
                typeof(IApplicationViewCollection).GUID);
            _virtualDesktopPinnedApps = (IVirtualDesktopPinnedApps)shell.QueryService(
                new Guid("B5A399E7-1C87-46B8-88E9-FC5747B171BD"),
                typeof(IVirtualDesktopPinnedApps).GUID);
        }
        catch (Exception ex)
        {
            _virtualDesktopManagerInternal = null;
            _applicationViewCollection = null;
            _virtualDesktopPinnedApps = null;
            _initError = (_initError is null ? "" : _initError + " | ") +
                         $"Internal COM init failed: {ex.GetType().Name}: {ex.Message}";
        }
    }

    public string GetDiagnostics()
    {
        var publicApi = _virtualDesktopManager is not null ? "OK" : "FAILED";
        var internalApi = _virtualDesktopManagerInternal is not null ? "OK" : "FAILED";
        var viewCollection = _applicationViewCollection is not null ? "OK" : "FAILED";
        var buildNumber = Environment.OSVersion.Version.Build;
        var guidSource = string.IsNullOrWhiteSpace(_guidConfiguration.Source) ? "defaults" : _guidConfiguration.Source;
        var guidUpdated = string.IsNullOrWhiteSpace(_guidConfiguration.LastUpdatedUtc) ? "n/a" : _guidConfiguration.LastUpdatedUtc;
        var result =
            $"Win build {buildNumber} | Public API: {publicApi} | Internal API: {internalApi} | ViewCollection: {viewCollection} | GUID source: {guidSource} ({guidUpdated})";
        if (_initError is not null)
        {
            result += $" | Error: {_initError}";
        }
        return result;
    }

    public IReadOnlyList<DesktopInfo> GetDesktops()
    {
        var desktopIds = ReadDesktopIdsFromRegistry();
        var currentDesktopId = GetCurrentDesktopId();

        if (desktopIds.Count == 0 && currentDesktopId.HasValue)
        {
            desktopIds.Add(currentDesktopId.Value);
        }

        var desktops = new List<DesktopInfo>(desktopIds.Count);
        for (var index = 0; index < desktopIds.Count; index++)
        {
            var id = desktopIds[index];
            var isCurrent = currentDesktopId.HasValue && currentDesktopId.Value == id;
            desktops.Add(new DesktopInfo(id, index + 1, ResolveDesktopName(id, index + 1), isCurrent));
        }

        return desktops;
    }

    public Guid? GetCurrentDesktopId()
    {
        return GetCurrentDesktopIdFromRegistry();
    }

    public Guid? GetWindowDesktopId(nint windowHandle)
    {
        if (windowHandle == nint.Zero)
        {
            return null;
        }

        if (_virtualDesktopManager is not null)
        {
            try
            {
                var desktopId = _virtualDesktopManager.GetWindowDesktopId(windowHandle);
                if (desktopId != Guid.Empty)
                {
                    return desktopId;
                }
            }
            catch
            {
            }
        }

        return null;
    }

    public bool IsWindowOnCurrentDesktop(nint windowHandle)
    {
        if (windowHandle == nint.Zero)
        {
            return false;
        }

        if (_virtualDesktopManager is not null)
        {
            try
            {
                return _virtualDesktopManager.IsWindowOnCurrentVirtualDesktop(windowHandle);
            }
            catch
            {
            }
        }

        return false;
    }

    public (bool Success, string? Error) MoveWindowToDesktop(nint windowHandle, Guid desktopId)
    {
        if (windowHandle == nint.Zero || desktopId == Guid.Empty)
        {
            return (false, "Invalid window handle or desktop ID.");
        }

        if (_virtualDesktopManagerInternal is null || _applicationViewCollection is null)
        {
            return (false, "Internal COM interfaces not available. " + (_initError ?? ""));
        }

        // Check if same-process window (MScholtes pattern)
        GetWindowThreadProcessId(windowHandle, out var processId);
        var isSameProcess = Process.GetCurrentProcess().Id == processId;

        if (isSameProcess)
        {
            // Same-process: try public API first, then internal
            try
            {
                _virtualDesktopManager?.MoveWindowToDesktop(windowHandle, ref desktopId);
                return (true, null);
            }
            catch
            {
                try
                {
                    _applicationViewCollection.GetViewForHwnd(windowHandle, out var view);
                    var target = _virtualDesktopManagerInternal.FindDesktop(ref desktopId);
                    _virtualDesktopManagerInternal.MoveViewToDesktop(view, target);
                    return (true, null);
                }
                catch (Exception ex)
                {
                    return (false, $"Same-process move failed: {ex.GetType().Name}: {ex.Message}");
                }
            }
        }

        // Cross-process: use internal API with MainWindowHandle fallback (MScholtes pattern)
        try
        {
            _applicationViewCollection.GetViewForHwnd(windowHandle, out var view);
            var target = _virtualDesktopManagerInternal.FindDesktop(ref desktopId);
            _virtualDesktopManagerInternal.MoveViewToDesktop(view, target);
            return (true, null);
        }
        catch (Exception ex)
        {
            // Fallback: try process MainWindowHandle instead
            try
            {
                var mainHandle = Process.GetProcessById((int)processId).MainWindowHandle;
                if (mainHandle != nint.Zero && mainHandle != windowHandle)
                {
                    _applicationViewCollection.GetViewForHwnd(mainHandle, out var view);
                    var target = _virtualDesktopManagerInternal.FindDesktop(ref desktopId);
                    _virtualDesktopManagerInternal.MoveViewToDesktop(view, target);
                    return (true, null);
                }
            }
            catch
            {
            }

            return (false, $"Cross-process move failed: {ex.GetType().Name}: {ex.Message}");
        }
    }

    public (bool Success, string? Error) SwitchToDesktop(Guid desktopId)
    {
        if (desktopId == Guid.Empty)
        {
            return (false, "Invalid desktop ID.");
        }

        if (_virtualDesktopManagerInternal is null)
        {
            return (false, "Internal COM interfaces not available. " + (_initError ?? ""));
        }

        try
        {
            var target = _virtualDesktopManagerInternal.FindDesktop(ref desktopId);
            _virtualDesktopManagerInternal.SwitchDesktop(target);
            return (true, null);
        }
        catch (Exception ex)
        {
            return (false, $"Switch desktop failed: {ex.GetType().Name}: {ex.Message}");
        }
    }

    public void PinWindow(nint windowHandle)
    {
        if (windowHandle == nint.Zero || _applicationViewCollection is null || _virtualDesktopPinnedApps is null)
        {
            return;
        }

        try
        {
            _applicationViewCollection.GetViewForHwnd(windowHandle, out var view);
            if (!_virtualDesktopPinnedApps.IsViewPinned(view))
            {
                _virtualDesktopPinnedApps.PinView(view);
            }
        }
        catch
        {
        }
    }

    public void UnpinWindow(nint windowHandle)
    {
        if (windowHandle == nint.Zero || _applicationViewCollection is null || _virtualDesktopPinnedApps is null)
        {
            return;
        }

        try
        {
            _applicationViewCollection.GetViewForHwnd(windowHandle, out var view);
            if (_virtualDesktopPinnedApps.IsViewPinned(view))
            {
                _virtualDesktopPinnedApps.UnpinView(view);
            }
        }
        catch
        {
        }
    }

    public bool IsWindowPinned(nint windowHandle)
    {
        if (windowHandle == nint.Zero || _applicationViewCollection is null || _virtualDesktopPinnedApps is null)
        {
            return false;
        }

        try
        {
            _applicationViewCollection.GetViewForHwnd(windowHandle, out var view);
            return _virtualDesktopPinnedApps.IsViewPinned(view);
        }
        catch
        {
            return false;
        }
    }

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint hWnd, out int processId);

    #region Registry helpers

    private static List<Guid> ReadDesktopIdsFromRegistry()
    {
        using var virtualDesktopsKey = Registry.CurrentUser.OpenSubKey(VirtualDesktopsRegistryPath);
        if (virtualDesktopsKey?.GetValue("VirtualDesktopIDs") is not byte[] rawIds || rawIds.Length < 16)
        {
            return new List<Guid>();
        }

        var ids = new List<Guid>(rawIds.Length / 16);
        for (var offset = 0; offset + 15 < rawIds.Length; offset += 16)
        {
            ids.Add(new Guid(rawIds.AsSpan(offset, 16)));
        }

        return ids.Distinct().ToList();
    }

    private static string ResolveDesktopName(Guid desktopId, int fallbackIndex)
    {
        using var desktopKey = OpenDesktopKey(desktopId);
        var nameValue = desktopKey?.GetValue("Name");

        if (nameValue is string directName && !string.IsNullOrWhiteSpace(directName))
        {
            return directName;
        }

        if (nameValue is byte[] binaryName && binaryName.Length > 0)
        {
            var decodedName = Encoding.Unicode.GetString(binaryName).TrimEnd('\0');
            if (!string.IsNullOrWhiteSpace(decodedName))
            {
                return decodedName;
            }
        }

        return $"Desktop {fallbackIndex}";
    }

    private static RegistryKey? OpenDesktopKey(Guid desktopId)
    {
        return Registry.CurrentUser.OpenSubKey($"{VirtualDesktopsRegistryPath}\\Desktops\\{desktopId:B}")
            ?? Registry.CurrentUser.OpenSubKey($"{VirtualDesktopsRegistryPath}\\Desktops\\{desktopId:D}");
    }

    private static Guid? GetCurrentDesktopIdFromRegistry()
    {
        var currentDesktopBytes = ReadCurrentDesktopIdBytes();
        if (currentDesktopBytes is null || currentDesktopBytes.Length != 16)
        {
            return null;
        }

        return new Guid(currentDesktopBytes);
    }

    private static byte[]? ReadCurrentDesktopIdBytes()
    {
        var currentSessionId = Process.GetCurrentProcess().SessionId;
        using var currentSessionKey = Registry.CurrentUser.OpenSubKey(
            $"{SessionInfoRegistryPath}\\{currentSessionId}\\VirtualDesktops");

        if (TryReadGuidBytes(currentSessionKey?.GetValue("CurrentVirtualDesktop"), out var directSessionDesktop))
        {
            return directSessionDesktop;
        }

        using var sessionInfoKey = Registry.CurrentUser.OpenSubKey(SessionInfoRegistryPath);
        if (sessionInfoKey is not null)
        {
            foreach (var sessionSubKeyName in sessionInfoKey.GetSubKeyNames())
            {
                using var virtualDesktopKey = Registry.CurrentUser.OpenSubKey(
                    $"{SessionInfoRegistryPath}\\{sessionSubKeyName}\\VirtualDesktops");

                if (TryReadGuidBytes(virtualDesktopKey?.GetValue("CurrentVirtualDesktop"), out var discoveredDesktop))
                {
                    return discoveredDesktop;
                }
            }
        }

        using var globalVirtualDesktopsKey = Registry.CurrentUser.OpenSubKey(VirtualDesktopsRegistryPath);
        return TryReadGuidBytes(globalVirtualDesktopsKey?.GetValue("CurrentVirtualDesktop"), out var globalDesktop)
            ? globalDesktop
            : null;
    }

    private static bool TryReadGuidBytes(object? value, out byte[] bytes)
    {
        if (value is byte[] rawBytes && rawBytes.Length >= 16)
        {
            bytes = rawBytes[..16];
            return true;
        }

        bytes = Array.Empty<byte>();
        return false;
    }

    #endregion

    #region COM interface definitions (from MScholtes/VirtualDesktop, Windows 11 22H2+/24H2)

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("6D5140C1-7436-11CE-8034-00AA006009FA")]
    private interface IServiceProvider10
    {
        [return: MarshalAs(UnmanagedType.IUnknown)]
        object QueryService(in Guid service, in Guid riid);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("A5CD92FF-29BE-454C-8D04-D82879FB3F1B")]
    private interface IVirtualDesktopManager
    {
        [return: MarshalAs(UnmanagedType.Bool)]
        bool IsWindowOnCurrentVirtualDesktop(nint topLevelWindow);

        Guid GetWindowDesktopId(nint topLevelWindow);

        void MoveWindowToDesktop(nint topLevelWindow, ref Guid desktopId);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("92CA9DCD-5622-4BBA-A805-5E9F541BD8C9")]
    private interface IObjectArray
    {
        void GetCount(out int count);
        void GetAt(int index, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object obj);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("372E1D3B-38D3-42E4-A15B-8AB2B178F513")]
    private interface IApplicationView
    {
        // IInspectable vtable slots (3 methods)
        int GetIids(out int iidCount, out nint iids);
        int GetRuntimeClassName(out nint className);
        int GetTrustLevel(out int trustLevel);

        // IApplicationView methods
        int SetFocus();
        int SwitchTo();
        int TryInvokeBack(nint callback);
        int GetThumbnailWindow(out nint hwnd);
        int GetMonitor(out nint immersiveMonitor);
        int GetVisibility(out int visibility);
        int SetCloak(int cloakType, int unknown);
        int GetPosition(ref Guid guid, out nint position);
        int SetPosition(ref nint position);
        int InsertAfterWindow(nint hwnd);
        int GetExtendedFramePosition(out long rect);
        int GetAppUserModelId([MarshalAs(UnmanagedType.LPWStr)] out string id);
        int SetAppUserModelId(string id);
        int IsEqualByAppUserModelId(string id, out int result);
        int GetViewState(out uint state);
        int SetViewState(uint state);
        int GetNeediness(out int neediness);
        int GetLastActivationTimestamp(out ulong timestamp);
        int SetLastActivationTimestamp(ulong timestamp);
        int GetVirtualDesktopId(out Guid guid);
        int SetVirtualDesktopId(ref Guid guid);
        int GetShowInSwitchers(out int flag);
        int SetShowInSwitchers(int flag);
        int GetScaleFactor(out int factor);
        int CanReceiveInput(out bool canReceiveInput);
        int GetCompatibilityPolicyType(out int flags);
        int SetCompatibilityPolicyType(int flags);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("1841C6D7-4F9D-42C0-AF41-8747538F10E5")]
    private interface IApplicationViewCollection
    {
        int GetViews(out IObjectArray array);
        int GetViewsByZOrder(out IObjectArray array);
        int GetViewsByAppUserModelId(string id, out IObjectArray array);
        int GetViewForHwnd(nint hwnd, out IApplicationView view);
        int GetViewForApplication(object application, out IApplicationView view);
        int GetViewForAppUserModelId(string id, out IApplicationView view);
        int GetViewInFocus(out nint view);
        int Unknown1(out nint view);
        void RefreshCollection();
        int RegisterForApplicationViewChanges(object listener, out int cookie);
        int UnregisterForApplicationViewChanges(int cookie);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("3F07F4BE-B107-441A-AF0F-39D82529072C")]
    private interface IVirtualDesktopInternal
    {
        bool IsViewVisible(IApplicationView view);
        Guid GetId();
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("53F5CA0B-158F-4124-900C-057158060B27")]
    private interface IVirtualDesktopManagerInternal
    {
        int GetCount();
        void MoveViewToDesktop(IApplicationView view, IVirtualDesktopInternal desktop);
        bool CanViewMoveDesktops(IApplicationView view);
        IVirtualDesktopInternal GetCurrentDesktop();
        void GetDesktops(out IObjectArray desktops);
        [PreserveSig]
        int GetAdjacentDesktop(IVirtualDesktopInternal from, int direction, out IVirtualDesktopInternal desktop);
        void SwitchDesktop(IVirtualDesktopInternal desktop);
        void SwitchDesktopAndMoveForegroundView(IVirtualDesktopInternal desktop);
        IVirtualDesktopInternal CreateDesktop();
        void MoveDesktop(IVirtualDesktopInternal desktop, int nIndex);
        void RemoveDesktop(IVirtualDesktopInternal desktop, IVirtualDesktopInternal fallback);
        IVirtualDesktopInternal FindDesktop(ref Guid desktopId);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("4CE81583-1E4C-4632-A621-07A53543148F")]
    private interface IVirtualDesktopPinnedApps
    {
        bool IsAppIdPinned(string appId);
        void PinAppID(string appId);
        void UnpinAppID(string appId);
        bool IsViewPinned(IApplicationView view);
        void PinView(IApplicationView view);
        void UnpinView(IApplicationView view);
    }

    #endregion
}
