using VirtualDesktopUtils.Models;
using System.Runtime.InteropServices;

namespace VirtualDesktopUtils.Services;

internal sealed class WindowMenuInjectionService : IDisposable
{
    private const uint EventSystemForeground = 0x0003;
    private const uint EventSystemMenuStart = 0x0004;
    private const uint EventObjectInvoked = 0x8013;
    private const uint WinEventOutOfContext = 0x0000;
    private const uint WinEventSkipOwnProcess = 0x0002;
    private const uint GaRoot = 0x0002;
    private const uint MfString = 0x0000;
    private const uint MfPopup = 0x0010;
    private const uint MfSeparator = 0x0800;

    private readonly VirtualDesktopService _virtualDesktopService;
    private readonly WindowDiscoveryService _windowDiscoveryService;
    private readonly Dictionary<int, MenuCommand> _menuCommands = new();
    private readonly object _sync = new();

    private WinEventDelegate? _winEventDelegate;
    private nint _foregroundHook;
    private nint _menuStartHook;
    private nint _objectInvokedHook;
    private nint _lastForegroundCandidateWindow;
    private int _nextCommandId = 0x1A00;
    private Action<string>? _statusSink;

    public WindowMenuInjectionService(
        VirtualDesktopService virtualDesktopService,
        WindowDiscoveryService windowDiscoveryService)
    {
        _virtualDesktopService = virtualDesktopService;
        _windowDiscoveryService = windowDiscoveryService;
    }

    public void Start(Action<string> statusSink)
    {
        _statusSink = statusSink;
        if (_foregroundHook != nint.Zero || _menuStartHook != nint.Zero || _objectInvokedHook != nint.Zero)
        {
            return;
        }

        _winEventDelegate = HandleWinEvent;
        _foregroundHook = SetWinEventHook(
            EventSystemForeground,
            EventSystemForeground,
            nint.Zero,
            _winEventDelegate,
            0,
            0,
            WinEventOutOfContext | WinEventSkipOwnProcess);
        _menuStartHook = SetWinEventHook(
            EventSystemMenuStart,
            EventSystemMenuStart,
            nint.Zero,
            _winEventDelegate,
            0,
            0,
            WinEventOutOfContext | WinEventSkipOwnProcess);
        _objectInvokedHook = SetWinEventHook(
            EventObjectInvoked,
            EventObjectInvoked,
            nint.Zero,
            _winEventDelegate,
            0,
            0,
            WinEventOutOfContext | WinEventSkipOwnProcess);
    }

    public void Stop()
    {
        if (_foregroundHook != nint.Zero)
        {
            _ = UnhookWinEvent(_foregroundHook);
            _foregroundHook = nint.Zero;
        }

        if (_menuStartHook != nint.Zero)
        {
            _ = UnhookWinEvent(_menuStartHook);
            _menuStartHook = nint.Zero;
        }

        if (_objectInvokedHook != nint.Zero)
        {
            _ = UnhookWinEvent(_objectInvokedHook);
            _objectInvokedHook = nint.Zero;
        }

        lock (_sync)
        {
            _menuCommands.Clear();
            _lastForegroundCandidateWindow = nint.Zero;
        }
    }

    public void Dispose()
    {
        Stop();
    }

    public void InvalidateCommands()
    {
        lock (_sync)
        {
            _menuCommands.Clear();
        }
    }

    public nint GetLastForegroundCandidateWindow()
    {
        lock (_sync)
        {
            return _lastForegroundCandidateWindow;
        }
    }

    private void HandleWinEvent(
        nint hookHandle,
        uint eventType,
        nint windowHandle,
        int objectId,
        int childId,
        uint eventThread,
        uint eventTime)
    {
        if (windowHandle == nint.Zero)
        {
            return;
        }

        try
        {
            var topLevelWindow = ResolveTopLevelWindow(windowHandle);
            if (eventType == EventSystemForeground)
            {
                lock (_sync)
                {
                    _lastForegroundCandidateWindow = topLevelWindow;
                }

                InjectMenu(topLevelWindow, resetMenu: true);
                return;
            }

            if (eventType == EventSystemMenuStart)
            {
                InjectMenu(topLevelWindow, resetMenu: false);
                return;
            }

            if (eventType == EventObjectInvoked)
            {
                HandleMenuCommand(childId);
            }
        }
        catch (KeyNotFoundException)
        {
            _statusSink?.Invoke("Desktop event handling failed; retry the action.");
        }
        catch (COMException)
        {
            _statusSink?.Invoke("Desktop event handling failed due to Windows desktop API state.");
        }
        catch (UnauthorizedAccessException)
        {
            _statusSink?.Invoke("Desktop event handling was denied; try running elevated.");
        }
        catch (InvalidOperationException)
        {
            _statusSink?.Invoke("Desktop event handling failed because the target window changed state.");
        }
        catch (ArgumentException)
        {
            _statusSink?.Invoke("Desktop event handling received an invalid target.");
        }
    }

    private void InjectMenu(nint windowHandle, bool resetMenu)
    {
        if (!_windowDiscoveryService.IsCandidateWindow(windowHandle))
        {
            return;
        }

        var desktops = _virtualDesktopService.GetDesktops();
        if (desktops.Count <= 1)
        {
            return;
        }

        if (!resetMenu && HasCommandsForWindow(windowHandle))
        {
            return;
        }

        if (resetMenu)
        {
            _ = GetSystemMenu(windowHandle, true);
        }

        var systemMenu = GetSystemMenu(windowHandle, false);
        if (systemMenu == nint.Zero)
        {
            return;
        }

        var popupMenu = CreatePopupMenu();
        if (popupMenu == nint.Zero)
        {
            return;
        }

        var windowDesktopId = _virtualDesktopService.GetWindowDesktopId(windowHandle);
        ClearWindowCommands(windowHandle);

        foreach (var desktop in desktops.Where(desktop => desktop.Id != windowDesktopId))
        {
            var commandId = Interlocked.Increment(ref _nextCommandId);
            if (!AppendMenu(popupMenu, MfString, (nuint)commandId, desktop.DisplayName))
            {
                continue;
            }

            lock (_sync)
            {
                _menuCommands[commandId] = new MenuCommand(windowHandle, desktop.Id, desktop.DisplayName);
            }
        }

        if (GetMenuItemCount(popupMenu) <= 0)
        {
            _ = DestroyMenu(popupMenu);
            return;
        }

        _ = AppendMenu(systemMenu, MfSeparator, 0, null);
        if (!AppendMenu(systemMenu, MfPopup, (nuint)popupMenu, "Move to virtual desktop"))
        {
            _ = DestroyMenu(popupMenu);
            return;
        }

        _ = DrawMenuBar(windowHandle);
    }

    private void HandleMenuCommand(int commandId)
    {
        if (commandId <= 0)
        {
            return;
        }

        MenuCommand? command = null;
        lock (_sync)
        {
            if (_menuCommands.TryGetValue(commandId, out var resolvedCommand))
            {
                command = resolvedCommand;
            }
        }

        if (!command.HasValue)
        {
            return;
        }

        var targetWindow = command.Value.WindowHandle;
        var (moved, error) = _virtualDesktopService.MoveWindowToDesktop(targetWindow, command.Value.DesktopId);

        if (moved)
        {
            ClearWindowCommands(targetWindow);
            _ = GetSystemMenu(targetWindow, true);
            _statusSink?.Invoke($"Moved window via system menu to {command.Value.DesktopName}.");
            return;
        }

        _statusSink?.Invoke($"System-menu move failed: {error}");
    }

    private bool HasCommandsForWindow(nint windowHandle)
    {
        lock (_sync)
        {
            return _menuCommands.Any(pair => pair.Value.WindowHandle == windowHandle);
        }
    }

    private void ClearWindowCommands(nint windowHandle)
    {
        lock (_sync)
        {
            var commandIds = _menuCommands
                .Where(pair => pair.Value.WindowHandle == windowHandle)
                .Select(pair => pair.Key)
                .ToList();

            foreach (var commandId in commandIds)
            {
                _menuCommands.Remove(commandId);
            }
        }
    }

    private static nint ResolveTopLevelWindow(nint windowHandle)
    {
        var topLevelWindow = GetAncestor(windowHandle, GaRoot);
        return topLevelWindow == nint.Zero ? windowHandle : topLevelWindow;
    }

    private readonly record struct MenuCommand(nint WindowHandle, Guid DesktopId, string DesktopName);

    private delegate void WinEventDelegate(
        nint hookHandle,
        uint eventType,
        nint windowHandle,
        int objectId,
        int childId,
        uint eventThread,
        uint eventTime);

    [DllImport("user32.dll")]
    private static extern nint SetWinEventHook(
        uint eventMin,
        uint eventMax,
        nint moduleHandle,
        WinEventDelegate eventDelegate,
        uint processId,
        uint threadId,
        uint flags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWinEvent(nint hookHandle);

    [DllImport("user32.dll")]
    private static extern nint GetSystemMenu(nint windowHandle, [MarshalAs(UnmanagedType.Bool)] bool revert);

    [DllImport("user32.dll")]
    private static extern nint CreatePopupMenu();

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyMenu(nint menuHandle);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AppendMenu(nint menuHandle, uint flags, nuint itemIdOrSubMenu, string? itemText);

    [DllImport("user32.dll")]
    private static extern int GetMenuItemCount(nint menuHandle);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DrawMenuBar(nint windowHandle);

    [DllImport("user32.dll")]
    private static extern nint GetAncestor(nint windowHandle, uint flags);
}
