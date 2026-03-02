using System.Runtime.InteropServices;
using System.Text;

namespace VirtualDesktopUtils.Services;

internal sealed class WindowDiscoveryService
{
    private const int DwmWindowAttributeCloaked = 14;
    private const int GwlExStyle = -20;
    private const uint WsExToolWindow = 0x00000080;
    private readonly VirtualDesktopService _virtualDesktopService;

    public WindowDiscoveryService(VirtualDesktopService virtualDesktopService)
    {
        _virtualDesktopService = virtualDesktopService;
    }

    public bool IsCandidateWindow(nint windowHandle)
    {
        if (windowHandle == nint.Zero || windowHandle == GetShellWindow())
        {
            return false;
        }

        if (!IsWindowVisible(windowHandle))
        {
            return false;
        }

        if (IsWindowCloaked(windowHandle))
        {
            return false;
        }

        var exStyle = (ulong)GetWindowLongPtr(windowHandle, GwlExStyle).ToInt64();
        if ((exStyle & WsExToolWindow) == WsExToolWindow)
        {
            return false;
        }

        return !string.IsNullOrWhiteSpace(GetWindowTitle(windowHandle));
    }

    /// <summary>
    /// Finds the topmost visible candidate window in Z-order that is on the current
    /// virtual desktop and does not belong to our own process.
    /// </summary>
    public nint FindTopmostCandidateWindowExcludingSelf()
    {
        var ownProcessId = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
        nint result = nint.Zero;

        EnumWindows((hwnd, _) =>
        {
            if (!IsCandidateWindow(hwnd))
            {
                return true;
            }

            GetWindowThreadProcessId(hwnd, out var windowProcessId);
            if ((uint)windowProcessId == ownProcessId)
            {
                return true;
            }

            if (!_virtualDesktopService.IsWindowOnCurrentDesktop(hwnd))
            {
                return true;
            }

            result = hwnd;
            return false; // stop enumeration — first match is topmost in Z-order
        }, nint.Zero);

        return result;
    }

    private static bool IsWindowCloaked(nint windowHandle)
    {
        if (DwmGetWindowAttribute(windowHandle, DwmWindowAttributeCloaked, out int cloaked, Marshal.SizeOf<int>()) != 0)
        {
            return false;
        }

        return cloaked != 0;
    }

    private static string GetWindowTitle(nint windowHandle)
    {
        var textLength = GetWindowTextLength(windowHandle);
        if (textLength <= 0)
        {
            return string.Empty;
        }

        var builder = new StringBuilder(textLength + 1);
        _ = GetWindowText(windowHandle, builder, builder.Capacity);
        return builder.ToString().Trim();
    }

    private static nint GetWindowLongPtr(nint windowHandle, int index)
    {
        if (nint.Size == 8)
        {
            return GetWindowLongPtr64(windowHandle, index);
        }

        return GetWindowLongPtr32(windowHandle, index);
    }

    private delegate bool EnumWindowsProc(nint windowHandle, nint lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(EnumWindowsProc enumProc, nint lParam);

    [DllImport("user32.dll")]
    private static extern nint GetShellWindow();

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(nint windowHandle);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(nint windowHandle, StringBuilder text, int maxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLength(nint windowHandle);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
    private static extern nint GetWindowLongPtr64(nint windowHandle, int index);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongW")]
    private static extern nint GetWindowLongPtr32(nint windowHandle, int index);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint hWnd, out int processId);

    [DllImport("dwmapi.dll")]
    private static extern int DwmGetWindowAttribute(nint windowHandle, int attribute, out int value, int valueSize);
}
