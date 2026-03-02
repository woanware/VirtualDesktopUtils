using System.Runtime.InteropServices;
using System.Windows.Interop;

namespace VirtualDesktopUtils.Services;

internal sealed class TrayIconService : IDisposable
{
    private const int WmApp = 0x8000;
    private const int TrayCallbackMessage = WmApp + 1;
    private const int NimAdd = 0x00;
    private const int NimModify = 0x01;
    private const int NimDelete = 0x02;
    private const int NifMessage = 0x01;
    private const int NifIcon = 0x02;
    private const int NifTip = 0x04;
    private const int NifInfo = 0x10;
    private const int NiifInfo = 0x01;
    private const int WmLButtonDblClk = 0x0203;
    private const int WmRButtonUp = 0x0205;

    private readonly Action _onShow;
    private readonly Action _onRefresh;
    private readonly Action _onExit;
    private readonly nint _iconHandle;
    private HwndSource? _hwndSource;
    private WindowTrayMenu? _activeMenu;
    private bool _disposed;

    public TrayIconService(
        Action onShowRequested,
        Action onRefreshRequested,
        Action onExitRequested
    )
    {
        _onShow = onShowRequested;
        _onRefresh = onRefreshRequested;
        _onExit = onExitRequested;

        _iconHandle = nint.Zero;
        if (!string.IsNullOrWhiteSpace(Environment.ProcessPath))
        {
            ExtractIconEx(Environment.ProcessPath!, 0, out _iconHandle, out _, 1);
        }

        if (_iconHandle == nint.Zero)
        {
            _iconHandle = LoadIcon(nint.Zero, new nint(32512));
        }

        _hwndSource = new HwndSource(
            new HwndSourceParameters("VirtualDesktopUtilsTray")
            {
                Width = 0,
                Height = 0,
                WindowStyle = 0,
            }
        );
        _hwndSource.AddHook(WndProc);

        AddTrayIcon();
    }

    public void ShowBalloonTip(string title, string text, int timeoutMs = 3000)
    {
        if (_hwndSource is null)
            return;

        var nid = CreateNotifyIconData();
        nid.uFlags = NifInfo;
        nid.szInfoTitle = title;
        nid.szInfo = text;
        nid.dwInfoFlags = NiifInfo;
        nid.uTimeout = timeoutMs;
        Shell_NotifyIcon(NimModify, ref nid);
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;

        RemoveTrayIcon();
        _hwndSource?.RemoveHook(WndProc);
        _hwndSource?.Dispose();
        _hwndSource = null;
    }

    private void AddTrayIcon()
    {
        var nid = CreateNotifyIconData();
        nid.uFlags = NifMessage | NifIcon | NifTip;
        nid.uCallbackMessage = TrayCallbackMessage;
        nid.hIcon = _iconHandle;
        nid.szTip = "VirtualDesktopUtils";
        Shell_NotifyIcon(NimAdd, ref nid);
    }

    private void RemoveTrayIcon()
    {
        var nid = CreateNotifyIconData();
        Shell_NotifyIcon(NimDelete, ref nid);
    }

    private NOTIFYICONDATA CreateNotifyIconData()
    {
        var nid = new NOTIFYICONDATA();
        nid.cbSize = Marshal.SizeOf(nid);
        nid.hWnd = _hwndSource!.Handle;
        nid.uID = 1;
        return nid;
    }

    private nint WndProc(nint hwnd, int msg, nint wParam, nint lParam, ref bool handled)
    {
        if (msg == TrayCallbackMessage)
        {
            var eventId = lParam.ToInt32() & 0xFFFF;

            if (eventId == WmLButtonDblClk)
            {
                _onShow();
                handled = true;
            }
            else if (eventId == WmRButtonUp)
            {
                ShowTrayMenu();
                handled = true;
            }
        }

        return nint.Zero;
    }

    private void ShowTrayMenu()
    {
        if (_activeMenu is not null)
        {
            try
            {
                _activeMenu.Close();
            }
            catch { }
            _activeMenu = null;
        }

        _activeMenu = new WindowTrayMenu(_onShow, _onRefresh, _onExit);
        _activeMenu.Closed += (_, _) => _activeMenu = null;
        _activeMenu.ShowAtCursor();
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NOTIFYICONDATA
    {
        public int cbSize;
        public nint hWnd;
        public int uID;
        public int uFlags;
        public int uCallbackMessage;
        public nint hIcon;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string szTip;
        public int dwState;
        public int dwStateMask;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string szInfo;
        public int uTimeout;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
        public string szInfoTitle;
        public int dwInfoFlags;
        public Guid guidItem;
        public nint hBalloonIcon;
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool Shell_NotifyIcon(int dwMessage, ref NOTIFYICONDATA lpData);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int ExtractIconEx(
        string lpszFile,
        int nIconIndex,
        out nint phiconLarge,
        out nint phiconSmall,
        int nIcons
    );

    [DllImport("user32.dll")]
    private static extern nint LoadIcon(nint hInstance, nint lpIconName);
}
