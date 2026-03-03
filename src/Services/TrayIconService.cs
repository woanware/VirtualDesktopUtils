using System.Drawing;
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
    private const int NimSetVersion = 0x04;
    private const int NifMessage = 0x01;
    private const int NifIcon = 0x02;
    private const int NifTip = 0x04;
    private const int NifInfo = 0x10;
    private const int NiifInfo = 0x01;
    private const int NotifyIconVersion4 = 4;
    private const int WmLButtonDblClk = 0x0203;
    private const int WmRButtonUp = 0x0205;

    private readonly HwndSource _hwndSource;
    private readonly Icon _appIcon;
    private readonly bool _ownsIcon;
    private readonly Action _onShow;
    private readonly Action _onRefresh;
    private readonly Action _onExit;
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

        var extractedIcon = !string.IsNullOrWhiteSpace(Environment.ProcessPath)
            ? Icon.ExtractAssociatedIcon(Environment.ProcessPath!)
            : null;
        _appIcon = extractedIcon ?? SystemIcons.Application;
        _ownsIcon = extractedIcon is not null;

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

    private void AddTrayIcon()
    {
        var notifyIconData = CreateNotifyIconData();
        notifyIconData.uFlags = NifMessage | NifIcon | NifTip;
        notifyIconData.uCallbackMessage = TrayCallbackMessage;
        notifyIconData.hIcon = _appIcon.Handle;
        notifyIconData.szTip = "VirtualDesktopUtils";

        if (!Shell_NotifyIcon(NimAdd, ref notifyIconData))
        {
            return;
        }

        notifyIconData.uVersion = NotifyIconVersion4;
        Shell_NotifyIcon(NimSetVersion, ref notifyIconData);
    }

    private void RemoveTrayIcon()
    {
        var notifyIconData = CreateNotifyIconData();
        _ = Shell_NotifyIcon(NimDelete, ref notifyIconData);
    }

    private NOTIFYICONDATA CreateNotifyIconData()
    {
        return new NOTIFYICONDATA
        {
            cbSize = Marshal.SizeOf<NOTIFYICONDATA>(),
            hWnd = _hwndSource.Handle,
            uID = 1,
            szTip = string.Empty,
            szInfo = string.Empty,
            szInfoTitle = string.Empty,
        };
    }

    private nint WndProc(nint hwnd, int msg, nint wParam, nint lParam, ref bool handled)
    {
        if (_disposed)
        {
            return nint.Zero;
        }

        if (msg != TrayCallbackMessage)
        {
            return nint.Zero;
        }

        var eventId = lParam.ToInt32() & 0xFFFF;
        switch (eventId)
        {
            case WmLButtonDblClk:
                _onShow();
                handled = true;
                break;
            case WmRButtonUp:
                ShowTrayMenu();
                handled = true;
                break;
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
            catch (ObjectDisposedException)
            {
            }
            catch (InvalidOperationException)
            {
            }
            _activeMenu = null;
        }

        _activeMenu = new WindowTrayMenu(_onShow, _onRefresh, _onExit);
        _activeMenu.Closed += (_, _) => _activeMenu = null;
        _activeMenu.ShowAtCursor();
    }

    public void ShowBalloonTip(string title, string text, int timeoutMs = 3000)
    {
        if (_disposed)
        {
            return;
        }

        var notifyIconData = CreateNotifyIconData();
        notifyIconData.uFlags = NifInfo;
        notifyIconData.szInfoTitle = title;
        notifyIconData.szInfo = text;
        notifyIconData.dwInfoFlags = NiifInfo;
        notifyIconData.uVersion = (uint)timeoutMs;

        _ = Shell_NotifyIcon(NimModify, ref notifyIconData);
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }
        _disposed = true;

        RemoveTrayIcon();
        _hwndSource.RemoveHook(WndProc);
        _hwndSource.Dispose();

        if (_activeMenu is not null)
        {
            try
            {
                _activeMenu.Close();
            }
            catch (ObjectDisposedException)
            {
            }
            catch (InvalidOperationException)
            {
            }
            _activeMenu = null;
        }

        if (_ownsIcon)
        {
            _appIcon.Dispose();
        }
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

        public uint uVersion;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
        public string szInfoTitle;

        public int dwInfoFlags;
        public Guid guidItem;
        public nint hBalloonIcon;
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool Shell_NotifyIcon(int dwMessage, ref NOTIFYICONDATA lpData);
}
