using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;

namespace VirtualDesktopUtils;

public partial class WindowTrayMenu : Window
{
    private readonly Action _onShow;
    private readonly Action _onRefresh;
    private readonly Action _onExit;
    private bool _closing;

    public WindowTrayMenu(Action onShow, Action onRefresh, Action onExit)
    {
        _onShow = onShow;
        _onRefresh = onRefresh;
        _onExit = onExit;
        InitializeComponent();
    }

    public void ShowAtCursor()
    {
        GetCursorPos(out var pt);

        // Position above and to the left of cursor (menu grows upward like Windows 11)
        Left = pt.X - 20;
        Top = pt.Y - ActualHeight - 10;

        Show();
        Activate();
    }

    protected override void OnContentRendered(EventArgs e)
    {
        base.OnContentRendered(e);
        // Reposition now that we know the actual size
        GetCursorPos(out var pt);
        Left = pt.X - 20;
        Top = pt.Y - ActualHeight + 8;
    }

    private void Show_Click(object sender, RoutedEventArgs e) { SafeClose(); _onShow(); }
    private void Refresh_Click(object sender, RoutedEventArgs e) { SafeClose(); _onRefresh(); }
    private void Exit_Click(object sender, RoutedEventArgs e) { SafeClose(); _onExit(); }

    private void Window_Deactivated(object? sender, EventArgs e) => SafeClose();

    private void Window_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == Key.Escape) SafeClose();
    }

    private void SafeClose()
    {
        if (_closing) return;
        _closing = true;
        Close();
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int X; public int Y; }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetCursorPos(out POINT lpPoint);
}
