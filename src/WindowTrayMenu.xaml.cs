using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Forms;

namespace VirtualDesktopUtils;

public partial class WindowTrayMenu : Window
{
    private readonly Action _onShow;
    private readonly Action _onRefresh;
    private readonly Action _onExit;
    private bool _closing;
    private bool _ready;
    private DateTime _shownAtUtc;

    public WindowTrayMenu(Action onShow, Action onRefresh, Action onExit)
    {
        _onShow = onShow;
        _onRefresh = onRefresh;
        _onExit = onExit;
        InitializeComponent();
    }

    public void ShowAtCursor()
    {
        _ready = false;
        _shownAtUtc = DateTime.UtcNow;

        GetCursorPos(out var pt);
        Left = 0;
        Top = 0;

        Show();
        UpdateLayout();

        var transformFromDevice = System.Windows.PresentationSource.FromVisual(this)?.CompositionTarget?.TransformFromDevice
            ?? System.Windows.Media.Matrix.Identity;
        var cursor = transformFromDevice.Transform(new System.Windows.Point(pt.X, pt.Y));
        var workingArea = Screen.FromPoint(new System.Drawing.Point(pt.X, pt.Y)).WorkingArea;
        var workingTopLeft = transformFromDevice.Transform(new System.Windows.Point(workingArea.Left, workingArea.Top));
        var workingBottomRight = transformFromDevice.Transform(new System.Windows.Point(workingArea.Right, workingArea.Bottom));
        var desiredLeft = cursor.X - ActualWidth + 8;
        var desiredTop = cursor.Y - ActualHeight - 8;
        var minLeft = workingTopLeft.X + 4d;
        var maxLeft = workingBottomRight.X - ActualWidth - 4d;
        var minTop = workingTopLeft.Y + 4d;
        var maxTop = workingBottomRight.Y - ActualHeight - 4d;

        Left = Math.Max(minLeft, Math.Min(desiredLeft, maxLeft));
        Top = Math.Max(minTop, Math.Min(desiredTop, maxTop));

        var hwnd = new System.Windows.Interop.WindowInteropHelper(this).Handle;
        if (hwnd != nint.Zero)
        {
            SetForegroundWindow(hwnd);
        }

        Activate();
        Focus();

        // Only enable deactivate-to-close after a short delay
        var timer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(200)
        };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            _ready = true;
        };
        timer.Start();
    }

    private void Show_Click(object sender, RoutedEventArgs e) { SafeClose(); _onShow(); }
    private void Refresh_Click(object sender, RoutedEventArgs e) { SafeClose(); _onRefresh(); }
    private void Exit_Click(object sender, RoutedEventArgs e) { SafeClose(); _onExit(); }

    private void Window_Deactivated(object? sender, EventArgs e)
    {
        if (!_ready)
        {
            return;
        }

        if ((DateTime.UtcNow - _shownAtUtc).TotalMilliseconds < 1200)
        {
            return;
        }

        if (IsMouseOver)
        {
            return;
        }

        SafeClose();
    }

    private void Window_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == Key.Escape) SafeClose();
    }

    private void SafeClose()
    {
        if (_closing) return;
        _closing = true;
        try { Close(); } catch { }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int X; public int Y; }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetCursorPos(out POINT lpPoint);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(nint hWnd);
}
