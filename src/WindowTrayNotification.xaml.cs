using System.Windows;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace VirtualDesktopUtils;

public partial class WindowTrayNotification : Window
{
    private static WindowTrayNotification? _activeNotification;
    private readonly DispatcherTimer _dismissTimer = new() { Interval = TimeSpan.FromSeconds(5) };
    private bool _isClosing;

    private WindowTrayNotification(string title, string message)
    {
        InitializeComponent();

        TitleTextBlock.Text = title;
        MessageTextBlock.Text = message;
        Opacity = 0;

        Loaded += WindowTrayNotification_OnLoaded;
        _dismissTimer.Tick += DismissTimer_OnTick;
    }

    public static void ShowNotification(string title, string message)
    {
        _activeNotification?.Close();

        var notification = new WindowTrayNotification(title, message);
        _activeNotification = notification;
        notification.Closed += (_, _) =>
        {
            if (ReferenceEquals(_activeNotification, notification))
            {
                _activeNotification = null;
            }
        };
        notification.Show();
    }

    protected override void OnClosed(EventArgs e)
    {
        _dismissTimer.Stop();
        base.OnClosed(e);
    }

    private void WindowTrayNotification_OnLoaded(object sender, RoutedEventArgs e)
    {
        PositionNearTray();
        BeginAnimation(OpacityProperty, new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(160)));
        _dismissTimer.Start();
    }

    private void DismissTimer_OnTick(object? sender, EventArgs e)
    {
        _dismissTimer.Stop();
        BeginCloseAnimation();
    }

    private void CloseButton_OnClick(object sender, RoutedEventArgs e)
    {
        _dismissTimer.Stop();
        BeginCloseAnimation();
    }

    private void BeginCloseAnimation()
    {
        if (_isClosing)
        {
            return;
        }

        _isClosing = true;
        var animation = new DoubleAnimation(0, TimeSpan.FromMilliseconds(140));
        animation.Completed += (_, _) => Close();
        BeginAnimation(OpacityProperty, animation);
    }

    private void PositionNearTray()
    {
        UpdateLayout();

        var workArea = SystemParameters.WorkArea;
        Left = Math.Max(workArea.Left, workArea.Right - ActualWidth - 16);
        Top = Math.Max(workArea.Top, workArea.Bottom - ActualHeight - 16);
    }
}
