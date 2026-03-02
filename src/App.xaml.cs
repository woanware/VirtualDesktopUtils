using System.ComponentModel;
using System.Threading;
using System.Windows;
using VirtualDesktopUtils.Services;

namespace VirtualDesktopUtils;

public partial class App : System.Windows.Application
{
    private static Mutex? _singleInstanceMutex;
    private WindowMain? _mainWindow;
    private TrayIconService? _trayIconService;
    private RuntimeConfigService? _runtimeConfigService;
    private bool _isExiting;
    private bool _hasShownTrayNotification;

    protected override async void OnStartup(StartupEventArgs e)
    {
        DispatcherUnhandledException += (_, args) =>
        {
            args.Handled = true;
        };

        _singleInstanceMutex = new Mutex(true, "VirtualDesktopUtils_SingleInstance", out var createdNew);
        if (!createdNew)
        {
            System.Windows.MessageBox.Show("VirtualDesktopUtils is already running.", "VirtualDesktopUtils",
                MessageBoxButton.OK, MessageBoxImage.Information);
            Shutdown();
            return;
        }

        base.OnStartup(e);
        ShutdownMode = ShutdownMode.OnExplicitShutdown;

        _runtimeConfigService = new RuntimeConfigService();
        string? startupStatus = null;

        if (_runtimeConfigService.IsGuidAutoUpdateOnStartupEnabled())
        {
            var (success, message) = await _runtimeConfigService.SyncGuidConfigFromUpstreamAsync();
            startupStatus = success
                ? message
                : $"GUID auto-update on startup failed: {message}";
        }

        _mainWindow = new WindowMain(_runtimeConfigService, startupStatus);
        _mainWindow.Closing += MainWindowOnClosing;
        _mainWindow.Show();

        _trayIconService = new TrayIconService(
            onShowRequested: ShowMainWindow,
            onRefreshRequested: () => _mainWindow?.RefreshData(),
            onExitRequested: ExitApplication);
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _trayIconService?.Dispose();
        _trayIconService = null;
        _runtimeConfigService = null;
        base.OnExit(e);
    }

    private void MainWindowOnClosing(object? sender, CancelEventArgs e)
    {
        if (_isExiting || _mainWindow is null)
        {
            return;
        }

        e.Cancel = true;
        _mainWindow.Hide();
        if (!_hasShownTrayNotification)
        {
            _hasShownTrayNotification = true;
            _trayIconService?.ShowBalloonTip("VirtualDesktopUtils", "Still running — double-click the tray icon to reopen.");
        }
    }

    private void ShowMainWindow()
    {
        _mainWindow?.ShowFromTray();
    }

    private void ExitApplication()
    {
        _isExiting = true;

        try
        {
            _mainWindow?.PrepareForExit();
        }
        catch
        {
            // Don't let cleanup failures block exit
        }

        try
        {
            _trayIconService?.Dispose();
            _trayIconService = null;
        }
        catch
        {
        }

        if (_mainWindow is not null)
        {
            _mainWindow.Close();
            _mainWindow = null;
        }

        Environment.Exit(0);
    }
}

