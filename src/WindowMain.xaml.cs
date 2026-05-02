using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using VirtualDesktopUtils.Models;
using VirtualDesktopUtils.Services;

namespace VirtualDesktopUtils;

public partial class WindowMain : Window
{
    private const int DwmwaUseImmersiveDarkMode = 20;
    private const int DwmwaUseImmersiveDarkModeLegacy = 19;
    private const int DwmwaCaptionColor = 35;
    private const int DwmwaTextColor = 36;
    private const int CaptionColorBgr = unchecked((int)0x002C2C2C);
    private const int CaptionTextColorBgr = unchecked((int)0x00F6EDE8);

    private readonly RuntimeConfigService _runtimeConfigService;
    private readonly VirtualDesktopService _virtualDesktopService;
    private readonly AppUpdateService _appUpdateService = new();
    private readonly GlobalHotkeyService _globalHotkeyService = new();
    private readonly ObservableCollection<RefreshIntervalOption> _refreshIntervals = new();
    private readonly System.Windows.Threading.DispatcherTimer _autoRefreshTimer;
    private static readonly TimeSpan StartupUpdateCheckInterval = TimeSpan.FromHours(6);
    private bool _suppressStartWithWindowsEvents;
    private bool _suppressAppUpdateOptionEvents;
    private bool _runtimeServicesStarted;

    // Captured hotkey state
    private uint _pickerModifiers;
    private uint _pickerVk;
    private uint _moveModifiers;

    internal WindowMain(RuntimeConfigService runtimeConfigService, string? startupStatusMessage)
    {
        _runtimeConfigService = runtimeConfigService;
        _virtualDesktopService = new VirtualDesktopService(_runtimeConfigService);

        InitializeComponent();

        _autoRefreshTimer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(3)
        };
        _autoRefreshTimer.Tick += (_, _) => AutoRefresh();

        RefreshIntervalComboBox.ItemsSource = _refreshIntervals;

        _refreshIntervals.Add(new RefreshIntervalOption("1 second", 1));
        _refreshIntervals.Add(new RefreshIntervalOption("2 seconds", 2));
        _refreshIntervals.Add(new RefreshIntervalOption("3 seconds", 3));
        _refreshIntervals.Add(new RefreshIntervalOption("5 seconds", 5));
        _refreshIntervals.Add(new RefreshIntervalOption("10 seconds", 10));
        RefreshIntervalComboBox.SelectedItem = _refreshIntervals.First(option => option.Seconds == 3);
        AutoRefreshCheckBox.IsChecked = true;
        _suppressStartWithWindowsEvents = true;
        StartWithWindowsCheckBox.IsChecked = _runtimeConfigService.IsStartWithWindowsEnabled();
        _suppressStartWithWindowsEvents = false;
        _suppressAppUpdateOptionEvents = true;
        AutoCheckAppUpdatesCheckBox.IsChecked = _runtimeConfigService.IsAppUpdateCheckOnStartupEnabled();
        _suppressAppUpdateOptionEvents = false;
        GuidAutoUpdateCheckBox.IsChecked = _runtimeConfigService.IsGuidAutoUpdateOnStartupEnabled();
        VersionTextBlock.Text = $"Version {GetDisplayVersion()}";
        AppUpdateStatusTextBlock.Text = BuildLastCheckedLabel(_runtimeConfigService.GetLastAppUpdateCheckUtc());

        LoadSavedHotkeys();

        Loaded += WindowMain_OnLoaded;
        SourceInitialized += WindowMain_OnSourceInitialized;
    }

    private void LoadSavedHotkeys()
    {
        var pickerConfig = _runtimeConfigService.LoadPickerHotkeyConfiguration();
        _pickerModifiers = pickerConfig.Modifiers;
        _pickerVk = pickerConfig.Vk;
        PickerHotkeyTextBox.Text = pickerConfig.DisplayText;

        var moveConfig = _runtimeConfigService.LoadMoveHotkeyConfiguration();
        _moveModifiers = moveConfig.Modifiers != 0
            ? moveConfig.Modifiers
            : GlobalHotkeyService.ModControl | GlobalHotkeyService.ModAlt;
        MoveHotkeyTextBox.Text = !string.IsNullOrWhiteSpace(moveConfig.DisplayText)
            ? moveConfig.DisplayText
            : "Ctrl+Alt";
    }

    private void AutoRefresh()
    {
        if (AutoRefreshCheckBox.IsChecked != true)
        {
            return;
        }

        RefreshRuntimeState();
    }

    public void RefreshData()
    {
        RefreshRuntimeState();
    }

    public void ShowFromTray()
    {
        if (!IsVisible)
        {
            Show();
        }

        if (WindowState == WindowState.Minimized)
        {
            WindowState = WindowState.Normal;
        }

        Activate();
        Topmost = true;
        Topmost = false;
        Focus();
    }

    public void StartHidden()
    {
        _ = new WindowInteropHelper(this).EnsureHandle();
        StartRuntimeServices();
    }

    public void PrepareForExit()
    {
        _autoRefreshTimer.Stop();
        _globalHotkeyService.Dispose();
    }

    public void UpdateStatus(string message)
    {
    }

    private void ApplyAutoRefreshSettings()
    {
        _autoRefreshTimer.Interval = TimeSpan.FromSeconds(GetSelectedRefreshSeconds());

        if (!_runtimeServicesStarted)
        {
            return;
        }

        if (AutoRefreshCheckBox.IsChecked == true)
        {
            if (!_autoRefreshTimer.IsEnabled)
            {
                _autoRefreshTimer.Start();
            }
        }
        else if (_autoRefreshTimer.IsEnabled)
        {
            _autoRefreshTimer.Stop();
        }
    }

    private int GetSelectedRefreshSeconds()
    {
        return RefreshIntervalComboBox.SelectedItem is RefreshIntervalOption option
            ? option.Seconds
            : 3;
    }

    private void WindowMain_OnLoaded(object sender, RoutedEventArgs e)
    {
        StartRuntimeServices();
    }

    private void StartRuntimeServices()
    {
        if (_runtimeServicesStarted)
        {
            return;
        }

        _runtimeServicesStarted = true;
        RefreshRuntimeState();

        ApplyAutoRefreshSettings();

        if (ShouldRunStartupUpdateCheck())
        {
            _ = CheckForAppUpdatesAsync(userInitiated: false);
        }
    }

    private void WindowMain_OnSourceInitialized(object? sender, EventArgs e)
    {
        ApplyNativeTitleBarTheme();
        RegisterPickerHotkey();
        RegisterMoveHotkeys();
    }

    private void AutoRefreshCheckBox_OnChecked(object sender, RoutedEventArgs e) => ApplyAutoRefreshSettings();
    private void AutoRefreshCheckBox_OnUnchecked(object sender, RoutedEventArgs e) => ApplyAutoRefreshSettings();

    private void StartWithWindowsCheckBox_OnChecked(object sender, RoutedEventArgs e) =>
        SetStartWithWindowsOption(true);
    private void StartWithWindowsCheckBox_OnUnchecked(object sender, RoutedEventArgs e) =>
        SetStartWithWindowsOption(false);

    private void AutoCheckAppUpdatesCheckBox_OnChecked(object sender, RoutedEventArgs e) =>
        SetAutoCheckAppUpdatesOption(true);
    private void AutoCheckAppUpdatesCheckBox_OnUnchecked(object sender, RoutedEventArgs e) =>
        SetAutoCheckAppUpdatesOption(false);

    private async void CheckForUpdatesButton_OnClick(object sender, RoutedEventArgs e)
    {
        await CheckForAppUpdatesAsync(userInitiated: true);
    }

    private void SetAutoCheckAppUpdatesOption(bool enabled)
    {
        if (_suppressAppUpdateOptionEvents)
        {
            return;
        }

        _runtimeConfigService.SetAppUpdateCheckOnStartupEnabled(enabled);
    }

    private async Task CheckForAppUpdatesAsync(bool userInitiated)
    {
        CheckForUpdatesButton.IsEnabled = false;
        try
        {
            AppUpdateStatusTextBlock.Text = "Checking for app updates…";

            var result = await _appUpdateService.CheckForUpdatesAsync();
            var checkedUtc = DateTime.UtcNow.ToString("O");
            _runtimeConfigService.SetLastAppUpdateCheckUtc(checkedUtc);

            if (result.Status == AppUpdateCheckStatus.Error)
            {
                AppUpdateStatusTextBlock.Text = result.Message;
                if (userInitiated)
                {
                    System.Windows.MessageBox.Show(
                        result.Message,
                        "VirtualDesktopUtils update check",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
            else
            {
                AppUpdateStatusTextBlock.Text = BuildLastCheckedLabel(checkedUtc);
            }

            if (result.Status == AppUpdateCheckStatus.UpdateAvailable && result.LatestRelease is not null)
            {
                var latestRelease = result.LatestRelease;
                var promptResult = System.Windows.MessageBox.Show(
                    $"A new version is available.\n\nCurrent: v{FormatVersionForDisplay(result.CurrentVersion)}\nLatest: v{latestRelease.VersionText}\n\nOpen the release page now?",
                    "VirtualDesktopUtils update available",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Information);

                if (promptResult == MessageBoxResult.Yes &&
                    !TryOpenUrl(latestRelease.ReleaseUrl, out var openError))
                {
                    System.Windows.MessageBox.Show(
                        $"Unable to open release page: {openError}",
                        "VirtualDesktopUtils",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
            else if (userInitiated && result.Status == AppUpdateCheckStatus.UpToDate)
            {
                System.Windows.MessageBox.Show(
                    result.Message,
                    "VirtualDesktopUtils update check",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
        }
        finally
        {
            CheckForUpdatesButton.IsEnabled = true;
        }
    }

    private bool ShouldRunStartupUpdateCheck()
    {
        if (!_runtimeConfigService.IsAppUpdateCheckOnStartupEnabled())
        {
            return false;
        }

        var lastCheckUtc = _runtimeConfigService.GetLastAppUpdateCheckUtc();
        if (string.IsNullOrWhiteSpace(lastCheckUtc))
        {
            return true;
        }

        if (!DateTime.TryParse(lastCheckUtc, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var parsedLastCheck))
        {
            return true;
        }

        return DateTime.UtcNow - parsedLastCheck >= StartupUpdateCheckInterval;
    }

    private static string BuildLastCheckedLabel(string lastCheckedUtc)
    {
        if (!DateTime.TryParse(lastCheckedUtc, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var parsedUtc))
        {
            return "Last update check: never";
        }

        return $"Last update check: {parsedUtc.ToLocalTime():g}";
    }

    private void SetStartWithWindowsOption(bool enabled)
    {
        if (_suppressStartWithWindowsEvents)
        {
            return;
        }

        var (success, message) = _runtimeConfigService.SetStartWithWindowsEnabled(enabled);
        if (success)
        {
            return;
        }

        _suppressStartWithWindowsEvents = true;
        StartWithWindowsCheckBox.IsChecked = !enabled;
        _suppressStartWithWindowsEvents = false;
        System.Windows.MessageBox.Show(
            message,
            "VirtualDesktopUtils",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
    }

    private void GuidAutoUpdateCheckBox_OnChecked(object sender, RoutedEventArgs e) =>
        _runtimeConfigService.SetGuidAutoUpdateOnStartupEnabled(true);
    private void GuidAutoUpdateCheckBox_OnUnchecked(object sender, RoutedEventArgs e) =>
        _runtimeConfigService.SetGuidAutoUpdateOnStartupEnabled(false);

    private async void UpdateGuidConfigButton_OnClick(object sender, RoutedEventArgs e)
    {
        UpdateGuidConfigButton.IsEnabled = false;
        await _runtimeConfigService.SyncGuidConfigFromUpstreamAsync();
        UpdateGuidConfigButton.IsEnabled = true;
    }

    private void RefreshIntervalComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e) =>
        ApplyAutoRefreshSettings();

    // --- Hotkey capture ---

    private void HotkeyTextBox_GotFocus(object sender, RoutedEventArgs e)
    {
        SuspendHotkeysForCapture();

        if (sender is System.Windows.Controls.TextBox tb)
        {
            tb.Background = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(0x50, 0x40, 0x20));
            tb.Text = "Press a key combo…";
        }
    }

    private void HotkeyTextBox_LostFocus(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.TextBox tb)
        {
            tb.Background = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(0x3A, 0x3A, 0x3A));

            // Restore saved text if user didn't press anything
            if (tb.Text == "Press a key combo…")
            {
                if (tb == PickerHotkeyTextBox)
                    tb.Text = _runtimeConfigService.LoadPickerHotkeyConfiguration().DisplayText;
                else if (tb == MoveHotkeyTextBox)
                    tb.Text = _runtimeConfigService.LoadMoveHotkeyConfiguration().DisplayText;
            }
        }

        ResumeHotkeysAfterCapture();
    }

    private void PickerHotkeyTextBox_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        e.Handled = true;
        var (modifiers, vk, display) = CaptureKeyCombo(e);
        if (modifiers == 0 || vk == 0) return;

        _pickerModifiers = modifiers;
        _pickerVk = vk;
        PickerHotkeyTextBox.Text = display;
        RegisterPickerHotkey();

        // Move focus away so the field doesn't re-capture
        Keyboard.ClearFocus();
        FocusManager.SetFocusedElement(this, this);
    }

    private void MoveHotkeyTextBox_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        e.Handled = true;

        var key = ResolveEventKey(e);
        var modifiers = CaptureModifiers(key);
        if (modifiers == 0) return;

        // For direct move, we only need the modifier (keys 1-9 are appended automatically)
        _moveModifiers = modifiers;
        MoveHotkeyTextBox.Text = ModifiersToDisplay(modifiers);
        RegisterMoveHotkeys();

        // Keep focus while user is still holding/adding modifier keys.
        if (!IsModifierKey(key))
        {
            Keyboard.ClearFocus();
            FocusManager.SetFocusedElement(this, this);
        }
    }

    private static (uint Modifiers, uint Vk, string Display) CaptureKeyCombo(System.Windows.Input.KeyEventArgs e)
    {
        var key = ResolveEventKey(e);

        // Ignore bare modifier presses
        if (IsModifierKey(key))
        {
            return (0, 0, "");
        }

        var modifiers = CaptureModifiers(key);
        if (modifiers == 0) return (0, 0, "");

        var vk = (uint)KeyInterop.VirtualKeyFromKey(key);
        var display = $"{ModifiersToDisplay(modifiers)}+{KeyToDisplayName(key)}";
        return (modifiers, vk, display);
    }

    private static Key ResolveEventKey(System.Windows.Input.KeyEventArgs e) =>
        e.Key == Key.System ? e.SystemKey : e.Key;

    private static bool IsModifierKey(Key key) =>
        key is Key.LeftCtrl or Key.RightCtrl
            or Key.LeftAlt or Key.RightAlt
            or Key.LeftShift or Key.RightShift
            or Key.LWin or Key.RWin;

    private static uint CaptureModifiers(Key key)
    {
        var modifiers = 0u;

        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control) || key is Key.LeftCtrl or Key.RightCtrl)
            modifiers |= GlobalHotkeyService.ModControl;
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Alt) || key is Key.LeftAlt or Key.RightAlt)
            modifiers |= GlobalHotkeyService.ModAlt;
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) || key is Key.LeftShift or Key.RightShift)
            modifiers |= GlobalHotkeyService.ModShift;
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Windows) || key is Key.LWin or Key.RWin)
            modifiers |= GlobalHotkeyService.ModWin;

        return modifiers;
    }

    private static string ModifiersToDisplay(uint modifiers)
    {
        var parts = new List<string>();
        if ((modifiers & GlobalHotkeyService.ModControl) != 0) parts.Add("Ctrl");
        if ((modifiers & GlobalHotkeyService.ModAlt) != 0) parts.Add("Alt");
        if ((modifiers & GlobalHotkeyService.ModShift) != 0) parts.Add("Shift");
        if ((modifiers & GlobalHotkeyService.ModWin) != 0) parts.Add("Win");
        return string.Join("+", parts);
    }

    private void SuspendHotkeysForCapture()
    {
        _globalHotkeyService.UnregisterPickerHotkey();
        _globalHotkeyService.UnregisterMoveHotkeys();
    }

    private void ResumeHotkeysAfterCapture()
    {
        if (!IsLoaded)
        {
            return;
        }

        RegisterPickerHotkey();
        RegisterMoveHotkeys();
    }

    private static string KeyToDisplayName(Key key) => key switch
    {
        Key.Space => "Space",
        Key.OemTilde => "~",
        Key.OemMinus => "-",
        Key.OemPlus => "=",
        Key.OemOpenBrackets => "[",
        Key.OemCloseBrackets => "]",
        Key.OemBackslash or Key.Oem5 => "\\",
        Key.OemSemicolon or Key.Oem1 => ";",
        Key.OemQuotes or Key.Oem7 => "'",
        Key.OemComma => ",",
        Key.OemPeriod => ".",
        Key.Oem2 => "/",
        _ => key.ToString()
    };

    // --- Hotkey registration ---

    private void RegisterPickerHotkey()
    {
        if (_pickerModifiers == 0 || _pickerVk == 0) return;
        var windowHandle = new WindowInteropHelper(this).Handle;
        if (windowHandle == nint.Zero) return;

        var display = PickerHotkeyTextBox.Text;
        var registered = _globalHotkeyService.RegisterPickerHotkey(windowHandle, _pickerModifiers, _pickerVk, OnPickerHotkeyTriggered);
        _runtimeConfigService.SavePickerHotkeyConfiguration(_pickerModifiers, _pickerVk, display);

        PickerHotkeyStatusTextBlock.Text = registered ? $"✓ {display}" : $"✗ {display} in use";
    }

    private void RegisterMoveHotkeys()
    {
        if (_moveModifiers == 0) return;
        var windowHandle = new WindowInteropHelper(this).Handle;
        if (windowHandle == nint.Zero) return;

        var display = MoveHotkeyTextBox.Text;
        var count = _globalHotkeyService.RegisterMoveHotkeys(windowHandle, _moveModifiers, OnMoveHotkeyTriggered);
        _runtimeConfigService.SaveMoveHotkeyConfiguration(_moveModifiers, display);

        MoveHotkeyStatusTextBlock.Text = count > 0
            ? $"✓ {display}+1‒9 ({count})"
            : $"✗ {display}+N in use";
    }

    private void OnMoveHotkeyTriggered(int desktopNumber)
    {
        var targetWindow = GlobalHotkeyService.GetForegroundWindow();
        var ownHandle = new WindowInteropHelper(this).Handle;
        if (targetWindow == ownHandle || targetWindow == nint.Zero) return;

        var desktops = _virtualDesktopService.GetDesktops();
        if (desktopNumber < 1 || desktopNumber > desktops.Count) return;

        var targetDesktop = desktops[desktopNumber - 1];
        if (targetDesktop.IsCurrent) return;

        var (moved, _) = _virtualDesktopService.MoveWindowToDesktop(targetWindow, targetDesktop.Id);
        if (moved)
        {
            _virtualDesktopService.SwitchToDesktop(targetDesktop.Id);
        }
    }

    private WindowPopUp? _activePicker;

    private void OnPickerHotkeyTriggered()
    {
        if (_activePicker is not null) return;

        var targetWindow = GlobalHotkeyService.GetForegroundWindow();
        var ownHandle = new WindowInteropHelper(this).Handle;
        if (targetWindow == ownHandle || targetWindow == nint.Zero) return;

        _activePicker = new WindowPopUp(_virtualDesktopService, targetWindow);
        _activePicker.Closed += (_, _) => _activePicker = null;
        _activePicker.Show();
    }

    private void RefreshRuntimeState()
    {
        _ = _virtualDesktopService.GetDesktops();
    }

    private static string GetDisplayVersion()
    {
        var processPath = Environment.ProcessPath;
        if (!string.IsNullOrWhiteSpace(processPath))
        {
            try
            {
                var fileVersion = FileVersionInfo.GetVersionInfo(processPath).FileVersion;
                if (!string.IsNullOrWhiteSpace(fileVersion))
                {
                    return fileVersion;
                }
            }
            catch (ArgumentException)
            {
            }
            catch (System.IO.IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        var assemblyVersion = typeof(WindowMain).Assembly.GetName().Version;
        return assemblyVersion?.ToString() ?? "unknown";
    }

    private static string FormatVersionForDisplay(Version version)
    {
        return version.Revision <= 0
            ? $"{version.Major}.{version.Minor}.{Math.Max(version.Build, 0)}"
            : version.ToString();
    }

    private static bool TryOpenUrl(string url, out string error)
    {
        try
        {
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            error = string.Empty;
            return true;
        }
        catch (Win32Exception ex)
        {
            error = ex.Message;
            return false;
        }
        catch (ObjectDisposedException ex)
        {
            error = ex.Message;
            return false;
        }
        catch (InvalidOperationException ex)
        {
            error = ex.Message;
            return false;
        }
    }

    private void ApplyNativeTitleBarTheme()
    {
        var windowHandle = new WindowInteropHelper(this).Handle;
        if (windowHandle == nint.Zero) return;

        var enabled = 1;
        _ = DwmSetWindowAttribute(windowHandle, DwmwaUseImmersiveDarkMode, ref enabled, sizeof(int));
        _ = DwmSetWindowAttribute(windowHandle, DwmwaUseImmersiveDarkModeLegacy, ref enabled, sizeof(int));

        var captionColor = CaptionColorBgr;
        _ = DwmSetWindowAttribute(windowHandle, DwmwaCaptionColor, ref captionColor, sizeof(int));

        var textColor = CaptionTextColorBgr;
        _ = DwmSetWindowAttribute(windowHandle, DwmwaTextColor, ref textColor, sizeof(int));
    }

    private sealed record RefreshIntervalOption(string Label, int Seconds);

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(
        nint hwnd,
        int dwAttribute,
        ref int pvAttribute,
        int cbAttribute);
}
