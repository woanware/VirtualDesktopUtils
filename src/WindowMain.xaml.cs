using System.Collections.ObjectModel;
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
    private readonly WindowDiscoveryService _windowDiscoveryService;
    private readonly WindowMenuInjectionService _windowMenuInjectionService;
    private readonly GlobalHotkeyService _globalHotkeyService = new();
    private readonly ObservableCollection<RefreshIntervalOption> _refreshIntervals = new();
    private readonly System.Windows.Threading.DispatcherTimer _autoRefreshTimer;
    private string _lastDesktopSignature = string.Empty;

    // Captured hotkey state
    private uint _pickerModifiers;
    private uint _pickerVk;
    private uint _moveModifiers;

    internal WindowMain(RuntimeConfigService runtimeConfigService, string? startupStatusMessage)
    {
        _runtimeConfigService = runtimeConfigService;
        _virtualDesktopService = new VirtualDesktopService(_runtimeConfigService);
        _windowDiscoveryService = new WindowDiscoveryService(_virtualDesktopService);
        _windowMenuInjectionService = new WindowMenuInjectionService(_virtualDesktopService, _windowDiscoveryService);

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
        GuidAutoUpdateCheckBox.IsChecked = _runtimeConfigService.IsGuidAutoUpdateOnStartupEnabled();
        ContextMenuCheckBox.IsChecked = _runtimeConfigService.IsContextMenuEnabled();

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

    public void PrepareForExit()
    {
        _autoRefreshTimer.Stop();
        _globalHotkeyService.Dispose();
        _windowMenuInjectionService.Stop();
    }

    public void UpdateStatus(string message)
    {
    }

    private void ApplyAutoRefreshSettings()
    {
        _autoRefreshTimer.Interval = TimeSpan.FromSeconds(GetSelectedRefreshSeconds());

        if (!IsLoaded)
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
        RefreshRuntimeState();

        if (ContextMenuCheckBox.IsChecked == true)
        {
            _windowMenuInjectionService.Start(UpdateStatus);
        }

        ApplyAutoRefreshSettings();
    }

    private void WindowMain_OnSourceInitialized(object? sender, EventArgs e)
    {
        ApplyNativeTitleBarTheme();
        RegisterPickerHotkey();
        RegisterMoveHotkeys();
    }

    private void AutoRefreshCheckBox_OnChecked(object sender, RoutedEventArgs e) => ApplyAutoRefreshSettings();
    private void AutoRefreshCheckBox_OnUnchecked(object sender, RoutedEventArgs e) => ApplyAutoRefreshSettings();

    private void ContextMenuCheckBox_OnChecked(object sender, RoutedEventArgs e)
    {
        _runtimeConfigService.SetContextMenuEnabled(true);
        if (IsLoaded) _windowMenuInjectionService.Start(UpdateStatus);
    }

    private void ContextMenuCheckBox_OnUnchecked(object sender, RoutedEventArgs e)
    {
        _runtimeConfigService.SetContextMenuEnabled(false);
        if (IsLoaded) _windowMenuInjectionService.Stop();
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
        var (modifiers, _, display) = CaptureKeyCombo(e);
        if (modifiers == 0) return;

        // For direct move, we only need the modifier (keys 1-9 are appended automatically)
        _moveModifiers = modifiers;
        MoveHotkeyTextBox.Text = display.Contains('+') ? display[..display.LastIndexOf('+')] : display;
        RegisterMoveHotkeys();

        Keyboard.ClearFocus();
        FocusManager.SetFocusedElement(this, this);
    }

    private static (uint Modifiers, uint Vk, string Display) CaptureKeyCombo(System.Windows.Input.KeyEventArgs e)
    {
        var key = e.Key == Key.System ? e.SystemKey : e.Key;

        // Ignore bare modifier presses
        if (key is Key.LeftCtrl or Key.RightCtrl or Key.LeftAlt or Key.RightAlt
            or Key.LeftShift or Key.RightShift or Key.LWin or Key.RWin)
        {
            return (0, 0, "");
        }

        uint modifiers = 0;
        var parts = new List<string>();

        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
        {
            modifiers |= GlobalHotkeyService.ModControl;
            parts.Add("Ctrl");
        }
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Alt))
        {
            modifiers |= GlobalHotkeyService.ModAlt;
            parts.Add("Alt");
        }
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
        {
            modifiers |= GlobalHotkeyService.ModShift;
            parts.Add("Shift");
        }
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Windows))
        {
            modifiers |= GlobalHotkeyService.ModWin;
            parts.Add("Win");
        }

        if (modifiers == 0) return (0, 0, "");

        var vk = (uint)KeyInterop.VirtualKeyFromKey(key);
        parts.Add(KeyToDisplayName(key));

        return (modifiers, vk, string.Join("+", parts));
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
            _windowMenuInjectionService.InvalidateCommands();
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

        _activePicker = new WindowPopUp(_virtualDesktopService, _windowMenuInjectionService, targetWindow);
        _activePicker.Closed += (_, _) => _activePicker = null;
        _activePicker.Show();
    }

    private void RefreshRuntimeState()
    {
        var desktops = _virtualDesktopService.GetDesktops();

        var signature = string.Join('|', desktops.Select(desktop => desktop.Id.ToString("D")));
        if (!string.Equals(_lastDesktopSignature, signature, StringComparison.Ordinal))
        {
            _lastDesktopSignature = signature;
            _windowMenuInjectionService.InvalidateCommands();
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
