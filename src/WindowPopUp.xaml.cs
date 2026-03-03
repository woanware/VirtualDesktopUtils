using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using VirtualDesktopUtils.Models;
using VirtualDesktopUtils.Services;

namespace VirtualDesktopUtils;

public partial class WindowPopUp : Window
{
    private readonly VirtualDesktopService _virtualDesktopService;
    private readonly WindowMenuInjectionService _windowMenuInjectionService;
    private readonly nint _targetWindowHandle;
    private List<PickerItem> _items = new();
    private bool _closing;

    internal WindowPopUp(
        VirtualDesktopService virtualDesktopService,
        WindowMenuInjectionService windowMenuInjectionService,
        nint targetWindowHandle)
    {
        _virtualDesktopService = virtualDesktopService;
        _windowMenuInjectionService = windowMenuInjectionService;
        _targetWindowHandle = targetWindowHandle;

        InitializeComponent();

        var windowTitle = GetWindowTitle(targetWindowHandle);
        WindowNameTextBlock.Text = string.IsNullOrWhiteSpace(windowTitle)
            ? "Unknown window"
            : windowTitle;

        PopulateDesktops();

        Loaded += (_, _) =>
        {
            ShowOnAllButton.Content = _isPinned ? "📌  Unpin from all desktops" : "📌  Pin to all desktops";

            var hwnd = new System.Windows.Interop.WindowInteropHelper(this).Handle;
            if (hwnd != nint.Zero)
            {
                SetForegroundWindow(hwnd);
            }

            Activate();
            Focus();
            if (DesktopListBox.Items.Count > 0)
            {
                DesktopListBox.SelectedIndex = 0;
                DesktopListBox.Focus();
            }
        };
    }

    private bool _isPinned;

    private void PopulateDesktops()
    {
        var desktops = _virtualDesktopService.GetDesktops();
        var windowDesktopId = _virtualDesktopService.GetWindowDesktopId(_targetWindowHandle);
        _isPinned = _virtualDesktopService.IsWindowPinned(_targetWindowHandle);

        var index = 0;
        _items = desktops
            .Where(d => d.Id != windowDesktopId)
            .Select(d => new PickerItem(++index, index.ToString(), d.DisplayName, d.Id))
            .ToList();

        DesktopListBox.ItemsSource = _items;
    }

    private void SafeClose()
    {
        if (_closing) return;
        _closing = true;
        try { Close(); } catch { }
    }

    private void MoveToItem(PickerItem item)
    {
        try
        {
            var (moved, _) = _virtualDesktopService.MoveWindowToDesktop(_targetWindowHandle, item.DesktopId);
            if (moved)
            {
                _windowMenuInjectionService.InvalidateCommands();
            }
        }
        catch
        {
        }

        SafeClose();
    }

    private void ShowOnAllButton_Click(object sender, RoutedEventArgs e)
    {
        PinAndClose();
    }

    private void PinAndClose()
    {
        try
        {
            if (_isPinned)
                _virtualDesktopService.UnpinWindow(_targetWindowHandle);
            else
                _virtualDesktopService.PinWindow(_targetWindowHandle);
            _windowMenuInjectionService.InvalidateCommands();
        }
        catch
        {
        }

        SafeClose();
    }

    private void Window_Deactivated(object? sender, EventArgs e)
    {
        SafeClose();
    }

    private void Window_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.Escape:
                Close();
                e.Handled = true;
                return;

            case Key.P:
                PinAndClose();
                e.Handled = true;
                return;

            case Key.Enter:
                if (DesktopListBox.SelectedItem is PickerItem selected)
                {
                    MoveToItem(selected);
                    e.Handled = true;
                }
                else if (ShowOnAllButton.IsFocused)
                {
                    PinAndClose();
                    e.Handled = true;
                }
                return;

            case Key.Tab:
                CycleSelection(Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) ? -1 : 1);
                e.Handled = true;
                return;

            case Key.Down:
                CycleSelection(1);
                e.Handled = true;
                return;

            case Key.Up:
                CycleSelection(-1);
                e.Handled = true;
                return;
        }

        // Number keys 1-9 for direct selection
        var number = e.Key switch
        {
            Key.D1 => 1, Key.D2 => 2, Key.D3 => 3, Key.D4 => 4, Key.D5 => 5,
            Key.D6 => 6, Key.D7 => 7, Key.D8 => 8, Key.D9 => 9,
            Key.NumPad1 => 1, Key.NumPad2 => 2, Key.NumPad3 => 3, Key.NumPad4 => 4,
            Key.NumPad5 => 5, Key.NumPad6 => 6, Key.NumPad7 => 7, Key.NumPad8 => 8,
            Key.NumPad9 => 9,
            _ => 0
        };

        if (number > 0)
        {
            var match = _items.FirstOrDefault(i => i.Number == number);
            if (match is not null)
            {
                MoveToItem(match);
                e.Handled = true;
            }
        }
    }

    private void DesktopListBox_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        // Let Window_PreviewKeyDown handle everything
    }

    private void DesktopListBox_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (DesktopListBox.SelectedItem is PickerItem item)
        {
            MoveToItem(item);
        }
    }

    private void CycleSelection(int direction)
    {
        var totalItems = _items.Count;
        // +1 for the "Show on all" button at the end
        var totalSlots = totalItems + 1;
        var currentIndex = DesktopListBox.SelectedIndex;

        // If "Show on all" is focused, treat as slot after last list item
        if (ShowOnAllButton.IsFocused)
        {
            currentIndex = totalItems;
        }

        var nextIndex = ((currentIndex + direction) % totalSlots + totalSlots) % totalSlots;

        if (nextIndex < totalItems)
        {
            DesktopListBox.SelectedIndex = nextIndex;
            DesktopListBox.Focus();
        }
        else
        {
            DesktopListBox.SelectedIndex = -1;
            ShowOnAllButton.Focus();
        }
    }

    private static string? GetWindowTitle(nint windowHandle)
    {
        var length = GetWindowTextLength(windowHandle);
        if (length <= 0)
        {
            return null;
        }

        var buffer = new char[length + 1];
        _ = GetWindowText(windowHandle, buffer, buffer.Length);
        return new string(buffer, 0, length);
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(nint hWnd, char[] lpString, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern int GetWindowTextLength(nint hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(nint hWnd);

    private sealed record PickerItem(int Number, string Shortcut, string DisplayName, Guid DesktopId);
}
