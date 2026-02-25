using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using WinUpdateManager.ViewModels;

namespace WinUpdateManager;

/// <summary>
/// MainWindow.xaml 交互逻辑
/// </summary>
public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;

    // --- DWM API for Mica/Dark Mode ---
    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

    private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
    private const int DWMWA_SYSTEMBACKDROP_TYPE = 38;
    private const int DWMWA_MICA_EFFECT = 1029;

    // --- User32 API for Acrylic (Win10) ---
    [DllImport("user32.dll")]
    internal static extern int SetWindowCompositionAttribute(IntPtr hwnd, ref WindowCompositionAttributeData data);

    [StructLayout(LayoutKind.Sequential)]
    internal struct WindowCompositionAttributeData
    {
        public WindowCompositionAttribute Attribute;
        public IntPtr Data;
        public int SizeOfData;
    }

    internal enum WindowCompositionAttribute
    {
        WCA_ACCENT_POLICY = 19
    }

    internal enum AccentState
    {
        ACCENT_DISABLED = 0,
        ACCENT_ENABLE_GRADIENT = 1,
        ACCENT_ENABLE_TRANSPARENTGRADIENT = 2,
        ACCENT_ENABLE_BLURBEHIND = 3,
        ACCENT_ENABLE_ACRYLICBLURBEHIND = 4,
        ACCENT_INVALID_STATE = 5
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct AccentPolicy
    {
        public AccentState AccentState;
        public int AccentFlags;
        public int GradientColor;
        public int AnimationId;
    }

    public MainWindow()
    {
        InitializeComponent();
        _viewModel = new MainViewModel();
        DataContext = _viewModel;
        Loaded += MainWindow_Loaded;
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        ApplyBackdrop();
    }

    private void ApplyBackdrop()
    {
        IntPtr hwnd = new WindowInteropHelper(this).Handle;

        // 1. Enable Immersive Dark Mode
        int isDarkMode = 1;
        DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref isDarkMode, sizeof(int));

        // 2. Apply Backdrop based on OS version
        var osVersion = Environment.OSVersion.Version;
        if (osVersion.Build >= 22523)
        {
            // Windows 11 22H2+ (Mica)
            int backdropType = 2; // DWMSBT_MAINWINDOW
            DwmSetWindowAttribute(hwnd, DWMWA_SYSTEMBACKDROP_TYPE, ref backdropType, sizeof(int));
        }
        else if (osVersion.Build >= 22000)
        {
            // Windows 11 21H2 (Mica)
            int enableMica = 1;
            DwmSetWindowAttribute(hwnd, DWMWA_MICA_EFFECT, ref enableMica, sizeof(int));
        }
        else if (osVersion.Build >= 17134)
        {
            // Windows 10 (Acrylic)
            var accent = new AccentPolicy
            {
                AccentState = AccentState.ACCENT_ENABLE_ACRYLICBLURBEHIND,
                GradientColor = (0x01 << 24) | (0x202020) // 0x01 opacity, #202020 background
            };

            int accentStructSize = Marshal.SizeOf(accent);
            IntPtr accentPtr = Marshal.AllocHGlobal(accentStructSize);
            Marshal.StructureToPtr(accent, accentPtr, false);

            var data = new WindowCompositionAttributeData
            {
                Attribute = WindowCompositionAttribute.WCA_ACCENT_POLICY,
                SizeOfData = accentStructSize,
                Data = accentPtr
            };

            SetWindowCompositionAttribute(hwnd, ref data);
            Marshal.FreeHGlobal(accentPtr);
        }
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        await _viewModel.InitializeAsync();
    }

    // --- Custom Title Bar Handlers ---
    private void MinimizeButton_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;
    
    private void MaximizeButton_Click(object sender, RoutedEventArgs e)
    {
        WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
    }
    
    private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();
}
