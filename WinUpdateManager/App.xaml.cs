using System.Security.Principal;
using System.Windows;

namespace WinUpdateManager;

/// <summary>
/// App.xaml 交互逻辑
/// </summary>
public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // 检查管理员权限
        var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);

        if (!principal.IsInRole(WindowsBuiltInRole.Administrator))
        {
            MessageBox.Show(
                "此工具需要以管理员权限运行。\n\n请右键点击程序 → 以管理员身份运行。",
                "权限不足",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            Shutdown(1);
            return;
        }
    }
}
