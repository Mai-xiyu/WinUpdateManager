using System.Diagnostics;
using System.Text;

namespace WinUpdateManager.Services;

/// <summary>
/// Windows Update Standalone Installer (wusa.exe) 服务
/// 通过 KB 编号直接卸载更新，作为 DISM 的备选方案
/// </summary>
public class WusaService
{
    /// <summary>
    /// 通过 KB 编号卸载更新
    /// </summary>
    public async Task<(bool Success, string Message)> UninstallByKBAsync(string kbArticleId)
    {
        if (string.IsNullOrEmpty(kbArticleId))
            return (false, "KB 编号为空");

        // 提取纯数字部分
        var kbNumber = kbArticleId.Replace("KB", "", StringComparison.OrdinalIgnoreCase).Trim();

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "wusa.exe",
                Arguments = $"/uninstall /kb:{kbNumber} /quiet /norestart",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi);
            if (process == null) return (false, "无法启动 wusa.exe");

            // wusa 超时控制（最长等待5分钟）
            var completed = await Task.Run(() => process.WaitForExit(300_000));
            
            if (!completed)
            {
                try { process.Kill(); } catch { }
                return (false, "wusa.exe 执行超时（5分钟）");
            }

            // wusa.exe 退出码:
            // 0 = 成功
            // 3010 = 成功，需要重启
            // 2359302 = 更新不适用
            // 1058 = 服务被禁用
            var exitCode = process.ExitCode;

            return exitCode switch
            {
                0    => (true,  "更新已成功卸载"),
                3010 => (true,  "更新已卸载，需要重启计算机以完成"),
                // 87 = ERROR_INVALID_PARAMETER: wusa 不支持此更新格式，需改用 DISM
                87   => (false, "wusa 不支持此更新格式 (0x57)——请改用 DISM 方式"),
                // 0x80240006 = WU_E_TOOMANYUPDATES / 不适用
                _ when (uint)exitCode == 0x80240006
                     => (false, "该更新不适用于此系统或已被卸载"),
                // 0x80240017 = WU_E_NOT_APPLICABLE
                _ when (uint)exitCode == 0x80240017
                     => (false, "此更新不可卸载"),
                // 0x80070005 = ACCESS_DENIED
                _ when (uint)exitCode == 0x80070005
                     => (false, "设指定典型: 权限不足 (0x80070005)，请以管理员身份运行"),
                // 0x80070002 = FILE_NOT_FOUND
                _ when (uint)exitCode == 0x80070002
                     => (false, "找不到指定的更新包 (0x80070002)"),
                _ => (false, $"wusa.exe 返回退出码: {exitCode} (0x{(uint)exitCode:X8})")
            };
        }
        catch (Exception ex)
        {
            return (false, $"执行 wusa.exe 失败: {ex.Message}");
        }
    }
}
