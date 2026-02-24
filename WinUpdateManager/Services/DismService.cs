using System.Diagnostics;
using System.Text;
using Microsoft.Win32;
using WinUpdateManager.Models;

namespace WinUpdateManager.Services;

/// <summary>
/// DISM 更新包管理服务
/// </summary>
public class DismService
{
    /// <summary>
    /// 获取 DISM 已安装包列表（用于 KB 编号匹配）
    /// </summary>
    public async Task<List<DismPackageInfo>> GetInstalledPackagesAsync()
    {
        var packages = new List<DismPackageInfo>();
        var (_, output) = await RunDismAsync("/Online /Get-Packages /English");
        if (string.IsNullOrEmpty(output)) return packages;

        var blocks = output.Split(new[] { "\r\n\r\n", "\n\n" }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var block in blocks)
        {
            var lines = block.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            var pkg = new DismPackageInfo();
            bool hasId = false;

            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                var colon = trimmed.IndexOf(':');
                if (colon < 0) continue;
                var key = trimmed[..colon].Trim();
                var val = trimmed[(colon + 1)..].Trim();

                if (key.Equals("Package Identity", StringComparison.OrdinalIgnoreCase))
                { pkg.PackageIdentity = val; hasId = !string.IsNullOrEmpty(val); }
                else if (key.Equals("State", StringComparison.OrdinalIgnoreCase))
                    pkg.State = val;
            }

            if (hasId) packages.Add(pkg);
        }
        return packages;
    }

    /// <summary>
    /// 从 CBS 注册表读取 RollupFix 包名列表。
    /// DISM /Get-Packages 会跳过某些 CBS 状态的包，注册表始终完整。
    /// </summary>
    public static List<string> GetCbsRollupFixPackages()
    {
        var result = new List<string>();
        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages");
            if (key == null) return result;

            foreach (var name in key.GetSubKeyNames())
            {
                if (name.Contains("RollupFix", StringComparison.OrdinalIgnoreCase))
                    result.Add(name);
            }
        }
        catch { }
        return result;
    }

    /// <summary>
    /// 匹配更新 → DISM 包
    /// 1) KB 编号匹配 DISM 包  2) 次版本号匹配 CBS RollupFix  3) WUSA 兜底
    /// </summary>
    public List<string> MatchPackagesToUpdates(
        List<UpdateItem> updates, List<DismPackageInfo> dismPkgs, List<string> cbsRollups)
    {
        var logs = new List<string>();

        foreach (var update in updates)
        {
            if (update.Category == UpdateCategory.Driver) continue;
            var label = $"{update.KBArticleID}({update.BuildVersion})";

            // ── 1. KB 编号匹配 ──
            if (!string.IsNullOrEmpty(update.KBArticleID))
            {
                var kbNum = update.KBArticleID.Replace("KB", "", StringComparison.OrdinalIgnoreCase);
                var m = dismPkgs.FirstOrDefault(p =>
                    p.PackageIdentity.Contains(kbNum, StringComparison.OrdinalIgnoreCase));
                if (m != null)
                {
                    update.PackageIdentity = m.PackageIdentity;
                    update.CanUninstall = true;
                    update.UninstallMethod = UninstallMethod.DISM;
                    logs.Add($"{label} -> DISM(KB): {m.PackageIdentity}");
                    continue;
                }
            }

            // ── 2. 次版本号匹配 CBS RollupFix ──
            // WUA "26200.7623" → ".7623." 匹配 CBS "...26100.7623.1.20"
            if (!string.IsNullOrEmpty(update.BuildVersion))
            {
                var parts = update.BuildVersion.Split('.');
                if (parts.Length >= 2)
                {
                    var pattern = $".{parts[1]}.";
                    var pkg = cbsRollups.FirstOrDefault(n =>
                        n.Contains(pattern, StringComparison.OrdinalIgnoreCase));

                    if (pkg != null)
                    {
                        update.PackageIdentity = pkg;
                        update.CanUninstall = true;
                        update.UninstallMethod = UninstallMethod.DISM;
                        logs.Add($"{label} -> DISM: {pkg}");
                        continue;
                    }

                    // 有版本号但 CBS 无此包 → 已被新累积更新取代
                    update.CanUninstall = false;
                    update.UninstallMethod = UninstallMethod.None;
                    logs.Add($"{label} -> 不可卸载(已被取代)");
                    continue;
                }
            }

            // ── 3. WUSA 兜底 ──
            if (!string.IsNullOrEmpty(update.KBArticleID))
            {
                update.CanUninstall = true;
                update.UninstallMethod = UninstallMethod.Wusa;
                logs.Add($"{label} -> WUSA");
            }
        }
        return logs;
    }

    /// <summary>移除指定更新包</summary>
    public async Task<(bool Success, string Message)> RemovePackageAsync(string packageIdentity)
    {
        if (string.IsNullOrEmpty(packageIdentity))
            return (false, "包标识为空");

        var (exitCode, output) = await RunDismAsync(
            $"/Online /Remove-Package /PackageName:\"{packageIdentity}\" /NoRestart /Quiet /English");

        return exitCode switch
        {
            0    => (true, "DISM 已成功移除包"),
            3010 => (true, "DISM 已移除包，需重启生效"),
            _ when (uint)exitCode == 0x800F0825 => (false, "DISM 找不到指定包"),
            _    => (false, $"DISM 错误 {exitCode} (0x{(uint)exitCode:X8})\n{output.Trim()}")
        };
    }

    private static async Task<(int ExitCode, string Output)> RunDismAsync(string arguments)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "DISM.exe",
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8
            };

            using var process = Process.Start(psi);
            if (process == null) return (-1, "无法启动 DISM");

            var stdout = await process.StandardOutput.ReadToEndAsync();
            var stderr = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            return (process.ExitCode, string.IsNullOrEmpty(stderr) ? stdout : $"{stdout}\n{stderr}");
        }
        catch (Exception ex)
        {
            return (-1, $"DISM 执行失败: {ex.Message}");
        }
    }
}

public class DismPackageInfo
{
    public string PackageIdentity { get; set; } = string.Empty;
    public string State { get; set; } = string.Empty;
}
