using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using WinUpdateManager.Models;

namespace WinUpdateManager.Services;

/// <summary>
/// PnPUtil 驱动管理服务
/// 用于枚举和删除第三方驱动包
/// </summary>
public class PnpUtilService
{
    /// <summary>
    /// 枚举所有第三方驱动包
    /// </summary>
    public async Task<List<DriverInfo>> EnumDriversAsync()
    {
        var drivers = new List<DriverInfo>();
        var output = await RunCommandAsync("pnputil", "/enum-drivers");

        if (string.IsNullOrEmpty(output)) return drivers;

        // 解析 pnputil 输出
        var blocks = Regex.Split(output, @"(?=发布名称|Published Name)", RegexOptions.IgnoreCase);

        foreach (var block in blocks)
        {
            if (string.IsNullOrWhiteSpace(block)) continue;

            var info = new DriverInfo();
            
            // 匹配 Published Name / 发布名称
            var nameMatch = Regex.Match(block, @"(?:Published Name|发布名称)\s*[:：]\s*(.+)", RegexOptions.IgnoreCase);
            if (nameMatch.Success)
                info.InfName = nameMatch.Groups[1].Value.Trim();

            // 匹配 Original Name / 原始名称
            var origMatch = Regex.Match(block, @"(?:Original Name|原始名称)\s*[:：]\s*(.+)", RegexOptions.IgnoreCase);
            if (origMatch.Success)
                info.OriginalName = origMatch.Groups[1].Value.Trim();

            // 匹配 Provider Name / 提供程序名称
            var provMatch = Regex.Match(block, @"(?:Provider Name|提供程序名称)\s*[:：]\s*(.+)", RegexOptions.IgnoreCase);
            if (provMatch.Success)
                info.Provider = provMatch.Groups[1].Value.Trim();

            // 匹配 Class Name / 类名
            var classMatch = Regex.Match(block, @"(?:Class Name|类名)\s*[:：]\s*(.+)", RegexOptions.IgnoreCase);
            if (classMatch.Success)
                info.ClassName = classMatch.Groups[1].Value.Trim();

            // 匹配 Driver Version and Date / 驱动程序版本和日期
            var verMatch = Regex.Match(block, @"(?:Driver Version|驱动程序版本和日期)\s*[:：]\s*(.+)", RegexOptions.IgnoreCase);
            if (verMatch.Success)
                info.VersionAndDate = verMatch.Groups[1].Value.Trim();

            // 匹配 Signer Name / 签名者名称
            var signMatch = Regex.Match(block, @"(?:Signer Name|签名者名称)\s*[:：]\s*(.+)", RegexOptions.IgnoreCase);
            if (signMatch.Success)
                info.Signer = signMatch.Groups[1].Value.Trim();

            if (!string.IsNullOrEmpty(info.InfName))
                drivers.Add(info);
        }

        return drivers;
    }

    /// <summary>
    /// 删除指定驱动包（强制卸载）
    /// </summary>
    public async Task<(bool Success, string Message)> DeleteDriverAsync(string infName)
    {
        if (string.IsNullOrEmpty(infName))
            return (false, "INF 文件名为空");

        // /force 强制删除 /uninstall 同时卸载设备
        var output = await RunCommandAsync("pnputil", $"/delete-driver {infName} /uninstall /force");

        bool success = output.Contains("成功", StringComparison.OrdinalIgnoreCase) ||
                       output.Contains("successfully", StringComparison.OrdinalIgnoreCase) ||
                       output.Contains("deleted", StringComparison.OrdinalIgnoreCase);

        return (success, output.Trim());
    }

    /// <summary>
    /// 尝试将更新历史中的驱动项与 pnputil 驱动列表匹配
    /// </summary>
    public void MatchDriversToUpdates(List<UpdateItem> driverUpdates, List<DriverInfo> drivers)
    {
        foreach (var update in driverUpdates)
        {
            if (update.Category != UpdateCategory.Driver) continue;

            // 尝试根据标题中的提供商和名称匹配
            var titleLower = update.Title.ToLowerInvariant();
            
            foreach (var drv in drivers)
            {
                var provider = drv.Provider?.ToLowerInvariant() ?? "";
                var className = drv.ClassName?.ToLowerInvariant() ?? "";
                var origName = drv.OriginalName?.ToLowerInvariant() ?? "";

                // 匹配策略：提供商+类名都在标题中出现
                if (!string.IsNullOrEmpty(provider) && titleLower.Contains(provider) &&
                    (!string.IsNullOrEmpty(className) && titleLower.Contains(className) ||
                     !string.IsNullOrEmpty(origName)))
                {
                    update.DriverInfName = drv.InfName;
                    update.DriverProvider = drv.Provider ?? "";
                    update.DriverClass = drv.ClassName ?? "";
                    update.DriverVersion = drv.VersionAndDate ?? "";
                    update.CanUninstall = true;
                    update.UninstallMethod = UninstallMethod.PnPUtil;
                    break;
                }
            }
        }
    }

    private static async Task<string> RunCommandAsync(string fileName, string arguments)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8
            };

            using var process = Process.Start(psi);
            if (process == null) return string.Empty;

            var output = await process.StandardOutput.ReadToEndAsync();
            var error = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            return string.IsNullOrEmpty(error) ? output : $"{output}\n{error}";
        }
        catch (Exception ex)
        {
            return $"执行命令失败: {ex.Message}";
        }
    }
}

/// <summary>
/// PnPUtil 驱动信息
/// </summary>
public class DriverInfo
{
    public string InfName { get; set; } = string.Empty;
    public string OriginalName { get; set; } = string.Empty;
    public string Provider { get; set; } = string.Empty;
    public string ClassName { get; set; } = string.Empty;
    public string VersionAndDate { get; set; } = string.Empty;
    public string Signer { get; set; } = string.Empty;
}
