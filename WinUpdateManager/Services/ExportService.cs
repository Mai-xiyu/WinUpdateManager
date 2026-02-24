using System.IO;
using System.Text;
using WinUpdateManager.Models;

namespace WinUpdateManager.Services;

/// <summary>
/// 导出服务：支持将更新列表导出为 CSV 和 TXT
/// </summary>
public class ExportService
{
    /// <summary>
    /// 导出为 CSV（UTF-8 BOM，兼容 Excel）
    /// </summary>
    public async Task ExportToCsvAsync(string filePath, IEnumerable<UpdateItem> items)
    {
        var sb = new StringBuilder();

        // CSV 表头
        sb.AppendLine("KB编号,更新名称,分类,安装日期,卸载方式,版本号,描述");

        foreach (var item in items)
        {
            sb.AppendLine(string.Join(",",
                EscapeCsv(item.KBArticleID),
                EscapeCsv(item.Title),
                EscapeCsv(item.CategoryDisplayName),
                EscapeCsv(item.InstalledDate.ToString("yyyy-MM-dd HH:mm:ss")),
                EscapeCsv(item.UninstallMethodDisplay),
                EscapeCsv(item.BuildVersion),
                EscapeCsv(item.Description)
            ));
        }

        // 写入 BOM，确保 Excel 正确识别 UTF-8
        await File.WriteAllTextAsync(filePath, sb.ToString(), new UTF8Encoding(true));
    }

    /// <summary>
    /// 导出为纯文本
    /// </summary>
    public async Task ExportToTxtAsync(string filePath, IEnumerable<UpdateItem> items)
    {
        var sb = new StringBuilder();

        sb.AppendLine("═══════════════════════════════════════════════════════════════");
        sb.AppendLine("                    Windows 更新列表导出报告");
        sb.AppendLine($"                    导出时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine("═══════════════════════════════════════════════════════════════");
        sb.AppendLine();

        var grouped = items.GroupBy(i => i.Category).OrderBy(g => g.Key);

        foreach (var group in grouped)
        {
            var catName = group.Key switch
            {
                UpdateCategory.Quality => "质量更新",
                UpdateCategory.Driver => "驱动程序更新",
                UpdateCategory.Definition => "定义更新",
                UpdateCategory.Other => "其他更新",
                _ => "未知"
            };

            sb.AppendLine($"── {catName} ({group.Count()} 项) ──────────────────────");
            sb.AppendLine();

            int idx = 1;
            foreach (var item in group.OrderByDescending(i => i.InstalledDate))
            {
                sb.AppendLine($"  [{idx:D3}] {item.Title}");
                if (!string.IsNullOrEmpty(item.KBArticleID))
                    sb.AppendLine($"        KB编号: {item.KBArticleID}");
                sb.AppendLine($"        安装日期: {item.InstalledDate:yyyy-MM-dd HH:mm}");
                sb.AppendLine($"        卸载方式: {item.UninstallMethodDisplay}");
                if (!string.IsNullOrEmpty(item.BuildVersion))
                    sb.AppendLine($"        版本号: {item.BuildVersion}");
                sb.AppendLine();
                idx++;
            }
        }

        sb.AppendLine("═══════════════════════════════════════════════════════════════");
        sb.AppendLine($"  合计: {items.Count()} 项更新");
        sb.AppendLine("═══════════════════════════════════════════════════════════════");

        await File.WriteAllTextAsync(filePath, sb.ToString(), Encoding.UTF8);
    }

    /// <summary>
    /// 导出操作日志
    /// </summary>
    public async Task ExportLogAsync(string filePath, IEnumerable<LogEntry> logs)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"操作日志 — {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine(new string('─', 60));

        foreach (var log in logs)
        {
            var prefix = log.IsError ? "[错误]" : "[信息]";
            sb.AppendLine($"{prefix} {log.Display}");
        }

        await File.WriteAllTextAsync(filePath, sb.ToString(), Encoding.UTF8);
    }

    private static string EscapeCsv(string value)
    {
        if (string.IsNullOrEmpty(value)) return "\"\"";
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
            return $"\"{value.Replace("\"", "\"\"")}\"";
        return $"\"{value}\"";
    }
}
