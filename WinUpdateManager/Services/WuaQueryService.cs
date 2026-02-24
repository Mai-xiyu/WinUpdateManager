using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using WinUpdateManager.Models;

namespace WinUpdateManager.Services;

/// <summary>
/// Windows Update Agent COM 接口查询服务
/// 通过动态 COM 互操作获取完整的更新历史记录
/// </summary>
public class WuaQueryService
{
    /// <summary>
    /// 获取所有已安装更新的完整历史记录
    /// </summary>
    public Task<List<UpdateItem>> GetUpdateHistoryAsync()
    {
        return Task.Run(() =>
        {
            var updates = new List<UpdateItem>();

            try
            {
                // 通过 ProgID 创建 COM 对象（不需要 COM 引用）
                var sessionType = Type.GetTypeFromProgID("Microsoft.Update.Session")
                    ?? throw new InvalidOperationException("无法创建 Windows Update Session COM 对象");
                dynamic session = Activator.CreateInstance(sessionType)!;
                dynamic searcher = session.CreateUpdateSearcher();

                int count = searcher.GetTotalHistoryCount();
                if (count == 0)
                {
                    Marshal.ReleaseComObject(searcher);
                    Marshal.ReleaseComObject(session);
                    return updates;
                }

                dynamic history = searcher.QueryHistory(0, count);

                for (int i = 0; i < history.Count; i++)
                {
                    dynamic entry = history[i];

                    // Operation: 1 = Installation, 2 = Uninstallation
                    int operation = (int)entry.Operation;
                    if (operation != 1) continue;

                    // ResultCode: 2 = Succeeded, 3 = SucceededWithErrors
                    int resultCode = (int)entry.ResultCode;
                    if (resultCode != 2 && resultCode != 3) continue;

                    string title = entry.Title?.ToString() ?? "未知更新";
                    DateTime date = (DateTime)entry.Date;
                    string description = entry.Description?.ToString() ?? string.Empty;
                    string supportUrl = entry.SupportUrl?.ToString() ?? string.Empty;

                    string updateId = string.Empty;
                    try { updateId = entry.UpdateIdentity?.UpdateID?.ToString() ?? string.Empty; } catch { }

                    var item = new UpdateItem
                    {
                        Title = title,
                        InstalledDate = date,
                        Description = description,
                        SupportUrl = supportUrl,
                        UpdateID = updateId,
                        KBArticleID = ExtractKB(title),
                        BuildVersion = ExtractBuildVersion(title),
                    };

                    // 分类判断
                    item.Category = CategorizeUpdate(entry, title);

                    updates.Add(item);
                }

                // 释放 COM 对象
                Marshal.ReleaseComObject(history);
                Marshal.ReleaseComObject(searcher);
                Marshal.ReleaseComObject(session);
            }
            catch (COMException ex)
            {
                System.Diagnostics.Debug.WriteLine($"WUA COM Error: {ex.Message}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"WUA Error: {ex.Message}");
            }

            return updates;
        });
    }

    /// <summary>
    /// 从标题中提取 KB 编号
    /// </summary>
    private static string ExtractKB(string? title)
    {
        if (string.IsNullOrEmpty(title)) return string.Empty;

        var match = Regex.Match(title, @"KB(\d{6,7})", RegexOptions.IgnoreCase);
        return match.Success ? $"KB{match.Groups[1].Value}" : string.Empty;
    }

    /// <summary>
    /// 从标题中提取版本号
    /// </summary>
    private static string ExtractBuildVersion(string? title)
    {
        if (string.IsNullOrEmpty(title)) return string.Empty;

        var match = Regex.Match(title, @"\((\d{5}\.\d+)\)");
        return match.Success ? match.Groups[1].Value : string.Empty;
    }

    /// <summary>
    /// 根据更新历史条目的分类信息判断更新类型
    /// </summary>
    private static UpdateCategory CategorizeUpdate(dynamic entry, string title)
    {
        var titleLower = title.ToLowerInvariant();

        try
        {
            // 尝试通过 Categories 集合判断
            dynamic? categories = entry.Categories;
            if (categories != null && (int)categories.Count > 0)
            {
                for (int i = 0; i < (int)categories.Count; i++)
                {
                    string catName = (categories[i].Name?.ToString() ?? string.Empty).ToLowerInvariant();
                    string catId = (categories[i].CategoryID?.ToString() ?? string.Empty);

                    // 定义更新（Windows Defender 病毒库等）
                    if (catName.Contains("definition") || catName.Contains("定义") ||
                        catId == "E0789628-CE08-4437-BE74-2495B842F43B")
                        return UpdateCategory.Definition;

                    // 驱动程序更新
                    if (catName.Contains("driver") || catName.Contains("驱动"))
                        return UpdateCategory.Driver;
                }
            }
        }
        catch
        {
            // COM 接口访问可能失败，忽略并根据标题判断
        }

        // 基于标题的后备判断 — 定义更新
        if (titleLower.Contains("definition update") || titleLower.Contains("security intelligence") ||
            titleLower.Contains("antimalware") ||
            (titleLower.Contains("defender") && (titleLower.Contains("定义") || titleLower.Contains("definition"))))
            return UpdateCategory.Definition;

        // 驱动更新
        if (titleLower.Contains("driver") || titleLower.Contains("firmware") ||
            (titleLower.Contains(" - ") && (titleLower.Contains("nvidia") || titleLower.Contains("intel") ||
            titleLower.Contains("amd") || titleLower.Contains("realtek") || titleLower.Contains("usb") ||
            titleLower.Contains("bluetooth") || titleLower.Contains("wi-fi") || titleLower.Contains("audio"))))
            return UpdateCategory.Driver;

        // 质量更新
        if (titleLower.Contains("cumulative update") || titleLower.Contains("security update") ||
            titleLower.Contains("安全更新") || titleLower.Contains("累积更新") ||
            titleLower.Contains("quality") || titleLower.Contains("servicing stack") ||
            titleLower.Contains(".net framework"))
            return UpdateCategory.Quality;

        return UpdateCategory.Other;
    }
}
