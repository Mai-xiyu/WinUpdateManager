using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace WinUpdateManager.Models;

/// <summary>
/// 更新分类枚举
/// </summary>
public enum UpdateCategory
{
    Quality,     // 质量更新
    Driver,      // 驱动程序更新
    Definition,  // 定义更新
    Other        // 其他更新
}

/// <summary>
/// 卸载方式枚举
/// </summary>
public enum UninstallMethod
{
    None,        // 不可卸载
    DISM,        // 通过 DISM 移除
    PnPUtil,     // 通过 PnPUtil 移除驱动
    Wusa,        // 通过 wusa.exe 移除
    Combined     // 组合方式
}

/// <summary>
/// 操作结果状态
/// </summary>
public enum OperationStatus
{
    Pending,
    InProgress,
    Success,
    Failed,
    Skipped
}

/// <summary>
/// Windows 更新项数据模型
/// </summary>
public class UpdateItem : INotifyPropertyChanged
{
    private bool _isSelected;
    private OperationStatus _operationStatus = OperationStatus.Pending;
    private string _statusMessage = string.Empty;

    /// <summary>更新标题</summary>
    public string Title { get; set; } = string.Empty;

    /// <summary>KB 文章编号（如 KB5077181）</summary>
    public string KBArticleID { get; set; } = string.Empty;

    /// <summary>安装日期</summary>
    public DateTime InstalledDate { get; set; }

    /// <summary>更新分类</summary>
    public UpdateCategory Category { get; set; }

    /// <summary>分类显示名称</summary>
    public string CategoryDisplayName => Category switch
    {
        UpdateCategory.Quality => "质量更新",
        UpdateCategory.Driver => "驱动程序更新",
        UpdateCategory.Definition => "定义更新",
        UpdateCategory.Other => "其他更新",
        _ => "未知"
    };

    /// <summary>更新描述</summary>
    public string Description { get; set; } = string.Empty;

    /// <summary>是否可卸载</summary>
    public bool CanUninstall { get; set; }

    /// <summary>卸载方式</summary>
    public UninstallMethod UninstallMethod { get; set; } = UninstallMethod.None;

    /// <summary>DISM 包名称标识</summary>
    public string PackageIdentity { get; set; } = string.Empty;

    /// <summary>PnPUtil 驱动 INF 文件名（如 oem12.inf）</summary>
    public string DriverInfName { get; set; } = string.Empty;

    /// <summary>驱动提供商</summary>
    public string DriverProvider { get; set; } = string.Empty;

    /// <summary>驱动类别</summary>
    public string DriverClass { get; set; } = string.Empty;

    /// <summary>驱动版本</summary>
    public string DriverVersion { get; set; } = string.Empty;

    /// <summary>更新大小（字节），-1 表示未知</summary>
    public long SizeBytes { get; set; } = -1;

    /// <summary>格式化的大小显示</summary>
    public string SizeDisplay => SizeBytes switch
    {
        < 0 => "未知",
        < 1024 => $"{SizeBytes} B",
        < 1024 * 1024 => $"{SizeBytes / 1024.0:F1} KB",
        < 1024L * 1024 * 1024 => $"{SizeBytes / (1024.0 * 1024):F1} MB",
        _ => $"{SizeBytes / (1024.0 * 1024 * 1024):F2} GB"
    };

    /// <summary>支持标识（如 x64-based Systems）</summary>
    public string SupportUrl { get; set; } = string.Empty;

    /// <summary>更新 ID（内部标识）</summary>
    public string UpdateID { get; set; } = string.Empty;

    /// <summary>构建版本号</summary>
    public string BuildVersion { get; set; } = string.Empty;

    /// <summary>是否为最新安全更新（高风险标记）</summary>
    public bool IsLatestSecurity { get; set; }

    /// <summary>是否选中（用于批量操作）</summary>
    public bool IsSelected
    {
        get => _isSelected;
        set { _isSelected = value; OnPropertyChanged(); }
    }

    /// <summary>操作状态</summary>
    public OperationStatus OperationStatus
    {
        get => _operationStatus;
        set { _operationStatus = value; OnPropertyChanged(); OnPropertyChanged(nameof(OperationStatusDisplay)); }
    }

    /// <summary>状态消息（成功/失败原因）</summary>
    public string StatusMessage
    {
        get => _statusMessage;
        set { _statusMessage = value; OnPropertyChanged(); }
    }

    /// <summary>操作状态显示文本</summary>
    public string OperationStatusDisplay => OperationStatus switch
    {
        OperationStatus.Pending => "",
        OperationStatus.InProgress => "执行中...",
        OperationStatus.Success => "已删除",
        OperationStatus.Failed => $"失败",
        OperationStatus.Skipped => "已跳过",
        _ => ""
    };

    /// <summary>卸载方式显示文本</summary>
    public string UninstallMethodDisplay => UninstallMethod switch
    {
        UninstallMethod.DISM => "DISM",
        UninstallMethod.PnPUtil => "PnPUtil",
        UninstallMethod.Wusa => "WUSA",
        UninstallMethod.Combined => "DISM/WUSA",
        UninstallMethod.None => "—",
        _ => "—"
    };

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

/// <summary>
/// 操作日志条目
/// </summary>
public class LogEntry
{
    public DateTime Timestamp { get; set; } = DateTime.Now;
    public string Message { get; set; } = string.Empty;
    public bool IsError { get; set; }
    public string Display => $"[{Timestamp:HH:mm:ss}] {Message}";
}
