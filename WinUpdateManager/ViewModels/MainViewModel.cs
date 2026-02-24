using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using Microsoft.Win32;
using WinUpdateManager.Models;
using WinUpdateManager.Services;

namespace WinUpdateManager.ViewModels;

/// <summary>
/// 主窗口 ViewModel — 核心业务逻辑
/// </summary>
public class MainViewModel : INotifyPropertyChanged
{
    // ─── 服务 ───
    private readonly WuaQueryService _wuaService = new();
    private readonly PnpUtilService _pnpService = new();
    private readonly DismService _dismService = new();
    private readonly WusaService _wusaService = new();
    private readonly RestorePointService _restoreService = new();
    private readonly ExportService _exportService = new();

    // ─── 所有更新（原始数据） ───
    private List<UpdateItem> _allUpdates = [];

    // ─── 分类列表（筛选后显示） ───
    public ObservableCollection<UpdateItem> QualityUpdates { get; } = [];
    public ObservableCollection<UpdateItem> DriverUpdates { get; } = [];
    public ObservableCollection<UpdateItem> DefinitionUpdates { get; } = [];
    public ObservableCollection<UpdateItem> OtherUpdates { get; } = [];

    // ─── 日志 ───
    public ObservableCollection<LogEntry> Logs { get; } = [];

    // ─── 属性 ───
    private bool _isLoading;
    public bool IsLoading { get => _isLoading; set { _isLoading = value; OnPropertyChanged(); } }

    private bool _isOperating;
    public bool IsOperating { get => _isOperating; set { _isOperating = value; OnPropertyChanged(); } }

    private string _searchText = string.Empty;
    public string SearchText
    {
        get => _searchText;
        set { _searchText = value; OnPropertyChanged(); ApplyFilter(); }
    }

    private string _statusText = "就绪";
    public string StatusText { get => _statusText; set { _statusText = value; OnPropertyChanged(); } }

    private int _totalCount;
    public int TotalCount { get => _totalCount; set { _totalCount = value; OnPropertyChanged(); } }

    private int _selectedCount;
    public int SelectedCount { get => _selectedCount; set { _selectedCount = value; OnPropertyChanged(); } }

    private double _progressValue;
    public double ProgressValue { get => _progressValue; set { _progressValue = value; OnPropertyChanged(); } }

    private double _progressMax = 100;
    public double ProgressMax { get => _progressMax; set { _progressMax = value; OnPropertyChanged(); } }

    private bool _isProgressVisible;
    public bool IsProgressVisible { get => _isProgressVisible; set { _isProgressVisible = value; OnPropertyChanged(); } }

    private int _selectedTabIndex;
    public int SelectedTabIndex { get => _selectedTabIndex; set { _selectedTabIndex = value; OnPropertyChanged(); } }

    // ─── 分类计数 ───
    private int _qualityCount;
    public int QualityCount { get => _qualityCount; set { _qualityCount = value; OnPropertyChanged(); } }

    private int _driverCount;
    public int DriverCount { get => _driverCount; set { _driverCount = value; OnPropertyChanged(); } }

    private int _definitionCount;
    public int DefinitionCount { get => _definitionCount; set { _definitionCount = value; OnPropertyChanged(); } }

    private int _otherCount;
    public int OtherCount { get => _otherCount; set { _otherCount = value; OnPropertyChanged(); } }

    // ─── 详情面板 ───
    private UpdateItem? _selectedUpdate;
    public UpdateItem? SelectedUpdate
    {
        get => _selectedUpdate;
        set { _selectedUpdate = value; OnPropertyChanged(); OnPropertyChanged(nameof(HasSelectedDetail)); }
    }
    public bool HasSelectedDetail => _selectedUpdate != null;

    // ─── 命令 ───
    public AsyncRelayCommand RefreshCommand { get; }
    public AsyncRelayCommand DeleteSelectedCommand { get; }
    public AsyncRelayCommand CreateRestorePointCommand { get; }
    public AsyncRelayCommand ExportCsvCommand { get; }
    public AsyncRelayCommand ExportTxtCommand { get; }
    public RelayCommand SelectAllCommand { get; }
    public RelayCommand DeselectAllCommand { get; }

    public MainViewModel()
    {
        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => !IsLoading && !IsOperating);
        DeleteSelectedCommand = new AsyncRelayCommand(DeleteSelectedAsync, () => !IsLoading && !IsOperating && SelectedCount > 0);
        CreateRestorePointCommand = new AsyncRelayCommand(CreateRestorePointAsync, () => !IsLoading && !IsOperating);
        ExportCsvCommand = new AsyncRelayCommand(ExportCsvAsync, () => TotalCount > 0);
        ExportTxtCommand = new AsyncRelayCommand(ExportTxtAsync, () => TotalCount > 0);
        SelectAllCommand = new RelayCommand(SelectAll);
        DeselectAllCommand = new RelayCommand(DeselectAll);
    }

    /// <summary>
    /// 初始化加载
    /// </summary>
    public async Task InitializeAsync()
    {
        await RefreshAsync();
    }

    /// <summary>
    /// 刷新所有更新数据
    /// </summary>
    private async Task RefreshAsync()
    {
        IsLoading = true;
        StatusText = "正在扫描更新历史...";
        AddLog("开始扫描 Windows 更新历史记录...");

        try
        {
            // 1. 通过 WUA COM 获取更新历史
            StatusText = "正在查询 Windows Update Agent...";
            _allUpdates = await _wuaService.GetUpdateHistoryAsync();
            AddLog($"WUA 返回 {_allUpdates.Count} 条更新记录");

            // 2. 获取 DISM 包 + CBS 注册表 RollupFix 列表
            StatusText = "正在查询 DISM 已安装包...";
            var dismPackages = await _dismService.GetInstalledPackagesAsync();
            var cbsRollups = DismService.GetCbsRollupFixPackages();
            AddLog($"DISM 包 {dismPackages.Count} 个，CBS RollupFix {cbsRollups.Count} 个");

            // 3. 获取 PnPUtil 驱动列表
            StatusText = "正在枚举驱动程序...";
            var drivers = await _pnpService.EnumDriversAsync();
            AddLog($"PnPUtil 返回 {drivers.Count} 个第三方驱动包");

            // 4. 匹配卸载方式
            StatusText = "正在匹配卸载信息...";
            var matchLogs = _dismService.MatchPackagesToUpdates(_allUpdates, dismPackages, cbsRollups);
            foreach (var log in matchLogs)
                AddLog($"[匹配] {log}");
            _pnpService.MatchDriversToUpdates(
                _allUpdates.Where(u => u.Category == UpdateCategory.Driver).ToList(), drivers);

            // 5. 标记最新安全更新
            var latestQuality = _allUpdates
                .Where(u => u.Category == UpdateCategory.Quality)
                .OrderByDescending(u => u.InstalledDate)
                .FirstOrDefault();
            if (latestQuality != null) latestQuality.IsLatestSecurity = true;

            // 6. 为每个更新项订阅选择变更
            foreach (var update in _allUpdates)
            {
                update.PropertyChanged += OnUpdateSelectionChanged;
            }

            // 7. 应用筛选并显示
            ApplyFilter();

            TotalCount = _allUpdates.Count;
            StatusText = $"扫描完成 — 共 {TotalCount} 项更新";
            AddLog($"扫描完成：质量更新 {QualityCount}、驱动 {DriverCount}、定义 {DefinitionCount}、其他 {OtherCount}");
        }
        catch (Exception ex)
        {
            StatusText = $"扫描失败: {ex.Message}";
            AddLog($"扫描失败: {ex.Message}", true);
        }
        finally
        {
            IsLoading = false;
        }
    }

    /// <summary>
    /// 应用搜索筛选
    /// </summary>
    private void ApplyFilter()
    {
        var keyword = SearchText.Trim().ToLowerInvariant();

        IEnumerable<UpdateItem> filtered = _allUpdates;

        if (!string.IsNullOrEmpty(keyword))
        {
            filtered = filtered.Where(u =>
                u.Title.Contains(keyword, StringComparison.OrdinalIgnoreCase) ||
                u.KBArticleID.Contains(keyword, StringComparison.OrdinalIgnoreCase) ||
                u.Description.Contains(keyword, StringComparison.OrdinalIgnoreCase)
            );
        }

        var list = filtered.OrderByDescending(u => u.InstalledDate).ToList();

        Application.Current.Dispatcher.Invoke(() =>
        {
            QualityUpdates.Clear();
            DriverUpdates.Clear();
            DefinitionUpdates.Clear();
            OtherUpdates.Clear();

            foreach (var item in list)
            {
                switch (item.Category)
                {
                    case UpdateCategory.Quality: QualityUpdates.Add(item); break;
                    case UpdateCategory.Driver: DriverUpdates.Add(item); break;
                    case UpdateCategory.Definition: DefinitionUpdates.Add(item); break;
                    case UpdateCategory.Other: OtherUpdates.Add(item); break;
                }
            }

            QualityCount = QualityUpdates.Count;
            DriverCount = DriverUpdates.Count;
            DefinitionCount = DefinitionUpdates.Count;
            OtherCount = OtherUpdates.Count;
        });

        UpdateSelectedCount();
    }

    /// <summary>
    /// 选中项变更时更新计数
    /// </summary>
    private void OnUpdateSelectionChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(UpdateItem.IsSelected))
            UpdateSelectedCount();
    }

    private void UpdateSelectedCount()
    {
        SelectedCount = _allUpdates.Count(u => u.IsSelected);
    }

    /// <summary>
    /// 全选当前 Tab 的更新
    /// </summary>
    private void SelectAll()
    {
        foreach (var item in GetCurrentTabItems())
            item.IsSelected = true;
    }

    /// <summary>
    /// 取消全选
    /// </summary>
    private void DeselectAll()
    {
        foreach (var item in _allUpdates)
            item.IsSelected = false;
    }

    /// <summary>
    /// 获取当前选中 Tab 的列表
    /// </summary>
    private ObservableCollection<UpdateItem> GetCurrentTabItems()
    {
        return SelectedTabIndex switch
        {
            0 => QualityUpdates,
            1 => DriverUpdates,
            2 => DefinitionUpdates,
            3 => OtherUpdates,
            _ => QualityUpdates
        };
    }

    /// <summary>
    /// 批量删除选中的更新
    /// </summary>
    private async Task DeleteSelectedAsync()
    {
        var selected = _allUpdates.Where(u => u.IsSelected && u.CanUninstall).ToList();
        if (selected.Count == 0)
        {
            MessageBox.Show("没有选中可卸载的更新项。\n\n提示: 没有 KB 编号或卸载方式标记为 - 的更新不支持卸载。",
                "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        // 安全警告
        var warningItems = selected.Where(u => u.IsLatestSecurity).ToList();
        var warningMsg = warningItems.Count > 0
            ? $"\n\n⚠ 警告：选中项包含最新安全更新，删除后可能影响系统安全！"
            : "";

        var confirm = MessageBox.Show(
            $"确定要删除以下 {selected.Count} 项更新？\n\n" +
            string.Join("\n", selected.Take(10).Select(u => $"  • {u.KBArticleID} {u.Title}")) +
            (selected.Count > 10 ? $"\n  ... 及其他 {selected.Count - 10} 项" : "") +
            warningMsg +
            "\n\n建议先创建还原点。此操作不可轻易撤销。",
            "确认删除", MessageBoxButton.YesNo, MessageBoxImage.Warning);

        if (confirm != MessageBoxResult.Yes) return;

        IsOperating = true;
        IsProgressVisible = true;
        ProgressMax = selected.Count;
        ProgressValue = 0;

        int successCount = 0, failCount = 0;
        AddLog($"开始批量删除 {selected.Count} 项更新...");

        foreach (var item in selected)
        {
            item.OperationStatus = OperationStatus.InProgress;
            StatusText = $"正在删除: {item.KBArticleID} {item.Title}...";

            try
            {
                var (success, message) = item.UninstallMethod switch
                {
                    UninstallMethod.PnPUtil  => await _pnpService.DeleteDriverAsync(item.DriverInfName),
                    UninstallMethod.DISM     => await _dismService.RemovePackageAsync(item.PackageIdentity),
                    UninstallMethod.Wusa     => await TryWusaWithDismFallback(item),
                    UninstallMethod.Combined => await TryCombinedUninstall(item),
                    _                        => (false, "不支持的卸载方式")
                };

                if (success)
                {
                    item.OperationStatus = OperationStatus.Success;
                    item.StatusMessage = message;
                    successCount++;
                    AddLog($"✅ 成功删除: {item.KBArticleID} {item.Title}");
                }
                else
                {
                    item.OperationStatus = OperationStatus.Failed;
                    item.StatusMessage = message;
                    failCount++;
                    AddLog($"❌ 删除失败: {item.KBArticleID} — {message}", true);
                }
            }
            catch (Exception ex)
            {
                item.OperationStatus = OperationStatus.Failed;
                item.StatusMessage = ex.Message;
                failCount++;
                AddLog($"❌ 异常: {item.KBArticleID} — {ex.Message}", true);
            }

            ProgressValue++;
        }

        IsOperating = false;
        StatusText = $"操作完成 — 成功 {successCount} 项，失败 {failCount} 项";
        AddLog($"批量删除完成：成功 {successCount}，失败 {failCount}");

        if (successCount > 0)
        {
            var restart = MessageBox.Show(
                $"成功删除 {successCount} 项更新。\n\n某些更新可能需要重启计算机才能完成卸载。是否现在重启？",
                "操作完成", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (restart == MessageBoxResult.Yes)
            {
                System.Diagnostics.Process.Start("shutdown", "/r /t 30 /c \"Windows 更新管理工具：正在应用更改，30秒后将重启...\"");
            }
        }
    }

    /// <summary>
    /// 组合卸载：先尝试 DISM，失败后尝试 WUSA（若 WUSA 返回 87 则再次扫描 DISM）
    /// </summary>
    private async Task<(bool Success, string Message)> TryCombinedUninstall(UpdateItem item)
    {
        // 先尝试 DISM
        if (!string.IsNullOrEmpty(item.PackageIdentity))
        {
            var (ok, msg) = await _dismService.RemovePackageAsync(item.PackageIdentity);
            if (ok) return (true, $"(DISM) {msg}");
        }

        // 再尝试 WUSA（内带 DISM 日期回退）
        if (!string.IsNullOrEmpty(item.KBArticleID))
        {
            return await TryWusaWithDismFallback(item);
        }

        return (false, "无可用的卸载方式");
    }

    /// <summary>
    /// 尝试 WUSA；若返回 error 87，则通过 CBS 注册表查找 DISM 包再尝试
    /// </summary>
    private async Task<(bool Success, string Message)> TryWusaWithDismFallback(UpdateItem item)
    {
        var (ok, msg) = await _wusaService.UninstallByKBAsync(item.KBArticleID);
        if (ok) return (true, $"(WUSA) {msg}");

        // Error 87 → 改用 DISM（通过 CBS 注册表查包名）
        if (msg.Contains("0x57") || msg.Contains("87") || msg.Contains("不支持此更新格式"))
        {
            AddLog($"{item.KBArticleID}: WUSA 错误 87，查找 CBS 包...");

            if (!string.IsNullOrEmpty(item.BuildVersion))
            {
                var parts = item.BuildVersion.Split('.');
                if (parts.Length >= 2)
                {
                    var pattern = $".{parts[1]}.";
                    var cbsPkg = DismService.GetCbsRollupFixPackages()
                        .FirstOrDefault(n => n.Contains(pattern, StringComparison.OrdinalIgnoreCase));

                    if (cbsPkg != null)
                    {
                        AddLog($"CBS 匹配: {cbsPkg}");
                        var (dismOk, dismMsg) = await _dismService.RemovePackageAsync(cbsPkg);
                        if (dismOk) { item.PackageIdentity = cbsPkg; return (true, $"(DISM) {dismMsg}"); }
                        return (false, $"DISM 失败: {dismMsg}");
                    }
                }
            }

            return (false, $"WUSA 错误 87，CBS 中未找到匹配包");
        }

        return (false, $"(WUSA) {msg}");
    }

    /// <summary>
    /// 创建还原点
    /// </summary>
    private async Task CreateRestorePointAsync()
    {
        StatusText = "正在创建系统还原点...";
        AddLog("正在创建系统还原点...");

        var description = $"WinUpdateManager 备份 — {DateTime.Now:yyyy-MM-dd HH:mm}";
        var (success, message) = await _restoreService.CreateRestorePointAsync(description);

        if (success)
        {
            StatusText = "还原点创建成功";
            AddLog($"✅ {message}");
            MessageBox.Show(message, "成功", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        else
        {
            StatusText = "还原点创建失败";
            AddLog($"❌ {message}", true);
            MessageBox.Show($"创建还原点失败：\n{message}\n\n请确认已启用系统保护（控制面板 → 系统 → 系统保护）。",
                "失败", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    /// <summary>
    /// 导出为 CSV
    /// </summary>
    private async Task ExportCsvAsync()
    {
        var dlg = new SaveFileDialog
        {
            Filter = "CSV 文件|*.csv",
            FileName = $"WindowsUpdates_{DateTime.Now:yyyyMMdd_HHmmss}.csv",
            Title = "导出更新列表 (CSV)"
        };

        if (dlg.ShowDialog() == true)
        {
            await _exportService.ExportToCsvAsync(dlg.FileName, _allUpdates);
            AddLog($"✅ 已导出 CSV: {dlg.FileName}");
            StatusText = $"已导出到 {dlg.FileName}";
        }
    }

    /// <summary>
    /// 导出为 TXT
    /// </summary>
    private async Task ExportTxtAsync()
    {
        var dlg = new SaveFileDialog
        {
            Filter = "文本文件|*.txt",
            FileName = $"WindowsUpdates_{DateTime.Now:yyyyMMdd_HHmmss}.txt",
            Title = "导出更新列表 (TXT)"
        };

        if (dlg.ShowDialog() == true)
        {
            await _exportService.ExportToTxtAsync(dlg.FileName, _allUpdates);
            AddLog($"✅ 已导出 TXT: {dlg.FileName}");
            StatusText = $"已导出到 {dlg.FileName}";
        }
    }

    /// <summary>
    /// 添加日志
    /// </summary>
    private void AddLog(string message, bool isError = false)
    {
        Application.Current.Dispatcher.Invoke(() =>
        {
            Logs.Add(new LogEntry { Message = message, IsError = isError });
            // 保持最新日志可见
        });
    }

    // ─── INotifyPropertyChanged ───
    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
