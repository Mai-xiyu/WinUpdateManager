using System.Management;

namespace WinUpdateManager.Services;

/// <summary>
/// 系统还原点服务
/// 通过 WMI 创建系统还原点
/// </summary>
public class RestorePointService
{
    /// <summary>
    /// 创建系统还原点
    /// </summary>
    /// <param name="description">还原点描述</param>
    /// <returns>是否创建成功及消息</returns>
    public async Task<(bool Success, string Message)> CreateRestorePointAsync(string description)
    {
        return await Task.Run(() =>
        {
            try
            {
                // 检查系统保护是否启用
                var oScope = new ManagementScope(@"\\localhost\root\default");
                var oPath = new ManagementPath("SystemRestore");
                var oGetIn = new ManagementClass(oScope, oPath, new ObjectGetOptions());

                var inParams = oGetIn.GetMethodParameters("CreateRestorePoint");

                // RestorePointType: 0 = APPLICATION_INSTALL, 12 = MODIFY_SETTINGS
                // EventType: 100 = BEGIN_SYSTEM_CHANGE
                inParams["Description"] = description;
                inParams["RestorePointType"] = 12;  // MODIFY_SETTINGS
                inParams["EventType"] = 100;         // BEGIN_SYSTEM_CHANGE

                var outParams = oGetIn.InvokeMethod("CreateRestorePoint", inParams, null);

                if (outParams != null)
                {
                    var returnValue = Convert.ToInt32(outParams["ReturnValue"]);
                    if (returnValue == 0)
                        return (true, $"还原点[{description}]已成功创建");
                    else
                        return (false, $"创建还原点失败，返回值: {returnValue}");
                }

                return (false, "创建还原点失败，无返回值");
            }
            catch (ManagementException ex)
            {
                return (false, $"WMI 错误: {ex.Message}。请确认系统保护已启用。");
            }
            catch (Exception ex)
            {
                return (false, $"创建还原点失败: {ex.Message}");
            }
        });
    }

    /// <summary>
    /// 检查系统保护是否启用
    /// </summary>
    public async Task<bool> IsSystemProtectionEnabledAsync()
    {
        return await Task.Run(() =>
        {
            try
            {
                var scope = new ManagementScope(@"\\localhost\root\default");
                var query = new ObjectQuery("SELECT * FROM SystemRestoreConfig");
                using var searcher = new ManagementObjectSearcher(scope, query);
                var results = searcher.Get();
                return results.Count > 0;
            }
            catch
            {
                return false;
            }
        });
    }
}
