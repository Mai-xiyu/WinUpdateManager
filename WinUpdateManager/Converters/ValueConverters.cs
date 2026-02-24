using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using WinUpdateManager.Models;

namespace WinUpdateManager.Converters;

/// <summary>
/// Bool → Visibility 转换器
/// </summary>
public class BoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        bool flag = value is bool b && b;
        bool invert = parameter is string s && s == "Invert";
        if (invert) flag = !flag;
        return flag ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => value is Visibility v && v == Visibility.Visible;
}

/// <summary>
/// 更新分类 → 图标颜色
/// </summary>
public class CategoryToColorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is UpdateCategory cat ? cat switch
        {
            UpdateCategory.Quality => new SolidColorBrush(Color.FromRgb(0, 120, 212)),    // 蓝色
            UpdateCategory.Driver => new SolidColorBrush(Color.FromRgb(16, 185, 129)),     // 绿色
            UpdateCategory.Definition => new SolidColorBrush(Color.FromRgb(245, 158, 11)), // 橙色
            UpdateCategory.Other => new SolidColorBrush(Color.FromRgb(139, 92, 246)),      // 紫色
            _ => new SolidColorBrush(Colors.Gray)
        } : new SolidColorBrush(Colors.Gray);
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// 操作状态 → 前景色
/// </summary>
public class OperationStatusToColorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is OperationStatus status ? status switch
        {
            OperationStatus.Success => new SolidColorBrush(Color.FromRgb(16, 185, 129)),
            OperationStatus.Failed => new SolidColorBrush(Color.FromRgb(239, 68, 68)),
            OperationStatus.InProgress => new SolidColorBrush(Color.FromRgb(59, 130, 246)),
            _ => new SolidColorBrush(Color.FromRgb(156, 163, 175))
        } : new SolidColorBrush(Color.FromRgb(156, 163, 175));
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// 日期格式化转换器
/// </summary>
public class DateFormatConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is DateTime dt)
        {
            if (dt == DateTime.MinValue) return "未知";
            return dt.ToString("yyyy/MM/dd HH:mm");
        }
        return "未知";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// 是否最新安全更新 → 警告色
/// </summary>
public class SecurityRiskConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool isLatest && isLatest)
            return new SolidColorBrush(Color.FromRgb(239, 68, 68)); // 红色警告
        return new SolidColorBrush(Color.FromRgb(229, 231, 235)); // 正常白色
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// 日志条目 → 颜色
/// </summary>
public class LogErrorToColorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool isError && isError)
            return new SolidColorBrush(Color.FromRgb(239, 68, 68));
        return new SolidColorBrush(Color.FromRgb(156, 163, 175));
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

/// <summary>
/// 数量 → 徽章显示（0时隐藏）
/// </summary>
public class CountToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is int count && count > 0)
            return Visibility.Visible;
        return Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}
