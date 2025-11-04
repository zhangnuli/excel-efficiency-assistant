using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEfficiencyAssistant.Models
{
    /// <summary>
    /// 通用数据模型和辅助类
    /// </summary>

    #region 基础模型

    /// <summary>
    /// 操作结果基类
    /// </summary>
    public abstract class OperationResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public DateTime StartTime { get; set; } = DateTime.Now;
        public DateTime EndTime { get; set; }
        public TimeSpan Duration => EndTime - StartTime;
        public List<string> Warnings { get; set; } = new List<string>();
        public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();

        public virtual void Complete()
        {
            EndTime = DateTime.Now;
        }
    }

    /// <summary>
    /// 分页参数
    /// </summary>
    public class PaginationParameters
    {
        public int PageIndex { get; set; } = 1;
        public int PageSize { get; set; } = 100;
        public string SortBy { get; set; }
        public bool SortDescending { get; set; } = false;

        public int Skip => (PageIndex - 1) * PageSize;
    }

    /// <summary>
    /// 分页结果
    /// </summary>
    public class PagedResult<T>
    {
        public List<T> Items { get; set; } = new List<T>();
        public int TotalCount { get; set; }
        public int PageIndex { get; set; }
        public int PageSize { get; set; }
        public int TotalPages => (int)Math.Ceiling((double)TotalCount / PageSize);
        public bool HasPreviousPage => PageIndex > 1;
        public bool HasNextPage => PageIndex < TotalPages;
    }

    #endregion

    #region Excel相关模型

    /// <summary>
    /// Excel区域信息
    /// </summary>
    public class ExcelRangeInfo
    {
        public string WorksheetName { get; set; }
        public string Address { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public int StartRow { get; set; }
        public int StartColumn { get; set; }
        public int EndRow { get; set; }
        public int EndColumn { get; set; }
        public object[,] Values { get; set; }

        public bool IsSingleRow => RowCount == 1;
        public bool IsSingleColumn => ColumnCount == 1;
        public bool IsSingleCell => RowCount == 1 && ColumnCount == 1;
    }

    /// <summary>
    /// Excel工作簿信息
    /// </summary>
    public class ExcelWorkbookInfo
    {
        public string Name { get; set; }
        public string FullName { get; set; }
        public bool ReadOnly { get; set; }
        public bool Saved { get; set; }
        public List<string> WorksheetNames { get; set; } = new List<string>();
        public Dictionary<string, ExcelRangeInfo> UsedRanges { get; set; } = new Dictionary<string, ExcelRangeInfo>();
    }

    /// <summary>
    /// 单元格信息
    /// </summary>
    public class CellInfo
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string Address { get; set; }
        public object Value { get; set; }
        public string Formula { get; set; }
        public string NumberFormat { get; set; }
        public System.Drawing.Color BackgroundColor { get; set; }
        public System.Drawing.Color FontColor { get; set; }
        public bool HasFormula => !string.IsNullOrEmpty(Formula);
        public bool IsEmpty => Value == null || Value.ToString().Trim() == "";
    }

    #endregion

    #region 配置和设置

    /// <summary>
    /// 应用程序设置
    /// </summary>
    public class ApplicationSettings
    {
        #region 性能设置
        public bool EnablePerformanceOptimization { get; set; } = true;
        public int MaxProcessingRows { get; set; } = 10000;
        public int BatchSize { get; set; } = 1000;
        public bool ShowProgressIndicator { get; set; } = true;
        public int ProgressUpdateInterval { get; set; } = 100;
        #endregion

        #region 界面设置
        public bool ShowTooltips { get; set; } = true;
        public bool EnableAnimations { get; set; } = true;
        public string Theme { get; set; } = "default";
        public System.Drawing.Color AccentColor { get; set; } = System.Drawing.Color.FromArgb(0, 120, 212);
        #endregion

        #region 数据处理设置
        public bool AutoBackup { get; set; } = true;
        public int BackupCount { get; set; } = 5;
        public bool SkipEmptyCells { get; set; } = true;
        public bool PreserveFormatting { get; set; } = false;
        #endregion

        #region 日志设置
        public bool EnableLogging { get; set; } = true;
        public LogLevel LogLevel { get; set; } = LogLevel.Information;
        public string LogPath { get; set; } = "";
        public int MaxLogFileSize { get; set; } = 10; // MB
        public int MaxLogFiles { get; set; } = 10;
        #endregion

        #region 安全设置
        public bool ConfirmBeforeExecution { get; set; } = true;
        public bool EnableUndo { get; set; } = true;
        public int MaxUndoOperations { get; set; } = 50;
        #endregion
    }

    /// <summary>
    /// 用户偏好设置
    /// </summary>
    public class UserPreferences
    {
        public string DefaultLanguage { get; set; } = "zh-CN";
        public bool EnableSounds { get; set; } = false;
        public bool AutoSave { get; set; } = true;
        public int AutoSaveInterval { get; set; } = 5; // 分钟
        public List<string> RecentFiles { get; set; } = new List<string>();
        public Dictionary<string, object> CustomPreferences { get; set; } = new Dictionary<string, object>();
    }

    #endregion

    #region 统计和报告

    /// <summary>
    /// 操作统计
    /// </summary>
    public class OperationStatistics
    {
        public int TotalOperations { get; set; }
        public int SuccessfulOperations { get; set; }
        public int FailedOperations { get; set; }
        public TimeSpan TotalProcessingTime { get; set; }
        public DateTime LastOperationTime { get; set; }
        public Dictionary<string, int> OperationCounts { get; set; } = new Dictionary<string, int>();
        public Dictionary<string, TimeSpan> OperationTimes { get; set; } = new Dictionary<string, TimeSpan>();

        public double SuccessRate => TotalOperations > 0 ? (double)SuccessfulOperations / TotalOperations * 100 : 0;
        public TimeSpan AverageProcessingTime => TotalOperations > 0 ? TimeSpan.FromTicks(TotalProcessingTime.Ticks / TotalOperations) : TimeSpan.Zero;
    }

    /// <summary>
    /// 性能指标
    /// </summary>
    public class PerformanceMetrics
    {
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan Duration => EndTime - StartTime;
        public long MemoryUsed { get; set; }
        public int ProcessedRows { get; set; }
        public int ProcessedColumns { get; set; }
        public int ProcessedCells => ProcessedRows * ProcessedColumns;
        public double CellsPerSecond => Duration.TotalSeconds > 0 ? ProcessedCells / Duration.TotalSeconds : 0;
        public Dictionary<string, object> AdditionalMetrics { get; set; } = new Dictionary<string, object>();
    }

    #endregion

    #region 验证和错误处理

    /// <summary>
    /// 验证结果
    /// </summary>
    public class ValidationResult
    {
        public bool IsValid { get; set; }
        public List<ValidationError> Errors { get; set; } = new List<ValidationError>();
        public List<ValidationWarning> Warnings { get; set; } = new List<ValidationWarning>();

        public bool HasErrors => Errors.Count > 0;
        public bool HasWarnings => Warnings.Count > 0;
    }

    /// <summary>
    /// 验证错误
    /// </summary>
    public class ValidationError
    {
        public string PropertyName { get; set; }
        public string ErrorMessage { get; set; }
        public object AttemptedValue { get; set; }
        public ValidationSeverity Severity { get; set; } = ValidationSeverity.Error;
    }

    /// <summary>
    /// 验证警告
    /// </summary>
    public class ValidationWarning
    {
        public string PropertyName { get; set; }
        public string WarningMessage { get; set; }
        public object Value { get; set; }
        public ValidationSeverity Severity { get; set; } = ValidationSeverity.Warning;
    }

    #endregion

    #region 事件和消息

    /// <summary>
    /// 应用程序事件
    /// </summary>
    public class ApplicationEvent
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public EventType Type { get; set; }
        public string Source { get; set; }
        public string Message { get; set; }
        public object Data { get; set; }
        public Dictionary<string, object> Context { get; set; } = new Dictionary<string, object>();
    }

    /// <summary>
    /// 进度信息
    /// </summary>
    public class ProgressInfo
    {
        public int CurrentValue { get; set; }
        public int MaximumValue { get; set; }
        public int Percentage => MaximumValue > 0 ? (int)((double)CurrentValue / MaximumValue * 100) : 0;
        public string Message { get; set; }
        public bool IsIndeterminate { get; set; } = false;
        public DateTime StartTime { get; set; } = DateTime.Now;
        public TimeSpan Elapsed => DateTime.Now - StartTime;
        public TimeSpan? EstimatedTimeRemaining { get; set; }

        public void Update(int current, string message = null)
        {
            CurrentValue = current;
            if (!string.IsNullOrEmpty(message))
                Message = message;

            // 计算预估剩余时间
            if (CurrentValue > 0 && MaximumValue > 0)
            {
                var avgTimePerItem = Elapsed.TotalMilliseconds / CurrentValue;
                var remainingItems = MaximumValue - CurrentValue;
                EstimatedTimeRemaining = TimeSpan.FromMilliseconds(avgTimePerItem * remainingItems);
            }
        }
    }

    #endregion

    #region 枚举

    /// <summary>
    /// 日志级别
    /// </summary>
    public enum LogLevel
    {
        Trace = 0,
        Debug = 1,
        Information = 2,
        Warning = 3,
        Error = 4,
        Critical = 5
    }

    /// <summary>
    /// 验证严重程度
    /// </summary>
    public enum ValidationSeverity
    {
        Info,
        Warning,
        Error,
        Critical
    }

    /// <summary>
    /// 事件类型
    /// </summary>
    public enum EventType
    {
        Information,
        Warning,
        Error,
        Success,
        Processing,
        Completed
    }

    /// <summary>
    /// 操作状态
    /// </summary>
    public enum OperationStatus
    {
        NotStarted,
        InProgress,
        Completed,
        Failed,
        Cancelled
    }

    #endregion

    #region 扩展方法

    /// <summary>
    /// 扩展方法类
    /// </summary>
    public static class Extensions
    {
        /// <summary>
        /// 转换为Excel区域信息
        /// </summary>
        public static ExcelRangeInfo ToRangeInfo(this Excel.Range range)
        {
            if (range == null) return null;

            return new ExcelRangeInfo
            {
                WorksheetName = range.Worksheet.Name,
                Address = range.Address,
                RowCount = range.Rows.Count,
                ColumnCount = range.Columns.Count,
                StartRow = range.Row,
                StartColumn = range.Column,
                EndRow = range.Row + range.Rows.Count - 1,
                EndColumn = range.Column + range.Columns.Count - 1,
                Values = range.Value2 as object[,]
            };
        }

        /// <summary>
        /// 安全转换为字符串
        /// </summary>
        public static string SafeToString(this object value)
        {
            return value?.ToString() ?? "";
        }

        /// <summary>
        /// 检查字符串是否为空或空白
        /// </summary>
        public static bool IsNullOrWhiteSpace(this string value)
        {
            return string.IsNullOrWhiteSpace(value);
        }

        /// <summary>
        /// 格式化时间跨度
        /// </summary>
        public static string ToFormattedString(this TimeSpan timeSpan)
        {
            if (timeSpan.TotalSeconds < 1)
                return $"{timeSpan.TotalMilliseconds:F0}ms";
            else if (timeSpan.TotalMinutes < 1)
                return $"{timeSpan.TotalSeconds:F1}s";
            else if (timeSpan.TotalHours < 1)
                return $"{timeSpan.TotalMinutes:F1}min";
            else
                return $"{timeSpan.TotalHours:F1}h";
        }

        /// <summary>
        /// 格式化文件大小
        /// </summary>
        public static string ToFileSizeString(this long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB", "TB" };
            int order = 0;
            double size = bytes;

            while (size >= 1024 && order < sizes.Length - 1)
            {
                order++;
                size /= 1024;
            }

            return $"{size:F1} {sizes[order]}";
        }
    }

    #endregion
}