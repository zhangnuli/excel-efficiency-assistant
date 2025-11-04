using System;
using System.IO;
using System.Text;
using ExcelEfficiencyAssistant.Models;

namespace ExcelEfficiencyAssistant.Services
{
    /// <summary>
    /// 日志服务 - 负责应用程序的日志记录和管理
    /// </summary>
    public static class LogService
    {
        #region 字段

        private static readonly string _logDirectory;
        private static readonly string _currentLogFile;
        private static readonly object _lockObject = new object();
        private static bool _isInitialized = false;

        #endregion

        #region 构造函数和初始化

        static LogService()
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            _logDirectory = Path.Combine(appDataPath, "ExcelEfficiencyAssistant", "Logs");

            // 确保日志目录存在
            if (!Directory.Exists(_logDirectory))
            {
                try
                {
                    Directory.CreateDirectory(_logDirectory);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"创建日志目录失败: {ex.Message}");
                }
            }

            // 当前日志文件路径
            var today = DateTime.Now.ToString("yyyy-MM-dd");
            _currentLogFile = Path.Combine(_logDirectory, $"ExcelEfficiencyAssistant_{today}.log");
        }

        /// <summary>
        /// 初始化日志服务
        /// </summary>
        public static void Initialize()
        {
            lock (_lockObject)
            {
                if (_isInitialized) return;

                try
                {
                    // 清理旧日志文件
                    CleanupOldLogFiles();

                    // 记录启动日志
                    Info("Excel效率助手日志服务初始化完成");

                    _isInitialized = true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"日志服务初始化失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 清理日志服务
        /// </summary>
        public static void Cleanup()
        {
            lock (_lockObject)
            {
                try
                {
                    if (_isInitialized)
                    {
                        Info("Excel效率助手日志服务正在关闭");
                        _isInitialized = false;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"日志服务清理失败: {ex.Message}");
                }
            }
        }

        #endregion

        #region 日志记录方法

        /// <summary>
        /// 记录跟踪日志
        /// </summary>
        public static void Trace(string message)
        {
            Log(LogLevel.Trace, message, null);
        }

        /// <summary>
        /// 记录跟踪日志
        /// </summary>
        public static void Trace(string message, Exception exception)
        {
            Log(LogLevel.Trace, message, exception);
        }

        /// <summary>
        /// 记录调试日志
        /// </summary>
        public static void Debug(string message)
        {
            Log(LogLevel.Debug, message, null);
        }

        /// <summary>
        /// 记录调试日志
        /// </summary>
        public static void Debug(string message, Exception exception)
        {
            Log(LogLevel.Debug, message, exception);
        }

        /// <summary>
        /// 记录信息日志
        /// </summary>
        public static void Info(string message)
        {
            Log(LogLevel.Information, message, null);
        }

        /// <summary>
        /// 记录信息日志
        /// </summary>
        public static void Info(string message, Exception exception)
        {
            Log(LogLevel.Information, message, exception);
        }

        /// <summary>
        /// 记录警告日志
        /// </summary>
        public static void Warning(string message)
        {
            Log(LogLevel.Warning, message, null);
        }

        /// <summary>
        /// 记录警告日志
        /// </summary>
        public static void Warning(string message, Exception exception)
        {
            Log(LogLevel.Warning, message, exception);
        }

        /// <summary>
        /// 记录错误日志
        /// </summary>
        public static void Error(string message)
        {
            Log(LogLevel.Error, message, null);
        }

        /// <summary>
        /// 记录错误日志
        /// </summary>
        public static void Error(string message, Exception exception)
        {
            Log(LogLevel.Error, message, exception);
        }

        /// <summary>
        /// 记录严重错误日志
        /// </summary>
        public static void Critical(string message)
        {
            Log(LogLevel.Critical, message, null);
        }

        /// <summary>
        /// 记录严重错误日志
        /// </summary>
        public static void Critical(string message, Exception exception)
        {
            Log(LogLevel.Critical, message, exception);
        }

        #endregion

        #region 核心日志方法

        /// <summary>
        /// 记录日志
        /// </summary>
        private static void Log(LogLevel level, string message, Exception exception)
        {
            try
            {
                // 检查是否启用日志
                if (!SettingsManager.CurrentSettings.EnableLogging)
                    return;

                // 检查日志级别
                if (level < SettingsManager.CurrentSettings.LogLevel)
                    return;

                // 生成日志条目
                var logEntry = CreateLogEntry(level, message, exception);

                // 写入文件
                WriteToFile(logEntry);

                // 输出到调试窗口
                System.Diagnostics.Debug.WriteLine(logEntry);

                // 如果是严重错误，同时写入Windows事件日志
                if (level >= LogLevel.Critical)
                {
                    WriteToEventLog(level, message, exception);
                }
            }
            catch (Exception ex)
            {
                // 日志记录失败时，输出到调试窗口
                System.Diagnostics.Debug.WriteLine($"日志记录失败: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"原始日志: [{level}] {message}");
            }
        }

        /// <summary>
        /// 创建日志条目
        /// </summary>
        private static string CreateLogEntry(LogLevel level, string message, Exception exception)
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            var threadId = System.Threading.Thread.CurrentThread.ManagedThreadId;
            var levelStr = level.ToString().ToUpper();

            var entry = new StringBuilder();
            entry.Append($"[{timestamp}] [{levelStr}] [T{threadId:D3}] ");

            // 添加消息
            if (!string.IsNullOrEmpty(message))
            {
                entry.Append(message);
            }

            // 添加异常信息
            if (exception != null)
            {
                entry.Append($" - 异常: {exception.Message}");
                if (!string.IsNullOrEmpty(exception.StackTrace))
                {
                    entry.AppendLine();
                    entry.Append($"堆栈跟踪: {exception.StackTrace}");
                }

                // 添加内部异常
                var innerException = exception.InnerException;
                var innerLevel = 1;
                while (innerException != null && innerLevel <= 3)
                {
                    entry.AppendLine();
                    entry.Append($"内部异常 {innerLevel}: {innerException.Message}");
                    innerException = innerException.InnerException;
                    innerLevel++;
                }
            }

            return entry.ToString();
        }

        /// <summary>
        /// 写入日志文件
        /// </summary>
        private static void WriteToFile(string logEntry)
        {
            lock (_lockObject)
            {
                try
                {
                    // 检查是否需要创建新的日志文件（跨天）
                    CheckAndCreateNewLogFile();

                    // 检查文件大小，如果超过限制则轮转
                    if (File.Exists(_currentLogFile))
                    {
                        var fileInfo = new FileInfo(_currentLogFile);
                        var maxSizeMB = SettingsManager.CurrentSettings.MaxLogFileSize;

                        if (fileInfo.Length > maxSizeMB * 1024 * 1024)
                        {
                            RotateLogFile();
                        }
                    }

                    // 写入日志
                    File.AppendAllText(_currentLogFile, logEntry + Environment.NewLine, Encoding.UTF8);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"写入日志文件失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 检查并创建新的日志文件
        /// </summary>
        private static void CheckAndCreateNewLogFile()
        {
            var today = DateTime.Now.ToString("yyyy-MM-dd");
            var expectedLogFile = Path.Combine(_logDirectory, $"ExcelEfficiencyAssistant_{today}.log");

            if (!string.Equals(_currentLogFile, expectedLogFile, StringComparison.OrdinalIgnoreCase))
            {
                // 更新当前日志文件路径
                _currentLogFile = expectedLogFile;

                // 记录日志文件切换
                try
                {
                    var message = $"日志文件已切换到: {_currentLogFile}";
                    File.AppendAllText(_currentLogFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [INFO] {message}{Environment.NewLine}");
                }
                catch
                {
                    // 忽略文件切换日志的错误
                }
            }
        }

        /// <summary>
        /// 轮转日志文件
        /// </summary>
        private static void RotateLogFile()
        {
            try
            {
                var timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
                var rotatedFile = Path.Combine(_logDirectory, $"ExcelEfficiencyAssistant_{timestamp}_rotated.log");

                // 移动当前日志文件
                if (File.Exists(_currentLogFile))
                {
                    File.Move(_currentLogFile, rotatedFile);
                }

                // 记录轮转信息
                var message = $"日志文件已轮转: {rotatedFile}";
                File.AppendAllText(_currentLogFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [INFO] {message}{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"日志文件轮转失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 写入Windows事件日志
        /// </summary>
        private static void WriteToEventLog(LogLevel level, string message, Exception exception)
        {
            try
            {
                if (!EventLog.SourceExists("ExcelEfficiencyAssistant"))
                {
                    EventLog.CreateEventSource("ExcelEfficiencyAssistant", "Application");
                }

                var eventLogEntry = new StringBuilder();
                eventLogEntry.AppendLine("Excel效率助手 - 严重错误");
                eventLogEntry.AppendLine($"时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                eventLogEntry.AppendLine($"级别: {level}");
                eventLogEntry.AppendLine($"消息: {message}");

                if (exception != null)
                {
                    eventLogEntry.AppendLine($"异常: {exception.Message}");
                    eventLogEntry.AppendLine($"堆栈: {exception.StackTrace}");
                }

                var eventLogType = level switch
                {
                    LogLevel.Critical => System.Diagnostics.EventLogEntryType.Error,
                    LogLevel.Error => System.Diagnostics.EventLogEntryType.Error,
                    LogLevel.Warning => System.Diagnostics.EventLogEntryType.Warning,
                    _ => System.Diagnostics.EventLogEntryType.Information
                };

                EventLog.WriteEntry("ExcelEfficiencyAssistant", eventLogEntry.ToString(), eventLogType);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"写入Windows事件日志失败: {ex.Message}");
            }
        }

        #endregion

        #region 日志管理

        /// <summary>
        /// 清理旧日志文件
        /// </summary>
        private static void CleanupOldLogFiles()
        {
            try
            {
                if (!Directory.Exists(_logDirectory)) return;

                var maxFiles = SettingsManager.CurrentSettings.MaxLogFiles;
                var logFiles = Directory.GetFiles(_logDirectory, "ExcelEfficiencyAssistant_*.log")
                                       .OrderByDescending(f => f)
                                       .ToList();

                if (logFiles.Count > maxFiles)
                {
                    var filesToDelete = logFiles.Skip(maxFiles);
                    foreach (var file in filesToDelete)
                    {
                        try
                        {
                            File.Delete(file);
                            System.Diagnostics.Debug.WriteLine($"已删除旧日志文件: {file}");
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"删除日志文件失败 {file}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"清理旧日志文件失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取最近的日志条目
        /// </summary>
        public static string GetRecentLogs(int lineCount = 100)
        {
            try
            {
                if (!File.Exists(_currentLogFile)) return "没有找到日志文件";

                var lines = File.ReadAllLines(_currentLogFile, Encoding.UTF8);
                var skipCount = Math.Max(0, lines.Length - lineCount);
                var recentLines = lines.Skip(skipCount);

                return string.Join(Environment.NewLine, recentLines);
            }
            catch (Exception ex)
            {
                return $"读取日志失败: {ex.Message}";
            }
        }

        /// <summary>
        /// 获取指定日期的日志
        /// </summary>
        public static string GetLogsByDate(DateTime date)
        {
            try
            {
                var dateStr = date.ToString("yyyy-MM-dd");
                var logFile = Path.Combine(_logDirectory, $"ExcelEfficiencyAssistant_{dateStr}.log");

                if (!File.Exists(logFile))
                    return $"没有找到 {dateStr} 的日志文件";

                return File.ReadAllText(logFile, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                return $"读取日志失败: {ex.Message}";
            }
        }

        /// <summary>
        /// 导出日志到指定文件
        /// </summary>
        public static bool ExportLogs(string filePath, DateTime? startDate = null, DateTime? endDate = null)
        {
            try
            {
                var allLogs = new StringBuilder();

                if (startDate.HasValue && endDate.HasValue)
                {
                    // 导出指定日期范围的日志
                    for (var date = startDate.Value.Date; date <= endDate.Value.Date; date = date.AddDays(1))
                    {
                        var dailyLogs = GetLogsByDate(date);
                        if (!string.IsNullOrEmpty(dailyLogs) && dailyLogs != $"没有找到 {date:yyyy-MM-dd} 的日志文件")
                        {
                            allLogs.AppendLine($"=== {date:yyyy-MM-dd} ===");
                            allLogs.AppendLine(dailyLogs);
                            allLogs.AppendLine();
                        }
                    }
                }
                else
                {
                    // 导出所有日志文件
                    var logFiles = Directory.GetFiles(_logDirectory, "ExcelEfficiencyAssistant_*.log")
                                           .OrderBy(f => f);

                    foreach (var logFile in logFiles)
                    {
                        var fileName = Path.GetFileNameWithoutExtension(logFile);
                        var dateStr = fileName.Replace("ExcelEfficiencyAssistant_", "");
                        allLogs.AppendLine($"=== {dateStr} ===");
                        allLogs.AppendLine(File.ReadAllText(logFile, Encoding.UTF8));
                        allLogs.AppendLine();
                    }
                }

                File.WriteAllText(filePath, allLogs.ToString(), Encoding.UTF8);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"导出日志失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 获取日志统计信息
        /// </summary>
        public static LogStatistics GetStatistics()
        {
            var stats = new LogStatistics();

            try
            {
                if (!Directory.Exists(_logDirectory)) return stats;

                var logFiles = Directory.GetFiles(_logDirectory, "ExcelEfficiencyAssistant_*.log");
                stats.TotalFiles = logFiles.Length;

                long totalSize = 0;
                DateTime lastModified = DateTime.MinValue;
                int totalLines = 0;

                foreach (var logFile in logFiles)
                {
                    var fileInfo = new FileInfo(logFile);
                    totalSize += fileInfo.Length;
                    if (fileInfo.LastWriteTime > lastModified)
                    {
                        lastModified = fileInfo.LastWriteTime;
                    }

                    try
                    {
                        var lines = File.ReadAllLines(logFile, Encoding.UTF8);
                        totalLines += lines.Length;

                        // 统计各级别日志数量
                        foreach (var line in lines)
                        {
                            if (line.Contains("[TRACE]")) stats.TraceCount++;
                            else if (line.Contains("[DEBUG]")) stats.DebugCount++;
                            else if (line.Contains("[INFO]")) stats.InfoCount++;
                            else if (line.Contains("[WARNING]")) stats.WarningCount++;
                            else if (line.Contains("[ERROR]")) stats.ErrorCount++;
                            else if (line.Contains("[CRITICAL]")) stats.CriticalCount++;
                        }
                    }
                    catch
                    {
                        // 忽略单个文件的读取错误
                    }
                }

                stats.TotalSizeBytes = totalSize;
                stats.TotalSizeFormatted = FormatFileSize(totalSize);
                stats.LastModified = lastModified;
                stats.TotalLines = totalLines;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取日志统计失败: {ex.Message}");
            }

            return stats;
        }

        /// <summary>
        /// 格式化文件大小
        /// </summary>
        private static string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            int order = 0;
            double size = bytes;

            while (size >= 1024 && order < sizes.Length - 1)
            {
                order++;
                size /= 1024;
            }

            return $"{size:F1} {sizes[order]}";
        }

        #endregion
    }

    #region 日志统计类

    /// <summary>
    /// 日志统计信息
    /// </summary>
    public class LogStatistics
    {
        public int TotalFiles { get; set; }
        public long TotalSizeBytes { get; set; }
        public string TotalSizeFormatted { get; set; }
        public int TotalLines { get; set; }
        public DateTime LastModified { get; set; }
        public int TraceCount { get; set; }
        public int DebugCount { get; set; }
        public int InfoCount { get; set; }
        public int WarningCount { get; set; }
        public int ErrorCount { get; set; }
        public int CriticalCount { get; set; }

        public int TotalEntries => TraceCount + DebugCount + InfoCount + WarningCount + ErrorCount + CriticalCount;
    }

    #endregion
}