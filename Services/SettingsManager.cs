using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelEfficiencyAssistant.Models;

namespace ExcelEfficiencyAssistant.Services
{
    /// <summary>
    /// 设置管理器 - 负责应用程序设置的加载、保存和管理
    /// </summary>
    public static class SettingsManager
    {
        #region 字段

        private static ApplicationSettings _currentSettings;
        private static UserPreferences _currentPreferences;
        private static readonly string _settingsFilePath;
        private static readonly string _preferencesFilePath;
        private static readonly object _lockObject = new object();

        #endregion

        #region 构造函数和初始化

        static SettingsManager()
        {
            // 初始化文件路径
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var appFolder = Path.Combine(appDataPath, "ExcelEfficiencyAssistant");

            // 确保目录存在
            if (!Directory.Exists(appFolder))
            {
                Directory.CreateDirectory(appFolder);
            }

            _settingsFilePath = Path.Combine(appFolder, "settings.json");
            _preferencesFilePath = Path.Combine(appFolder, "preferences.json");
        }

        /// <summary>
        /// 初始化设置管理器
        /// </summary>
        public static void Initialize()
        {
            lock (_lockObject)
            {
                LoadSettings();
                LoadPreferences();
            }
        }

        #endregion

        #region 属性

        /// <summary>
        /// 当前应用程序设置
        /// </summary>
        public static ApplicationSettings CurrentSettings
        {
            get
            {
                if (_currentSettings == null)
                {
                    lock (_lockObject)
                    {
                        if (_currentSettings == null)
                        {
                            _currentSettings = GetDefaultSettings();
                        }
                    }
                }
                return _currentSettings;
            }
        }

        /// <summary>
        /// 当前用户偏好设置
        /// </summary>
        public static UserPreferences CurrentPreferences
        {
            get
            {
                if (_currentPreferences == null)
                {
                    lock (_lockObject)
                    {
                        if (_currentPreferences == null)
                        {
                            _currentPreferences = GetDefaultPreferences();
                        }
                    }
                }
                return _currentPreferences;
            }
        }

        #endregion

        #region 设置管理

        /// <summary>
        /// 获取默认应用程序设置
        /// </summary>
        private static ApplicationSettings GetDefaultSettings()
        {
            return new ApplicationSettings
            {
                EnablePerformanceOptimization = true,
                MaxProcessingRows = 10000,
                BatchSize = 1000,
                ShowProgressIndicator = true,
                ProgressUpdateInterval = 100,
                ShowTooltips = true,
                EnableAnimations = true,
                Theme = "default",
                AccentColor = System.Drawing.Color.FromArgb(0, 120, 212),
                AutoBackup = true,
                BackupCount = 5,
                SkipEmptyCells = true,
                PreserveFormatting = false,
                EnableLogging = true,
                LogLevel = LogLevel.Information,
                LogPath = "",
                MaxLogFileSize = 10,
                MaxLogFiles = 10,
                ConfirmBeforeExecution = true,
                EnableUndo = true,
                MaxUndoOperations = 50,
                ShowWelcomeMessage = true,
                ShowExcelAlerts = false,
                ShowStatusBarInfo = true,
                EnableSmartDoubleClick = false,
                ShowTaskPane = true
            };
        }

        /// <summary>
        /// 获取默认用户偏好设置
        /// </summary>
        private static UserPreferences GetDefaultPreferences()
        {
            return new UserPreferences
            {
                DefaultLanguage = "zh-CN",
                EnableSounds = false,
                AutoSave = true,
                AutoSaveInterval = 5,
                RecentFiles = new System.Collections.Generic.List<string>(),
                CustomPreferences = new System.Collections.Generic.Dictionary<string, object>()
            };
        }

        /// <summary>
        /// 加载应用程序设置
        /// </summary>
        private static void LoadSettings()
        {
            try
            {
                if (File.Exists(_settingsFilePath))
                {
                    var json = File.ReadAllText(_settingsFilePath);
                    var options = new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true,
                        Converters = { new JsonStringEnumConverter() }
                    };

                    var loadedSettings = JsonSerializer.Deserialize<ApplicationSettings>(json, options);
                    if (loadedSettings != null)
                    {
                        _currentSettings = loadedSettings;
                        LogService.Info("应用程序设置加载成功");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                LogService.Error("加载应用程序设置失败", ex);
            }

            // 如果加载失败，使用默认设置
            _currentSettings = GetDefaultSettings();
            LogService.Info("使用默认应用程序设置");
        }

        /// <summary>
        /// 加载用户偏好设置
        /// </summary>
        private static void LoadPreferences()
        {
            try
            {
                if (File.Exists(_preferencesFilePath))
                {
                    var json = File.ReadAllText(_preferencesFilePath);
                    var options = new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true,
                        Converters = { new JsonStringEnumConverter() }
                    };

                    var loadedPreferences = JsonSerializer.Deserialize<UserPreferences>(json, options);
                    if (loadedPreferences != null)
                    {
                        _currentPreferences = loadedPreferences;
                        LogService.Info("用户偏好设置加载成功");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                LogService.Error("加载用户偏好设置失败", ex);
            }

            // 如果加载失败，使用默认设置
            _currentPreferences = GetDefaultPreferences();
            LogService.Info("使用默认用户偏好设置");
        }

        /// <summary>
        /// 保存应用程序设置
        /// </summary>
        public static void SaveSettings()
        {
            SaveSettings(_currentSettings);
        }

        /// <summary>
        /// 保存指定的应用程序设置
        /// </summary>
        public static void SaveSettings(ApplicationSettings settings)
        {
            if (settings == null) return;

            try
            {
                lock (_lockObject)
                {
                    var options = new JsonSerializerOptions
                    {
                        WriteIndented = true,
                        PropertyNameCaseInsensitive = true,
                        Converters = { new JsonStringEnumConverter() }
                    };

                    var json = JsonSerializer.Serialize(settings, options);
                    File.WriteAllText(_settingsFilePath, json);

                    _currentSettings = settings;
                    LogService.Info("应用程序设置保存成功");
                }
            }
            catch (Exception ex)
            {
                LogService.Error("保存应用程序设置失败", ex);
            }
        }

        /// <summary>
        /// 保存用户偏好设置
        /// </summary>
        public static void SavePreferences()
        {
            SavePreferences(_currentPreferences);
        }

        /// <summary>
        /// 保存指定的用户偏好设置
        /// </summary>
        public static void SavePreferences(UserPreferences preferences)
        {
            if (preferences == null) return;

            try
            {
                lock (_lockObject)
                {
                    var options = new JsonSerializerOptions
                    {
                        WriteIndented = true,
                        PropertyNameCaseInsensitive = true,
                        Converters = { new JsonStringEnumConverter() }
                    };

                    var json = JsonSerializer.Serialize(preferences, options);
                    File.WriteAllText(_preferencesFilePath, json);

                    _currentPreferences = preferences;
                    LogService.Info("用户偏好设置保存成功");
                }
            }
            catch (Exception ex)
            {
                LogService.Error("保存用户偏好设置失败", ex);
            }
        }

        /// <summary>
        /// 重置为默认设置
        /// </summary>
        public static void ResetToDefaults()
        {
            try
            {
                lock (_lockObject)
                {
                    _currentSettings = GetDefaultSettings();
                    _currentPreferences = GetDefaultPreferences();

                    SaveSettings(_currentSettings);
                    SavePreferences(_currentPreferences);

                    LogService.Info("设置已重置为默认值");
                }
            }
            catch (Exception ex)
            {
                LogService.Error("重置设置失败", ex);
            }
        }

        /// <summary>
        /// 导出设置到指定文件
        /// </summary>
        public static bool ExportSettings(string filePath)
        {
            try
            {
                var exportData = new
                {
                    Settings = _currentSettings,
                    Preferences = _currentPreferences,
                    ExportTime = DateTime.Now,
                    Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "1.0.0"
                };

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNameCaseInsensitive = true,
                    Converters = { new JsonStringEnumConverter() }
                };

                var json = JsonSerializer.Serialize(exportData, options);
                File.WriteAllText(filePath, json);

                LogService.Info($"设置已导出到: {filePath}");
                return true;
            }
            catch (Exception ex)
            {
                LogService.Error("导出设置失败", ex);
                return false;
            }
        }

        /// <summary>
        /// 从指定文件导入设置
        /// </summary>
        public static bool ImportSettings(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    LogService.Warning($"导入文件不存在: {filePath}");
                    return false;
                }

                var json = File.ReadAllText(filePath);
                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                    Converters = { new JsonStringEnumConverter() }
                };

                using (var document = JsonDocument.Parse(json))
                {
                    var root = document.RootElement;

                    // 导入应用程序设置
                    if (root.TryGetProperty("Settings", out var settingsElement))
                    {
                        var settings = JsonSerializer.Deserialize<ApplicationSettings>(settingsElement.GetRawText(), options);
                        if (settings != null)
                        {
                            _currentSettings = settings;
                        }
                    }

                    // 导入用户偏好设置
                    if (root.TryGetProperty("Preferences", out var preferencesElement))
                    {
                        var preferences = JsonSerializer.Deserialize<UserPreferences>(preferencesElement.GetRawText(), options);
                        if (preferences != null)
                        {
                            _currentPreferences = preferences;
                        }
                    }
                }

                LogService.Info($"设置已从文件导入: {filePath}");
                return true;
            }
            catch (Exception ex)
            {
                LogService.Error("导入设置失败", ex);
                return false;
            }
        }

        #endregion

        #region 设置访问器

        /// <summary>
        /// 更新设置项
        /// </summary>
        public static void UpdateSetting(Action<ApplicationSettings> updateAction)
        {
            if (updateAction == null) return;

            try
            {
                lock (_lockObject)
                {
                    updateAction(_currentSettings);
                    SaveSettings(_currentSettings);
                }
            }
            catch (Exception ex)
            {
                LogService.Error("更新设置失败", ex);
            }
        }

        /// <summary>
        /// 更新偏好设置项
        /// </summary>
        public static void UpdatePreference(Action<UserPreferences> updateAction)
        {
            if (updateAction == null) return;

            try
            {
                lock (_lockObject)
                {
                    updateAction(_currentPreferences);
                    SavePreferences(_currentPreferences);
                }
            }
            catch (Exception ex)
            {
                LogService.Error("更新偏好设置失败", ex);
            }
        }

        /// <summary>
        /// 获取自定义偏好设置
        /// </summary>
        public static T GetCustomPreference<T>(string key, T defaultValue = default)
        {
            try
            {
                if (_currentPreferences.CustomPreferences.TryGetValue(key, out var value))
                {
                    if (value is T typedValue)
                    {
                        return typedValue;
                    }

                    // 尝试转换类型
                    try
                    {
                        return (T)Convert.ChangeType(value, typeof(T));
                    }
                    catch
                    {
                        return defaultValue;
                    }
                }
            }
            catch (Exception ex)
            {
                LogService.Error($"获取自定义偏好设置失败: {key}", ex);
            }

            return defaultValue;
        }

        /// <summary>
        /// 设置自定义偏好设置
        /// </summary>
        public static void SetCustomPreference<T>(string key, T value)
        {
            try
            {
                lock (_lockObject)
                {
                    _currentPreferences.CustomPreferences[key] = value;
                    SavePreferences(_currentPreferences);
                }
            }
            catch (Exception ex)
            {
                LogService.Error($"设置自定义偏好设置失败: {key}", ex);
            }
        }

        #endregion

        #region 验证和诊断

        /// <summary>
        /// 验证设置的有效性
        /// </summary>
        public static ValidationResult ValidateSettings()
        {
            var result = new ValidationResult { IsValid = true };

            try
            {
                var settings = CurrentSettings;

                // 验证数值范围
                if (settings.MaxProcessingRows <= 0)
                {
                    result.IsValid = false;
                    result.Errors.Add(new ValidationError
                    {
                        PropertyName = nameof(settings.MaxProcessingRows),
                        ErrorMessage = "最大处理行数必须大于0",
                        AttemptedValue = settings.MaxProcessingRows
                    });
                }

                if (settings.BatchSize <= 0)
                {
                    result.IsValid = false;
                    result.Errors.Add(new ValidationError
                    {
                        PropertyName = nameof(settings.BatchSize),
                        ErrorMessage = "批处理大小必须大于0",
                        AttemptedValue = settings.BatchSize
                    });
                }

                if (settings.ProgressUpdateInterval <= 0)
                {
                    result.IsValid = false;
                    result.Errors.Add(new ValidationError
                    {
                        PropertyName = nameof(settings.ProgressUpdateInterval),
                        ErrorMessage = "进度更新间隔必须大于0",
                        AttemptedValue = settings.ProgressUpdateInterval
                    });
                }

                // 验证日志设置
                if (settings.EnableLogging && settings.MaxLogFileSize <= 0)
                {
                    result.IsValid = false;
                    result.Errors.Add(new ValidationError
                    {
                        PropertyName = nameof(settings.MaxLogFileSize),
                        ErrorMessage = "启用日志时，最大日志文件大小必须大于0",
                        AttemptedValue = settings.MaxLogFileSize
                    });
                }

                // 验证备份设置
                if (settings.AutoBackup && settings.BackupCount <= 0)
                {
                    result.IsValid = false;
                    result.Errors.Add(new ValidationError
                    {
                        PropertyName = nameof(settings.BackupCount),
                        ErrorMessage = "启用自动备份时，备份数量必须大于0",
                        AttemptedValue = settings.BackupCount
                    });
                }

                // 验证撤销设置
                if (settings.EnableUndo && settings.MaxUndoOperations <= 0)
                {
                    result.IsValid = false;
                    result.Errors.Add(new ValidationError
                    {
                        PropertyName = nameof(settings.MaxUndoOperations),
                        ErrorMessage = "启用撤销功能时，最大撤销操作数必须大于0",
                        AttemptedValue = settings.MaxUndoOperations
                    });
                }

                // 添加警告
                if (settings.MaxProcessingRows > 100000)
                {
                    result.Warnings.Add(new ValidationWarning
                    {
                        PropertyName = nameof(settings.MaxProcessingRows),
                        WarningMessage = "最大处理行数较大，可能影响性能",
                        Value = settings.MaxProcessingRows
                    });
                }

                if (settings.AutoSaveInterval < 1)
                {
                    result.Warnings.Add(new ValidationWarning
                    {
                        PropertyName = nameof(CurrentPreferences.AutoSaveInterval),
                        WarningMessage = "自动保存间隔过短，可能影响性能",
                        Value = CurrentPreferences.AutoSaveInterval
                    });
                }
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.Errors.Add(new ValidationError
                {
                    PropertyName = "Settings",
                    ErrorMessage = "验证设置时发生异常",
                    AttemptedValue = ex.Message
                });
            }

            return result;
        }

        /// <summary>
        /// 获取设置诊断信息
        /// </summary>
        public static string GetDiagnosticInfo()
        {
            try
            {
                var info = new System.Text.StringBuilder();
                info.AppendLine("=== Excel效率助手设置诊断信息 ===");
                info.AppendLine($"诊断时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                info.AppendLine($"设置文件路径: {_settingsFilePath}");
                info.AppendLine($"偏好设置文件路径: {_preferencesFilePath}");
                info.AppendLine();

                // 文件信息
                info.AppendLine("=== 文件状态 ===");
                info.AppendLine($"设置文件存在: {File.Exists(_settingsFilePath)}");
                if (File.Exists(_settingsFilePath))
                {
                    var fileInfo = new FileInfo(_settingsFilePath);
                    info.AppendLine($"设置文件大小: {fileInfo.Length} 字节");
                    info.AppendLine($"设置文件修改时间: {fileInfo.LastWriteTime:yyyy-MM-dd HH:mm:ss}");
                }

                info.AppendLine($"偏好设置文件存在: {File.Exists(_preferencesFilePath)}");
                if (File.Exists(_preferencesFilePath))
                {
                    var fileInfo = new FileInfo(_preferencesFilePath);
                    info.AppendLine($"偏好设置文件大小: {fileInfo.Length} 字节");
                    info.AppendLine($"偏好设置文件修改时间: {fileInfo.LastWriteTime:yyyy-MM-dd HH:mm:ss}");
                }

                info.AppendLine();

                // 当前设置摘要
                info.AppendLine("=== 当前设置摘要 ===");
                var settings = CurrentSettings;
                info.AppendLine($"性能优化: {(settings.EnablePerformanceOptimization ? "启用" : "禁用")}");
                info.AppendLine($"最大处理行数: {settings.MaxProcessingRows:N0}");
                info.AppendLine($"批处理大小: {settings.BatchSize:N0}");
                info.AppendLine($"显示进度指示器: {(settings.ShowProgressIndicator ? "启用" : "禁用")}");
                info.AppendLine($"启用日志: {(settings.EnableLogging ? "启用" : "禁用")}");
                info.AppendLine($"日志级别: {settings.LogLevel}");
                info.AppendLine($"自动备份: {(settings.AutoBackup ? "启用" : "禁用")}");
                info.AppendLine($"显示欢迎消息: {(settings.ShowWelcomeMessage ? "启用" : "禁用")}");

                info.AppendLine();

                // 验证结果
                info.AppendLine("=== 设置验证 ===");
                var validation = ValidateSettings();
                info.AppendLine($"设置有效: {(validation.IsValid ? "是" : "否")}");
                info.AppendLine($"错误数量: {validation.Errors.Count}");
                info.AppendLine($"警告数量: {validation.Warnings.Count}");

                if (validation.HasErrors)
                {
                    info.AppendLine("错误详情:");
                    foreach (var error in validation.Errors)
                    {
                        info.AppendLine($"  - {error.PropertyName}: {error.ErrorMessage}");
                    }
                }

                if (validation.HasWarnings)
                {
                    info.AppendLine("警告详情:");
                    foreach (var warning in validation.Warnings)
                    {
                        info.AppendLine($"  - {warning.PropertyName}: {warning.WarningMessage}");
                    }
                }

                return info.ToString();
            }
            catch (Exception ex)
            {
                return $"获取诊断信息失败: {ex.Message}";
            }
        }

        #endregion
    }
}