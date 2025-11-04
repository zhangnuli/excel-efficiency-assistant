using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace ExcelEfficiencyAssistant
{
    /// <summary>
    /// Excel效率助手主插件类
    /// 插件的入口点，负责初始化、事件处理和生命周期管理
    /// </summary>
    public partial class ThisAddIn
    {
        #region 字段和属性

        private EfficiencyRibbon _ribbon;
        private Excel.Application _application;
        private bool _isInitialized = false;

        /// <summary>
        /// 获取Excel应用程序实例
        /// </summary>
        public Excel.Application Application => _application;

        #endregion

        #region 插件生命周期事件

        /// <summary>
        /// 插件启动事件
        /// </summary>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                InitializePlugin();
            }
            catch (Exception ex)
            {
                LogError("插件启动失败", ex);
                ShowStartupError(ex);
            }
        }

        /// <summary>
        /// 插件关闭事件
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                CleanupPlugin();
            }
            catch (Exception ex)
            {
                LogError("插件关闭时出错", ex);
            }
        }

        #endregion

        #region 初始化和清理

        /// <summary>
        /// 初始化插件
        /// </summary>
        private void InitializePlugin()
        {
            if (_isInitialized) return;

            try
            {
                // 获取Excel应用程序实例
                _application = this.Application;

                // 初始化设置管理器
                SettingsManager.Initialize();

                // 初始化日志服务
                LogService.Initialize();

                // 订阅Excel应用程序级别事件
                SubscribeToApplicationEvents();

                // 创建自定义任务窗格
                CreateTaskPanes();

                // 标记为已初始化
                _isInitialized = true;

                LogInfo("Excel效率助手插件启动成功");

                // 显示欢迎消息（仅在首次启动时）
                if (SettingsManager.CurrentSettings.ShowWelcomeMessage)
                {
                    ShowWelcomeMessage();
                }
            }
            catch (Exception ex)
            {
                LogError("插件初始化失败", ex);
                throw;
            }
        }

        /// <summary>
        /// 清理插件资源
        /// </summary>
        private void CleanupPlugin()
        {
            try
            {
                // 取消事件订阅
                UnsubscribeFromApplicationEvents();

                // 保存设置
                SettingsManager.SaveSettings();

                // 清理任务窗格
                DisposeTaskPanes();

                // 清理日志服务
                LogService.Cleanup();

                _isInitialized = false;

                LogInfo("Excel效率助手插件已关闭");
            }
            catch (Exception ex)
            {
                LogError("插件清理时出错", ex);
            }
        }

        #endregion

        #region Excel事件处理

        /// <summary>
        /// 订阅Excel应用程序事件
        /// </summary>
        private void SubscribeToApplicationEvents()
        {
            try
            {
                // 工作簿事件
                _application.WorkbookOpen += Application_WorkbookOpen;
                _application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
                _application.NewWorkbook += Application_NewWorkbook;

                // 工作表事件
                _application.SheetSelectionChange += Application_SheetSelectionChange;
                _application.SheetBeforeDoubleClick += Application_SheetBeforeDoubleClick;
                _application.SheetBeforeRightClick += Application_SheetBeforeRightClick;

                // 应用程序事件
                _application.WindowActivate += Application_WindowActivate;
                _application.WindowDeactivate += Application_WindowDeactivate;

                LogInfo("已订阅Excel应用程序事件");
            }
            catch (Exception ex)
            {
                LogError("订阅Excel事件失败", ex);
            }
        }

        /// <summary>
        /// 取消订阅Excel应用程序事件
        /// </summary>
        private void UnsubscribeFromApplicationEvents()
        {
            try
            {
                if (_application != null)
                {
                    // 工作簿事件
                    _application.WorkbookOpen -= Application_WorkbookOpen;
                    _application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
                    _application.NewWorkbook -= Application_NewWorkbook;

                    // 工作表事件
                    _application.SheetSelectionChange -= Application_SheetSelectionChange;
                    _application.SheetBeforeDoubleClick -= Application_SheetBeforeDoubleClick;
                    _application.SheetBeforeRightClick -= Application_SheetBeforeRightClick;

                    // 应用程序事件
                    _application.WindowActivate -= Application_WindowActivate;
                    _application.WindowDeactivate -= Application_WindowDeactivate;

                    LogInfo("已取消订阅Excel应用程序事件");
                }
            }
            catch (Exception ex)
            {
                LogError("取消订阅Excel事件失败", ex);
            }
        }

        #region 事件处理程序

        /// <summary>
        /// 工作簿打开事件
        /// </summary>
        private void Application_WorkbookOpen(Excel.Workbook workbook)
        {
            try
            {
                LogInfo($"工作簿已打开: {workbook.Name}");

                // 检查工作簿是否需要特殊处理
                CheckWorkbookForSpecialHandling(workbook);

                // 更新最近使用的文件列表
                UpdateRecentFiles(workbook.FullName);
            }
            catch (Exception ex)
            {
                LogError("处理工作簿打开事件失败", ex);
            }
        }

        /// <summary>
        /// 工作簿关闭前事件
        /// </summary>
        private void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            try
            {
                LogInfo($"工作簿即将关闭: {workbook.Name}");

                // 如果工作簿有未保存的更改，提示用户
                if (!workbook.Saved)
                {
                    var result = MessageBox.Show(
                        $"工作簿 '{workbook.Name}' 有未保存的更改，是否保存？",
                        "Excel效率助手",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question);

                    switch (result)
                    {
                        case DialogResult.Yes:
                            workbook.Save();
                            break;
                        case DialogResult.No:
                            workbook.Saved = true; // 跳过保存提示
                            break;
                        case DialogResult.Cancel:
                            cancel = true;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("处理工作簿关闭前事件失败", ex);
            }
        }

        /// <summary>
        /// 新工作簿事件
        /// </summary>
        private void Application_NewWorkbook(Excel.Workbook workbook)
        {
            try
            {
                LogInfo("创建了新工作簿");

                // 为新工作簿应用默认设置
                ApplyDefaultSettingsToWorkbook(workbook);
            }
            catch (Exception ex)
            {
                LogError("处理新工作簿事件失败", ex);
            }
        }

        /// <summary>
        /// 选择变化事件
        /// </summary>
        private void Application_SheetSelectionChange(object sheet, Excel.Range target)
        {
            try
            {
                // 可以在这里更新状态栏或任务窗格
                UpdateStatusBarInfo(target);
            }
            catch (Exception ex)
            {
                LogError("处理选择变化事件失败", ex);
            }
        }

        /// <summary>
        /// 双击事件
        /// </summary>
        private void Application_SheetBeforeDoubleClick(object sheet, Excel.Range target, ref bool cancel)
        {
            try
            {
                // 如果启用了智能双击功能
                if (SettingsManager.CurrentSettings.EnableSmartDoubleClick)
                {
                    HandleSmartDoubleClick(target, ref cancel);
                }
            }
            catch (Exception ex)
            {
                LogError("处理双击事件失败", ex);
            }
        }

        /// <summary>
        /// 右键事件
        /// </summary>
        private void Application_SheetBeforeRightClick(object sheet, Excel.Range target, ref bool cancel)
        {
            try
            {
                // 可以在这里扩展右键菜单功能
                LogDebug($"右键点击: {target.Address}");
            }
            catch (Exception ex)
            {
                LogError("处理右键事件失败", ex);
            }
        }

        /// <summary>
        /// 窗口激活事件
        /// </summary>
        private void Application_WindowActivate(Excel.Workbook workbook, Excel.Window window)
        {
            try
            {
                LogInfo($"窗口已激活: {workbook.Name}");

                // 更新任务窗格状态
                UpdateTaskPanesState(workbook);
            }
            catch (Exception ex)
            {
                LogError("处理窗口激活事件失败", ex);
            }
        }

        /// <summary>
        /// 窗口失活事件
        /// </summary>
        private void Application_WindowDeactivate(Excel.Workbook workbook, Excel.Window window)
        {
            try
            {
                LogDebug($"窗口失活: {workbook.Name}");
            }
            catch (Exception ex)
            {
                LogError("处理窗口失活事件失败", ex);
            }
        }

        #endregion

        #endregion

        #region 任务窗格管理

        /// <summary>
        /// 创建自定义任务窗格
        /// </summary>
        private void CreateTaskPanes()
        {
            try
            {
                // 创建效率助手任务窗格
                var efficiencyPane = new UI.TaskPanes.EfficiencyTaskPane();
                var customTaskPane = this.CustomTaskPanes.Add(efficiencyPane, "Excel效率助手");
                customTaskPane.Visible = SettingsManager.CurrentSettings.ShowTaskPane;
                customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                customTaskPane.Width = 300;

                LogInfo("任务窗格创建成功");
            }
            catch (Exception ex)
            {
                LogError("创建任务窗格失败", ex);
            }
        }

        /// <summary>
        /// 清理任务窗格
        /// </summary>
        private void DisposeTaskPanes()
        {
            try
            {
                foreach (Microsoft.Office.Tools.CustomTaskPane pane in this.CustomTaskPanes)
                {
                    if (pane.Control != null)
                    {
                        pane.Control.Dispose();
                    }
                }

                this.CustomTaskPanes.Clear();
                LogInfo("任务窗格已清理");
            }
            catch (Exception ex)
            {
                LogError("清理任务窗格失败", ex);
            }
        }

        /// <summary>
        /// 更新任务窗格状态
        /// </summary>
        private void UpdateTaskPanesState(Excel.Workbook workbook)
        {
            try
            {
                // 根据工作簿状态更新任务窗格内容
                foreach (Microsoft.Office.Tools.CustomTaskPane pane in this.CustomTaskPanes)
                {
                    if (pane.Control is UI.TaskPanes.EfficiencyTaskPane efficiencyPane)
                    {
                        efficiencyPane.UpdateWorkbookInfo(workbook);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("更新任务窗格状态失败", ex);
            }
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 检查工作簿是否需要特殊处理
        /// </summary>
        private void CheckWorkbookForSpecialHandling(Excel.Workbook workbook)
        {
            try
            {
                // 检查是否是特定类型的文件
                var fileName = workbook.Name.ToLowerInvariant();

                if (fileName.Contains("report") || fileName.Contains("报告"))
                {
                    LogInfo("检测到报告文件，应用报告优化设置");
                    // 可以应用特定于报告的设置
                }

                if (fileName.Contains("data") || fileName.Contains("数据"))
                {
                    LogInfo("检测到数据文件，启用数据分析功能");
                    // 可以启用数据分析相关的功能
                }
            }
            catch (Exception ex)
            {
                LogError("检查工作簿特殊处理失败", ex);
            }
        }

        /// <summary>
        /// 应用默认设置到工作簿
        /// </summary>
        private void ApplyDefaultSettingsToWorkbook(Excel.Workbook workbook)
        {
            try
            {
                // 设置默认计算模式
                _application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

                // 设置默认显示选项
                _application.DisplayAlerts = SettingsManager.CurrentSettings.ShowExcelAlerts;

                LogInfo("已应用默认设置到新工作簿");
            }
            catch (Exception ex)
            {
                LogError("应用默认设置失败", ex);
            }
        }

        /// <summary>
        /// 更新状态栏信息
        /// </summary>
        private void UpdateStatusBarInfo(Excel.Range target)
        {
            try
            {
                if (target != null && SettingsManager.CurrentSettings.ShowStatusBarInfo)
                {
                    var info = $"选中区域: {target.Rows.Count} 行 × {target.Columns.Count} 列";
                    _application.StatusBar = $"Excel效率助手 | {info}";
                }
            }
            catch (Exception ex)
            {
                LogError("更新状态栏信息失败", ex);
            }
        }

        /// <summary>
        /// 处理智能双击
        /// </summary>
        private void HandleSmartDoubleClick(Excel.Range target, ref bool cancel)
        {
            try
            {
                // 示例：双击单元格时自动应用格式
                if (target != null && target.Cells.Count == 1)
                {
                    var value = target.Value2;
                    if (value != null && IsEmailAddress(value.ToString()))
                    {
                        // 如果是邮箱地址，可以创建邮件链接
                        target.Hyperlinks.Add(target, $"mailto:{value}", Type.Missing, Type.Missing, Type.Missing);
                        cancel = true; // 取消默认的双击行为
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("处理智能双击失败", ex);
            }
        }

        /// <summary>
        /// 检查是否是邮箱地址
        /// </summary>
        private bool IsEmailAddress(string text)
        {
            try
            {
                return System.Text.RegularExpressions.Regex.IsMatch(
                    text,
                    @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$");
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 更新最近使用的文件列表
        /// </summary>
        private void UpdateRecentFiles(string fullPath)
        {
            try
            {
                if (!string.IsNullOrEmpty(fullPath))
                {
                    var recentFiles = SettingsManager.CurrentPreferences.RecentFiles;
                    recentFiles.Remove(fullPath);
                    recentFiles.Insert(0, fullPath);

                    // 保留最近10个文件
                    while (recentFiles.Count > 10)
                    {
                        recentFiles.RemoveAt(recentFiles.Count - 1);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("更新最近文件列表失败", ex);
            }
        }

        /// <summary>
        /// 显示欢迎消息
        /// </summary>
        private void ShowWelcomeMessage()
        {
            try
            {
                var result = MessageBox.Show(
                    "欢迎使用 Excel效率助手！\n\n" +
                    "这是一个专为提高您的工作效率而设计的Excel插件。\n\n" +
                    "主要功能：\n" +
                    "• 🔗 智能数据匹配\n" +
                    "• 🎨 专业表格美化\n" +
                    "• 📝 批量文本处理\n\n" +
                    "是否查看新手指南？",
                    "Excel效率助手",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);

                if (result == DialogResult.Yes)
                {
                    // 打开新手指南
                    var guideDialog = new UI.Dialogs.BeginnerGuideDialog();
                    guideDialog.ShowDialog();
                }

                // 下次不再显示
                SettingsManager.CurrentSettings.ShowWelcomeMessage = false;
            }
            catch (Exception ex)
            {
                LogError("显示欢迎消息失败", ex);
            }
        }

        /// <summary>
        /// 显示启动错误
        /// </summary>
        private void ShowStartupError(Exception ex)
        {
            try
            {
                MessageBox.Show(
                    $"Excel效率助手启动失败：\n\n{ex.Message}\n\n" +
                    "请检查Excel版本兼容性或联系技术支持。",
                    "启动错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch
            {
                // 如果连消息框都无法显示，只能记录到系统日志
                System.Diagnostics.EventLog.WriteEntry(
                    "Excel效率助手",
                    $"插件启动失败: {ex}",
                    System.Diagnostics.EventLogEntryType.Error);
            }
        }

        #endregion

        #region 日志方法

        /// <summary>
        /// 记录信息日志
        /// </summary>
        private void LogInfo(string message)
        {
            try
            {
                LogService.Info($"[ThisAddIn] {message}");
            }
            catch
            {
                // 忽略日志错误
            }
        }

        /// <summary>
        /// 记录调试日志
        /// </summary>
        private void LogDebug(string message)
        {
            try
            {
                LogService.Debug($"[ThisAddIn] {message}");
            }
            catch
            {
                // 忽略日志错误
            }
        }

        /// <summary>
        /// 记录错误日志
        /// </summary>
        private void LogError(string message, Exception ex)
        {
            try
            {
                LogService.Error($"[ThisAddIn] {message}", ex);
            }
            catch
            {
                // 忽略日志错误
            }
        }

        #endregion

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}