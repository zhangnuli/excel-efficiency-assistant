using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

// ReSharper disable CheckNamespace
namespace ExcelEfficiencyAssistant
{
    /// <summary>
    /// Excel效率助手功能区界面类
    /// 处理所有Ribbon按钮的事件响应和UI状态管理
    /// </summary>
    [ComVisible(true)]
    public class EfficiencyRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbonUI;
        private readonly Dictionary<string, bool> _controlStates;
        private Excel.Application _application;

        #region 构造函数和初始化

        public EfficiencyRibbon()
        {
            _controlStates = new Dictionary<string, bool>();
            InitializeControlStates();
        }

        /// <summary>
        /// 初始化控件状态
        /// </summary>
        private void InitializeControlStates()
        {
            // 默认启用的控件
            var enabledControls = new[]
            {
                "btnSmartMatch", "btnBatchVLookup", "btnMatchSettings",
                "btnSmartBeautify", "menuTemplateGallery", "btnQuickTools",
                "menuCaseConversion", "menuSpaceHandling", "menuSmartExtraction", "menuBatchOperations",
                "btnBeginnerGuide", "btnSettings", "btnAbout"
            };

            foreach (var control in enabledControls)
            {
                _controlStates[control] = true;
            }
        }

        #endregion

        #region IRibbonExtensibility 成员

        /// <summary>
        /// 加载Ribbon XML
        /// </summary>
        public string GetCustomUI(string ribbonID)
        {
            try
            {
                return GetResourceText("ExcelEfficiencyAssistant.Ribbon.EfficiencyRibbon.xml");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"加载Ribbon XML失败: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// Ribbon加载事件
        /// </summary>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbonUI = ribbonUI;

            // 获取Excel应用程序实例
            _application = Globals.ThisAddIn.Application;

            // 订阅Excel事件
            SubscribeToExcelEvents();
        }

        #endregion

        #region 回调方法 - 控件可见性

        /// <summary>
        /// 获取上下文选项卡的可见性
        /// </summary>
        public bool GetContextualTabVisible(Office.IRibbonControl control)
        {
            try
            {
                // 检查是否有选定的区域
                var selection = _application.Selection as Excel.Range;
                return selection != null && selection.Cells.Count > 1;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 获取控件的启用状态
        /// </summary>
        public bool GetControlEnabled(Office.IRibbonControl control)
        {
            return _controlStates.GetValueOrDefault(control.Id, true);
        }

        #endregion

        #region 数据匹配功能

        /// <summary>
        /// 智能匹配
        /// </summary>
        public async void OnSmartMatch(Office.IRibbonControl control)
        {
            try
            {
                var selection = GetSelectedRange();
                if (selection == null)
                {
                    ShowMessage("请先选择要匹配的数据区域", "提示");
                    return;
                }

                // 显示智能匹配对话框
                var matchDialog = new UI.Dialogs.DataMatcherDialog(_application, selection);
                var dialogResult = matchDialog.ShowDialog();

                if (dialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    // 执行数据匹配
                    var matcher = new Core.DataMatcher.DataMatcherEngine(_application);
                    var result = await matcher.SmartMatchAsync(selection, matchDialog.MatchOptions);

                    if (result.Success)
                    {
                        ShowMessage($"数据匹配完成！成功匹配 {result.MatchedRows} 行数据", "匹配成功");
                    }
                    else
                    {
                        ShowMessage($"数据匹配失败：{result.ErrorMessage}", "匹配失败");
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("智能匹配失败", ex);
            }
        }

        /// <summary>
        /// 批量VLOOKUP
        /// </summary>
        public async void OnBatchVLookup(Office.IRibbonControl control)
        {
            try
            {
                var selection = GetSelectedRange();
                if (selection == null)
                {
                    ShowMessage("请先选择要执行VLOOKUP的数据区域", "提示");
                    return;
                }

                // 显示VLOOKUP对话框
                var vlookupDialog = new UI.Dialogs.VLookupDialog(_application, selection);
                var dialogResult = vlookupDialog.ShowDialog();

                if (dialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    // 执行批量VLOOKUP
                    var matcher = new Core.DataMatcher.DataMatcherEngine(_application);
                    var result = await matcher.BatchVLookupAsync(
                        vlookupDialog.LookupRange,
                        vlookupDialog.TableArray,
                        vlookupDialog.ColumnIndex,
                        vlookupDialog.ExactMatch);

                    if (result.Success)
                    {
                        ShowMessage($"批量VLOOKUP完成！处理 {result.TotalRows} 行，成功匹配 {result.MatchedRows} 行", "VLOOKUP成功");
                    }
                    else
                    {
                        ShowMessage($"批量VLOOKUP失败：{result.ErrorMessage}", "VLOOKUP失败");
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("批量VLOOKUP失败", ex);
            }
        }

        /// <summary>
        /// 匹配设置
        /// </summary>
        public void OnMatchSettings(Office.IRibbonControl control)
        {
            try
            {
                var settingsDialog = new UI.Dialogs.MatchSettingsDialog();
                settingsDialog.ShowDialog();
            }
            catch (Exception ex)
            {
                ShowError("打开匹配设置失败", ex);
            }
        }

        #endregion

        #region 表格美化功能

        /// <summary>
        /// 智能美化
        /// </summary>
        public async void OnSmartBeautify(Office.IRibbonControl control)
        {
            try
            {
                var selection = GetSelectedRange();
                if (selection == null)
                {
                    ShowMessage("请先选择要美化的表格区域", "提示");
                    return;
                }

                // 执行智能美化
                var beautifier = new Core.Beautifier.TableBeautifier(_application);
                var result = await beautifier.SmartBeautifyAsync(selection);

                if (result.Success)
                {
                    ShowMessage($"表格美化完成！使用模板：{result.TemplateApplied}\n{result.RecommendedReason}", "美化成功");
                }
                else
                {
                    ShowMessage($"表格美化失败：{result.ErrorMessage}", "美化失败");
                }
            }
            catch (Exception ex)
            {
                ShowError("智能美化失败", ex);
            }
        }

        /// <summary>
        /// 应用模板
        /// </summary>
        public async void OnApplyTemplate(Office.IRibbonControl control)
        {
            try
            {
                var selection = GetSelectedRange();
                if (selection == null)
                {
                    ShowMessage("请先选择要应用模板的表格区域", "提示");
                    return;
                }

                var templateName = control.Tag?.ToString();
                if (string.IsNullOrEmpty(templateName))
                {
                    ShowMessage("模板信息错误", "错误");
                    return;
                }

                // 应用指定模板
                var beautifier = new Core.Beautifier.TableBeautifier(_application);
                var result = await beautifier.ApplyTemplateAsync(selection, templateName);

                if (result.Success)
                {
                    ShowMessage($"模板应用成功！处理了 {result.ProcessedCells} 个单元格", "美化成功");
                }
                else
                {
                    ShowMessage($"模板应用失败：{result.ErrorMessage}", "美化失败");
                }
            }
            catch (Exception ex)
            {
                ShowError("应用模板失败", ex);
            }
        }

        /// <summary>
        /// 快速美化工具
        /// </summary>
        public async void OnQuickTools(Office.IRibbonControl control)
        {
            try
            {
                var selection = GetSelectedRange();
                if (selection == null)
                {
                    ShowMessage("请先选择要处理的数据区域", "提示");
                    return;
                }

                // 显示快速工具对话框
                var quickToolsDialog = new UI.Dialogs.QuickBeautifyDialog();
                var dialogResult = quickToolsDialog.ShowDialog();

                if (dialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    var beautifier = new Core.Beautifier.TableBeautifier(_application);
                    var result = await beautifier.QuickBeautifyAsync(selection, quickToolsDialog.SelectedTool);

                    if (result.Success)
                    {
                        ShowMessage($"快速美化完成！处理了 {result.ProcessedCells} 个单元格", "美化成功");
                    }
                    else
                    {
                        ShowMessage($"快速美化失败：{result.ErrorMessage}", "美化失败");
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("快速美化工具失败", ex);
            }
        }

        #endregion

        #region 文本处理功能

        /// <summary>
        /// 文本处理
        /// </summary>
        public async void OnTextProcess(Office.IRibbonControl control)
        {
            try
            {
                var selection = GetSelectedRange();
                if (selection == null)
                {
                    ShowMessage("请先选择要处理的文本数据", "提示");
                    return;
                }

                var operationTag = control.Tag?.ToString();
                if (string.IsNullOrEmpty(operationTag))
                {
                    ShowMessage("操作类型错误", "错误");
                    return;
                }

                // 解析操作类型
                if (Enum.TryParse<Core.TextProcessor.TextOperation>(operationTag, out var operation))
                {
                    var processor = new Core.TextProcessor.TextProcessor(_application);
                    var result = await processor.ProcessTextAsync(selection, operation);

                    if (result.Success)
                    {
                        ShowMessage($"文本处理完成！处理了 {result.ProcessedCells} 个单元格", "处理成功");
                    }
                    else
                    {
                        ShowMessage($"文本处理失败：{result.ErrorMessage}", "处理失败");
                    }
                }
                else
                {
                    ShowMessage("不支持的操作类型", "错误");
                }
            }
            catch (Exception ex)
            {
                ShowError("文本处理失败", ex);
            }
        }

        /// <summary>
        /// 添加前缀
        /// </summary>
        public async void OnAddPrefix(Office.IRibbonControl control)
        {
            try
            {
                var selection = GetSelectedRange();
                if (selection == null) return;

                var prefixDialog = new UI.Dialogs.TextInputDialog("请输入要添加的前缀：", "添加前缀");
                if (prefixDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var processor = new Core.TextProcessor.TextProcessor(_application);
                    var result = await processor.AddAffixAsync(selection, prefixDialog.InputText);

                    if (result.Success)
                    {
                        ShowMessage($"前缀添加完成！处理了 {result.ProcessedCells} 个单元格", "添加成功");
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("添加前缀失败", ex);
            }
        }

        /// <summary>
        /// 添加后缀
        /// </summary>
        public async void OnAddSuffix(Office.IRibbonControl control)
        {
            try
            {
                var selection = GetSelectedRange();
                if (selection == null) return;

                var suffixDialog = new UI.Dialogs.TextInputDialog("请输入要添加的后缀：", "添加后缀");
                if (suffixDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var processor = new Core.TextProcessor.TextProcessor(_application);
                    var result = await processor.AddAffixAsync(selection, "", suffixDialog.InputText);

                    if (result.Success)
                    {
                        ShowMessage($"后缀添加完成！处理了 {result.ProcessedCells} 个单元格", "添加成功");
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("添加后缀失败", ex);
            }
        }

        /// <summary>
        /// 批量替换
        /// </summary>
        public async void OnBatchReplace(Office.IRibbonControl control)
        {
            try
            {
                var selection = GetSelectedRange();
                if (selection == null) return;

                var replaceDialog = new UI.Dialogs.ReplaceDialog();
                if (replaceDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var processor = new Core.TextProcessor.TextProcessor(_application);
                    var result = await processor.BatchReplaceAsync(
                        selection,
                        replaceDialog.FindText,
                        replaceDialog.ReplaceText,
                        replaceDialog.MatchCase,
                        replaceDialog.WholeWord);

                    if (result.Success)
                    {
                        ShowMessage($"批量替换完成！处理了 {result.ProcessedCells} 个单元格", "替换成功");
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("批量替换失败", ex);
            }
        }

        /// <summary>
        /// 拆分列
        /// </summary>
        public async void OnSplitColumn(Office.IRibbonControl control)
        {
            try
            {
                var selection = GetSelectedRange();
                if (selection == null) return;

                var splitDialog = new UI.Dialogs.SplitColumnDialog();
                if (splitDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var processor = new Core.TextProcessor.TextProcessor(_application);
                    var result = await processor.SplitColumnAsync(selection, splitDialog.Delimiter, splitDialog.SplitToNewColumns);

                    if (result.Success)
                    {
                        ShowMessage($"列拆分完成！最大列数：{result.MaxColumns}", "拆分成功");
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("拆分列失败", ex);
            }
        }

        /// <summary>
        /// 合并列
        /// </summary>
        public async void OnMergeColumns(Office.IRibbonControl control)
        {
            try
            {
                var selection = GetSelectedRange();
                if (selection == null) return;

                var mergeDialog = new UI.Dialogs.MergeColumnsDialog();
                if (mergeDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var processor = new Core.TextProcessor.TextProcessor(_application);
                    var result = await processor.MergeColumnsAsync(selection, mergeDialog.Delimiter);

                    if (result.Success)
                    {
                        ShowMessage($"列合并完成！处理了 {result.ProcessedCells} 个单元格", "合并成功");
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("合并列失败", ex);
            }
        }

        #endregion

        #region 帮助和设置

        /// <summary>
        /// 新手指南
        /// </summary>
        public void OnBeginnerGuide(Office.IRibbonControl control)
        {
            try
            {
                var guideDialog = new UI.Dialogs.BeginnerGuideDialog();
                guideDialog.ShowDialog();
            }
            catch (Exception ex)
            {
                ShowError("打开新手指南失败", ex);
            }
        }

        /// <summary>
        /// 设置
        /// </summary>
        public void OnSettings(Office.IRibbonControl control)
        {
            try
            {
                var settingsDialog = new UI.Dialogs.SettingsDialog();
                var result = settingsDialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    // 保存设置
                    SettingsManager.SaveSettings(settingsDialog.CurrentSettings);
                    ShowMessage("设置已保存", "设置");
                }
            }
            catch (Exception ex)
            {
                ShowError("打开设置失败", ex);
            }
        }

        /// <summary>
        /// 关于
        /// </summary>
        public void OnAbout(Office.IRibbonControl control)
        {
            try
            {
                var aboutDialog = new UI.Dialogs.AboutDialog();
                aboutDialog.ShowDialog();
            }
            catch (Exception ex)
            {
                ShowError("打开关于对话框失败", ex);
            }
        }

        #endregion

        #region 上下文菜单事件

        /// <summary>
        /// 上下文匹配
        /// </summary>
        public async void OnContextMatch(Office.IRibbonControl control)
        {
            await OnSmartMatch(control);
        }

        /// <summary>
        /// 上下文美化
        /// </summary>
        public async void OnContextFormat(Office.IRibbonControl control)
        {
            await OnSmartBeautify(control);
        }

        /// <summary>
        /// 上下文文本处理
        /// </summary>
        public async void OnContextText(Office.IRibbonControl control)
        {
            // 显示文本操作选择对话框
            var textDialog = new UI.Dialogs.TextOperationDialog();
            if (textDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var processor = new Core.TextProcessor.TextProcessor(_application);
                var selection = GetSelectedRange();
                var result = await processor.ProcessTextAsync(selection, textDialog.SelectedOperation);

                if (result.Success)
                {
                    ShowMessage($"文本处理完成！", "处理成功");
                }
            }
        }

        /// <summary>
        /// 上下文文本操作
        /// </summary>
        public async void OnContextTextAction(Office.IRibbonControl control)
        {
            await OnTextProcess(control);
        }

        /// <summary>
        /// 上下文拆分列
        /// </summary>
        public async void OnContextSplitColumn(Office.IRibbonControl control)
        {
            await OnSplitColumn(control);
        }

        /// <summary>
        /// 上下文合并列
        /// </summary>
        public async void OnContextMergeColumns(Office.IRibbonControl control)
        {
            await OnMergeColumns(control);
        }

        #endregion

        #region Backstage 事件

        /// <summary>
        /// 新手教程
        /// </summary>
        public void OnBackstageTutorial(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Process.Start("https://docs.excel-assistant.com/tutorial");
            }
            catch (Exception ex)
            {
                ShowError("无法打开教程页面", ex);
            }
        }

        /// <summary>
        /// 示例文件
        /// </summary>
        public void OnBackstageSamples(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Process.Start("https://docs.excel-assistant.com/samples");
            }
            catch (Exception ex)
            {
                ShowError("无法打开示例文件页面", ex);
            }
        }

        /// <summary>
        /// 模板下载
        /// </summary>
        public void OnBackstageTemplates(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Process.Start("https://templates.excel-assistant.com");
            }
            catch (Exception ex)
            {
                ShowError("无法打开模板下载页面", ex);
            }
        }

        /// <summary>
        /// 使用文档
        /// </summary>
        public void OnBackstageDocumentation(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Process.Start("https://docs.excel-assistant.com");
            }
            catch (Exception ex)
            {
                ShowError("无法打开文档页面", ex);
            }
        }

        /// <summary>
        /// 视频教程
        /// </summary>
        public void OnBackstageVideos(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Process.Start("https://videos.excel-assistant.com");
            }
            catch (Exception ex)
            {
                ShowError("无法打开视频教程页面", ex);
            }
        }

        /// <summary>
        /// 常见问题
        /// </summary>
        public void OnBackstageFAQ(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Process.Start("https://faq.excel-assistant.com");
            }
            catch (Exception ex)
            {
                ShowError("无法打开常见问题页面", ex);
            }
        }

        /// <summary>
        /// 联系我们
        /// </summary>
        public void OnBackstageContact(Office.IRibbonControl control)
        {
            try
            {
                var contactDialog = new UI.Dialogs.ContactDialog();
                contactDialog.ShowDialog();
            }
            catch (Exception ex)
            {
                ShowError("打开联系方式失败", ex);
            }
        }

        /// <summary>
        /// 意见反馈
        /// </summary>
        public void OnBackstageFeedback(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Process.Start("https://feedback.excel-assistant.com");
            }
            catch (Exception ex)
            {
                ShowError("无法打开反馈页面", ex);
            }
        }

        /// <summary>
        /// 检查更新
        /// </summary>
        public void OnBackstageUpdate(Office.IRibbonControl control)
        {
            try
            {
                // TODO: 实现更新检查逻辑
                ShowMessage("当前已是最新版本", "检查更新");
            }
            catch (Exception ex)
            {
                ShowError("检查更新失败", ex);
            }
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 获取选定区域
        /// </summary>
        private Excel.Range GetSelectedRange()
        {
            try
            {
                return _application.Selection as Excel.Range;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 订阅Excel事件
        /// </summary>
        private void SubscribeToExcelEvents()
        {
            try
            {
                _application.SheetSelectionChange += Application_SheetSelectionChange;
                _application.SheetActivate += Application_SheetActivate;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"订阅Excel事件失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 选择变化事件
        /// </summary>
        private void Application_SheetSelectionChange(object sheet, Excel.Range target)
        {
            try
            {
                // 更新上下文选项卡状态
                _ribbonUI.InvalidateControl("tabContextualTools");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"处理选择变化事件失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 工作表激活事件
        /// </summary>
        private void Application_SheetActivate(object sheet)
        {
            try
            {
                // 更新所有控件状态
                _ribbonUI.Invalidate();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"处理工作表激活事件失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 显示消息框
        /// </summary>
        private void ShowMessage(string message, string title)
        {
            System.Windows.Forms.MessageBox.Show(message, title, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

        /// <summary>
        /// 显示错误消息
        /// </summary>
        private void ShowError(string message, Exception ex)
        {
            var fullMessage = $"{message}\n\n详细信息：{ex.Message}";
            System.Windows.Forms.MessageBox.Show(fullMessage, "错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
        }

        /// <summary>
        /// 获取资源文本
        /// </summary>
        private static string GetResourceText(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    using (var reader = new System.IO.StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            return "";
        }

        #endregion

        #region 快速操作（上下文选项卡）

        /// <summary>
        /// 快速匹配
        /// </summary>
        public async void OnQuickMatch(Office.IRibbonControl control)
        {
            await OnSmartMatch(control);
        }

        /// <summary>
        /// 快速美化
        /// </summary>
        public async void OnQuickFormat(Office.IRibbonControl control)
        {
            await OnSmartBeautify(control);
        }

        /// <summary>
        /// 快速文本
        /// </summary>
        public async void OnQuickText(Office.IRibbonControl control)
        {
            await OnContextText(control);
        }

        /// <summary>
        /// 快速美化
        /// </summary>
        public async void OnQuickBeautify(Office.IRibbonControl control)
        {
            await OnQuickTools(control);
        }

        #endregion
    }
}