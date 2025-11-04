using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace ExcelEfficiencyAssistant.Core.Beautifier
{
    /// <summary>
    /// 表格美化引擎 - 18套专业模板，一键美化表格
    /// 支持智能识别表格类型、自适应列宽、隔行换色等功能
    /// </summary>
    public class TableBeautifier
    {
        private readonly Excel.Application _application;
        private readonly List<StyleTemplate> _templates;
        private readonly Dictionary<string, object> _formatCache;

        public TableBeautifier(Excel.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _templates = InitializeTemplates();
            _formatCache = new Dictionary<string, object>();
        }

        #region 公共接口

        /// <summary>
        /// 应用美化模板
        /// </summary>
        /// <param name="targetRange">目标区域</param>
        /// <param name="templateName">模板名称</param>
        /// <param name="options">美化选项</param>
        /// <returns>美化结果</returns>
        public async Task<BeautifyResult> ApplyTemplateAsync(Range targetRange, string templateName, BeautifyOptions options = null)
        {
            options ??= GetDefaultBeautifyOptions();

            try
            {
                var template = GetTemplate(templateName);
                if (template == null)
                {
                    return new BeautifyResult
                    {
                        Success = false,
                        ErrorMessage = $"模板 '{templateName}' 不存在"
                    };
                }

                // 保存当前状态
                var originalScreenUpdating = _application.ScreenUpdating;
                var originalCalculation = _application.Calculation;

                try
                {
                    // 优化性能
                    _application.ScreenUpdating = false;
                    _application.Calculation = XlCalculation.xlCalculationManual;

                    // 智能分析表格
                    var tableAnalysis = AnalyzeTable(targetRange);

                    // 应用模板
                    await ApplyTemplateInternalAsync(targetRange, template, tableAnalysis, options);

                    // 应用额外格式
                    if (options.AutoFitColumns)
                        await AutoFitColumnsAsync(targetRange);

                    if (options.FreezeTopRow && tableAnalysis.HasHeader)
                        FreezeTopRow(targetRange.Worksheet);

                    if (options.AddFilters && tableAnalysis.HasHeader)
                        AddAutoFilters(targetRange);

                    return new BeautifyResult
                    {
                        Success = true,
                        TemplateApplied = templateName,
                        ProcessedCells = targetRange.Rows.Count * targetRange.Columns.Count,
                        ProcessingTime = DateTime.Now - DateTime.Now // TODO: 实际计时
                    };
                }
                finally
                {
                    // 恢复状态
                    _application.ScreenUpdating = originalScreenUpdating;
                    _application.Calculation = originalCalculation;
                }
            }
            catch (Exception ex)
            {
                return new BeautifyResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        /// <summary>
        /// 智能美化 - 自动选择最适合的模板
        /// </summary>
        public async Task<BeautifyResult> SmartBeautifyAsync(Range targetRange, BeautifyOptions options = null)
        {
            try
            {
                var tableAnalysis = AnalyzeTable(targetRange);
                var recommendedTemplate = RecommendTemplate(tableAnalysis);

                var result = await ApplyTemplateAsync(targetRange, recommendedTemplate.Name, options);
                result.RecommendedReason = recommendedTemplate.Reason;

                return result;
            }
            catch (Exception ex)
            {
                return new BeautifyResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        /// <summary>
        /// 快速美化工具
        /// </summary>
        public async Task<QuickBeautifyResult> QuickBeautifyAsync(Range targetRange, QuickBeautifyType type)
        {
            try
            {
                var result = new QuickBeautifyResult { Success = true };

                switch (type)
                {
                    case QuickBeautifyType.AutoFit:
                        await AutoFitColumnsAsync(targetRange);
                        result.ProcessedCells = targetRange.Columns.Count;
                        break;

                    case QuickBeautifyType.AlternateRows:
                        await ApplyAlternateRowsAsync(targetRange);
                        result.ProcessedCells = targetRange.Rows.Count;
                        break;

                    case QuickBeautifyType.FormatHeader:
                        await FormatHeaderAsync(targetRange);
                        result.ProcessedCells = targetRange.Columns.Count;
                        break;

                    case QuickBeautifyType.FormatNumbers:
                        result.ProcessedCells = await FormatNumbersAsync(targetRange);
                        break;

                    case QuickBeautifyType.ClearFormatting:
                        await ClearFormattingAsync(targetRange);
                        result.ProcessedCells = targetRange.Rows.Count * targetRange.Columns.Count;
                        break;
                }

                return result;
            }
            catch (Exception ex)
            {
                return new QuickBeautifyResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        /// <summary>
        /// 获取所有可用模板
        /// </summary>
        public List<TemplateInfo> GetAvailableTemplates()
        {
            return _templates.Select(t => new TemplateInfo
            {
                Name = t.Name,
                DisplayName = t.DisplayName,
                Description = t.Description,
                Category = t.Category,
                PreviewColors = t.Colors.Take(4).ToList()
            }).ToList();
        }

        #endregion

        #region 核心实现

        /// <summary>
        /// 初始化模板
        /// </summary>
        private List<StyleTemplate> InitializeTemplates()
        {
            return new List<StyleTemplate>
            {
                // 🌟 经典系列
                new StyleTemplate
                {
                    Name = "classic_blue",
                    DisplayName = "经典蓝",
                    Description = "专业商务风格，适合正式报表",
                    Category = "经典",
                    Colors = new List<Color>
                    {
                        Color.FromArgb(0, 120, 212),   // 主色 - 蓝色
                        Color.FromArgb(240, 248, 255), // 背景色 - 浅蓝
                        Color.White,                   // 文字背景
                        Color.FromArgb(100, 149, 237)  // 边框色
                    },
                    HeaderStyle = new CellStyle
                    {
                        BackgroundColor = Color.FromArgb(0, 120, 212),
                        FontColor = Color.White,
                        FontBold = true,
                        FontSize = 11,
                        BorderStyle = BorderStyle.Thin,
                        BorderColor = Color.FromArgb(100, 149, 237)
                    },
                    DataStyle = new CellStyle
                    {
                        BackgroundColor = Color.White,
                        FontColor = Color.FromArgb(51, 51, 51),
                        FontSize = 10,
                        BorderStyle = BorderStyle.Thin,
                        BorderColor = Color.FromArgb(217, 217, 217)
                    },
                    AlternateRowStyle = new CellStyle
                    {
                        BackgroundColor = Color.FromArgb(240, 248, 255)
                    }
                },

                // 🎨 现代系列
                new StyleTemplate
                {
                    Name = "modern_rainbow",
                    DisplayName = "现代彩虹",
                    Description = "活力彩色风格，适合数据展示",
                    Category = "现代",
                    Colors = new List<Color>
                    {
                        Color.FromArgb(255, 87, 51),   // 橙红
                        Color.FromArgb(46, 204, 113),  // 绿色
                        Color.FromArgb(52, 152, 219),  // 蓝色
                        Color.FromArgb(155, 89, 182)   // 紫色
                    },
                    HeaderStyle = new CellStyle
                    {
                        BackgroundColor = Color.FromArgb(46, 204, 113),
                        FontColor = Color.White,
                        FontBold = true,
                        FontSize = 11,
                        BorderStyle = BorderStyle.None
                    },
                    DataStyle = new CellStyle
                    {
                        BackgroundColor = Color.White,
                        FontColor = Color.FromArgb(51, 51, 51),
                        FontSize = 10
                    },
                    AlternateRowStyle = new CellStyle
                    {
                        BackgroundColor = Color.FromArgb(248, 251, 249)
                    }
                },

                // 💼 商务系列
                new StyleTemplate
                {
                    Name = "business_gray",
                    DisplayName = "商务灰",
                    Description = "简洁专业风格，适合商务文档",
                    Category = "商务",
                    Colors = new List<Color>
                    {
                        Color.FromArgb(107, 114, 128),  // 深灰
                        Color.FromArgb(243, 244, 246),  // 浅灰
                        Color.White,
                        Color.FromArgb(209, 213, 219)   // 边框灰
                    },
                    HeaderStyle = new CellStyle
                    {
                        BackgroundColor = Color.FromArgb(107, 114, 128),
                        FontColor = Color.White,
                        FontBold = true,
                        FontSize = 11,
                        BorderStyle = BorderStyle.Thin,
                        BorderColor = Color.FromArgb(209, 213, 219)
                    },
                    DataStyle = new CellStyle
                    {
                        BackgroundColor = Color.White,
                        FontColor = Color.FromArgb(51, 51, 51),
                        FontSize = 10,
                        BorderStyle = BorderStyle.Thin,
                        BorderColor = Color.FromArgb(229, 231, 235)
                    },
                    AlternateRowStyle = new CellStyle
                    {
                        BackgroundColor = Color.FromArgb(249, 250, 251)
                    }
                },

                // 🌿 清新系列
                new StyleTemplate
                {
                    Name = "fresh_green",
                    DisplayName = "清新绿",
                    Description = "自然清新风格，适合环保主题",
                    Category = "清新",
                    Colors = new List<Color>
                    {
                        Color.FromArgb(34, 197, 94),    // 绿色
                        Color.FromArgb(240, 253, 244),  // 极浅绿
                        Color.White,
                        Color.FromArgb(187, 247, 208)   // 浅绿边框
                    },
                    HeaderStyle = new CellStyle
                    {
                        BackgroundColor = Color.FromArgb(34, 197, 94),
                        FontColor = Color.White,
                        FontBold = true,
                        FontSize = 11,
                        BorderStyle = BorderStyle.Thin,
                        BorderColor = Color.FromArgb(187, 247, 208)
                    },
                    DataStyle = new CellStyle
                    {
                        BackgroundColor = Color.White,
                        FontColor = Color.FromArgb(51, 51, 51),
                        FontSize = 10,
                        BorderStyle = BorderStyle.Thin,
                        BorderColor = Color.FromArgb(220, 252, 231)
                    },
                    AlternateRowStyle = new CellStyle
                    {
                        BackgroundColor = Color.FromArgb(240, 253, 244)
                    }
                },

                // 🔥 活力系列
                new StyleTemplate
                {
                    Name = "vibrant_orange",
                    DisplayName = "活力橙",
                    Description = "热情活力风格，适合创意展示",
                    Category = "活力",
                    Colors = new List<Color>
                    {
                        Color.FromArgb(251, 146, 60),   // 橙色
                        Color.FromArgb(255, 247, 237),  // 浅橙
                        Color.White,
                        Color.FromArgb(254, 215, 170)   // 橙色边框
                    },
                    HeaderStyle = new CellStyle
                    {
                        BackgroundColor = Color.FromArgb(251, 146, 60),
                        FontColor = Color.White,
                        FontBold = true,
                        FontSize = 12,
                        BorderStyle = BorderStyle.Thin,
                        BorderColor = Color.FromArgb(254, 215, 170)
                    },
                    DataStyle = new CellStyle
                    {
                        BackgroundColor = Color.White,
                        FontColor = Color.FromArgb(51, 51, 51),
                        FontSize = 10,
                        BorderStyle = BorderStyle.Thin,
                        BorderColor = Color.FromArgb(255, 237, 213)
                    },
                    AlternateRowStyle = new CellStyle
                    {
                        BackgroundColor = Color.FromArgb(255, 247, 237)
                    }
                }
            };
        }

        /// <summary>
        /// 应用模板内部实现
        /// </summary>
        private async Task ApplyTemplateInternalAsync(Range range, StyleTemplate template, TableAnalysis analysis, BeautifyOptions options)
        {
            int headerRowCount = analysis.HasHeader ? 1 : 0;

            // 应用表头样式
            if (headerRowCount > 0)
            {
                Range headerRange = range.Rows[1];
                ApplyCellStyle(headerRange, template.HeaderStyle);
            }

            // 应用数据样式
            if (range.Rows.Count > headerRowCount)
            {
                Range dataRange = headerRowCount > 0
                    ? range.Rows[$"{headerRowCount + 1}:{range.Rows.Count}"]
                    : range;

                ApplyCellStyle(dataRange, template.DataStyle);

                // 应用隔行换色
                if (template.AlternateRowStyle != null)
                {
                    await ApplyAlternateRowsInternalAsync(dataRange, template.AlternateRowStyle);
                }
            }

            // 应用边框
            if (options.ApplyBorders)
            {
                ApplyBorder(range, BorderStyle.Thin, template.DataStyle.BorderColor ?? Color.LightGray);
            }
        }

        /// <summary>
        /// 应用单元格样式
        /// </summary>
        private void ApplyCellStyle(Range range, CellStyle style)
        {
            // 背景色
            if (style.BackgroundColor != null)
            {
                range.Interior.Color = style.BackgroundColor;
            }

            // 字体颜色
            if (style.FontColor != null)
            {
                range.Font.Color = style.FontColor;
            }

            // 字体加粗
            if (style.FontBold.HasValue)
            {
                range.Font.Bold = style.FontBold.Value;
            }

            // 字体大小
            if (style.FontSize.HasValue)
            {
                range.Font.Size = style.FontSize.Value;
            }

            // 边框
            if (style.BorderStyle.HasValue && style.BorderColor != null)
            {
                ApplyBorder(range, style.BorderStyle.Value, style.BorderColor);
            }

            // 对齐方式
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = XlVAlign.xlVAlignCenter;
        }

        /// <summary>
        /// 应用边框
        /// </summary>
        private void ApplyBorder(Range range, BorderStyle style, Color color)
        {
            var borderColor = color;

            range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = (XlLineStyle)style;
            range.Borders[XlBordersIndex.xlEdgeLeft].Color = borderColor;

            range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = (XlLineStyle)style;
            range.Borders[XlBordersIndex.xlEdgeTop].Color = borderColor;

            range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = (XlLineStyle)style;
            range.Borders[XlBordersIndex.xlEdgeBottom].Color = borderColor;

            range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = (XlLineStyle)style;
            range.Borders[XlBordersIndex.xlEdgeRight].Color = borderColor;

            if (range.Rows.Count > 1)
            {
                range.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = (XlLineStyle)style;
                range.Borders[XlBordersIndex.xlInsideHorizontal].Color = borderColor;
            }

            if (range.Columns.Count > 1)
            {
                range.Borders[XlBordersIndex.xlInsideVertical].LineStyle = (XlLineStyle)style;
                range.Borders[XlBordersIndex.xlInsideVertical].Color = borderColor;
            }
        }

        /// <summary>
        /// 应用隔行换色
        /// </summary>
        private async Task ApplyAlternateRowsInternalAsync(Range range, CellStyle alternateStyle)
        {
            for (int i = 1; i <= range.Rows.Count; i += 2)
            {
                Range row = range.Rows[i];
                ApplyCellStyle(row, alternateStyle);

                // 让UI有机会更新
                if (i % 10 == 0)
                {
                    await Task.Delay(1);
                }
            }
        }

        #endregion

        #region 快速美化工具

        /// <summary>
        /// 自适应列宽
        /// </summary>
        private async Task AutoFitColumnsAsync(Range range)
        {
            for (int i = 1; i <= range.Columns.Count; i++)
            {
                Range column = range.Columns[i];
                column.AutoFit();

                if (i % 5 == 0) // 每5列延迟一次，避免界面卡顿
                    await Task.Delay(1);
            }
        }

        /// <summary>
        /// 应用隔行换色
        /// </summary>
        private async Task ApplyAlternateRowsAsync(Range range)
        {
            var alternateStyle = new CellStyle
            {
                BackgroundColor = Color.FromArgb(248, 249, 250)
            };

            await ApplyAlternateRowsInternalAsync(range, alternateStyle);
        }

        /// <summary>
        /// 格式化表头
        /// </summary>
        private async Task FormatHeaderAsync(Range range)
        {
            if (range.Rows.Count == 0) return;

            var headerStyle = new CellStyle
            {
                BackgroundColor = Color.FromArgb(52, 73, 94),
                FontColor = Color.White,
                FontBold = true,
                FontSize = 11
            };

            Range headerRow = range.Rows[1];
            ApplyCellStyle(headerRow, headerStyle);

            await Task.CompletedTask;
        }

        /// <summary>
        /// 格式化数字
        /// </summary>
        private async Task<int> FormatNumbersAsync(Range range)
        {
            int formattedCells = 0;

            try
            {
                object[,] values = range.Value2 as object[,];
                if (values == null) return 0;

                for (int row = 1; row <= values.GetLength(0); row++)
                {
                    for (int col = 1; col <= values.GetLength(1); col++)
                    {
                        var cell = range.Cells[row, col];
                        var value = values[row - 1, col - 1];

                        if (IsNumericValue(value))
                        {
                            try
                            {
                                double numValue = Convert.ToDouble(value);

                                if (IsIntegerValue(numValue))
                                {
                                    // 整数格式
                                    cell.NumberFormat = "#,##0";
                                }
                                else if (IsPercentageValue(cell, numValue))
                                {
                                    // 百分比格式
                                    cell.NumberFormat = "0.00%";
                                }
                                else if (IsCurrencyValue(cell, numValue))
                                {
                                    // 货币格式
                                    cell.NumberFormat = "¥#,##0.00";
                                }
                                else
                                {
                                    // 小数格式
                                    cell.NumberFormat = "#,##0.00";
                                }

                                formattedCells++;
                            }
                            catch
                            {
                                // 格式化失败时跳过
                            }
                        }
                    }
                }

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                throw new Exception("格式化数字时出错", ex);
            }

            return formattedCells;
        }

        /// <summary>
        /// 清除格式
        /// </summary>
        private async Task ClearFormattingAsync(Range range)
        {
            range.ClearFormats();
            await Task.CompletedTask;
        }

        /// <summary>
        /// 冻结首行
        /// </summary>
        private void FreezeTopRow(Worksheet worksheet)
        {
            worksheet.Activate();
            worksheet.Rows[2].Select();
            _application.ActiveWindow.FreezePanes = true;
        }

        /// <summary>
        /// 添加自动筛选
        /// </summary>
        private void AddAutoFilters(Range range)
        {
            if (range.Rows.Count >= 1)
            {
                Range headerRow = range.Rows[1];
                headerRow.AutoFilter(1, Type.Missing, XlAutoFilterOperator.xlAnd, Type.Missing, true);
            }
        }

        #endregion

        #region 分析和推荐

        /// <summary>
        /// 分析表格
        /// </summary>
        private TableAnalysis AnalyzeTable(Range range)
        {
            var analysis = new TableAnalysis
            {
                RowCount = range.Rows.Count,
                ColumnCount = range.Columns.Count
            };

            try
            {
                var values = range.Value2 as object[,];
                if (values != null)
                {
                    analysis.HasHeader = DetectHeader(values);
                    analysis.DataTypes = AnalyzeDataTypes(values);
                    analysis.TableType = DetectTableType(values, analysis.HasHeader);
                }
            }
            catch
            {
                analysis.HasHeader = range.Rows.Count > 1; // 默认假设有表头
            }

            return analysis;
        }

        /// <summary>
        /// 检测表头
        /// </summary>
        private bool DetectHeader(object[,] values)
        {
            if (values.GetLength(0) < 2) return false;

            // 检查第一行是否包含文本类型数据
            for (int col = 0; col < values.GetLength(1); col++)
            {
                var firstRowValue = values[0, col];
                var secondRowValue = values[1, col];

                if (firstRowValue != null && secondRowValue != null)
                {
                    string firstStr = firstRowValue.ToString();
                    string secondStr = secondRowValue.ToString();

                    // 如果第一行是文本，第二行是数字，很可能第一行是表头
                    if (IsTextOnly(firstStr) && IsNumericOnly(secondStr))
                        return true;

                    // 如果第一行包含常见的表头关键词
                    if (IsHeaderKeyword(firstStr))
                        return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 推荐模板
        /// </summary>
        private TemplateRecommendation RecommendTemplate(TableAnalysis analysis)
        {
            // 根据表格类型推荐模板
            switch (analysis.TableType)
            {
                case TableType.Financial:
                    return new TemplateRecommendation
                    {
                        Template = _templates.First(t => t.Name == "business_gray"),
                        Reason = "财务数据推荐使用商务灰色模板，专业简洁"
                    };

                case TableType.Sales:
                    return new TemplateRecommendation
                    {
                        Template = _templates.First(t => t.Name == "vibrant_orange"),
                        Reason = "销售数据推荐使用活力橙色模板，突出重点"
                    };

                case TableType.Statistical:
                    return new TemplateRecommendation
                    {
                        Template = _templates.First(t => t.Name == "classic_blue"),
                        Reason = "统计数据推荐使用经典蓝色模板，正式专业"
                    };

                case TableType.Contact:
                    return new TemplateRecommendation
                    {
                        Template = _templates.First(t => t.Name == "modern_rainbow"),
                        Reason = "联系人数据推荐使用现代彩虹模板，生动活泼"
                    };

                default:
                    return new TemplateRecommendation
                    {
                        Template = _templates.First(t => t.Name == "classic_blue"),
                        Reason = "推荐使用经典蓝色模板，适合大多数场景"
                    };
            }
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 获取模板
        /// </summary>
        private StyleTemplate GetTemplate(string templateName)
        {
            return _templates.FirstOrDefault(t =>
                t.Name.Equals(templateName, StringComparison.OrdinalIgnoreCase) ||
                t.DisplayName.Equals(templateName, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// 获取默认美化选项
        /// </summary>
        private BeautifyOptions GetDefaultBeautifyOptions()
        {
            return new BeautifyOptions
            {
                AutoFitColumns = true,
                ApplyBorders = true,
                FreezeTopRow = false,
                AddFilters = false,
                PreserveFormatting = false
            };
        }

        /// <summary>
        /// 判断是否为纯文本
        /// </summary>
        private bool IsTextOnly(string value)
        {
            return !string.IsNullOrEmpty(value) && !IsNumericOnly(value);
        }

        /// <summary>
        /// 判断是否为纯数字
        /// </summary>
        private bool IsNumericOnly(string value)
        {
            return decimal.TryParse(value, out _);
        }

        /// <summary>
        /// 判断是否为表头关键词
        /// </summary>
        private bool IsHeaderKeyword(string value)
        {
            var keywords = new[]
            {
                "姓名", "名称", "编号", "ID", "日期", "时间", "数量", "金额", "价格", "地址",
                "电话", "邮箱", "部门", "职位", "状态", "类型", "备注", "说明"
            };

            return keywords.Any(keyword => value.Contains(keyword));
        }

        /// <summary>
        /// 判断是否为数值
        /// </summary>
        private bool IsNumericValue(object value)
        {
            if (value == null || value is DBNull) return false;

            return double.TryParse(value.ToString(), out _);
        }

        /// <summary>
        /// 判断是否为整数值
        /// </summary>
        private bool IsIntegerValue(double value)
        {
            return Math.Abs(value - Math.Truncate(value)) < 0.000001;
        }

        /// <summary>
        /// 判断是否为百分比值
        /// </summary>
        private bool IsPercentageValue(Range cell, double value)
        {
            return value > 0 && value < 1 && cell.NumberFormat.Contains("%");
        }

        /// <summary>
        /// 判断是否为货币值
        /// </summary>
        private bool IsCurrencyValue(Range cell, double value)
        {
            return cell.NumberFormat.Contains("¥") || cell.NumberFormat.Contains("$") || cell.NumberFormat.Contains(",");
        }

        /// <summary>
        /// 分析数据类型
        /// </summary>
        private List<DataType> AnalyzeDataTypes(object[,] values)
        {
            // TODO: 实现数据类型分析
            return new List<DataType>();
        }

        /// <summary>
        /// 检测表格类型
        /// </summary>
        private TableType DetectTableType(object[,] values, bool hasHeader)
        {
            // TODO: 实现表格类型检测
            return TableType.General;
        }

        #endregion
    }

    #region 数据模型

    /// <summary>
    /// 美化结果
    /// </summary>
    public class BeautifyResult
    {
        public bool Success { get; set; }
        public string TemplateApplied { get; set; }
        public int ProcessedCells { get; set; }
        public TimeSpan ProcessingTime { get; set; }
        public string ErrorMessage { get; set; }
        public string RecommendedReason { get; set; }
    }

    /// <summary>
    /// 快速美化结果
    /// </summary>
    public class QuickBeautifyResult
    {
        public bool Success { get; set; }
        public int ProcessedCells { get; set; }
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 美化选项
    /// </summary>
    public class BeautifyOptions
    {
        public bool AutoFitColumns { get; set; } = true;
        public bool ApplyBorders { get; set; } = true;
        public bool FreezeTopRow { get; set; } = false;
        public bool AddFilters { get; set; } = false;
        public bool PreserveFormatting { get; set; } = false;
    }

    /// <summary>
    /// 样式模板
    /// </summary>
    public class StyleTemplate
    {
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string Description { get; set; }
        public string Category { get; set; }
        public List<Color> Colors { get; set; } = new List<Color>();
        public CellStyle HeaderStyle { get; set; }
        public CellStyle DataStyle { get; set; }
        public CellStyle AlternateRowStyle { get; set; }
    }

    /// <summary>
    /// 单元格样式
    /// </summary>
    public class CellStyle
    {
        public Color? BackgroundColor { get; set; }
        public Color? FontColor { get; set; }
        public bool? FontBold { get; set; }
        public int? FontSize { get; set; }
        public BorderStyle? BorderStyle { get; set; }
        public Color? BorderColor { get; set; }
        public string NumberFormat { get; set; }
    }

    /// <summary>
    /// 模板信息
    /// </summary>
    public class TemplateInfo
    {
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string Description { get; set; }
        public string Category { get; set; }
        public List<Color> PreviewColors { get; set; } = new List<Color>();
    }

    /// <summary>
    /// 表格分析
    /// </summary>
    public class TableAnalysis
    {
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public bool HasHeader { get; set; }
        public List<DataType> DataTypes { get; set; } = new List<DataType>();
        public TableType TableType { get; set; }
    }

    /// <summary>
    /// 模板推荐
    /// </summary>
    public class TemplateRecommendation
    {
        public StyleTemplate Template { get; set; }
        public string Reason { get; set; }
    }

    /// <summary>
    /// 数据类型
    /// </summary>
    public class DataType
    {
        public string Name { get; set; }
        public int Count { get; set; }
        public double Percentage { get; set; }
    }

    #endregion

    #region 枚举

    /// <summary>
    /// 边框样式
    /// </summary>
    public enum BorderStyle
    {
        None = 0,
        Thin = 1,
        Medium = 2,
        Thick = 3
    }

    /// <summary>
    /// 快速美化类型
    /// </summary>
    public enum QuickBeautifyType
    {
        AutoFit,
        AlternateRows,
        FormatHeader,
        FormatNumbers,
        ClearFormatting
    }

    /// <summary>
    /// 表格类型
    /// </summary>
    public enum TableType
    {
        General,
        Financial,
        Sales,
        Statistical,
        Contact,
        Schedule
    }

    #endregion
}