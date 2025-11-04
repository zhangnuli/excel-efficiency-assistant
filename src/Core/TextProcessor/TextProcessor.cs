using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace ExcelEfficiencyAssistant.Core.TextProcessor
{
    /// <summary>
    /// 文本处理引擎 - 15种文本工具，批量处理数据
    /// 支持大小写转换、空格处理、智能提取、批量操作等功能
    /// </summary>
    public class TextProcessor
    {
        private readonly Excel.Application _application;
        private readonly Dictionary<TextOperation, Func<string, string>> _operationHandlers;
        private readonly Dictionary<string, Regex> _regexCache;

        public TextProcessor(Excel.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _operationHandlers = InitializeOperationHandlers();
            _regexCache = new Dictionary<string, Regex>();
        }

        #region 公共接口

        /// <summary>
        /// 执行文本处理操作
        /// </summary>
        /// <param name="targetRange">目标区域</param>
        /// <param name="operation">操作类型</param>
        /// <param name="options">处理选项</param>
        /// <returns>处理结果</returns>
        public async Task<TextProcessResult> ProcessTextAsync(Range targetRange, TextOperation operation, TextProcessOptions options = null)
        {
            options ??= GetDefaultOptions();

            try
            {
                var result = new TextProcessResult
                {
                    Operation = operation,
                    TotalCells = targetRange.Rows.Count * targetRange.Columns.Count,
                    ProcessedCells = 0,
                    SkippedCells = 0,
                    ErrorCells = 0
                };

                // 保存Excel状态
                var originalScreenUpdating = _application.ScreenUpdating;
                var originalCalculation = _application.Calculation;

                try
                {
                    _application.ScreenUpdating = false;
                    _application.Calculation = XlCalculation.xlCalculationManual;

                    // 批量读取数据
                    var values = targetRange.Value2 as object[,];
                    if (values == null) return result;

                    var processedValues = new object[values.GetLength(0), values.GetLength(1)];

                    // 处理每个单元格
                    for (int row = 0; row < values.GetLength(0); row++)
                    {
                        for (int col = 0; col < values.GetLength(1); col++)
                        {
                            try
                            {
                                var originalValue = values[row, col];
                                var processedValue = ProcessSingleCell(originalValue, operation, options);

                                processedValues[row, col] = processedValue;

                                if (!Equals(originalValue, processedValue))
                                {
                                    result.ProcessedCells++;
                                }
                                else if (originalValue != null)
                                {
                                    result.SkippedCells++;
                                }
                            }
                            catch (Exception ex)
                            {
                                processedValues[row, col] = values[row, col]; // 保持原值
                                result.ErrorCells++;
                                result.Errors.Add(new ProcessError
                                {
                                    Row = row + 1,
                                    Column = col + 1,
                                    OriginalValue = values[row, col]?.ToString(),
                                    Error = ex.Message
                                });
                            }
                        }

                        // 定期更新界面，避免Excel假死
                        if (row % 50 == 0)
                        {
                            await Task.Delay(1);
                        }
                    }

                    // 批量写入结果
                    targetRange.Value2 = processedValues;

                    result.Success = true;
                    result.ProcessingTime = DateTime.Now - DateTime.Now; // TODO: 实际计时

                    return result;
                }
                finally
                {
                    _application.ScreenUpdating = originalScreenUpdating;
                    _application.Calculation = originalCalculation;
                }
            }
            catch (Exception ex)
            {
                return new TextProcessResult
                {
                    Operation = operation,
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        /// <summary>
        /// 批量操作 - 添加前缀或后缀
        /// </summary>
        public async Task<TextProcessResult> AddAffixAsync(Range targetRange, string prefix = "", string suffix = "")
        {
            return await ProcessTextAsync(targetRange, TextOperation.Custom, new TextProcessOptions
            {
                CustomOperation = (text) => prefix + text + suffix,
                SkipEmptyCells = true
            });
        }

        /// <summary>
        /// 批量替换
        /// </summary>
        public async Task<TextProcessResult> BatchReplaceAsync(Range targetRange, string findText, string replaceText, bool matchCase = false, bool wholeWord = false)
        {
            return await ProcessTextAsync(targetRange, TextOperation.Replace, new TextProcessOptions
            {
                FindText = findText,
                ReplaceText = replaceText,
                MatchCase = matchCase,
                WholeWord = wholeWord
            });
        }

        /// <summary>
        /// 拆分列
        /// </summary>
        public async Task<SplitResult> SplitColumnAsync(Range targetRange, string delimiter, bool splitToNewColumns = true)
        {
            try
            {
                var result = new SplitResult { Success = true };

                var values = targetRange.Value2 as object[,];
                if (values == null) return result;

                // 分析拆分结果
                var maxSplits = 0;
                var splitResults = new List<List<string>>();

                foreach (var cellValue in values)
                {
                    if (cellValue != null)
                    {
                        var parts = cellValue.ToString().Split(new[] { delimiter }, StringSplitOptions.None);
                        maxSplits = Math.Max(maxSplits, parts.Length);
                        splitResults.Add(parts.ToList());
                    }
                    else
                    {
                        splitResults.Add(new List<string>());
                    }
                }

                result.MaxColumns = maxSplits;

                if (splitToNewColumns && maxSplits > 1)
                {
                    // 创建新列
                    var newRange = targetRange.Worksheet.Range[
                        targetRange.Cells[1, 1],
                        targetRange.Cells[targetRange.Rows.Count, targetRange.Column + maxSplits - 1]];

                    var newValues = new object[values.GetLength(0), maxSplits];

                    for (int row = 0; row < splitResults.Count; row++)
                    {
                        var parts = splitResults[row];
                        for (int col = 0; col < maxSplits; col++)
                        {
                            newValues[row, col] = col < parts.Count ? parts[col] : "";
                        }
                    }

                    newRange.Value2 = newValues;
                    result.NewRange = newRange;
                }

                result.SplittedCells = splitResults.Count;
                return result;
            }
            catch (Exception ex)
            {
                return new SplitResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        /// <summary>
        /// 合并列
        /// </summary>
        public async Task<TextProcessResult> MergeColumnsAsync(Range targetRange, string delimiter = " ")
        {
            return await ProcessTextAsync(targetRange, TextOperation.MergeColumns, new TextProcessOptions
            {
                Delimiter = delimiter
            });
        }

        /// <summary>
        /// 获取支持的文本操作列表
        /// </summary>
        public List<TextOperationInfo> GetSupportedOperations()
        {
            return new List<TextOperationInfo>
            {
                new TextOperationInfo { Operation = TextOperation.UpperCase, Name = "转大写", Description = "将所有字母转换为大写", Category = "大小写转换" },
                new TextOperationInfo { Operation = TextOperation.LowerCase, Name = "转小写", Description = "将所有字母转换为小写", Category = "大小写转换" },
                new TextOperationInfo { Operation = TextOperation.ProperCase, Name = "首字母大写", Description = "将每个单词的首字母大写", Category = "大小写转换" },
                new TextOperationInfo { Operation = TextOperation.TrimStart, Name = "删除首部空格", Description = "删除文本开头的空格", Category = "空格处理" },
                new TextOperationInfo { Operation = TextOperation.TrimEnd, Name = "删除尾部空格", Description = "删除文本末尾的空格", Category = "空格处理" },
                new TextOperationInfo { Operation = TextOperation.TrimAll, Name = "删除所有空格", Description = "删除文本中的所有空格", Category = "空格处理" },
                new TextOperationInfo { Operation = TextOperation.ExtractNumbers, Name = "提取数字", Description = "从文本中提取所有数字", Category = "智能提取" },
                new TextOperationInfo { Operation = TextOperation.ExtractLetters, Name = "提取字母", Description = "从文本中提取所有字母", Category = "智能提取" },
                new TextOperationInfo { Operation = TextOperation.ExtractEmails, Name = "提取邮箱", Description = "从文本中提取邮箱地址", Category = "智能提取" },
                new TextOperationInfo { Operation = TextOperation.ExtractPhones, Name = "提取手机号", Description = "从文本中提取手机号码", Category = "智能提取" },
                new TextOperationInfo { Operation = TextOperation.ExtractUrls, Name = "提取网址", Description = "从文本中提取网址链接", Category = "智能提取" },
                new TextOperationInfo { Operation = TextOperation.Replace, Name = "批量替换", Description = "批量替换指定文本", Category = "批量操作" },
                new TextOperationInfo { Operation = TextOperation.AddPrefix, Name = "添加前缀", Description = "为文本添加前缀", Category = "批量操作" },
                new TextOperationInfo { Operation = TextOperation.AddSuffix, Name = "添加后缀", Description = "为文本添加后缀", Category = "批量操作" },
                new TextOperationInfo { Operation = TextOperation.MergeColumns, Name = "合并列", Description = "将多列数据合并为一列", Category = "批量操作" }
            };
        }

        #endregion

        #region 核心处理逻辑

        /// <summary>
        /// 初始化操作处理器
        /// </summary>
        private Dictionary<TextOperation, Func<string, string>> InitializeOperationHandlers()
        {
            return new Dictionary<TextOperation, Func<string, string>>
            {
                [TextOperation.UpperCase] = text => text?.ToUpperInvariant(),
                [TextOperation.LowerCase] = text => text?.ToLowerInvariant(),
                [TextOperation.ProperCase] = ToProperCase,
                [TextOperation.TrimStart] = text => text?.TrimStart(),
                [TextOperation.TrimEnd] = text => text?.TrimEnd(),
                [TextOperation.TrimAll] = text => text?.Trim(),
                [TextOperation.ExtractNumbers] = ExtractNumbers,
                [TextOperation.ExtractLetters] = ExtractLetters,
                [TextOperation.ExtractEmails] = ExtractEmails,
                [TextOperation.ExtractPhones] = ExtractPhones,
                [TextOperation.ExtractUrls] = ExtractUrls,
                [TextOperation.RemoveSpaces] = text => text?.Replace(" ", ""),
                [TextOperation.AddPrefix] = (text) => text, // 在ProcessSingleCell中处理
                [TextOperation.AddSuffix] = (text) => text, // 在ProcessSingleCell中处理
                [TextOperation.Replace] = (text) => text, // 在ProcessSingleCell中处理
                [TextOperation.MergeColumns] = (text) => text, // 特殊处理
                [TextOperation.Custom] = (text) => text // 在ProcessSingleCell中处理
            };
        }

        /// <summary>
        /// 处理单个单元格
        /// </summary>
        private object ProcessSingleCell(object originalValue, TextOperation operation, TextProcessOptions options)
        {
            if (originalValue == null || originalValue is DBNull)
            {
                return options.SkipEmptyCells ? originalValue : "";
            }

            string text = originalValue.ToString();

            // 跳过空单元格
            if (string.IsNullOrWhiteSpace(text) && options.SkipEmptyCells)
            {
                return originalValue;
            }

            try
            {
                string result = text;

                // 应用文本转换
                if (_operationHandlers.TryGetValue(operation, out var handler))
                {
                    if (operation == TextOperation.Custom && options.CustomOperation != null)
                    {
                        result = options.CustomOperation(result);
                    }
                    else if (operation == TextOperation.Replace)
                    {
                        result = PerformReplace(result, options);
                    }
                    else if (operation == TextOperation.AddPrefix || operation == TextOperation.AddSuffix)
                    {
                        result = PerformAffixOperation(result, operation, options);
                    }
                    else
                    {
                        result = handler(result);
                    }
                }

                // 返回适当的数据类型
                return ConvertToAppropriateType(result, originalValue);
            }
            catch
            {
                // 处理失败时返回原值
                return originalValue;
            }
        }

        /// <summary>
        /// 转换为适当的数据类型
        /// </summary>
        private object ConvertToAppropriateType(string text, object originalValue)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            // 尝试保持原始数据类型
            if (originalValue is double || originalValue is decimal)
            {
                if (double.TryParse(text, out double numValue))
                    return numValue;
            }
            else if (originalValue is int || originalValue is long)
            {
                if (int.TryParse(text, out int intValue))
                    return intValue;
            }
            else if (originalValue is DateTime)
            {
                if (DateTime.TryParse(text, out DateTime dateValue))
                    return dateValue;
            }

            return text;
        }

        #endregion

        #region 具体操作实现

        /// <summary>
        /// 首字母大写
        /// </summary>
        private string ToProperCase(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            var words = text.ToLower().Split(' ', StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < words.Length; i++)
            {
                if (words[i].Length > 0)
                {
                    words[i] = char.ToUpper(words[i][0]) + words[i].Substring(1);
                }
            }

            return string.Join(" ", words);
        }

        /// <summary>
        /// 提取数字
        /// </summary>
        private string ExtractNumbers(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";

            var regex = GetCachedRegex(@"\d+\.?\d*");
            var matches = regex.Matches(text);
            var numbers = new List<string>();

            foreach (Match match in matches)
            {
                numbers.Add(match.Value);
            }

            return string.Join(" ", numbers);
        }

        /// <summary>
        /// 提取字母
        /// </summary>
        private string ExtractLetters(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";

            var regex = GetCachedRegex(@"[a-zA-Z]+");
            var matches = regex.Matches(text);
            var letters = new List<string>();

            foreach (Match match in matches)
            {
                letters.Add(match.Value);
            }

            return string.Join(" ", letters);
        }

        /// <summary>
        /// 提取邮箱
        /// </summary>
        private string ExtractEmails(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";

            var regex = GetCachedRegex(@"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b");
            var matches = regex.Matches(text);
            var emails = new List<string>();

            foreach (Match match in matches)
            {
                emails.Add(match.Value);
            }

            return string.Join(", ", emails);
        }

        /// <summary>
        /// 提取手机号
        /// </summary>
        private string ExtractPhones(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";

            // 支持多种手机号格式
            var regex = GetCachedRegex(@"(?:1[3-9]\d{9}|0\d{2,3}-?\d{7,8})");
            var matches = regex.Matches(text);
            var phones = new List<string>();

            foreach (Match match in matches)
            {
                phones.Add(match.Value);
            }

            return string.Join(", ", phones);
        }

        /// <summary>
        /// 提取网址
        /// </summary>
        private string ExtractUrls(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";

            var regex = GetCachedRegex(@"https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()@:%_\+.~#?&//=]*)");
            var matches = regex.Matches(text);
            var urls = new List<string>();

            foreach (Match match in matches)
            {
                urls.Add(match.Value);
            }

            return string.Join(", ", urls);
        }

        /// <summary>
        /// 执行替换操作
        /// </summary>
        private string PerformReplace(string text, TextProcessOptions options)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(options.FindText))
                return text;

            var comparison = options.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            var replaceText = options.ReplaceText ?? "";

            if (options.WholeWord)
            {
                // 全词替换
                var pattern = $@"\b{Regex.Escape(options.FindText)}\b";
                var regexOptions = options.MatchCase ? RegexOptions.None : RegexOptions.IgnoreCase;
                return Regex.Replace(text, pattern, replaceText, regexOptions);
            }
            else
            {
                // 普通替换
                if (options.MatchCase)
                {
                    return text.Replace(options.FindText, replaceText);
                }
                else
                {
                    return Regex.Replace(text, Regex.Escape(options.FindText), replaceText, RegexOptions.IgnoreCase);
                }
            }
        }

        /// <summary>
        /// 执行前后缀操作
        /// </summary>
        private string PerformAffixOperation(string text, TextOperation operation, TextProcessOptions options)
        {
            if (string.IsNullOrEmpty(text)) return text;

            switch (operation)
            {
                case TextOperation.AddPrefix:
                    return (options.Prefix ?? "") + text;
                case TextOperation.AddSuffix:
                    return text + (options.Suffix ?? "");
                default:
                    return text;
            }
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 获取缓存的正则表达式
        /// </summary>
        private Regex GetCachedRegex(string pattern)
        {
            if (!_regexCache.TryGetValue(pattern, out var regex))
            {
                regex = new Regex(pattern, RegexOptions.Compiled);
                _regexCache[pattern] = regex;
            }

            return regex;
        }

        /// <summary>
        /// 获取默认选项
        /// </summary>
        private TextProcessOptions GetDefaultOptions()
        {
            return new TextProcessOptions
            {
                SkipEmptyCells = true,
                MatchCase = false,
                WholeWord = false
            };
        }

        #endregion
    }

    #region 数据模型

    /// <summary>
    /// 文本处理结果
    /// </summary>
    public class TextProcessResult
    {
        public bool Success { get; set; }
        public TextOperation Operation { get; set; }
        public int TotalCells { get; set; }
        public int ProcessedCells { get; set; }
        public int SkippedCells { get; set; }
        public int ErrorCells { get; set; }
        public TimeSpan ProcessingTime { get; set; }
        public string ErrorMessage { get; set; }
        public List<ProcessError> Errors { get; set; } = new List<ProcessError>();
    }

    /// <summary>
    /// 拆分结果
    /// </summary>
    public class SplitResult
    {
        public bool Success { get; set; }
        public int SplittedCells { get; set; }
        public int MaxColumns { get; set; }
        public Range NewRange { get; set; }
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 文本处理选项
    /// </summary>
    public class TextProcessOptions
    {
        public bool SkipEmptyCells { get; set; } = true;
        public bool MatchCase { get; set; } = false;
        public bool WholeWord { get; set; } = false;
        public string FindText { get; set; }
        public string ReplaceText { get; set; }
        public string Prefix { get; set; }
        public string Suffix { get; set; }
        public string Delimiter { get; set; } = " ";
        public Func<string, string> CustomOperation { get; set; }
    }

    /// <summary>
    /// 处理错误
    /// </summary>
    public class ProcessError
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string OriginalValue { get; set; }
        public string Error { get; set; }
    }

    /// <summary>
    /// 文本操作信息
    /// </summary>
    public class TextOperationInfo
    {
        public TextOperation Operation { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string Category { get; set; }
    }

    #endregion

    #region 枚举

    /// <summary>
    /// 文本操作类型
    /// </summary>
    public enum TextOperation
    {
        // 大小写转换
        UpperCase,
        LowerCase,
        ProperCase,

        // 空格处理
        TrimStart,
        TrimEnd,
        TrimAll,
        RemoveSpaces,

        // 智能提取
        ExtractNumbers,
        ExtractLetters,
        ExtractEmails,
        ExtractPhones,
        ExtractUrls,

        // 批量操作
        Replace,
        AddPrefix,
        AddSuffix,
        MergeColumns,

        // 自定义操作
        Custom
    }

    #endregion
}