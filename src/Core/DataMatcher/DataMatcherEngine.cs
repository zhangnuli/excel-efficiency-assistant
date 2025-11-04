using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace ExcelEfficiencyAssistant.Core.DataMatcher
{
    /// <summary>
    /// 数据匹配引擎 - 智能VLOOKUP功能
    /// 支持自动识别主键列、智能匹配源数据、批量处理
    /// </summary>
    public class DataMatcherEngine
    {
        private readonly Excel.Application _application;
        private readonly List<MatchingRule> _matchingRules;
        private readonly Dictionary<string, object> _cache;

        public DataMatcherEngine(Excel.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _matchingRules = InitializeDefaultRules();
            _cache = new Dictionary<string, object>();
        }

        #region 公共接口

        /// <summary>
        /// 智能匹配数据 - 主要入口点
        /// </summary>
        /// <param name="targetRange">目标区域</param>
        /// <param name="matchOptions">匹配选项</param>
        /// <returns>匹配结果</returns>
        public async Task<MatchResult> SmartMatchAsync(Range targetRange, MatchOptions matchOptions = null)
        {
            matchOptions ??= GetDefaultMatchOptions();

            try
            {
                // 1. 分析目标区域
                var analysis = AnalyzeTargetRange(targetRange);

                // 2. 智能识别主键列
                var keyColumn = DetectKeyColumn(analysis);

                // 3. 扫描工作簿查找匹配源
                var matchSources = ScanForMatchSources(keyColumn, matchOptions);

                // 4. 生成匹配建议
                var suggestions = GenerateMatchSuggestions(analysis, matchSources);

                // 5. 执行匹配
                var result = await ExecuteMatchAsync(targetRange, suggestions, matchOptions);

                return result;
            }
            catch (Exception ex)
            {
                throw new DataMatcherException("数据匹配失败", ex);
            }
        }

        /// <summary>
        /// 批量VLOOKUP操作
        /// </summary>
        public async Task<BatchMatchResult> BatchVLookupAsync(Range lookupRange, Range tableArray, int colIndex, bool exactMatch = true)
        {
            var result = new BatchMatchResult();

            try
            {
                var lookupValues = GetRangeValues(lookupRange);
                var tableData = GetRangeValues(tableArray);
                var lookupColumn = 0; // 假设第一列是查找列

                // 构建查找字典
                var lookupDict = BuildLookupDictionary(tableData, lookupColumn);

                // 执行批量查找
                var matchedValues = new object[lookupValues.GetLength(0), 1];
                int matchCount = 0, notFoundCount = 0;

                for (int i = 0; i < lookupValues.GetLength(0); i++)
                {
                    var lookupValue = lookupValues[i, 0];
                    if (lookupValue != null && lookupDict.TryGetValue(lookupValue, out var matchRow))
                    {
                        matchedValues[i, 0] = matchRow[Math.Min(colIndex - 1, matchRow.Length - 1)];
                        matchCount++;
                    }
                    else
                    {
                        matchedValues[i, 0] = exactMatch ? "#N/A" : "";
                        notFoundCount++;
                    }
                }

                // 写入结果
                var resultRange = lookupRange.Offset(0, 1);
                resultRange.Value2 = matchedValues;

                result.TotalRows = lookupValues.GetLength(0);
                result.MatchedRows = matchCount;
                result.NotFoundRows = notFoundCount;
                result.Success = true;

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
        }

        #endregion

        #region 核心算法

        /// <summary>
        /// 分析目标区域
        /// </summary>
        private RangeAnalysis AnalyzeTargetRange(Range range)
        {
            var analysis = new RangeAnalysis
            {
                Range = range,
                RowCount = range.Rows.Count,
                ColumnCount = range.Columns.Count,
                Values = GetRangeValues(range),
                UsedRange = range.Worksheet.UsedRange
            };

            // 分析数据类型
            analysis.ColumnTypes = AnalyzeColumnTypes(analysis.Values);

            // 分析表头
            analysis.HasHeader = DetectHeader(analysis.Values);

            // 分析数据质量
            analysis.DataQuality = AnalyzeDataQuality(analysis.Values);

            return analysis;
        }

        /// <summary>
        /// 检测主键列
        /// </summary>
        private KeyColumnDetection DetectKeyColumn(RangeAnalysis analysis)
        {
            var candidates = new List<KeyColumnCandidate>();

            for (int col = 0; col < analysis.ColumnCount; col++)
            {
                var candidate = AnalyzeColumnAsKey(analysis.Values, col, analysis.HasHeader);
                if (candidate.Score > 0.3) // 只考虑得分较高的列
                {
                    candidates.Add(candidate);
                }
            }

            var best = candidates.OrderByDescending(c => c.Score).FirstOrDefault();
            return new KeyColumnDetection
            {
                ColumnIndex = best?.ColumnIndex ?? -1,
                Confidence = best?.Score ?? 0,
                ColumnName = best?.ColumnName,
                Reason = best?.Reason,
                Alternatives = candidates.Where(c => c != best).ToList()
            };
        }

        /// <summary>
        /// 分析列作为主键的潜力
        /// </summary>
        private KeyColumnCandidate AnalyzeColumnAsKey(object[,] values, int columnIndex, bool hasHeader)
        {
            var candidate = new KeyColumnCandidate
            {
                ColumnIndex = columnIndex,
                ColumnName = hasHeader && values.GetLength(0) > 0 ? values[0, columnIndex]?.ToString() : $"列{columnIndex + 1}"
            };

            int startRow = hasHeader ? 1 : 0;
            int totalRows = values.GetLength(0) - startRow;
            var uniqueValues = new HashSet<object>();
            int nullCount = 0, numericCount = 0, textCount = 0, idPatternCount = 0;

            for (int row = startRow; row < values.GetLength(0); row++)
            {
                var value = values[row, columnIndex];

                if (value == null || value.ToString().Trim() == "")
                {
                    nullCount++;
                }
                else
                {
                    uniqueValues.Add(value);

                    var strValue = value.ToString();

                    if (IsNumericPattern(strValue))
                        numericCount++;
                    else
                        textCount++;

                    if (IsIdPattern(strValue))
                        idPatternCount++;
                }
            }

            // 计算唯一性得分
            double uniquenessScore = (double)uniqueValues.Count / totalRows;

            // 计算完整性得分
            double completenessScore = 1.0 - (double)nullCount / totalRows;

            // 计算类型一致性得分
            double typeConsistencyScore = Math.Max(numericCount, textCount) / (double)totalRows;

            // 计算ID模式得分
            double idPatternScore = (double)idPatternCount / totalRows;

            // 综合得分
            candidate.Score = (uniquenessScore * 0.4) +
                             (completenessScore * 0.3) +
                             (typeConsistencyScore * 0.2) +
                             (idPatternScore * 0.1);

            candidate.UniqueCount = uniqueValues.Count;
            candidate.NullCount = nullCount;
            candidate.TypeConsistency = typeConsistencyScore;

            // 生成判断原因
            candidate.Reason = GenerateKeyColumnReason(candidate, uniquenessScore, completenessScore, typeConsistencyScore);

            return candidate;
        }

        /// <summary>
        /// 扫描匹配源
        /// </summary>
        private List<MatchSource> ScanForMatchSources(KeyColumnDetection keyColumn, MatchOptions options)
        {
            var sources = new List<MatchSource>();

            foreach (Worksheet worksheet in _application.Worksheets)
            {
                if (worksheet == keyColumn.Column?.Worksheet) continue;

                try
                {
                    var usedRange = worksheet.UsedRange;
                    if (usedRange == null) continue;

                    var analysis = AnalyzeTargetRange(usedRange);

                    // 查找包含相似列的表格
                    var matchingColumns = FindMatchingColumns(keyColumn, analysis);

                    foreach (var matchCol in matchingColumns)
                    {
                        var source = new MatchSource
                        {
                            Worksheet = worksheet,
                            Range = analysis.UsedRange,
                            MatchingColumnIndex = matchCol.ColumnIndex,
                            ColumnName = matchCol.ColumnName,
                            MatchScore = matchCol.MatchScore,
                            AvailableColumns = GetAvailableColumns(analysis.Values),
                            RowCount = analysis.RowCount
                        };

                        if (source.MatchScore > 0.3) // 只考虑匹配度较高的源
                        {
                            sources.Add(source);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 记录错误但继续处理其他工作表
                    System.Diagnostics.Debug.WriteLine($"扫描工作表 {worksheet.Name} 时出错: {ex.Message}");
                }
            }

            return sources.OrderByDescending(s => s.MatchScore).ToList();
        }

        /// <summary>
        /// 查找匹配列
        /// </summary>
        private List<ColumnMatch> FindMatchingColumns(KeyColumnDetection keyColumn, RangeAnalysis analysis)
        {
            var matches = new List<ColumnMatch>();

            for (int col = 0; col < analysis.ColumnCount; col++)
            {
                var match = CalculateColumnMatch(keyColumn, analysis, col);
                if (match.MatchScore > 0.2)
                {
                    matches.Add(match);
                }
            }

            return matches;
        }

        /// <summary>
        /// 计算列匹配度
        /// </summary>
        private ColumnMatch CalculateColumnMatch(KeyColumnDetection keyColumn, RangeAnalysis analysis, int columnIndex)
        {
            var match = new ColumnMatch
            {
                ColumnIndex = columnIndex,
                ColumnName = analysis.HasHeader && analysis.Values.GetLength(0) > 0
                    ? analysis.Values[0, columnIndex]?.ToString()
                    : $"列{columnIndex + 1}"
            };

            // 名称相似度
            double nameSimilarity = CalculateStringSimilarity(keyColumn.ColumnName, match.ColumnName);

            // 数据类型相似度
            double typeSimilarity = CalculateDataTypeSimilarity(keyColumn, analysis, columnIndex);

            // 值重叠度（抽样检查）
            double valueOverlap = CalculateValueOverlap(keyColumn, analysis, columnIndex);

            match.MatchScore = (nameSimilarity * 0.3) + (typeSimilarity * 0.3) + (valueOverlap * 0.4);

            return match;
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 获取区域值
        /// </summary>
        private object[,] GetRangeValues(Range range)
        {
            object value = range.Value2;

            if (value is object[,] value2D)
                return value2D;

            // 处理单行或单列的情况
            if (range.Rows.Count == 1 || range.Columns.Count == 1)
            {
                var value1D = value as object[,] ?? new object[1, 1] { { value } };
                return value1D;
            }

            return new object[1, 1] { { value } };
        }

        /// <summary>
        /// 构建查找字典
        /// </summary>
        private Dictionary<object, object[]> BuildLookupDictionary(object[,] tableData, int keyColumnIndex)
        {
            var dict = new Dictionary<object, object[]>();

            for (int row = 0; row < tableData.GetLength(0); row++)
            {
                var key = tableData[row, keyColumnIndex];
                if (key != null && !dict.ContainsKey(key))
                {
                    var rowData = new object[tableData.GetLength(1)];
                    for (int col = 0; col < tableData.GetLength(1); col++)
                    {
                        rowData[col] = tableData[row, col];
                    }
                    dict[key] = rowData;
                }
            }

            return dict;
        }

        /// <summary>
        /// 计算字符串相似度
        /// </summary>
        private double CalculateStringSimilarity(string str1, string str2)
        {
            if (string.IsNullOrEmpty(str1) || string.IsNullOrEmpty(str2))
                return 0;

            str1 = str1.ToLowerInvariant();
            str2 = str2.ToLowerInvariant();

            // 完全匹配
            if (str1 == str2) return 1.0;

            // 包含关系
            if (str1.Contains(str2) || str2.Contains(str1)) return 0.8;

            // Levenshtein距离
            int maxLength = Math.Max(str1.Length, str2.Length);
            if (maxLength == 0) return 1.0;

            int distance = LevenshteinDistance(str1, str2);
            return 1.0 - (double)distance / maxLength;
        }

        /// <summary>
        /// 计算Levenshtein距离
        /// </summary>
        private int LevenshteinDistance(string s1, string s2)
        {
            int[,] matrix = new int[s1.Length + 1, s2.Length + 1];

            for (int i = 0; i <= s1.Length; i++)
                matrix[i, 0] = i;

            for (int j = 0; j <= s2.Length; j++)
                matrix[0, j] = j;

            for (int i = 1; i <= s1.Length; i++)
            {
                for (int j = 1; j <= s2.Length; j++)
                {
                    int cost = s1[i - 1] == s2[j - 1] ? 0 : 1;
                    matrix[i, j] = Math.Min(
                        Math.Min(matrix[i - 1, j] + 1, matrix[i, j - 1] + 1),
                        matrix[i - 1, j - 1] + cost);
                }
            }

            return matrix[s1.Length, s2.Length];
        }

        /// <summary>
        /// 检查是否为数字模式
        /// </summary>
        private bool IsNumericPattern(string value)
        {
            return decimal.TryParse(value, out _) ||
                   System.Text.RegularExpressions.Regex.IsMatch(value, @"^\d+$");
        }

        /// <summary>
        /// 检查是否为ID模式
        /// </summary>
        private bool IsIdPattern(string value)
        {
            // 常见ID模式：纯数字、字母数字组合、带分隔符的编号等
            return System.Text.RegularExpressions.Regex.IsMatch(value, @"^\d+$") || // 纯数字
                   System.Text.RegularExpressions.Regex.IsMatch(value, @"^[A-Za-z0-9]+$") || // 字母数字
                   System.Text.RegularExpressions.Regex.IsMatch(value, @"^\d{4}-\d{2}-\d{2}$") || // 日期格式
                   System.Text.RegularExpressions.Regex.IsMatch(value, @"^[A-Z]{2,4}\d{4,6}$"); // 常见编号格式
        }

        /// <summary>
        /// 初始化默认匹配规则
        /// </summary>
        private List<MatchingRule> InitializeDefaultRules()
        {
            return new List<MatchingRule>
            {
                new MatchingRule { Name = "ID匹配", Pattern = @"^\d+$", Weight = 0.9 },
                new MatchingRule { Name = "姓名匹配", Pattern = @"^[\u4e00-\u9fa5]{2,4}$", Weight = 0.8 },
                new MatchingRule { Name = "手机号匹配", Pattern = @"^1[3-9]\d{9}$", Weight = 0.95 },
                new MatchingRule { Name = "邮箱匹配", Pattern = @"^\w+@\w+\.\w+$", Weight = 0.9 },
                new MatchingRule { Name = "日期匹配", Pattern = @"^\d{4}-\d{2}-\d{2}$", Weight = 0.85 }
            };
        }

        /// <summary>
        /// 获取默认匹配选项
        /// </summary>
        private MatchOptions GetDefaultMatchOptions()
        {
            return new MatchOptions
            {
                SearchScope = SearchScope.Workbook,
                MatchType = MatchType.Exact,
                CaseSensitive = false,
                TrimWhitespace = true,
                IgnoreErrors = true,
                MaxResults = 1000
            };
        }

        #endregion

        #region 未实现的方法（需要根据实际需求完善）

        private ColumnType[] AnalyzeColumnTypes(object[,] values)
        {
            // TODO: 实现列类型分析
            return new ColumnType[0];
        }

        private bool DetectHeader(object[,] values)
        {
            // TODO: 实现表头检测逻辑
            return values.GetLength(0) > 1;
        }

        private DataQuality AnalyzeDataQuality(object[,] values)
        {
            // TODO: 实现数据质量分析
            return new DataQuality { CompletenessScore = 0.9, ConsistencyScore = 0.85 };
        }

        private double CalculateDataTypeSimilarity(KeyColumnDetection keyColumn, RangeAnalysis analysis, int columnIndex)
        {
            // TODO: 实现数据类型相似度计算
            return 0.7;
        }

        private double CalculateValueOverlap(KeyColumnDetection keyColumn, RangeAnalysis analysis, int columnIndex)
        {
            // TODO: 实现值重叠度计算（抽样检查）
            return 0.6;
        }

        private List<string> GetAvailableColumns(object[,] values)
        {
            // TODO: 实现可用列信息获取
            return new List<string>();
        }

        private List<MatchSuggestion> GenerateMatchSuggestions(RangeAnalysis analysis, List<MatchSource> sources)
        {
            // TODO: 实现匹配建议生成
            return new List<MatchSuggestion>();
        }

        private async Task<MatchResult> ExecuteMatchAsync(Range targetRange, List<MatchSuggestion> suggestions, MatchOptions options)
        {
            // TODO: 实现匹配执行逻辑
            return new MatchResult { Success = true, MatchedRows = 0 };
        }

        private string GenerateKeyColumnReason(KeyColumnCandidate candidate, double uniqueness, double completeness, double consistency)
        {
            var reasons = new List<string>();

            if (uniqueness > 0.8) reasons.Add("唯一性高");
            if (completeness > 0.9) reasons.Add("完整性好");
            if (consistency > 0.8) reasons.Add("类型一致");

            return reasons.Count > 0 ? string.Join("，", reasons) : "一般";
        }

        #endregion
    }

    #region 数据模型类

    /// <summary>
    /// 匹配结果
    /// </summary>
    public class MatchResult
    {
        public bool Success { get; set; }
        public int MatchedRows { get; set; }
        public int TotalRows { get; set; }
        public string ErrorMessage { get; set; }
        public List<MatchDetail> Details { get; set; } = new List<MatchDetail>();
    }

    /// <summary>
    /// 批量匹配结果
    /// </summary>
    public class BatchMatchResult
    {
        public bool Success { get; set; }
        public int TotalRows { get; set; }
        public int MatchedRows { get; set; }
        public int NotFoundRows { get; set; }
        public string ErrorMessage { get; set; }
        public TimeSpan ExecutionTime { get; set; }
    }

    /// <summary>
    /// 匹配选项
    /// </summary>
    public class MatchOptions
    {
        public SearchScope SearchScope { get; set; } = SearchScope.Workbook;
        public MatchType MatchType { get; set; } = MatchType.Exact;
        public bool CaseSensitive { get; set; } = false;
        public bool TrimWhitespace { get; set; } = true;
        public bool IgnoreErrors { get; set; } = true;
        public int MaxResults { get; set; } = 1000;
    }

    /// <summary>
    /// 区域分析结果
    /// </summary>
    public class RangeAnalysis
    {
        public Range Range { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public object[,] Values { get; set; }
        public Range UsedRange { get; set; }
        public bool HasHeader { get; set; }
        public ColumnType[] ColumnTypes { get; set; }
        public DataQuality DataQuality { get; set; }
    }

    /// <summary>
    /// 主键列检测结果
    /// </summary>
    public class KeyColumnDetection
    {
        public int ColumnIndex { get; set; }
        public double Confidence { get; set; }
        public string ColumnName { get; set; }
        public string Reason { get; set; }
        public List<KeyColumnCandidate> Alternatives { get; set; } = new List<KeyColumnCandidate>();
        public Range Column { get; set; }
    }

    /// <summary>
    /// 主键列候选
    /// </summary>
    public class KeyColumnCandidate
    {
        public int ColumnIndex { get; set; }
        public string ColumnName { get; set; }
        public double Score { get; set; }
        public string Reason { get; set; }
        public int UniqueCount { get; set; }
        public int NullCount { get; set; }
        public double TypeConsistency { get; set; }
    }

    /// <summary>
    /// 匹配源
    /// </summary>
    public class MatchSource
    {
        public Worksheet Worksheet { get; set; }
        public Range Range { get; set; }
        public int MatchingColumnIndex { get; set; }
        public string ColumnName { get; set; }
        public double MatchScore { get; set; }
        public List<string> AvailableColumns { get; set; } = new List<string>();
        public int RowCount { get; set; }
    }

    /// <summary>
    /// 列匹配结果
    /// </summary>
    public class ColumnMatch
    {
        public int ColumnIndex { get; set; }
        public string ColumnName { get; set; }
        public double MatchScore { get; set; }
    }

    /// <summary>
    /// 匹配建议
    /// </summary>
    public class MatchSuggestion
    {
        public MatchSource Source { get; set; }
        public List<ColumnMapping> Mappings { get; set; } = new List<ColumnMapping>();
        public double Confidence { get; set; }
        public string Reason { get; set; }
    }

    /// <summary>
    /// 列映射
    /// </summary>
    public class ColumnMapping
    {
        public int SourceColumnIndex { get; set; }
        public int TargetColumnIndex { get; set; }
        public string SourceColumnName { get; set; }
        public string TargetColumnName { get; set; }
        public double MatchScore { get; set; }
    }

    /// <summary>
    /// 匹配详情
    /// </summary>
    public class MatchDetail
    {
        public int RowIndex { get; set; }
        public object LookupValue { get; set; }
        public object MatchedValue { get; set; }
        public MatchStatus Status { get; set; }
        public string Message { get; set; }
    }

    /// <summary>
    /// 匹配规则
    /// </summary>
    public class MatchingRule
    {
        public string Name { get; set; }
        public string Pattern { get; set; }
        public double Weight { get; set; }
    }

    /// <summary>
    /// 列类型
    /// </summary>
    public class ColumnType
    {
        public Type DataType { get; set; }
        public double ConsistencyScore { get; set; }
        public List<string> SampleValues { get; set; } = new List<string>();
    }

    /// <summary>
    /// 数据质量
    /// </summary>
    public class DataQuality
    {
        public double CompletenessScore { get; set; }
        public double ConsistencyScore { get; set; }
        public double AccuracyScore { get; set; }
        public List<string> Issues { get; set; } = new List<string>();
    }

    #endregion

    #region 枚举

    /// <summary>
    /// 搜索范围
    /// </summary>
    public enum SearchScope
    {
        Worksheet,
        Workbook
    }

    /// <summary>
    /// 匹配类型
    /// </summary>
    public enum MatchType
    {
        Exact,
        Fuzzy,
        Partial
    }

    /// <summary>
    /// 匹配状态
    /// </summary>
    public enum MatchStatus
    {
        Success,
        NotFound,
        Error,
        Duplicate
    }

    #endregion

    #region 自定义异常

    /// <summary>
    /// 数据匹配异常
    /// </summary>
    public class DataMatcherException : Exception
    {
        public DataMatcherException(string message) : base(message) { }
        public DataMatcherException(string message, Exception innerException) : base(message, innerException) { }
    }

    #endregion
}