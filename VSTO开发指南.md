# VSTO Excel插件开发完整指南

## 🎯 从零开始创建VSTO项目

由于我们无法直接生成.sln文件，需要在Visual Studio中手动创建项目。以下是完整步骤：

---

## 📋 步骤一：安装必要软件

### 1. 安装Visual Studio 2022

#### 下载地址
- Community版（免费）: https://visualstudio.microsoft.com/zh-hans/downloads/
- 或直接下载: https://visualstudio.microsoft.com/zh-hans/thank-you-downloading-visual-studio/?sku=Community

#### 安装时选择的工作负载
安装时必须勾选：
- ✅ **.NET桌面开发**
- ✅ **Office/SharePoint 开发**

#### 可选组件
- ✅ .NET 6.0 Runtime
- ✅ .NET Framework 4.8 开发工具
- ✅ Office开发工具

### 2. 安装Office

需要安装以下之一：
- Microsoft Office 2016/2019/2021
- Microsoft 365 (推荐)

---

## 🚀 步骤二：在Visual Studio中创建VSTO项目

### 1. 创建新项目

1. 打开 Visual Studio 2022
2. 点击 **"创建新项目"**
3. 搜索 **"Excel VSTO Add-in"** 或 **"Excel 加载项"**
   - 如果找不到，说明没有安装"Office/SharePoint 开发"工作负载
   - 需要返回Visual Studio Installer安装
4. 选择 **"Excel VSTO Add-in"**
5. 点击 **"下一步"**

### 2. 配置项目

```
项目名称: ExcelEfficiencyAssistant
位置: D:\项目代码存放\2025\excel插件
解决方案名称: ExcelEfficiencyAssistant
框架: .NET 6.0 (或 .NET Framework 4.8)
```

点击 **"创建"**

### 3. 选择Office版本

在弹出的向导中：
- Office 版本: **Excel 2016** (向下兼容)
- 点击 **"完成"**

---

## 📁 步骤三：项目创建后的初始结构

Visual Studio会自动创建以下文件：

```
ExcelEfficiencyAssistant/
├── Properties/
│   ├── AssemblyInfo.cs
│   └── Settings.settings
├── ThisAddIn.cs                    # 👈 插件主入口
├── ThisAddIn.Designer.cs
├── ExcelEfficiencyAssistant.csproj
└── packages.config
```

---

## 🎨 步骤四：添加Ribbon界面

### 1. 添加Ribbon（功能区）

1. 右键点击项目 → **添加** → **新建项**
2. 选择 **"功能区(可视化设计器)"**
3. 名称: `EfficiencyRibbon.cs`
4. 点击 **"添加"**

### 2. 设计Ribbon界面

Visual Studio会打开可视化设计器：

#### 添加选项卡
1. 从工具箱拖拽 **"Tab"** 到设计器
2. 设置属性:
   - Name: `tabEfficiency`
   - Label: `效率助手`
   - ControlId: `EfficiencyTab`

#### 添加组
1. 拖拽 **"Group"** 到选项卡
2. 设置属性:
   - Name: `groupDataMatch`
   - Label: `数据匹配`

#### 添加按钮
1. 拖拽 **"Button"** 到组
2. 设置属性:
   - Name: `btnSmartMatch`
   - Label: `智能匹配`
   - ControlSize: `Large`
   - ShowImage: `True`

### 3. 添加按钮图标

#### 准备图标（32x32 PNG）
我们需要创建图标文件，您可以：
- 使用assets文件夹中已有的icon.png
- 或者下载图标库: https://icons8.com

#### 导入图标
1. 右键项目 → 添加 → 现有项
2. 选择图标文件
3. 设置 **"生成操作"** 为 **"嵌入的资源"**

#### 设置按钮图标
```csharp
// 在EfficiencyRibbon.cs的代码中
private void EfficiencyRibbon_Load(object sender, RibbonUIEventArgs e)
{
    // 加载图标
    btnSmartMatch.Image = Properties.Resources.SmartMatchIcon;
}
```

---

## 💻 步骤五：创建项目文件夹结构

在Visual Studio解决方案资源管理器中：

1. 右键项目 → 添加 → 新建文件夹

创建以下文件夹：
```
ExcelEfficiencyAssistant/
├── Core/
│   ├── DataMatcher/
│   ├── Beautifier/
│   └── TextProcessor/
├── UI/
│   ├── Dialogs/
│   └── TaskPanes/
├── Services/
├── Data/
│   ├── Database/
│   └── Models/
├── Helpers/
└── Resources/
    ├── Templates/
    └── Icons/
```

---

## 📝 步骤六：添加NuGet包

### 方法1: 使用NuGet包管理器

1. 右键项目 → **管理NuGet程序包**
2. 点击 **"浏览"**
3. 搜索并安装以下包：

```
Microsoft.EntityFrameworkCore.Sqlite (7.0.14)
Newtonsoft.Json (13.0.3)
AutoMapper (12.0.1)
Serilog (3.1.1)
Serilog.Sinks.File (5.0.0)
```

### 方法2: 使用Package Manager Console

1. 工具 → NuGet包管理器 → 程序包管理器控制台
2. 运行以下命令：

```powershell
Install-Package Microsoft.EntityFrameworkCore.Sqlite -Version 7.0.14
Install-Package Newtonsoft.Json -Version 13.0.3
Install-Package AutoMapper -Version 12.0.1
Install-Package Serilog -Version 3.1.1
Install-Package Serilog.Sinks.File -Version 5.0.0
```

---

## 🔧 步骤七：编写核心代码

### 1. 修改ThisAddIn.cs

```csharp
using System;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Serilog;

namespace ExcelEfficiencyAssistant
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 初始化日志
            InitializeLogger();

            Log.Information("Excel效率助手启动...");

            try
            {
                // 初始化服务
                InitializeServices();

                Log.Information("插件初始化完成");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "插件初始化失败");
                System.Windows.Forms.MessageBox.Show(
                    $"插件初始化失败: {ex.Message}",
                    "Excel效率助手",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Log.Information("Excel效率助手关闭");
            Log.CloseAndFlush();
        }

        private void InitializeLogger()
        {
            string logPath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ExcelEfficiencyAssistant",
                "Logs",
                "log-.txt");

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.File(logPath,
                    rollingInterval: RollingInterval.Day,
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();
        }

        private void InitializeServices()
        {
            // TODO: 初始化服务
            // var settingsManager = new SettingsManager();
            // var templateManager = new TemplateManager();
        }

        #region VSTO 生成的代码

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
```

### 2. 创建数据匹配引擎

创建文件: `Core/DataMatcher/DataMatcherEngine.cs`

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Serilog;

namespace ExcelEfficiencyAssistant.Core.DataMatcher
{
    /// <summary>
    /// 数据匹配核心引擎
    /// </summary>
    public class DataMatcherEngine
    {
        /// <summary>
        /// 智能数据匹配
        /// </summary>
        public MatchResult SmartMatch(Excel.Range targetRange)
        {
            Log.Information("开始智能匹配...");

            var stopwatch = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                // 1. 分析目标区域
                Log.Debug("分析目标区域");
                var analysis = AnalyzeRange(targetRange);

                // 2. 智能识别主键列
                Log.Debug("识别主键列");
                var keyColumn = DetectKeyColumn(analysis);

                // 3. 扫描工作簿查找匹配源
                Log.Debug("扫描匹配源");
                var sources = FindMatchingSources(keyColumn, targetRange.Worksheet);

                if (sources == null || sources.Count == 0)
                {
                    Log.Warning("未找到可匹配的数据源");
                    return MatchResult.NoSourceFound();
                }

                // 4. 生成匹配建议
                Log.Debug("生成匹配建议");
                var suggestion = GenerateBestSuggestion(sources, keyColumn);

                // 5. 执行匹配
                Log.Debug("执行匹配");
                var result = ExecuteMatch(targetRange, suggestion);

                stopwatch.Stop();
                result.Duration = stopwatch.ElapsedMilliseconds;

                Log.Information($"匹配完成: 成功{result.MatchedCount}行, 失败{result.UnmatchedCount}行, 耗时{result.Duration}ms");

                return result;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "智能匹配失败");
                return MatchResult.Error(ex.Message);
            }
        }

        /// <summary>
        /// 分析数据区域
        /// </summary>
        private RangeAnalysis AnalyzeRange(Excel.Range range)
        {
            object[,] data = range.Value2 as object[,];

            if (data == null)
            {
                throw new ArgumentException("目标区域没有数据");
            }

            return new RangeAnalysis
            {
                RowCount = data.GetLength(0),
                ColumnCount = data.GetLength(1),
                Data = data,
                HasHeader = DetectHeader(data)
            };
        }

        /// <summary>
        /// 检测是否有标题行
        /// </summary>
        private bool DetectHeader(object[,] data)
        {
            if (data.GetLength(0) < 2) return false;

            // 检查第一行是否全是文本
            for (int col = 1; col <= data.GetLength(1); col++)
            {
                var value = data[1, col];
                if (value == null) return false;
                if (value is double || value is int) return false;
            }

            return true;
        }

        /// <summary>
        /// 智能检测主键列
        /// </summary>
        private ColumnInfo DetectKeyColumn(RangeAnalysis analysis)
        {
            var candidates = new List<ColumnCandidate>();

            int startRow = analysis.HasHeader ? 2 : 1;

            for (int col = 1; col <= analysis.ColumnCount; col++)
            {
                int score = CalculateKeyScore(analysis.Data, col, startRow, analysis.HasHeader);

                candidates.Add(new ColumnCandidate
                {
                    ColumnIndex = col,
                    ColumnName = analysis.HasHeader ? analysis.Data[1, col]?.ToString() : $"列{col}",
                    Score = score
                });
            }

            var best = candidates.OrderByDescending(c => c.Score).First();

            return new ColumnInfo
            {
                Index = best.ColumnIndex,
                Name = best.ColumnName,
                Confidence = best.Score
            };
        }

        /// <summary>
        /// 计算列作为主键的得分
        /// </summary>
        private int CalculateKeyScore(object[,] data, int colIndex, int startRow, bool hasHeader)
        {
            int score = 0;

            // 1. 检查列名（如果有标题）
            if (hasHeader)
            {
                string colName = data[1, colIndex]?.ToString()?.ToLower() ?? "";

                if (colName.Contains("id")) score += 50;
                else if (colName.Contains("编号")) score += 50;
                else if (colName.Contains("序号")) score += 30;
                else if (colName.Contains("代码")) score += 30;
                else if (colName.Contains("code")) score += 30;
            }

            // 2. 检查唯一性
            var uniqueValues = new HashSet<string>();
            int totalRows = data.GetLength(0) - (hasHeader ? 1 : 0);

            for (int row = startRow; row <= data.GetLength(0); row++)
            {
                string value = data[row, colIndex]?.ToString();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    uniqueValues.Add(value);
                }
            }

            double uniqueRatio = (double)uniqueValues.Count / totalRows;
            score += (int)(uniqueRatio * 50);

            // 3. 检查数据类型一致性
            bool isAllNumeric = true;
            bool isAllText = true;

            for (int row = startRow; row <= Math.Min(startRow + 100, data.GetLength(0)); row++)
            {
                var value = data[row, colIndex];
                if (value != null)
                {
                    if (value is double || value is int)
                        isAllText = false;
                    else
                        isAllNumeric = false;
                }
            }

            if (isAllNumeric || isAllText) score += 10;

            return score;
        }

        /// <summary>
        /// 查找可匹配的数据源
        /// </summary>
        private List<DataSource> FindMatchingSources(ColumnInfo keyColumn, Excel.Worksheet currentSheet)
        {
            var sources = new List<DataSource>();

            Excel.Workbook workbook = currentSheet.Parent as Excel.Workbook;

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name == currentSheet.Name) continue;

                try
                {
                    Excel.Range usedRange = sheet.UsedRange;
                    if (usedRange.Rows.Count < 2) continue;

                    object[,] data = usedRange.Value2 as object[,];
                    if (data == null) continue;

                    // 查找匹配列
                    for (int col = 1; col <= data.GetLength(1); col++)
                    {
                        string headerName = data[1, col]?.ToString() ?? "";

                        // 简单的名称匹配
                        if (IsSimilarColumnName(headerName, keyColumn.Name))
                        {
                            sources.Add(new DataSource
                            {
                                SheetName = sheet.Name,
                                MatchColumnIndex = col,
                                MatchColumnName = headerName,
                                RowCount = data.GetLength(0) - 1,
                                ColumnCount = data.GetLength(1)
                            });
                            break;
                        }
                    }
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                }
            }

            return sources;
        }

        /// <summary>
        /// 判断列名是否相似
        /// </summary>
        private bool IsSimilarColumnName(string name1, string name2)
        {
            if (string.IsNullOrWhiteSpace(name1) || string.IsNullOrWhiteSpace(name2))
                return false;

            name1 = name1.ToLower().Trim();
            name2 = name2.ToLower().Trim();

            return name1 == name2 ||
                   name1.Contains(name2) ||
                   name2.Contains(name1);
        }

        /// <summary>
        /// 生成最佳匹配建议
        /// </summary>
        private MatchSuggestion GenerateBestSuggestion(List<DataSource> sources, ColumnInfo keyColumn)
        {
            // 简单选择第一个源（后续可以增加智能选择逻辑）
            var bestSource = sources.OrderByDescending(s => s.RowCount).First();

            return new MatchSuggestion
            {
                SourceSheet = bestSource.SheetName,
                SourceMatchColumn = bestSource.MatchColumnIndex,
                TargetKeyColumn = keyColumn.Index,
                ReturnColumns = Enumerable.Range(1, bestSource.ColumnCount)
                    .Where(i => i != bestSource.MatchColumnIndex)
                    .Take(3) // 默认返回前3列
                    .ToList()
            };
        }

        /// <summary>
        /// 执行匹配
        /// </summary>
        private MatchResult ExecuteMatch(Excel.Range targetRange, MatchSuggestion suggestion)
        {
            Excel.Workbook workbook = targetRange.Worksheet.Parent as Excel.Workbook;
            Excel.Worksheet sourceSheet = workbook.Worksheets[suggestion.SourceSheet] as Excel.Worksheet;

            try
            {
                // 读取源数据
                Excel.Range sourceRange = sourceSheet.UsedRange;
                object[,] sourceData = sourceRange.Value2 as object[,];

                // 构建索引
                var index = BuildMatchIndex(sourceData, suggestion.SourceMatchColumn);

                // 读取目标数据
                object[,] targetData = targetRange.Value2 as object[,];

                // 执行匹配
                var result = PerformMatch(targetData, index, suggestion, targetRange);

                return result;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceSheet);
            }
        }

        /// <summary>
        /// 构建匹配索引
        /// </summary>
        private Dictionary<string, object[]> BuildMatchIndex(object[,] data, int keyCol)
        {
            var index = new Dictionary<string, object[]>();

            for (int row = 2; row <= data.GetLength(0); row++)
            {
                string key = data[row, keyCol]?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(key) && !index.ContainsKey(key))
                {
                    object[] rowData = new object[data.GetLength(1)];
                    for (int col = 1; col <= data.GetLength(1); col++)
                    {
                        rowData[col - 1] = data[row, col];
                    }
                    index[key] = rowData;
                }
            }

            return index;
        }

        /// <summary>
        /// 执行匹配操作
        /// </summary>
        private MatchResult PerformMatch(
            object[,] targetData,
            Dictionary<string, object[]> index,
            MatchSuggestion suggestion,
            Excel.Range targetRange)
        {
            int matchedCount = 0;
            int unmatchedCount = 0;

            // 准备结果数组
            int rowCount = targetData.GetLength(0) - 1; // 减去标题行
            int colCount = suggestion.ReturnColumns.Count;
            object[,] results = new object[rowCount, colCount];

            // 执行匹配
            for (int row = 2; row <= targetData.GetLength(0); row++)
            {
                string key = targetData[row, suggestion.TargetKeyColumn]?.ToString()?.Trim();

                if (!string.IsNullOrEmpty(key) && index.ContainsKey(key))
                {
                    var sourceRow = index[key];
                    for (int i = 0; i < suggestion.ReturnColumns.Count; i++)
                    {
                        results[row - 2, i] = sourceRow[suggestion.ReturnColumns[i] - 1];
                    }
                    matchedCount++;
                }
                else
                {
                    unmatchedCount++;
                }
            }

            // 写入结果到Excel
            int targetCol = targetData.GetLength(1) + 1;
            Excel.Range resultRange = targetRange.Worksheet.Cells[2, targetCol] as Excel.Range;
            resultRange = resultRange.Resize[rowCount, colCount];
            resultRange.Value2 = results;

            return MatchResult.Success(matchedCount, unmatchedCount);
        }
    }

    #region 数据模型

    public class RangeAnalysis
    {
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public object[,] Data { get; set; }
        public bool HasHeader { get; set; }
    }

    public class ColumnCandidate
    {
        public int ColumnIndex { get; set; }
        public string ColumnName { get; set; }
        public int Score { get; set; }
    }

    public class ColumnInfo
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public int Confidence { get; set; }
    }

    public class DataSource
    {
        public string SheetName { get; set; }
        public int MatchColumnIndex { get; set; }
        public string MatchColumnName { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
    }

    public class MatchSuggestion
    {
        public string SourceSheet { get; set; }
        public int SourceMatchColumn { get; set; }
        public int TargetKeyColumn { get; set; }
        public List<int> ReturnColumns { get; set; }
    }

    public class MatchResult
    {
        public bool Success { get; set; }
        public int MatchedCount { get; set; }
        public int UnmatchedCount { get; set; }
        public long Duration { get; set; }
        public string ErrorMessage { get; set; }

        public static MatchResult Success(int matched, int unmatched)
        {
            return new MatchResult
            {
                Success = true,
                MatchedCount = matched,
                UnmatchedCount = unmatched
            };
        }

        public static MatchResult Error(string message)
        {
            return new MatchResult
            {
                Success = false,
                ErrorMessage = message
            };
        }

        public static MatchResult NoSourceFound()
        {
            return new MatchResult
            {
                Success = false,
                ErrorMessage = "未找到可匹配的数据源"
            };
        }
    }

    #endregion
}
```

---

## ▶️ 步骤八：运行和调试

### 1. 首次运行

1. 按 **F5** 启动调试
2. Visual Studio会自动：
   - 编译项目
   - 注册插件
   - 启动Excel
   - 加载插件

### 2. 验证插件加载

在Excel中：
1. 查看顶部功能区
2. 应该能看到 **"效率助手"** 选项卡
3. 点击可以看到你添加的按钮

### 3. 调试技巧

#### 设置断点
```csharp
public void OnSmartMatch(IRibbonControl control)
{
    // 在这里设置断点 ← 点击左侧边栏添加红点
    var engine = new DataMatcherEngine();
    // ...
}
```

#### 查看日志
日志文件位置：
```
C:\Users\你的用户名\AppData\Roaming\ExcelEfficiencyAssistant\Logs\
```

#### 实时监视
在调试时：
1. 调试 → 窗口 → 即时窗口
2. 可以输入变量名查看值
3. 可以执行C#代码

---

## 📦 步骤九：测试功能

### 创建测试数据

在Excel中创建两个工作表：

#### Sheet1（订单表）- 需要匹配的数据
```
| 订单号  | 日期       | 数量 |
|---------|-----------|------|
| A001    | 2024-01-01| 100  |
| A002    | 2024-01-02| 200  |
| A003    | 2024-01-03| 150  |
```

#### Sheet2（产品表）- 数据源
```
| 订单号  | 产品名称 | 单价 |
|---------|----------|------|
| A001    | 键盘     | 99   |
| A002    | 鼠标     | 59   |
| A003    | 显示器   | 999  |
```

### 测试智能匹配

1. 选中Sheet1的数据
2. 点击 **"智能匹配"** 按钮
3. 应该自动将产品名称和单价匹配填充到Sheet1

---

## 🎉 完成！

现在你已经有了一个基础的VSTO Excel插件！

### 下一步可以做什么：

1. ✅ 添加更多Ribbon按钮
2. ✅ 实现表格美化功能
3. ✅ 实现文本处理功能
4. ✅ 创建任务窗格
5. ✅ 添加对话框界面
6. ✅ 打包发布

---

## 📝 常见问题和解决方案

### Q: 找不到"Excel VSTO Add-in"模板？
**A:** 需要安装"Office/SharePoint开发"工作负载
1. 打开Visual Studio Installer
2. 点击"修改"
3. 勾选"Office/SharePoint开发"
4. 点击"修改"安装

### Q: 编译错误：找不到Microsoft.Office.Interop.Excel？
**A:** 添加COM引用
1. 右键项目 → 添加 → 引用
2. COM → 类型库
3. 找到"Microsoft Excel 16.0 Object Library"
4. 勾选并确定

### Q: Excel启动但看不到插件？
**A:** 检查信任中心设置
1. Excel → 文件 → 选项 → 信任中心
2. 信任中心设置 → 加载项
3. 取消勾选"要求应用程序加载项由受信任的发布者签名"
4. 重启Excel

### Q: 如何卸载插件？
**A:**
1. 控制面板 → 程序和功能
2. 找到"ExcelEfficiencyAssistant"
3. 右键卸载

---

## 🚀 准备好了吗？

现在打开Visual Studio 2022，按照上面的步骤创建你的第一个VSTO Excel插件吧！

有任何问题随时查看这个指南或查阅官方文档：
- https://docs.microsoft.com/zh-cn/visualstudio/vsto/

**祝开发顺利！** 🎉
