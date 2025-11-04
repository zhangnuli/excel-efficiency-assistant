# GitHub Codespace 使用说明

## ⚠️ 重要提示

**当前遇到的问题**：
- VSTO和Office Interop组件**只能在Windows环境**中编译和运行
- GitHub Codespace基于Linux，**无法编译包含Office COM引用的代码**
- 所有包含`Microsoft.Office.Interop.Excel`的文件都会导致编译错误

## 🎯 解决方案

### 方案A：在Codespace中开发算法（推荐）

**适用于**：核心算法开发和测试

在Codespace中，我们只运行独立的测试程序：

```bash
# 运行不依赖Excel的测试程序
cd /workspaces/excel-efficiency-assistant
dotnet run
```

**功能**：
- ✅ 测试核心算法逻辑
- ✅ 开发数据处理算法
- ✅ 编写文档
- ✅ 版本控制和协作
- ❌ 无法编译完整的VSTO插件
- ❌ 无法测试Excel集成

### 方案B：使用本地Windows环境

**适用于**：完整VSTO开发

需要安装Visual Studio 2022（可以安装到D盘）：

1. **安装VS 2022到D盘**
   ```
   - 下载Visual Studio Installer
   - 在安装时选择安装位置为 D:\Program Files\
   - 选择工作负载：Office/SharePoint开发
   ```

2. **打开项目**
   ```
   - 从GitHub克隆代码
   - 用VS 2022打开 ExcelEfficiencyAssistant.sln
   - 完整编译和调试VSTO插件
   ```

### 方案C：混合开发模式（最佳实践）

结合两种方式的优势：

1. **在Codespace中**：
   - 开发核心算法
   - 编写文档
   - 版本控制
   - 团队协作

2. **在本地Windows中**：
   - 集成到VSTO
   - 测试Excel功能
   - 打包发布

## 📁 项目文件说明

### 可以在Codespace中运行的文件

```
✅ Program.cs            - 测试程序（无Excel依赖）
✅ README.md            - 项目文档
✅ GitHub-Codespace开发指南.md
✅ .devcontainer/       - Codespace配置
```

### 需要Windows环境的文件

```
❌ src/Core/DataMatcher/DataMatcherEngine.cs      - 使用Excel.Range
❌ src/Core/Beautifier/TableBeautifier.cs         - 使用Excel.Range
❌ src/Core/TextProcessor/TextProcessor.cs        - 使用Excel.Range
❌ Ribbon/EfficiencyRibbon.cs                     - 使用Office.IRibbonUI
❌ ThisAddIn.cs                                   - VSTO入口点
❌ Services/SettingsManager.cs                    - 部分功能依赖Windows
❌ Services/LogService.cs                         - 部分功能依赖Windows
```

## 🚀 在Codespace中快速开始

### 1. 运行测试程序

```bash
# 当前项目只运行Program.cs
dotnet run
```

**预期输出**：
```
🚀 Excel效率助手 Pro - Codespace版本
=====================================

版本: v1.0.0
运行环境: Unix 6.0.428
...
✅ 程序运行完成！
```

### 2. 开发独立算法

创建不依赖Excel的算法文件：

```bash
# 创建独立算法目录
mkdir -p src-standalone/Algorithms

# 编写纯算法代码（不使用Excel对象）
code src-standalone/Algorithms/StringMatcher.cs
```

### 3. 编写单元测试

```bash
# 创建测试项目
dotnet new xunit -o Tests
cd Tests
dotnet add reference ../ExcelEfficiencyAssistant.csproj
```

## 🔧 调整项目以适配Codespace

### 临时方案：排除VSTO文件

修改`.csproj`文件，排除Windows专用文件：

```xml
<ItemGroup>
  <!-- 排除VSTO相关文件 -->
  <Compile Remove="src/Core/**/*.cs" />
  <Compile Remove="Ribbon/**/*.cs" />
  <Compile Remove="ThisAddIn.cs" />
  <Compile Remove="Services/**/*.cs" />
</ItemGroup>
```

### 长期方案：分离项目

创建两个项目：
1. **ExcelEfficiencyAssistant.Core** - 纯算法（跨平台）
2. **ExcelEfficiencyAssistant.VSTO** - VSTO集成（仅Windows）

## 📊 当前状态总结

| 环境 | 可用功能 | 限制 |
|------|---------|------|
| **GitHub Codespace** | ✅ 算法开发<br>✅ 文档编写<br>✅ 版本控制 | ❌ 无法编译VSTO<br>❌ 无法测试Excel集成 |
| **Windows本地** | ✅ 完整VSTO开发<br>✅ Excel集成测试<br>✅ 打包发布 | ❌ 需要安装VS<br>❌ 需要C盘空间 |

## 💡 推荐的工作流程

1. **在Codespace中**：
   ```bash
   # 开发算法
   git pull
   code src-standalone/
   dotnet test
   git commit && git push
   ```

2. **同步到本地Windows**：
   ```bash
   # 在本地
   git pull
   # 在VS中打开项目
   # 集成算法到VSTO
   # 编译测试
   ```

3. **发布**：
   ```
   在Windows本地打包VSTO插件
   发布到GitHub Releases
   ```

## 🎯 下一步建议

### 立即可行：
1. ✅ 使用Codespace编写文档
2. ✅ 设计算法架构
3. ✅ 编写独立测试代码

### 需要Windows环境：
1. ❌ 编译完整VSTO项目
2. ❌ 测试Excel集成
3. ❌ 打包安装程序

## ❓ 常见问题

**Q: 为什么不能在Codespace中编译VSTO？**
A: VSTO依赖Windows特有的COM组件和.NET Framework，Linux环境无法支持。

**Q: 我的C盘空间不够怎么办？**
A: Visual Studio 2022可以安装到D盘或其他盘符，安装时选择自定义路径即可。

**Q: 可以完全在Codespace中开发吗？**
A: 可以开发核心算法和编写文档，但最终编译VSTO插件必须在Windows环境中进行。

**Q: 有没有纯云端的解决方案？**
A: 可以考虑使用Windows虚拟机服务（如Azure Windows VM），但需要付费。

---

**总结**：GitHub Codespace适合算法开发和文档编写，完整的VSTO开发需要本地Windows环境。