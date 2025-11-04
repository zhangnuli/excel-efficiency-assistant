# Windows服务器 VSTO开发完整指南

## 🎯 概述

有了Windows服务器，您可以进行完整的VSTO开发！本指南将帮您在Windows服务器上搭建开发环境。

## 📋 前置要求

### 服务器配置建议
- **操作系统**：Windows Server 2016/2019/2022 或 Windows 10/11
- **内存**：8GB+ （推荐16GB）
- **硬盘**：至少20GB可用空间（用于VS安装）
- **Office版本**：Excel 2016/2019/2021/365

### 网络要求
- 能够访问GitHub
- 能够下载NuGet包
- （可选）远程桌面访问

---

## 🚀 第一步：安装Visual Studio 2022

### 1. 下载安装程序

访问：https://visualstudio.microsoft.com/zh-hans/downloads/

选择：**Visual Studio 2022 Community**（免费版）

或者使用PowerShell下载：
```powershell
# 下载VS 2022 Community安装程序
Invoke-WebRequest -Uri "https://aka.ms/vs/17/release/vs_community.exe" -OutFile "D:\Downloads\vs_installer.exe"
```

### 2. 自定义安装路径

```
运行安装程序后：
1. 点击"修改"或"安装"
2. 在"安装位置"选项卡中
3. 选择安装路径，例如：
   D:\Program Files\Microsoft Visual Studio\2022\Community
```

### 3. 选择工作负载

**必需勾选**：
- ✅ **.NET 桌面开发**
  - .NET Framework 4.8 开发工具
  - C# 和 Visual Basic

- ✅ **Office/SharePoint 开发** ⭐（最重要！）
  - Office 开发工具
  - VSTO（Visual Studio Tools for Office）
  - Office 互操作程序集

**可选勾选**：
- ASP.NET 和 Web 开发
- 数据存储和处理

### 4. 等待安装完成

- 预计需要 30-60 分钟
- 需要下载约 10-15GB 数据
- 安装后重启计算机

---

## 🔧 第二步：配置开发环境

### 1. 安装Git

```powershell
# 下载Git for Windows
Invoke-WebRequest -Uri "https://github.com/git-for-windows/git/releases/download/v2.43.0.windows.1/Git-2.43.0-64-bit.exe" -OutFile "D:\Downloads\git-installer.exe"

# 或访问：https://git-scm.com/download/win
```

安装时选择默认选项即可。

### 2. 配置Git

```powershell
# 打开PowerShell或命令提示符
git config --global user.name "您的名字"
git config --global user.email "您的邮箱"

# 配置SSH（可选，用于免密推送）
ssh-keygen -t ed25519 -C "your_email@example.com"
# 将 ~/.ssh/id_ed25519.pub 添加到GitHub SSH Keys
```

### 3. 安装Excel（如果未安装）

VSTO开发需要本地安装Excel：
- Office 2016 或更高版本
- Office 365

---

## 📦 第三步：克隆和打开项目

### 1. 克隆GitHub仓库

```powershell
# 选择一个项目目录
cd D:\Projects

# 克隆仓库
git clone https://github.com/zhangnuli/excel-efficiency-assistant.git
cd excel-efficiency-assistant

# 查看文件
dir
```

### 2. 用Visual Studio打开项目

**方法A - 通过命令行**：
```powershell
# 打开解决方案文件
start ExcelEfficiencyAssistant.sln
```

**方法B - 通过VS界面**：
```
1. 启动 Visual Studio 2022
2. 点击"打开项目或解决方案"
3. 浏览到：D:\Projects\excel-efficiency-assistant\ExcelEfficiencyAssistant.sln
4. 点击"打开"
```

---

## 🔨 第四步：修复项目配置

当前项目配置是为Codespace优化的，需要调整为Windows VSTO模式。

### 1. 修改项目文件

打开 `ExcelEfficiencyAssistant.csproj`，修改为：

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net6.0-windows</TargetFramework>
    <OutputType>Library</OutputType>
    <UseWindowsForms>true</UseWindowsForms>
    <AssemblyTitle>Excel效率助手 Pro</AssemblyTitle>
    <AssemblyDescription>Excel数据处理和美化工具</AssemblyDescription>
    <AssemblyVersion>1.0.0.0</AssemblyVersion>
    <Copyright>Copyright © 2025 Excel效率助手团队</Copyright>
  </PropertyGroup>

  <ItemGroup>
    <!-- Office Interop 引用 -->
    <PackageReference Include="Microsoft.Office.Interop.Excel" Version="15.0.4795.1001" />
    <PackageReference Include="Microsoft.Vbe.Interop" Version="15.0.4795.1001" />

    <!-- 其他依赖 -->
    <PackageReference Include="System.Text.Json" Version="7.0.0" />
    <PackageReference Include="System.Drawing.Common" Version="7.0.0" />
    <PackageReference Include="Microsoft.Extensions.Logging" Version="7.0.0" />
  </ItemGroup>

</Project>
```

### 2. 还原NuGet包

在Visual Studio中：
```
工具 -> NuGet包管理器 -> 程序包管理器控制台

执行命令：
Update-Package -reinstall
```

或者在解决方案资源管理器中：
```
右键解决方案 -> 还原 NuGet 包
```

---

## 🏗️ 第五步：构建项目

### 1. 设置启动项目

```
1. 在解决方案资源管理器中
2. 右键 ExcelEfficiencyAssistant 项目
3. 选择"设为启动项目"
```

### 2. 构建解决方案

```
生成 -> 生成解决方案 (Ctrl+Shift+B)
```

**预期结果**：
```
生成成功
========== 生成: 1 成功，0 失败，0 最新，0 跳过 ==========
```

### 3. 处理编译错误

如果遇到错误，常见问题和解决方案：

#### 错误1：缺少Office引用
```
错误：找不到类型或命名空间名称"Office"
解决：安装 Microsoft.Office.Interop.Excel NuGet包
```

#### 错误2：目标框架不匹配
```
错误：项目需要 .NET Framework 4.8
解决：修改 TargetFramework 为 net48 或 net6.0-windows
```

#### 错误3：VSTO运行时缺失
```
错误：需要Visual Studio Tools for Office Runtime
解决：下载安装 VSTO Runtime
https://aka.ms/VSTORuntimeDownload
```

---

## 🧪 第六步：调试和测试

### 1. 调试配置

在Visual Studio中：
```
1. 点击工具栏的调试按钮（或按F5）
2. 首次运行会提示选择调试器
3. 选择"Excel"作为宿主应用程序
```

### 2. 设置断点

```
1. 在代码行号左侧点击，设置断点（红点）
2. 按F5启动调试
3. Excel会自动打开并加载插件
4. 触发功能时会在断点处暂停
```

### 3. 测试功能

```
1. Excel打开后，查看"开发工具"或"加载项"选项卡
2. 应该能看到"Excel效率助手"功能区
3. 测试各个功能按钮
4. 查看输出窗口的日志信息
```

---

## 📊 第七步：性能优化和发布

### 1. 发布配置

```
生成 -> 配置管理器
选择"Release"配置
```

### 2. 发布VSTO插件

```
生成 -> 发布 ExcelEfficiencyAssistant

配置发布选项：
- 发布位置：D:\Publish\ExcelEfficiencyAssistant
- 安装URL：（可选）网络共享路径
- 更新策略：检查更新频率
```

### 3. 生成安装程序

Visual Studio会生成：
- setup.exe - 安装程序
- .vsto文件 - 清单文件
- Application Files 文件夹 - 应用程序文件

---

## 🔄 开发工作流程

### 日常开发循环

```bash
# 1. 拉取最新代码
git pull

# 2. 在VS中开发和测试
# （编写代码、调试、测试）

# 3. 提交更改
git add .
git commit -m "功能描述"
git push

# 4. 定期发布测试版本
# 生成 -> 发布
```

### 与Codespace协作

```
在Codespace中：
- 开发核心算法
- 编写文档
- 版本控制

在Windows服务器中：
- 集成到VSTO
- Excel环境测试
- 打包发布
```

---

## ❓ 常见问题

### Q1: Visual Studio可以安装到D盘吗？
**A:** 可以！安装时选择自定义路径即可。

### Q2: 需要安装哪个版本的Office？
**A:** Excel 2016及以上任何版本都可以，Office 365也支持。

### Q3: 如何远程访问Windows服务器？
**A:** 使用Windows远程桌面连接（mstsc）或第三方工具如TeamViewer。

### Q4: 编译错误如何解决？
**A:**
1. 检查是否安装了Office/SharePoint开发工作负载
2. 确认NuGet包已正确还原
3. 查看输出窗口的详细错误信息

### Q5: 如何将插件部署给其他用户？
**A:**
1. 使用VS发布功能生成安装程序
2. 将setup.exe分发给用户
3. 用户运行setup.exe即可安装

---

## 🎯 下一步

现在您可以：

1. ✅ **立即开始**：在Windows服务器上安装Visual Studio 2022
2. ✅ **克隆项目**：从GitHub获取代码
3. ✅ **修复配置**：调整项目文件以支持VSTO
4. ✅ **开始开发**：实现核心功能
5. ✅ **测试调试**：在真实Excel环境中测试
6. ✅ **打包发布**：生成安装程序

---

## 📞 获取帮助

- **Visual Studio文档**：https://docs.microsoft.com/zh-cn/visualstudio/
- **VSTO文档**：https://docs.microsoft.com/zh-cn/visualstudio/vsto/
- **Office开发文档**：https://docs.microsoft.com/zh-cn/office/dev/add-ins/

---

**祝您开发顺利！** 🚀

*有了Windows服务器，您现在拥有了完整的VSTO开发能力！*