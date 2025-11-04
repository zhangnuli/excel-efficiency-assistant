# Excel 效率助手

一个强大的 Excel 插件，提供数据匹配等快捷功能，帮助提高工作效率。

## 主要功能

### 🎯 数据匹配功能
- 类似 VLOOKUP 的可视化操作界面
- 在两个工作表之间快速匹配和填充数据
- 自动处理匹配失败的情况
- 支持大规模数据处理

## 技术栈

- **Office.js**: Office Add-in API
- **React 18**: 用户界面框架
- **TypeScript**: 类型安全的开发
- **Fluent UI**: Microsoft 官方 UI 组件库
- **Webpack**: 模块打包工具

## 开发环境要求

- Node.js 14.x 或更高版本
- npm 或 yarn
- Excel (Windows/Mac/Online)

## 安装依赖

```bash
npm install
```

## 开发

### 启动开发服务器

```bash
npm run dev-server
```

### 在 Excel 中加载插件

1. 首次运行需要安装开发证书：
```bash
npx office-addin-dev-certs install
```

2. 启动插件：
```bash
npm start
```

这将自动打开 Excel 并加载插件。

### 构建生产版本

```bash
npm run build
```

## 使用说明

### 数据匹配功能

1. **选择源数据表**：包含要查找数据的工作表
2. **选择目标数据表**：需要填充数据的工作表
3. **选择匹配字段**：两个表中用于匹配的关键列（如 ID、编号）
4. **选择填充字段**：从源表中获取的数据列
5. **点击"开始匹配"**：系统会自动在目标表中添加新列并填充匹配的数据

### 示例场景

假设你有两个工作表：

**员工基本信息表**（源表）：
| 员工ID | 姓名 | 部门 | 职位 |
|--------|------|------|------|
| E001   | 张三 | 技术部 | 工程师 |
| E002   | 李四 | 销售部 | 经理 |

**员工考勤表**（目标表）：
| 员工ID | 日期 | 出勤状态 |
|--------|------|----------|
| E001   | 2024-01-01 | 正常 |
| E002   | 2024-01-01 | 正常 |

使用数据匹配功能，可以快速将"部门"或"姓名"等信息从基本信息表匹配填充到考勤表中。

## 项目结构

```
excel插件/
├── src/
│   ├── taskpane/
│   │   ├── components/
│   │   │   ├── App.tsx          # 主应用组件
│   │   │   └── DataMatcher.tsx  # 数据匹配组件
│   │   ├── taskpane.html        # 任务窗格 HTML
│   │   ├── taskpane.tsx         # 任务窗格入口
│   │   └── taskpane.css         # 样式文件
│   └── commands/
│       ├── commands.html
│       └── commands.ts
├── assets/                      # 图标资源
├── manifest.xml                 # Office Add-in 清单文件
├── webpack.config.js            # Webpack 配置
├── tsconfig.json               # TypeScript 配置
└── package.json                # 项目依赖
```

## 故障排除

### 插件无法加载
- 确保开发服务器正在运行（`npm run dev-server`）
- 检查是否安装了开发证书
- 尝试清除 Office 缓存

### 数据匹配失败
- 确保源表和目标表都有数据
- 检查匹配字段的数据格式是否一致
- 查看控制台错误信息

## 许可证

MIT

