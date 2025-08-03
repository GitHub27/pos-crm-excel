# POS CRM Electron 应用

这是一个基于 Electron 构建的 POS CRM 数据导入工具，用于处理和管理客户数据的 Excel 文件。

## 功能特性

- 📊 **Excel 文件导入**: 支持 .xlsx 和 .xls 格式
- 📤 **数据导出**: 将处理后的数据导出为 Excel 文件
- 🎨 **现代化界面**: 美观的用户界面设计
- 📈 **数据统计**: 实时显示数据统计信息
- 🖱️ **拖拽上传**: 支持拖拽文件上传
- 📋 **数据预览**: 表格形式预览导入的数据

## 安装和运行

### 前置要求

- Node.js (版本 14 或更高)
- npm 或 yarn

### 安装依赖

```bash
# 进入项目目录
cd electron-app

# 安装依赖
npm install
```

### 运行应用

```bash
# 开发模式运行
npm run dev

# 生产模式运行
npm start
```

### 构建应用

```bash
# 构建可执行文件
npm run build

# 构建并生成安装包
npm run dist
```

## 使用说明

### 导入数据

1. 点击「导入Excel文件」按钮或直接拖拽文件到上传区域
2. 选择要导入的 Excel 文件（.xlsx 或 .xls 格式）
3. 系统会自动解析文件并显示数据预览
4. 查看统计信息了解数据质量

### 导出数据

1. 导入数据后，「导出Excel文件」按钮会被激活
2. 点击按钮选择保存位置
3. 系统会生成包含所有数据的 Excel 文件

### 菜单功能

- **文件菜单**:
  - 导入Excel (Ctrl/Cmd + O)
  - 导出Excel (Ctrl/Cmd + S)
  - 退出应用 (Ctrl/Cmd + Q)

- **编辑菜单**: 标准的编辑操作（撤销、重做、复制、粘贴等）

- **视图菜单**: 缩放、全屏、开发者工具等

- **帮助菜单**: 关于信息

## 项目结构

```
electron-app/
├── main.js          # 主进程文件
├── index.html       # 应用界面
├── renderer.js      # 渲染进程脚本
├── package.json     # 项目配置
└── README.md        # 说明文档
```

## 技术栈

- **Electron**: 跨平台桌面应用框架
- **Node.js**: JavaScript 运行时
- **XLSX**: Excel 文件处理库
- **HTML/CSS/JavaScript**: 前端技术

## 开发说明

### 主要文件说明

- `main.js`: Electron 主进程，负责窗口管理、菜单创建、文件操作等
- `index.html`: 应用的用户界面
- `renderer.js`: 渲染进程脚本，处理用户交互和数据显示

### 自定义配置

可以在 `package.json` 中的 `build` 字段修改应用的构建配置：

- `appId`: 应用标识符
- `productName`: 产品名称
- `directories.output`: 构建输出目录

## 故障排除

### 常见问题

1. **应用无法启动**
   - 确保已安装所有依赖：`npm install`
   - 检查 Node.js 版本是否符合要求

2. **Excel 文件无法导入**
   - 确认文件格式为 .xlsx 或 .xls
   - 检查文件是否损坏
   - 确保文件没有被其他程序占用

3. **构建失败**
   - 清除 node_modules 并重新安装：`rm -rf node_modules && npm install`
   - 检查系统是否有足够的磁盘空间

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request 来改进这个项目。

## 更新日志

### v1.0.0
- 初始版本发布
- 基本的 Excel 导入导出功能
- 现代化用户界面
- 数据统计和预览功能