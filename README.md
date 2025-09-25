# Excel Function Maker | Excel函数制作器

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.6%2B-blue?style=flat-square&logo=python" alt="Python Version">
  <img src="https://img.shields.io/badge/Platform-Windows%7CmacOS%7CLinux-lightgrey?style=flat-square" alt="Platform">
  <img src="https://img.shields.io/badge/License-MIT-green?style=flat-square" alt="License">
  <img src="https://img.shields.io/badge/Language-Chinese%7CEnglish-red?style=flat-square" alt="Language">
</p>

A user-friendly GUI tool for generating Excel functions with Chinese/English bilingual support.

一个用户友好的Excel函数生成工具，支持中英文双语界面。

## 🌟 Features | 功能特性

### 🎯 Core Features | 核心功能
- **13 Excel Functions** | **13种Excel函数支持**
  - Mathematical: SUM, AVERAGE, COUNT
  - Conditional: COUNTIF, SUMIF, IF  
  - Lookup: VLOOKUP, HLOOKUP
  - Text: CONCATENATE, LEFT, RIGHT, MID
  - Advanced: CHAR_COUNT (Character Statistics)

### 🌐 Bilingual Interface | 双语界面
- **Chinese/English Switch** | **中英文切换**
- **Auto-save Language Preference** | **自动保存语言偏好**
- **Complete Bilingual Documentation** | **完整双语文档**

### 💡 Smart Features | 智能功能
- **Real-time Preview** | **实时预览**
- **One-click Copy** | **一键复制**
- **Auto Quote Detection** | **智能引号添加**
- **Cell Reference Validation** | **单元格引用验证**

## 🚀 Quick Start | 快速开始

### Installation | 安装

```bash
# Clone the repository | 克隆仓库
git clone https://github.com/yourusername/excel-function-maker.git

# Navigate to project directory | 进入项目目录
cd excel-function-maker

# Install dependencies | 安装依赖
pip install -r requirements.txt

# Run the application | 运行应用
python excel_function_maker.py
```

### For Windows Users | Windows用户
```bash
# Use the batch file | 使用批处理文件
run.bat
```

## 📖 Usage | 使用方法

### Basic Usage | 基本用法
1. **Select Function** | **选择函数**: Choose from the dropdown menu
2. **Input Parameters** | **输入参数**: Fill in the required parameters
3. **Preview Result** | **预览结果**: See real-time function generation
4. **Copy to Excel** | **复制到Excel**: One-click copy to clipboard

### Language Switch | 语言切换
- Click the language selector in the top-right corner
- Choose "中文" or "English"
- Interface updates instantly

在右上角点击语言选择器，选择"中文"或"English"，界面立即切换。

## 📋 Supported Functions | 支持的函数

| Function | 中文名称 | Description | 描述 |
|----------|----------|-------------|------|
| SUM | 求和函数 | Calculate sum of range | 计算范围总和 |
| AVERAGE | 平均值函数 | Calculate average | 计算平均值 |
| COUNT | 计数函数 | Count numbers | 计算数字个数 |
| COUNTIF | 条件计数 | Conditional count | 条件计数 |
| SUMIF | 条件求和 | Conditional sum | 条件求和 |
| IF | 条件判断 | Conditional logic | 条件判断 |
| VLOOKUP | 垂直查找 | Vertical lookup | 垂直查找 |
| HLOOKUP | 水平查找 | Horizontal lookup | 水平查找 |
| CONCATENATE | 文本连接 | Join text strings | 连接文本 |
| LEFT | 左取字符 | Left characters | 左侧字符 |
| RIGHT | 右取字符 | Right characters | 右侧字符 |
| MID | 中间字符 | Middle characters | 中间字符 |
| CHAR_COUNT | 字符统计 | Character statistics | 字符统计 |

## ⭐ Special Feature: CHAR_COUNT | 特色功能：字符统计

The CHAR_COUNT function generates complex character counting formulas:

CHAR_COUNT函数可以生成复杂的字符统计公式：

**Example | 示例:**
```excel
="a"&SUMPRODUCT(LEN(A1:J1)-LEN(SUBSTITUTE(A1:J1,"a","")))&"b"&SUMPRODUCT(LEN(A1:J1)-LEN(SUBSTITUTE(A1:J1,"b","")))
```

**Output | 输出:** `a3b2` (meaning "a" appears 3 times, "b" appears 2 times)

## 🛠️ Development | 开发

### Project Structure | 项目结构
```
excel-function-maker/
├── excel_function_maker.py    # Main application | 主程序
├── requirements.txt           # Dependencies | 依赖文件
├── README.md                 # Documentation | 项目文档
├── run.bat                   # Windows launcher | Windows启动脚本
├── docs/                     # Additional docs | 附加文档
└── screenshots/              # Screenshots | 截图
```

### Building Executable | 构建可执行文件

To create a standalone executable:

创建独立可执行文件：

```bash
# Install PyInstaller | 安装PyInstaller
pip install pyinstaller

# Build executable | 构建可执行文件
pyinstaller --onefile --windowed --name "Excel函数制作器" excel_function_maker.py
```

**Note**: Executable files are not included in this repository due to GitHub's file size limits.

**注意**: 由于GitHub文件大小限制，可执行文件未包含在此仓库中。

## 🎨 Screenshots | 截图

### Chinese Interface | 中文界面
![Chinese Interface](screenshots/chinese_interface.png)

### English Interface | 英文界面  
![English Interface](screenshots/english_interface.png)

### Function Generation | 函数生成
![Function Generation](screenshots/function_generation.png)

## 🤝 Contributing | 贡献

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

欢迎贡献代码！请查看我们的[贡献指南](CONTRIBUTING.md)了解详情。

### How to Contribute | 如何贡献
1. Fork the repository | Fork仓库
2. Create a feature branch | 创建功能分支
3. Make your changes | 做出更改
4. Add tests if applicable | 添加测试（如适用）
5. Submit a pull request | 提交拉取请求

## 📝 License | 许可证

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

本项目采用MIT许可证 - 查看[LICENSE](LICENSE)文件了解详情。

## 🙏 Acknowledgments | 致谢

- Built with Python and Tkinter | 使用Python和Tkinter构建
- Icon design inspired by Excel | 图标设计灵感来源于Excel
- Thanks to all contributors | 感谢所有贡献者

## 📊 Statistics | 统计

![GitHub stars](https://img.shields.io/github/stars/yourusername/excel-function-maker?style=social)
![GitHub forks](https://img.shields.io/github/forks/yourusername/excel-function-maker?style=social)
![GitHub issues](https://img.shields.io/github/issues/yourusername/excel-function-maker)
![GitHub pull requests](https://img.shields.io/github/issues-pr/yourusername/excel-function-maker)

---

<p align="center">
Made with ❤️ by [Your Name] | 由[您的姓名]用❤️制作
</p>

<p align="center">
<a href="#top">Back to Top | 返回顶部</a>
</p>

