# 电子天平数据处理工具

这是一个 Windows 下使用的 Python Excel 处理工具，用于处理电子天平导出的原始 `.xls` 数据，并生成新的 `.xlsm` 结果文件。程序优先使用 Microsoft Excel COM 自动化，因此可以处理旧版 `.xls` 文件。

## 功能特点

- 输入文件：`.xls`
- 输出文件：新的 `.xlsm`
- 不覆盖原始输入文件
- 支持纯水和污染物两种实验模式
- 通过图形界面选择模式、输入参数和选择文件
- 输出文件名会按处理模式自动追加后缀
- 如果同名输出文件已存在，会自动追加时间戳，避免覆盖已有结果

## 当前仓库文件

```text
.
├── process_excel-1.py   # 主程序
├── requirements.txt     # Python 依赖
├── 使用说明.doc          # 使用说明文档
├── README.md            # 本说明
├── AGENTS.md            # 开发约束说明
└── .gitignore           # Git 忽略规则
```

处理原始数据时，核心只需要：

```text
process_excel-1.py
requirements.txt
```

## 运行环境

请在 Windows 电脑上运行，并确保已安装：

- Windows 10 或 Windows 11
- Python 3.10 或更高版本
- Microsoft Excel 桌面版
- `pywin32`

本工具依赖 Excel COM，因此不适合在没有 Microsoft Excel 桌面版的环境中运行。

## 安装依赖

进入项目目录后执行：

```powershell
py -m pip install -r requirements.txt
```

如果安装 `pywin32` 后仍出现 COM 相关错误，可以尝试：

```powershell
py -m pywin32_postinstall -install
```

## 启动程序

在项目目录中运行：

```powershell
py .\process_excel-1.py
```

程序启动后会弹出窗口，按提示选择实验模式、输入参数，并选择电子天平导出的 `.xls` 原始数据文件。

## 输入文件要求

输入文件需要满足：

- 文件扩展名为 `.xls`
- 文件来自电子天平导出的原始数据
- 原始数据主要位于 A/B 两列
- A 列通常为时间信息
- B 列通常为质量或称量数据

如果输入文件结构不符合要求，程序会给出错误提示。

## 输出结果

程序会生成新的 `.xlsm` 文件，不会覆盖原始 `.xls` 文件。

输出文件名示例：

```text
原文件名-纯水计算后.xlsm
原文件名-污染物计算后.xlsm
```

如果同名输出文件已存在，会自动追加时间戳，例如：

```text
原文件名-纯水计算后-20260510_120000.xlsm
```

## PyInstaller 打包

如果需要打包成 Windows 可执行文件，先安装 PyInstaller：

```powershell
py -m pip install pyinstaller
```

然后执行：

```powershell
py -m PyInstaller --onefile --noconsole .\process_excel-1.py
```

打包完成后，可执行文件位于：

```text
dist\process_excel-1.exe
```

目标电脑仍然需要安装 Microsoft Excel 桌面版，因为程序需要调用 Excel COM。

如果需要保留命令行窗口查看报错，可以去掉 `--noconsole`：

```powershell
py -m PyInstaller --onefile .\process_excel-1.py
```

## 本地生成文件

以下文件或目录属于本地运行、测试或打包产物，不需要上传到 GitHub：

```text
__pycache__/
debug_run_log.txt
*.xlsm
build/
dist/
*.spec
```

## 常见问题

### 提示找不到 pywin32 或 win32com

先确认依赖已安装：

```powershell
py -m pip install -r requirements.txt
```

如果仍然报错，执行：

```powershell
py -m pywin32_postinstall -install
```

### 提示 Excel COM 相关错误

请确认当前电脑已经安装 Microsoft Excel 桌面版，并且 Excel 可以正常打开。

### 双击 exe 没有明显反应

如果使用 `--noconsole` 打包，错误信息可能不会显示。建议临时使用以下命令重新打包，保留控制台窗口查看报错：

```powershell
py -m PyInstaller --onefile .\process_excel-1.py
```

### 输出文件没有覆盖原文件

这是预期行为。程序设计目标就是不覆盖原始 `.xls` 文件。如果同名结果文件已经存在，程序会自动生成带时间戳的新文件。
