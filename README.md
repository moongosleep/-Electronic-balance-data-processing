# Electronic Balance Data Processing

电子天平数据处理工具。项目面向 Windows 环境，使用 Microsoft Excel COM 自动化读取电子天平导出的 `.xls` 文件，并生成新的 `.xlsm` 结果文件。

## 项目特点

- 输入格式：`.xls`
- 输出格式：`.xlsm`
- 不覆盖原始输入文件
- 优先使用 Excel COM 自动化，适合处理旧版 Excel `.xls` 文件
- 单文件主程序，便于直接运行和打包
- 支持图形界面选择文件
- 支持纯水和污染物两类计算流程
- 输出文件名会按处理模式自动追加后缀
- 当目标输出文件已存在时，会自动追加时间戳，避免覆盖已有结果

## 仓库文件

```text
.
├── process_excel-1.py   # 主程序
├── requirements.txt     # Python 依赖
├── test_excels.xls      # 示例输入文件
├── AGENTS.md            # 项目开发约束说明
├── .gitignore           # Git 忽略规则
└── README.md            # 使用说明
```

说明：本仓库不上传本地批量自检脚本、运行日志、测试报告、缓存目录、Excel 输出结果和 PyInstaller 打包产物。

## 运行环境

请在 Windows 电脑上运行本工具，并确保已经安装：

- Windows 10 或 Windows 11
- Python 3.10 或更高版本
- Microsoft Excel 桌面版
- `pywin32`

本工具依赖 Microsoft Excel COM，因此不适合在没有 Excel 桌面版的环境中运行，例如普通 Linux 服务器、无桌面环境的 CI 环境等。

## 安装 Python

如果电脑还没有安装 Python，请从 Python 官网下载安装：

https://www.python.org/downloads/windows/

安装时建议勾选：

```text
Add python.exe to PATH
```

安装完成后，在 PowerShell 中检查：

```powershell
py --version
```

## 下载项目

可以通过 Git 克隆仓库：

```powershell
git clone https://github.com/moongosleep/-Electronic-balance-data-processing.git
cd -Electronic-balance-data-processing
```

也可以在 GitHub 页面点击 `Code`，下载 ZIP 后解压。

## 安装依赖

进入项目目录后执行：

```powershell
py -m pip install -r requirements.txt
```

如果安装 `pywin32` 后仍出现 COM 相关错误，可以尝试执行：

```powershell
py -m pywin32_postinstall -install
```

## 启动程序

在项目目录中运行：

```powershell
py .\process_excel-1.py
```

程序启动后会弹出窗口，按界面提示选择电子天平导出的 `.xls` 文件并执行处理。

## 输入文件要求

输入文件应满足以下条件：

- 文件扩展名为 `.xls`
- 文件来自电子天平导出数据
- 需要处理的工作表名称应符合程序内置规则
- 原始数据主要位于 A/B 列
- A 列通常为时间信息
- B 列通常为质量或称量数据

如果输入文件结构不符合要求，程序会通过错误提示说明失败原因。

## 输出结果

程序会在不覆盖原始文件的前提下生成新的 `.xlsm` 文件。

输出文件名通常基于输入文件名和处理模式生成，例如：

```text
原文件名-纯水计算后.xlsm
原文件名-污染物计算后.xlsm
```

如果同名输出文件已经存在，程序会自动追加时间戳，例如：

```text
原文件名-纯水计算后-20260509_120000.xlsm
```

## 常见运行方式

直接运行脚本：

```powershell
py .\process_excel-1.py
```

如果系统中有多个 Python 版本，也可以指定 Python 启动器版本：

```powershell
py -3 .\process_excel-1.py
```

## PyInstaller 打包

如果需要把脚本打包成 Windows 可执行文件，可以先安装 PyInstaller：

```powershell
py -m pip install pyinstaller
```

然后执行打包命令：

```powershell
py -m PyInstaller --onefile --noconsole .\process_excel-1.py
```

打包完成后，可执行文件会生成在：

```text
dist\process_excel-1.exe
```

可以把 `dist` 目录中的 `.exe` 文件复制到其他 Windows 电脑运行。目标电脑仍需要安装 Microsoft Excel，因为程序需要调用 Excel COM。

## 打包参数说明

- `--onefile`：打包为单个 `.exe` 文件
- `--noconsole`：运行时不显示命令行窗口，适合图形界面程序
- `.\process_excel-1.py`：主程序入口文件

如果需要保留命令行窗口查看报错，可以去掉 `--noconsole`：

```powershell
py -m PyInstaller --onefile .\process_excel-1.py
```

## 本地生成文件

以下文件或目录属于本地运行产物，不会提交到 GitHub：

- `__pycache__/`
- `debug_run_log.txt`
- `batch_test_report.csv`
- `batch_test_report.txt`
- `batch_test_outputs/`
- `*.xlsm`
- `build/`
- `dist/`
- `*.spec`

## 常见问题

### 提示找不到 pywin32 或 win32com

先确认依赖已经安装：

```powershell
py -m pip install -r requirements.txt
```

如果仍然报错，执行：

```powershell
py -m pywin32_postinstall -install
```

### 提示 Excel COM 相关错误

请确认当前电脑已经安装 Microsoft Excel 桌面版，并且 Excel 能够正常打开。

### 双击 exe 没有明显反应

如果使用了 `--noconsole` 打包，错误信息可能不会显示。建议临时使用以下命令重新打包，保留控制台窗口查看报错：

```powershell
py -m PyInstaller --onefile .\process_excel-1.py
```

### 输出文件没有覆盖原文件

这是预期行为。程序设计目标就是不覆盖原始 `.xls` 文件。如果同名结果文件已存在，程序会自动生成带时间戳的新文件。

## 开发说明

项目当前优先保证 Windows 下可运行，主逻辑集中在 `process_excel-1.py` 中。后续如果需要扩展规则，建议先明确输入表结构、计算规则和输出格式，再修改主脚本。
