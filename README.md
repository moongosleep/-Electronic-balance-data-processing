# 电子天平数据处理工具

这是一个 Windows 下使用的 Python Excel 处理工具，用于处理电子天平导出的原始 `.xls` 数据，并生成新的 `.xlsm` 结果文件。程序优先使用 Microsoft Excel COM 自动化，因此可以处理旧版 `.xls` 文件。

## 功能特点

- 输入文件：`.xls`
- 输出文件：新的 `.xlsm`
- 不覆盖原始输入文件
- 支持纯水和污染物两种实验模式
- 通过图形界面选择模式、输入参数和选择文件
- 污染物模式会生成 G:M 结果区、O:V 分析区和原有散点图工作表
- 污染物模式新增用户指定平均点散点分析：结果写入 AA:AB，并新建单独散点图工作表
- 输出文件名会按处理模式自动追加后缀；如果同名输出文件已存在，会自动追加时间戳，避免覆盖已有结果

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

## 使用流程

纯水模式：

1. 选择“纯水”。
2. 输入时间s和膜面积cm2。
3. 选择 `.xls` 原始数据文件。
4. 输入前半段压力和后半段压力。
5. 程序生成纯水计算结果并另存为新的 `.xlsm` 文件。

污染物模式：

1. 选择“污染物”。
2. 输入时间s和膜面积cm2。
3. 选择 `.xls` 原始数据文件。
4. 输入运行压力bar。
5. 输入纯水渗透性LMH/bar。
6. 在“散点平均点设置”窗口输入“每几分钟作为一个平均点”和“一共需要几个平均点”。
7. 程序生成原有污染物结果，并额外生成 AA:AB 用户指定平均点分析区和独立散点图工作表。

如果在新增的“散点平均点设置”窗口点击取消或关闭窗口，程序会跳过新增 AA/AB 和新散点图，但仍继续生成原有污染物结果。

## 输入文件要求

输入文件需要满足：

- 文件扩展名为 `.xls`
- 文件来自电子天平导出的原始数据
- 工作簿中需要有名为“电子天平输出”的工作表
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

## 污染物模式新增平均点散点分析

新增功能只在污染物模式中生效，纯水模式不受影响。

输入校验：

- “每几分钟作为一个平均点”必须是大于 0 的数字，可以是整数或小数，例如 `2`、`2.5`。
- “一共需要几个平均点”必须是大于 0 的整数，例如 `5`。
- 每组点数按 `每几分钟作为一个平均点 * 60 / 时间s` 计算。
- 如果分钟数无法被当前采样时间整除，会提示“平均分钟数无法被当前采样时间整除，请重新输入。”，并回到同一个输入窗口重新填写。

AA:AB 输出：

- Z 列保留为空白间隔。
- AA2：`{总分钟数}min-{平均点数}散点-渗透性`
- AB2：`{总分钟数}min-{平均点数}散点-归一化通量`
- AA3：用户输入的纯水渗透性LMH/bar。
- AB3：`=AA3/$AA$3`
- AA4 起：从 C3 开始按用户指定点数连续分组，对每组 C 列液滴质量差求平均后计算渗透性。
- AB4 起：按 `AA行/$AA$3` 归一化。

数据不足时，程序会报错并停止本次处理：

```text
当前C列有效数据不足，无法生成指定数量的平均点，请检查原始数据长度。
```

新增散点图：

- 工作表名称：`{总分钟数}min-{平均点数}散点图`
- 图表标题：`{总分钟数}min-{平均点数}散点`
- 图表数据来自 AB 列归一化通量，包含 AB3 的基准点 1。
- X 轴为序号 `1, 2, 3, ...`，Y 轴为归一化通量。

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
text-excel.xls
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

### 污染物模式没有弹出新增窗口

请确认运行的是当前项目目录下的 `process_excel-1.py`，路径应为：

```text
D:\AI\codex-all\excel处理工具\process_excel-1.py
```

新增窗口只在选择“污染物”后出现，并且位于运行压力和纯水渗透性LMH/bar输入之后。

### 双击 exe 没有明显反应

如果使用 `--noconsole` 打包，错误信息可能不会显示。建议临时使用以下命令重新打包，保留控制台窗口查看报错：

```powershell
py -m PyInstaller --onefile .\process_excel-1.py
```

### 输出文件没有覆盖原文件

这是预期行为。程序设计目标就是不覆盖原始 `.xls` 文件。如果同名结果文件已经存在，程序会自动生成带时间戳的新文件。
