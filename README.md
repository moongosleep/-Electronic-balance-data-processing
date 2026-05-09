# Electronic balance data processing

Windows 下使用 Microsoft Excel COM 自动化处理电子天平导出的 `.xls` 文件，并输出新的 `.xlsm` 文件，不覆盖原文件。

## 文件说明

- `process_excel-1.py`：主程序，保持单文件实现。
- `batch_self_check.py`：批量自检脚本，用于对示例输入执行纯水/污染物两种模式检查。
- `test_excels.xls`：示例输入文件。

运行后生成的 `.xlsm`、测试报告、调试日志、`__pycache__`、`build/`、`dist/` 等文件不会提交到仓库。

## 环境要求

- Windows
- Python 3.10 或更高版本
- 已安装 Microsoft Excel
- Python 依赖：`pywin32`

## 安装依赖

```powershell
py -m pip install -r requirements.txt
```

如果安装后仍提示 COM 相关错误，可以执行：

```powershell
py -m pywin32_postinstall -install
```

## 运行

```powershell
py .\process_excel-1.py
```

程序会通过窗口选择输入 `.xls` 文件，并在不覆盖原文件的前提下输出新的 `.xlsm` 文件。

## 批量自检

```powershell
py .\batch_self_check.py
```

自检会读取 `test_excels.xls` 或 `test_excels` 目录下的 Excel 文件，并生成本地测试输出与报告。

## PyInstaller 打包

先安装打包依赖：

```powershell
py -m pip install pyinstaller
```

打包为单文件 GUI 程序：

```powershell
py -m PyInstaller --onefile --noconsole .\process_excel-1.py
```

打包完成后，可执行文件位于 `dist` 目录。`dist` 目录属于生成产物，不提交到仓库。
