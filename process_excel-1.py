# -*- coding: utf-8 -*-
"""
process_excel.py

Windows 下 Excel 实验数据处理工具。

依赖安装：
    py -m pip install pywin32

如果安装 pywin32 后仍提示 COM 相关错误，可尝试：
    py -m pywin32_postinstall -install

可选打包依赖：
    py -m pip install pyinstaller

打包示例：
    pyinstaller --onefile --noconsole process_excel.py

设计说明：
    1. 本脚本优先并实际使用 Microsoft Excel COM 自动化处理文件。
    2. 不使用 openpyxl 读取 .xls。openpyxl 不支持旧版 .xls 格式，且本需求要求
       输入 .xls、输出 .xlsm，并且用户电脑已安装 Excel，使用 Excel COM 更符合场景。
    3. 输出文件保存前会检查目标文件是否存在；如果存在，会自动追加时间戳，
       避免覆盖已有结果文件。

假设点：
    - 需求说明输入文件只有一个工作表，但错误处理只明确要求“找不到
      电子天平输出 工作表”时报错。本脚本只处理该工作表；如果工作簿里
      还有其他工作表，会原样保留但不参与计算。
    - A 列时间只用于保留和格式校验，不参与后续通量公式。时间校验接受 Excel
      日期/时间值、Excel 序列数，以及常见日期时间文本。
    - 输入文件按需求应只有 A/B 两列原始数据。输出时本脚本会生成并覆盖
      C:E 与 G:M 计算区域，以保证结果区没有旧内容残留。
"""

from __future__ import annotations

import atexit
import math
import os
import sys
import traceback
import unicodedata
from datetime import date, datetime, time as datetime_time
from decimal import Decimal, InvalidOperation
from numbers import Real
from pathlib import Path
from typing import Callable, Iterable, List, NamedTuple, Optional, Sequence, Tuple


try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog
except ImportError:
    tk = None
    filedialog = None
    messagebox = None
    simpledialog = None


PYWIN32_IMPORT_ERROR = None
try:
    import pythoncom
    import pywintypes
    import win32com.client as win32
except ImportError as exc:
    PYWIN32_IMPORT_ERROR = exc
    pythoncom = None
    pywintypes = None
    win32 = None


# =========================
# 可配置参数
# =========================

SHEET_NAME = "电子天平输出"

MODE_PURE_WATER = "纯水"
MODE_POLLUTANT = "污染物"

# 输出文件名：原文件名 + 模式后缀 + .xlsm
PURE_WATER_OUTPUT_NAME_SUFFIX = "-纯水计算后"
POLLUTANT_OUTPUT_NAME_SUFFIX = "-污染物计算后"
OUTPUT_EXTENSION = ".xlsm"
MAX_OUTPUT_CANDIDATE_COUNTER = 1000

# 以后如果“每 5 个取平均”或“每 10 个取平均”的规则变化，只需要改这里。
AVG_GROUP_1 = 5
AVG_GROUP_2 = 10

# 右侧结果区固定参数。
DEFAULT_TIME_SECONDS = 4.0
DEFAULT_MEMBRANE_AREA_CM2 = 12.57

# 压力切换点自动识别规则：
# 至少前面已有 5 个 H 值时，才允许判断；
# 当前连续 3 个 H 值平均值 < 前面所有 H 值平均值的 85%。
PRESSURE_SWITCH_MIN_PREVIOUS_COUNT = 5
PRESSURE_SWITCH_CURRENT_COUNT = 3
PRESSURE_SWITCH_RATIO = 0.85

# 最终校验规则：
# 只检查当前 K 列等于后半段压力的结果行；
# 阈值 = 当前所有有效 M 值平均值 * 1.15。
FINAL_PERMEABILITY_OUTLIER_RATIO = 1.15


# =========================
# Excel COM 常量
# =========================

# Excel 文件格式：52 = xlOpenXMLWorkbookMacroEnabled，即 .xlsm。
XL_FILE_FORMAT_XLSM = 52

# Excel 方向与插入常量。
XL_UP = -4162
XL_SHIFT_DOWN = -4121

# Excel 自动计算。
XL_CALCULATION_AUTOMATIC = -4105

# Excel 图表常量。
XL_XY_SCATTER = -4169
XL_CATEGORY_AXIS = 1
XL_VALUE_AXIS = 2


# =========================
# 列号常量
# =========================

COL_TIME = 1
COL_WEIGHT = 2
COL_DIFF = 3
COL_AVG_GROUP_1 = 4
COL_AVG_GROUP_2 = 5

COL_RESULT_NO = 7
COL_RESULT_WEIGHT = 8
COL_RESULT_TIME = 9
COL_RESULT_AREA = 10
COL_RESULT_PRESSURE = 11
COL_RESULT_FLUX = 12
COL_RESULT_PERMEABILITY = 13

COL_POLLUTANT_5_POINT_PERMEABILITY = 15
COL_POLLUTANT_5_POINT_NORMALIZED = 16
COL_POLLUTANT_15_GROUP_PERMEABILITY = 17
COL_POLLUTANT_15_GROUP_NORMALIZED = 18
COL_POLLUTANT_10_GROUP_PERMEABILITY = 19
COL_POLLUTANT_10_GROUP_NORMALIZED = 20
COL_POLLUTANT_5_GROUP_PERMEABILITY = 21
COL_POLLUTANT_5_GROUP_NORMALIZED = 22
COL_STANDARD_10MIN_PERMEABILITY = 24
COL_STANDARD_10MIN_NORMALIZED = 25
COL_CUSTOM_AVERAGE_PERMEABILITY = 27
COL_CUSTOM_AVERAGE_NORMALIZED = 28

STANDARD_10MIN_REQUIRED_COUNT = 150
STANDARD_10MIN_GROUP_SIZE = 30
STANDARD_10MIN_GROUP_COUNT = 5


LEFT_HEADERS = {
    COL_TIME: "时间",
    COL_WEIGHT: "每隔四秒的液滴重量",
    COL_DIFF: "液滴质量差",
    COL_AVG_GROUP_1: "每五个取平均",
    COL_AVG_GROUP_2: "每十个取平均",
}

RIGHT_HEADERS = {
    COL_RESULT_NO: "编号",
    COL_RESULT_WEIGHT: "重量g",
    COL_RESULT_TIME: "时间s",
    COL_RESULT_AREA: "膜面积cm2",
    COL_RESULT_PRESSURE: "运行压力bar",
    COL_RESULT_FLUX: "通量LMH",
    COL_RESULT_PERMEABILITY: "渗透性LMHbar-1",
}

POLLUTANT_ANALYSIS_HEADERS = {
    COL_POLLUTANT_5_POINT_PERMEABILITY: "5点渗透性",
    COL_POLLUTANT_5_POINT_NORMALIZED: "5点归一化通量",
    COL_POLLUTANT_15_GROUP_PERMEABILITY: "15组渗透性",
    COL_POLLUTANT_15_GROUP_NORMALIZED: "15组归一化通量",
    COL_POLLUTANT_10_GROUP_PERMEABILITY: "10组渗透性",
    COL_POLLUTANT_10_GROUP_NORMALIZED: "10组归一化通量",
    COL_POLLUTANT_5_GROUP_PERMEABILITY: "5组渗透性",
    COL_POLLUTANT_5_GROUP_NORMALIZED: "5组归一化通量",
}

POLLUTANT_STANDARD_10MIN_HEADERS = {
    COL_STANDARD_10MIN_PERMEABILITY: "标准十分钟渗透性",
    COL_STANDARD_10MIN_NORMALIZED: "标准十分钟归一化通量",
}


TIME_TEXT_FORMATS = (
    "%H:%M:%S",
    "%H:%M",
    "%Y-%m-%d %H:%M:%S",
    "%Y/%m/%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y/%m/%d %H:%M",
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%m/%d/%Y %H:%M:%S",
    "%m/%d/%Y %H:%M",
    "%m/%d/%Y",
)


class ExcelProcessError(Exception):
    """业务处理错误：用于给用户显示清晰、可理解的错误信息。"""


class UserCancelledError(Exception):
    """用户主动取消文件选择或压力输入。"""


class SourceDataInfo(NamedTuple):
    """原始数据校验后的关键信息。"""

    last_b_row: int
    calculation_start_row: int


class PressureSettings(NamedTuple):
    """\u4e24\u6bb5\u8fd0\u884c\u538b\u529b\u8bbe\u7f6e\u3002"""

    first_segment_pressure: float
    second_segment_pressure: float


class RuntimeSettings(NamedTuple):
    """\u8fd0\u884c\u65f6\u8f93\u5165\u7684\u516c\u5171\u5b9e\u9a8c\u53c2\u6570\u3002"""

    time_seconds: float
    membrane_area_cm2: float


class CustomAverageSettings(NamedTuple):
    """污染物模式用户指定平均点散点分析配置。"""

    minutes_per_point: float
    point_count: int
    points_per_group: int
    total_minutes: float
    name_prefix: str


class TestModeOptions(NamedTuple):
    """\u6279\u91cf\u81ea\u67e5\u65f6\u7528\u4e8e\u6ce8\u5165\u7684\u975e\u4ea4\u4e92\u53c2\u6570\u3002"""

    experiment_mode: str
    runtime_settings: RuntimeSettings
    pressure_settings: Optional[PressureSettings] = None
    pressure_bar: Optional[float] = None
    pure_water_flux: Optional[float] = None
    manual_switch_index: Optional[int] = None
    output_directory: Optional[Path] = None
    allow_xlsm_input: bool = False
    show_messages: bool = False
    custom_average_settings: Optional[CustomAverageSettings] = None

class OutputLayoutResult(NamedTuple):
    """输出布局生成后的摘要信息。"""

    result_count: int
    pressure_switch_index: int
    pressure_switch_auto_detected: bool


class PollutantAnalysisResult(NamedTuple):
    """污染物模式 O:V 分析区生成摘要。"""

    five_point_count: int
    fifteen_group_count: int
    ten_group_count: int
    five_group_count: int


def _build_mode_detail(
    mode: str,
    runtime_settings: RuntimeSettings,
    layout_result: OutputLayoutResult,
    **kwargs: object,
) -> str:
    """构造完成提示中的实验模式详情。"""
    if mode == MODE_PURE_WATER:
        if layout_result.pressure_switch_index > 0:
            switch_source = "自动识别" if layout_result.pressure_switch_auto_detected else "手动输入"
            return (
                f"实验模式：{MODE_PURE_WATER}\n"
                f"时间s：{runtime_settings.time_seconds:g}\n"
                f"膜面积cm2：{runtime_settings.membrane_area_cm2:g}\n"
                f"结果行数：{layout_result.result_count}\n"
                f"压力切换编号：{layout_result.pressure_switch_index}（{switch_source}）"
            )

        return (
            f"实验模式：{MODE_PURE_WATER}\n"
            f"时间s：{runtime_settings.time_seconds:g}\n"
            f"膜面积cm2：{runtime_settings.membrane_area_cm2:g}\n"
            f"结果行数：{layout_result.result_count}\n"
            "压力模式：单一压力"
        )

    if mode == MODE_POLLUTANT:
        pressure_bar = kwargs["pressure_bar"]
        pure_water_flux = kwargs["pure_water_flux"]
        return (
            f"实验模式：{MODE_POLLUTANT}\n"
            f"时间s：{runtime_settings.time_seconds:g}\n"
            f"膜面积cm2：{runtime_settings.membrane_area_cm2:g}\n"
            f"结果行数：{layout_result.result_count}\n"
            f"运行压力bar：{pressure_bar:g}\n"
            f"纯水通量：{pure_water_flux:g}"
        )

    raise ExcelProcessError(f"未知实验模式：{mode}")


_TK_ROOT = None

DEBUG_LOG_PATH = Path(__file__).resolve().with_name("debug_run_log.txt")
DEBUG_LOG_BUFFER_LIMIT = 20
_DEBUG_LOG_BUFFER: List[str] = []


def debug_repr(value) -> str:
    """把调试值转成短文本，避免日志被超长内容撑满。"""
    try:
        text = str(value)
    except Exception:
        text = repr(value)

    if len(text) > 500:
        text = text[:497] + "..."
    return text


def flush_debug_log_buffer() -> None:
    """将缓冲的调试日志一次性写入文件，减少频繁 open/close。"""
    if not _DEBUG_LOG_BUFFER:
        return

    try:
        DEBUG_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with DEBUG_LOG_PATH.open("a", encoding="utf-8") as handle:
            handle.write("".join(_DEBUG_LOG_BUFFER))
        _DEBUG_LOG_BUFFER.clear()
    except Exception:
        pass


atexit.register(flush_debug_log_buffer)


def debug_log_event(event: str, **fields) -> None:
    """
    写入增量排查日志。

    日志只用于确认手动运行、批量自查、打包 exe 是否走同一份代码和同一套参数。
    写日志失败不影响主流程，避免调试功能反过来干扰 Excel 处理。
    """
    try:
        lines = [
            f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {event}",
            f"  __file__={Path(__file__).resolve()}",
            f"  sys.executable={sys.executable}",
            f"  cwd={os.getcwd()}",
        ]
        for key, value in fields.items():
            lines.append(f"  {key}={debug_repr(value)}")
        _DEBUG_LOG_BUFFER.append("\n".join(lines) + "\n\n")
        if len(_DEBUG_LOG_BUFFER) >= DEBUG_LOG_BUFFER_LIMIT:
            flush_debug_log_buffer()
    except Exception:
        pass


def get_cell_debug_state(worksheet, row: int, column: int) -> str:
    """读取单元格的值和公式，用于判断公式是否真实写入。"""
    try:
        cell = worksheet.Cells(row, column)
        value = cell.Value
        formula = cell.Formula
        return f"Value={debug_repr(value)}; Formula={debug_repr(formula)}"
    except Exception as exc:
        return f"<read failed: {debug_repr(exc)}>"


def get_tk_root():
    """创建一个隐藏的 Tk 根窗口，用于文件选择、输入框和消息框。"""
    global _TK_ROOT

    if tk is None:
        return None

    if _TK_ROOT is None:
        _TK_ROOT = tk.Tk()
        _TK_ROOT.withdraw()
        try:
            _TK_ROOT.attributes("-topmost", True)
        except Exception:
            # 某些环境不支持 topmost，不影响核心功能。
            pass

    return _TK_ROOT


def _force_window_to_foreground(root: object) -> None:
    """尽量把 Tk 窗口带到 Windows 前台。"""
    try:
        import ctypes

        hwnd = int(root.winfo_id())
        ctypes.windll.user32.ShowWindow(hwnd, 9)  # SW_RESTORE
        ctypes.windll.user32.SetForegroundWindow(hwnd)
    except Exception:
        pass


def _center_and_focus_window(window: object, min_width: int, min_height: int) -> None:
    """居中 Tk 窗口，并尽量把它聚焦到前台。"""
    window.update_idletasks()
    width = max(window.winfo_reqwidth(), min_width)
    height = max(window.winfo_reqheight(), min_height)
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = max((screen_width - width) // 2, 0)
    y = max((screen_height - height) // 2, 0)
    window.geometry(f"{width}x{height}+{x}+{y}")

    try:
        window.lift()
        window.focus_force()
        window.update()

        # VS Code 集成终端有时不会把 Tk 窗口主动带到前台。
        # 因此这里在 Windows 下额外调用 Win32 API，尽量把模式选择窗口推到前台。
        _force_window_to_foreground(window)
    except Exception:
        pass


def show_user_message(title: str, message: str, is_error: bool = False) -> None:
    """同时兼顾命令行和双击运行：打印信息，并尽量弹出消息框。"""
    stream = sys.stderr if is_error else sys.stdout
    try:
        print(f"{title}: {message}", file=stream)
    except Exception:
        pass

    root = get_tk_root()
    if root is not None and messagebox is not None:
        try:
            if is_error:
                messagebox.showerror(title, message, parent=root)
            else:
                messagebox.showinfo(title, message, parent=root)
        except Exception:
            pass


def select_experiment_mode() -> str:
    """
    弹出模式选择界面，让用户二选一：纯水 / 污染物。

    这里使用 tkinter Checkbutton，界面上会显示勾选状态；两个选项共用
    同一个 StringVar，因此同一时间只能选中一个模式。
    """
    global _TK_ROOT

    if tk is None:
        raise ExcelProcessError("当前 Python 环境没有可用的 tkinter，无法显示模式选择界面。")

    selected = {"mode": None, "cancelled": False}

    # 最稳妥的做法：直接用 Tk 主窗口本身显示模式选择。
    # 不使用“隐藏 root + Toplevel”，避免窗口挂在隐藏父窗口后面导致用户看不到。
    if _TK_ROOT is None:
        root = tk.Tk()
        _TK_ROOT = root
    else:
        root = _TK_ROOT

    for child in root.winfo_children():
        child.destroy()

    root.title("选择实验模式")
    root.resizable(False, False)
    root.deiconify()

    mode_var = tk.StringVar(master=root, value="")

    try:
        root.attributes("-topmost", True)
    except Exception:
        pass

    tk.Label(
        root,
        text="请选择实验模式：",
        anchor="center",
        font=("Microsoft YaHei UI", 11),
    ).grid(row=0, column=0, columnspan=2, padx=24, pady=(18, 10))

    pure_check = tk.Checkbutton(
        root,
        text=MODE_PURE_WATER,
        variable=mode_var,
        onvalue=MODE_PURE_WATER,
        offvalue="",
        indicatoron=True,
        width=12,
        anchor="w",
    )
    pure_check.grid(row=1, column=0, padx=(24, 12), pady=10, sticky="w")

    pollutant_check = tk.Checkbutton(
        root,
        text=MODE_POLLUTANT,
        variable=mode_var,
        onvalue=MODE_POLLUTANT,
        offvalue="",
        indicatoron=True,
        width=12,
        anchor="w",
    )
    pollutant_check.grid(row=1, column=1, padx=(12, 24), pady=10, sticky="w")

    def confirm() -> None:
        mode = mode_var.get()
        if mode not in (MODE_PURE_WATER, MODE_POLLUTANT):
            if messagebox is not None:
                messagebox.showwarning("请选择模式", "请先勾选“纯水”或“污染物”。", parent=root)
            return
        selected["mode"] = mode
        root.quit()

    def cancel() -> None:
        selected["cancelled"] = True
        root.quit()

    tk.Button(root, text="确定", width=12, command=confirm).grid(
        row=2,
        column=0,
        columnspan=2,
        pady=(8, 18),
    )
    root.protocol("WM_DELETE_WINDOW", cancel)

    _center_and_focus_window(root, 320, 150)

    print("正在显示模式选择窗口；如果没有看到，请查看任务栏或按 Alt+Tab 切换窗口。")
    root.mainloop()

    try:
        root.attributes("-topmost", False)
        root.withdraw()
    except Exception:
        pass

    if selected["cancelled"] or selected["mode"] is None:
        raise UserCancelledError("已取消选择实验模式，程序已终止。")

    return str(selected["mode"])


def normalize_text(value: str) -> str:
    """把全角数字、全角小数点等字符规范化，方便用户输入。"""
    return unicodedata.normalize("NFKC", value).strip()


def strip_outer_quotes(value: str) -> str:
    """处理手动输入路径时可能带上的首尾引号。"""
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
        return value[1:-1]
    return value


def parse_input_args(argv: Sequence[str]) -> Optional[Path]:
    """
    解析拖拽或命令行参数。

    用法 1：拖拽一个 .xls 到脚本或 exe 上，argv 中会带文件路径。
    用法 2：没有参数时返回 None，后续弹出文件选择框。
    """
    args = list(argv[1:])

    if not args:
        return None

    if args[0] in ("-h", "--help", "/?"):
        print_usage()
        raise UserCancelledError("已显示帮助，未处理文件。")

    if len(args) > 1:
        raise ExcelProcessError("一次只支持处理一个 Excel 文件，请只拖入一个 .xls 文件。")

    return Path(strip_outer_quotes(args[0])).expanduser()


def print_usage() -> None:
    """打印简单用法。"""
    usage = (
        "用法：\n"
        "  1. 拖拽 .xls 文件到 process_excel.py 或打包后的 exe 上。\n"
        "  2. 双击运行，然后在文件选择框中选择 .xls 文件。\n"
        "\n"
        "依赖安装：\n"
        "  py -m pip install pywin32\n"
        "\n"
        "打包示例：\n"
        "  py -m pip install pyinstaller\n"
        "  pyinstaller --onefile --noconsole process_excel.py\n"
    )
    print(usage)


def select_input_file() -> Path:
    """没有拖拽文件时，弹出文件选择框；无 GUI 时退回命令行输入。"""
    root = get_tk_root()

    if root is not None and filedialog is not None:
        selected = filedialog.askopenfilename(
            parent=root,
            title="请选择原始 Excel .xls 文件",
            filetypes=[
                ("Excel 97-2003 工作簿 (*.xls)", "*.xls"),
                ("所有文件", "*.*"),
            ],
        )
        if not selected:
            raise UserCancelledError("已取消选择文件，程序未处理任何文件。")
        return Path(selected)

    # 极少数环境没有 tkinter 时使用命令行输入。
    text = input("请输入 .xls 文件完整路径，直接回车取消：").strip()
    if not text:
        raise UserCancelledError("已取消输入文件路径，程序未处理任何文件。")
    return Path(strip_outer_quotes(text)).expanduser()


def validate_input_file(input_path: Path, allow_xlsm_input: bool = False) -> Path:
    """?????????????????????? Excel ???"""
    try:
        resolved = input_path.resolve(strict=False)
    except Exception:
        resolved = input_path

    if not resolved.exists():
        raise ExcelProcessError(f"????????{resolved}")

    if not resolved.is_file():
        raise ExcelProcessError(f"?????????{resolved}")

    allowed_suffixes = {".xls"}
    if allow_xlsm_input:
        allowed_suffixes.add(".xlsm")

    if resolved.suffix.lower() not in allowed_suffixes:
        expected_text = ".xls ? .xlsm" if allow_xlsm_input else ".xls"
        raise ExcelProcessError(
            f"??????? {expected_text} ???????????{resolved.suffix or '???'}"
        )

    return resolved


def parse_positive_float(text: str, value_name: str = "????") -> float:
    """???????????"""
    normalized = normalize_text(text)

    if "," in normalized and "." not in normalized:
        normalized = normalized.replace(",", ".")

    try:
        value = float(normalized)
    except ValueError as exc:
        raise ExcelProcessError(f"{value_name}?????????1.5") from exc

    if not math.isfinite(value) or value <= 0:
        raise ExcelProcessError(f"{value_name}????? 0 ????")

    return value


def parse_positive_int(text: str, value_name: str) -> int:
    """解析必须大于 0 的整数输入。"""
    normalized = normalize_text(text)

    try:
        value = int(normalized)
    except ValueError as exc:
        raise ExcelProcessError(f"{value_name}必须是大于 0 的整数。") from exc

    if value <= 0:
        raise ExcelProcessError(f"{value_name}必须是大于 0 的整数。")

    return value


def format_default_numeric_text(value: float) -> str:
    """\u628a\u9ed8\u8ba4\u6570\u503c\u683c\u5f0f\u5316\u4e3a\u66f4\u9002\u5408\u663e\u793a\u5728\u8f93\u5165\u6846\u91cc\u7684\u6587\u672c\u3002"""
    if float(value).is_integer():
        return str(int(value))
    return f"{value:g}"


def get_time_and_area_inputs() -> RuntimeSettings:
    """\u5728\u540c\u4e00\u4e2a\u7a97\u53e3\u4e2d\u83b7\u53d6\u65f6\u95f4s\u548c\u819c\u9762\u79efcm2\u3002"""
    root = get_tk_root()
    if root is None or tk is None:
        raise ExcelProcessError("\u5f53\u524d Python \u73af\u5883\u6ca1\u6709\u53ef\u7528\u7684 tkinter\uff0c\u65e0\u6cd5\u663e\u793a\u5b9e\u9a8c\u53c2\u6570\u8f93\u5165\u7a97\u53e3\u3002")

    selected = {"settings": None, "cancelled": False}
    dialog = tk.Toplevel(root)
    dialog.title("\u8bf7\u8f93\u5165\u5b9e\u9a8c\u53c2\u6570")
    dialog.resizable(False, False)
    dialog.grab_set()

    try:
        dialog.attributes("-topmost", True)
    except Exception:
        pass

    tk.Label(dialog, text="\u8bf7\u586b\u5199\u672c\u6b21\u5b9e\u9a8c\u53c2\u6570\uff1a", anchor="w").grid(
        row=0,
        column=0,
        columnspan=2,
        padx=18,
        pady=(16, 10),
        sticky="w",
    )
    tk.Label(dialog, text="\u65f6\u95f4s\uff1a", anchor="e", width=12).grid(
        row=1,
        column=0,
        padx=(18, 8),
        pady=6,
        sticky="e",
    )
    tk.Label(dialog, text="\u819c\u9762\u79efcm2\uff1a", anchor="e", width=12).grid(
        row=2,
        column=0,
        padx=(18, 8),
        pady=6,
        sticky="e",
    )

    time_var = tk.StringVar(master=dialog, value=format_default_numeric_text(DEFAULT_TIME_SECONDS))
    area_var = tk.StringVar(master=dialog, value=format_default_numeric_text(DEFAULT_MEMBRANE_AREA_CM2))

    time_entry = tk.Entry(dialog, textvariable=time_var, width=20)
    area_entry = tk.Entry(dialog, textvariable=area_var, width=20)
    time_entry.grid(row=1, column=1, padx=(0, 18), pady=6, sticky="we")
    area_entry.grid(row=2, column=1, padx=(0, 18), pady=6, sticky="we")

    button_frame = tk.Frame(dialog)
    button_frame.grid(row=3, column=0, columnspan=2, pady=(10, 16))

    def submit() -> None:
        try:
            settings = RuntimeSettings(
                time_seconds=parse_positive_float(time_var.get(), "\u65f6\u95f4s"),
                membrane_area_cm2=parse_positive_float(area_var.get(), "\u819c\u9762\u79efcm2"),
            )
        except ExcelProcessError as exc:
            if messagebox is not None:
                messagebox.showwarning("\u8f93\u5165\u65e0\u6548", str(exc), parent=dialog)
            return

        selected["settings"] = settings
        dialog.destroy()

    def cancel() -> None:
        selected["cancelled"] = True
        dialog.destroy()

    tk.Button(button_frame, text="\u786e\u5b9a", width=10, command=submit).pack(side="right", padx=(8, 0))
    tk.Button(button_frame, text="\u53d6\u6d88", width=10, command=cancel).pack(side="right")

    dialog.protocol("WM_DELETE_WINDOW", cancel)
    time_entry.bind("<Return>", lambda _event: submit())
    area_entry.bind("<Return>", lambda _event: submit())

    _center_and_focus_window(dialog, 320, 170)

    time_entry.focus_set()
    root.wait_window(dialog)

    if selected["cancelled"] or selected["settings"] is None:
        raise UserCancelledError("\u5df2\u53d6\u6d88\u8f93\u5165\u65f6\u95f4s\u548c\u819c\u9762\u79efcm2\uff0c\u7a0b\u5e8f\u5df2\u7ec8\u6b62\u3002")

    return selected["settings"]


def get_positive_float_input(
    dialog_title: str,
    prompt_text: str,
    value_name: str,
    cli_prompt_text: str,
) -> float:
    """\u83b7\u53d6\u4e00\u4e2a\u6b63\u6570\u8f93\u5165\uff0c\u53ef\u7528\u4e8e\u538b\u529b\u6216\u7eaf\u6c34\u901a\u91cf\u3002"""
    root = get_tk_root()

    if root is not None and simpledialog is not None:
        while True:
            text = simpledialog.askstring(dialog_title, prompt_text, parent=root)
            if text is None:
                raise UserCancelledError(f"\u5df2\u53d6\u6d88\u8f93\u5165{value_name}\uff0c\u7a0b\u5e8f\u5df2\u7ec8\u6b62\u3002")

            try:
                return parse_positive_float(text, value_name)
            except ExcelProcessError as exc:
                if messagebox is not None:
                    messagebox.showwarning("\u8f93\u5165\u65e0\u6548", str(exc), parent=root)

    while True:
        text = input(cli_prompt_text).strip()
        if not text:
            raise UserCancelledError(f"\u5df2\u53d6\u6d88\u8f93\u5165{value_name}\uff0c\u7a0b\u5e8f\u5df2\u7ec8\u6b62\u3002")
        try:
            return parse_positive_float(text, value_name)
        except ExcelProcessError as exc:
            print(f"\u8f93\u5165\u65e0\u6548\uff1a{exc}")


def get_pressure_input(label: str) -> float:
    """\u83b7\u53d6\u538b\u529b\u8f93\u5165\uff0c\u754c\u9762\u660e\u786e\u663e\u793a\u5355\u4f4d\u4e3a bar\u3002"""
    return get_positive_float_input(
        dialog_title="\u8fd0\u884c\u538b\u529b",
        prompt_text=f"\u8bf7\u8f93\u5165{label}\uff08bar\uff09\uff1a",
        value_name=label,
        cli_prompt_text=f"\u8bf7\u8f93\u5165{label}\uff08bar\uff09\uff0c\u76f4\u63a5\u56de\u8f66\u53d6\u6d88\uff1a",
    )


def get_pure_water_flux_input() -> float:
    """\u6c61\u67d3\u7269\u6a21\u5f0f\u4f7f\u7528\uff1a\u83b7\u53d6\u7528\u4e8e\u5f52\u4e00\u5316\u7684\u7eaf\u6c34\u901a\u91cf\u3002"""
    return get_positive_float_input(
        dialog_title="\u7eaf\u6c34\u901a\u91cf",
        prompt_text="\u8bf7\u8f93\u5165\u7eaf\u6c34\u901a\u91cf\uff1a",
        value_name="\u7eaf\u6c34\u901a\u91cf",
        cli_prompt_text="\u8bf7\u8f93\u5165\u7eaf\u6c34\u901a\u91cf\uff0c\u76f4\u63a5\u56de\u8f66\u53d6\u6d88\uff1a",
    )


def format_decimal_for_name(value: Decimal) -> str:
    """把 Decimal 格式化成适合写入表头和工作表名的短文本。"""
    normalized = value.normalize()
    if normalized == normalized.to_integral_value():
        return str(int(normalized))
    return format(normalized, "f").rstrip("0").rstrip(".")


def build_custom_average_settings(
    minutes_per_point: float,
    point_count: int,
    runtime_settings: RuntimeSettings,
) -> CustomAverageSettings:
    """根据用户输入计算每组点数，并生成 AA/AB 与图表共用的名称前缀。"""
    try:
        minutes_decimal = Decimal(str(minutes_per_point))
        time_decimal = Decimal(str(runtime_settings.time_seconds))
        total_minutes_decimal = minutes_decimal * Decimal(point_count)
        source_seconds_decimal = minutes_decimal * Decimal("60")
        remainder = source_seconds_decimal % time_decimal
    except (InvalidOperation, ZeroDivisionError) as exc:
        raise ExcelProcessError("平均分钟数无法被当前采样时间整除，请重新输入。") from exc

    if remainder != 0:
        raise ExcelProcessError("平均分钟数无法被当前采样时间整除，请重新输入。")

    points_per_group_decimal = source_seconds_decimal / time_decimal
    if points_per_group_decimal <= 0 or points_per_group_decimal != points_per_group_decimal.to_integral_value():
        raise ExcelProcessError("平均分钟数无法被当前采样时间整除，请重新输入。")

    points_per_group = int(points_per_group_decimal)
    total_minutes_text = format_decimal_for_name(total_minutes_decimal)
    name_prefix = f"{total_minutes_text}min-{point_count}散点"

    return CustomAverageSettings(
        minutes_per_point=minutes_per_point,
        point_count=point_count,
        points_per_group=points_per_group,
        total_minutes=float(total_minutes_decimal),
        name_prefix=name_prefix,
    )


def get_pollutant_custom_average_settings(
    runtime_settings: RuntimeSettings,
) -> Optional[CustomAverageSettings]:
    """
    获取污染物模式“用户指定平均点”配置。

    该窗口属于新增功能：用户取消或关闭时仅跳过新增 AA/AB 与新图表，
    污染物模式原有 G:M、O:V、X:Y 和既有散点图继续生成。
    """
    root = get_tk_root()

    if root is not None and tk is not None:
        selected = {"settings": None, "cancelled": False}
        dialog = tk.Toplevel(root)
        dialog.title("散点平均点设置")
        dialog.resizable(False, False)
        dialog.grab_set()

        try:
            dialog.attributes("-topmost", True)
        except Exception:
            pass

        tk.Label(dialog, text="请填写新增散点分析参数：", anchor="w").grid(
            row=0,
            column=0,
            columnspan=2,
            padx=18,
            pady=(16, 10),
            sticky="w",
        )
        tk.Label(dialog, text="每几分钟作为一个平均点：", anchor="e").grid(
            row=1,
            column=0,
            padx=(18, 8),
            pady=6,
            sticky="e",
        )
        tk.Label(dialog, text="一共需要几个平均点：", anchor="e").grid(
            row=2,
            column=0,
            padx=(18, 8),
            pady=6,
            sticky="e",
        )

        minutes_var = tk.StringVar(master=dialog)
        point_count_var = tk.StringVar(master=dialog)
        minutes_entry = tk.Entry(dialog, textvariable=minutes_var, width=18)
        point_count_entry = tk.Entry(dialog, textvariable=point_count_var, width=18)
        minutes_entry.grid(row=1, column=1, padx=(0, 18), pady=6, sticky="we")
        point_count_entry.grid(row=2, column=1, padx=(0, 18), pady=6, sticky="we")

        button_frame = tk.Frame(dialog)
        button_frame.grid(row=3, column=0, columnspan=2, pady=(10, 16))

        def submit() -> None:
            try:
                minutes_per_point = parse_positive_float(
                    minutes_var.get(),
                    "每几分钟作为一个平均点",
                )
                point_count = parse_positive_int(
                    point_count_var.get(),
                    "一共需要几个平均点",
                )
                settings = build_custom_average_settings(
                    minutes_per_point,
                    point_count,
                    runtime_settings,
                )
            except ExcelProcessError as exc:
                if messagebox is not None:
                    messagebox.showwarning("输入无效", str(exc), parent=dialog)
                return

            selected["settings"] = settings
            dialog.destroy()

        def cancel() -> None:
            selected["cancelled"] = True
            dialog.destroy()

        tk.Button(button_frame, text="确定", width=10, command=submit).pack(side="right", padx=(8, 0))
        tk.Button(button_frame, text="取消", width=10, command=cancel).pack(side="right")

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        minutes_entry.bind("<Return>", lambda _event: submit())
        point_count_entry.bind("<Return>", lambda _event: submit())

        _center_and_focus_window(dialog, 380, 180)

        minutes_entry.focus_set()
        root.wait_window(dialog)

        if selected["cancelled"]:
            return None
        return selected["settings"]

    while True:
        minutes_text = input("请输入每几分钟作为一个平均点，直接回车跳过新增散点分析：").strip()
        if not minutes_text:
            return None
        point_count_text = input("请输入一共需要几个平均点，直接回车跳过新增散点分析：").strip()
        if not point_count_text:
            return None

        try:
            minutes_per_point = parse_positive_float(minutes_text, "每几分钟作为一个平均点")
            point_count = parse_positive_int(point_count_text, "一共需要几个平均点")
            return build_custom_average_settings(
                minutes_per_point,
                point_count,
                runtime_settings,
            )
        except ExcelProcessError as exc:
            print(f"输入无效：{exc}")


def get_pressure_inputs() -> PressureSettings:
    """\u4f9d\u6b21\u83b7\u53d6\u524d\u534a\u6bb5\u538b\u529b\u548c\u540e\u534a\u6bb5\u538b\u529b\u3002"""
    first_pressure = get_pressure_input("\u524d\u534a\u6bb5\u538b\u529b")
    second_pressure = get_pressure_input("\u540e\u534a\u6bb5\u538b\u529b")
    return PressureSettings(
        first_segment_pressure=first_pressure,
        second_segment_pressure=second_pressure,
    )

def is_blank_value(value) -> bool:
    """判断 Excel 单元格值是否为空。"""
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def format_value_for_message(value) -> str:
    """把异常值格式化到错误信息中，避免信息过长。"""
    text = repr(value)
    if len(text) > 80:
        text = text[:77] + "..."
    return text


def _parse_float_cell_value(value, row_label: str) -> float:
    """按统一规则解析 Excel 单元格里的浮点数。"""
    if isinstance(value, bool):
        raise ExcelProcessError(f"{row_label}格式异常：{format_value_for_message(value)}")

    if isinstance(value, Real):
        number = float(value)
    elif isinstance(value, str):
        text = normalize_text(value)
        if "," in text and "." not in text:
            text = text.replace(",", ".")
        try:
            number = float(text)
        except ValueError as exc:
            raise ExcelProcessError(
                f"{row_label}格式异常：{format_value_for_message(value)}"
            ) from exc
    else:
        raise ExcelProcessError(f"{row_label}格式异常：{format_value_for_message(value)}")

    if not math.isfinite(number):
        raise ExcelProcessError(f"{row_label}不是有效数字：{format_value_for_message(value)}")

    return number


def parse_weight_value(value, row_number: int) -> float:
    """校验并解析 B 列重量。"""
    return _parse_float_cell_value(value, f"第 {row_number} 行 B 列重量")


def validate_time_value(value, row_number: int) -> None:
    """校验 A 列时间格式。"""
    if isinstance(value, bool):
        raise ExcelProcessError(f"第 {row_number} 行 A 列时间格式异常：{format_value_for_message(value)}")

    if isinstance(value, (datetime, date, datetime_time)):
        return

    # Excel 中时间常常是序列数，例如 0.5 表示中午 12 点。
    if isinstance(value, Real):
        if math.isfinite(float(value)):
            return
        raise ExcelProcessError(f"第 {row_number} 行 A 列时间不是有效数值：{format_value_for_message(value)}")

    if isinstance(value, str):
        text = normalize_text(value)
        if not text:
            raise ExcelProcessError(f"第 {row_number} 行 A 列时间为空。")

        # 有些设备可能把时间导出成纯数字文本。
        try:
            number = float(text.replace(",", "."))
            if math.isfinite(number):
                return
        except ValueError:
            pass

        for fmt in TIME_TEXT_FORMATS:
            try:
                datetime.strptime(text, fmt)
                return
            except ValueError:
                continue

        raise ExcelProcessError(f"第 {row_number} 行 A 列时间格式异常：{format_value_for_message(value)}")

    raise ExcelProcessError(f"第 {row_number} 行 A 列时间格式异常：{format_value_for_message(value)}")


def is_com_error(exc: Exception) -> bool:
    """判断异常是否来自 pywin32 COM。"""
    return pywintypes is not None and isinstance(exc, pywintypes.com_error)


def ensure_pywin32_available() -> None:
    """确认 pywin32 已安装。"""
    if PYWIN32_IMPORT_ERROR is not None:
        raise ExcelProcessError(
            "未能导入 pywin32，无法启动 Excel COM。\n"
            "请先安装依赖：py -m pip install pywin32\n"
            "如果仍失败，可尝试：py -m pywin32_postinstall -install\n"
            f"原始错误：{PYWIN32_IMPORT_ERROR}"
        )


def start_excel_application():
    """启动独立的 Excel COM 实例。"""
    ensure_pywin32_available()

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        try:
            excel.Calculation = XL_CALCULATION_AUTOMATIC
        except Exception:
            # 无法设置自动计算时不阻断流程，后面仍会主动 Calculate。
            pass
        return excel
    except Exception as exc:
        raise ExcelProcessError(
            "Excel COM 启动失败。请确认 Windows 已安装 Microsoft Excel，"
            "并且当前 Python 位数与 Office/系统环境可正常使用 COM。"
        ) from exc


def open_workbook_readonly(excel, input_path: Path):
    """以只读方式打开原始 .xls，避免误保存原文件。"""
    try:
        # 参数顺序：Filename, UpdateLinks, ReadOnly。
        return excel.Workbooks.Open(str(input_path), 0, True)
    except Exception as exc:
        raise ExcelProcessError(
            f"无法打开输入文件：{input_path}\n"
            "请确认文件未损坏，且不是受保护视图中无法访问的文件。"
        ) from exc


def get_worksheet_by_name(workbook, sheet_name: str):
    """按名称查找工作表，不存在时给出清晰错误。"""
    sheet_names = []
    try:
        count = int(workbook.Worksheets.Count)
        for index in range(1, count + 1):
            sheet = workbook.Worksheets(index)
            sheet_names.append(str(sheet.Name))
            if str(sheet.Name) == sheet_name:
                return sheet
    except Exception as exc:
        raise ExcelProcessError("读取工作表列表失败。") from exc

    raise ExcelProcessError(
        f"找不到名为“{sheet_name}”的工作表。\n"
        f"当前工作表：{', '.join(sheet_names) if sheet_names else '无'}"
    )


def find_last_nonblank_row_in_column(worksheet, column_number: int) -> int:
    """
    按 Excel 规则查找指定列最后一个非空单元格所在行。

    返回 0 表示整列没有有效数据。
    """
    try:
        row_count = int(worksheet.Rows.Count)
        candidate = int(worksheet.Cells(row_count, column_number).End(XL_UP).Row)
    except Exception as exc:
        raise ExcelProcessError("无法读取工作表行数或定位 B 列最后数据行。") from exc

    # End(xlUp) 在整列为空时通常会返回 1，所以还要检查实际值。
    for row_number in range(candidate, 0, -1):
        try:
            value = worksheet.Cells(row_number, column_number).Value
        except Exception as exc:
            raise ExcelProcessError(f"读取第 {row_number} 行 B 列失败。") from exc

        if not is_blank_value(value):
            return row_number

    return 0


def coerce_two_column_values(values: object, row_count: int) -> List[Tuple[object, object]]:
    """
    将 Excel Range.Value 统一为 [(A值, B值), ...]。

    多单元格区域通常返回 tuple[tuple]，这里做一层防御式转换。
    """
    rows: List[Tuple[object, object]] = []

    if row_count == 1:
        if isinstance(values, tuple) and len(values) == 1 and isinstance(values[0], tuple):
            row = values[0]
        else:
            row = values

        if not isinstance(row, tuple):
            raise ExcelProcessError("读取 A/B 列数据失败：返回值结构异常。")
        if len(row) < 2:
            raise ExcelProcessError("读取 A/B 列数据失败：列数不足。")
        rows.append((row[0], row[1]))
        return rows

    if not isinstance(values, tuple):
        raise ExcelProcessError("读取 A/B 列数据失败：返回值结构异常。")

    for row in values:
        if not isinstance(row, tuple) or len(row) < 2:
            raise ExcelProcessError("读取 A/B 列数据失败：某一行列数不足。")
        rows.append((row[0], row[1]))

    return rows


def find_calculation_start_row(weights: Sequence[float]) -> int:
    """
    根据 B 列开头连续 0 区段确定正式计算起点。

    规则：
        1. 只检查 B 列从第 1 行开始的连续 0。
        2. 如果开头就是非零数据，则从第 1 行开始计算。
        3. 如果开头存在连续 0，则取这段连续 0 的最后一个 0 所在行。
        4. 如果 B 列全是 0，没有后续非零数据，则无法计算，明确报错。

    注意：
        后续数据中再次出现 0 不会影响起点判断，因为本函数在遇到第一个
        非零值后就停止判断开头 0 区段。
    """
    if not weights:
        raise ExcelProcessError("B 列没有有效数据，无法确定有效数据结束行。")

    if weights[0] != 0:
        return 1

    last_leading_zero_row = 1
    for row_number, weight in enumerate(weights[1:], start=2):
        if weight != 0:
            return last_leading_zero_row
        last_leading_zero_row = row_number

    raise ExcelProcessError(
        "B 列全是 0，没有开头连续 0 之后的非零重量数据，无法计算液滴质量差。"
    )


def validate_source_data(worksheet) -> SourceDataInfo:
    """
    校验原始 A/B 列，并返回原始数据范围与正式计算起点。

    有效数据结束行规则严格使用 B 列最后一个非空单元格。
    中间空行、A 列空、B 列空、时间格式异常、重量格式异常都会报错。
    正式计算起点按“B 列开头连续 0 区段的最后一个 0”判定。
    """
    last_b_row = find_last_nonblank_row_in_column(worksheet, COL_WEIGHT)
    if last_b_row <= 0:
        raise ExcelProcessError("B 列没有有效数据，无法确定有效数据结束行。")

    try:
        values = worksheet.Range(
            worksheet.Cells(1, COL_TIME),
            worksheet.Cells(last_b_row, COL_WEIGHT),
        ).Value
    except Exception as exc:
        raise ExcelProcessError("读取 A/B 列数据失败。") from exc

    rows = coerce_two_column_values(values, last_b_row)
    weights: List[float] = []

    for index, (time_value, weight_value) in enumerate(rows, start=1):
        time_blank = is_blank_value(time_value)
        weight_blank = is_blank_value(weight_value)

        if time_blank and weight_blank:
            raise ExcelProcessError(
                f"第 {index} 行 A/B 列同时为空，原始数据中间出现空行，数据不连续。"
            )
        if time_blank:
            raise ExcelProcessError(f"第 {index} 行 A 列时间为空，数据不连续。")
        if weight_blank:
            raise ExcelProcessError(f"第 {index} 行 B 列重量为空，数据不连续。")

        validate_time_value(time_value, index)
        weights.append(parse_weight_value(weight_value, index))

    calculation_start_row = find_calculation_start_row(weights)
    return SourceDataInfo(
        last_b_row=last_b_row,
        calculation_start_row=calculation_start_row,
    )


def iter_output_candidates(
    input_path: Path,
    output_name_suffix: str,
    output_directory: Optional[Path] = None,
) -> Iterable[Path]:
    """???????????"""
    target_directory = (
        Path(output_directory).resolve(strict=False)
        if output_directory is not None
        else input_path.parent.resolve(strict=False)
    )
    base_stem = f"{input_path.stem}{output_name_suffix}"
    yield target_directory / f"{base_stem}{OUTPUT_EXTENSION}"

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    yield target_directory / f"{base_stem}-{timestamp}{OUTPUT_EXTENSION}"

    for counter in range(1, MAX_OUTPUT_CANDIDATE_COUNTER + 1):
        yield target_directory / f"{base_stem}-{timestamp}-{counter}{OUTPUT_EXTENSION}"

    raise ExcelProcessError("?????????????")


def reserve_unique_output_path(
    input_path: Path,
    output_name_suffix: str,
    output_directory: Optional[Path] = None,
) -> Path:
    """选择一个不会覆盖已有文件的输出路径。"""
    try:
        input_resolved = input_path.resolve(strict=False)
    except Exception:
        input_resolved = input_path

    target_directory = (
        Path(output_directory).resolve(strict=False)
        if output_directory is not None
        else input_path.parent.resolve(strict=False)
    )
    try:
        target_directory.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise ExcelProcessError(f"?????????{target_directory}") from exc

    for candidate in iter_output_candidates(input_path, output_name_suffix, target_directory):
        try:
            candidate_resolved = candidate.resolve(strict=False)
        except Exception:
            candidate_resolved = candidate

        if candidate_resolved == input_resolved:
            continue

        if candidate.exists():
            continue

        return candidate

    raise ExcelProcessError("?????????????")


def remove_reserved_output_file(output_path: Path) -> None:
    """保存失败时删除本程序创建的占位文件或失败残留文件。"""
    try:
        if output_path.exists():
            output_path.unlink()
    except Exception:
        # 删除失败不再抛出，以免掩盖真正的保存错误。
        pass


def prepare_output_path_for_excel_save(output_path: Path) -> None:
    """
    Excel SaveAs 前做最终防覆盖检查。

    reserve_unique_output_path 现在只选择路径，不再创建 0 字节占位文件，
    避免用户在处理尚未完成时看到“空白 xlsm”。如果保存前路径突然出现，
    说明有其他程序或上一次运行留下了文件；此时停止保存，避免覆盖。
    """
    if not output_path.exists():
        return

    raise ExcelProcessError(
        f"输出文件在保存前已经存在：{output_path}\n"
        "为避免覆盖已有结果，程序已停止保存。"
    )


def get_file_size_text(path: Optional[Path]) -> str:
    """返回文件大小调试文本。"""
    if path is None:
        return "<None>"
    try:
        if not path.exists():
            return "<missing>"
        return str(path.stat().st_size)
    except Exception as exc:
        return f"<stat failed: {debug_repr(exc)}>"


def trim_rows_before_calculation_start(worksheet, calculation_start_row: int) -> None:
    """
    删除正式计算起点之前的静置阶段数据。

    删除后：
        原始 calculation_start_row 会变成工作表第 1 行；
        随后再插入表头，正式起点会出现在输出表第 2 行。

    这样第一条质量差公式自然位于 C3：
        B3 - B2
    即“最后一个开头 0 的下一行重量 - 最后一个开头 0 的重量”。
    """
    if calculation_start_row <= 1:
        return

    try:
        worksheet.Rows(f"1:{calculation_start_row - 1}").Delete()
    except Exception as exc:
        raise ExcelProcessError("删除正式计算起点之前的静置阶段数据失败。") from exc


def clear_generated_regions(worksheet, last_output_data_row: int) -> None:
    """清空将由本脚本生成的区域，避免旧内容残留。"""
    clear_to_row = max(2, last_output_data_row)
    try:
        worksheet.Range(
            worksheet.Cells(2, COL_DIFF),
            worksheet.Cells(clear_to_row, COL_AVG_GROUP_2),
        ).ClearContents()
        worksheet.Range(
            worksheet.Cells(2, COL_RESULT_NO),
            worksheet.Cells(clear_to_row, COL_RESULT_PERMEABILITY),
        ).ClearContents()
    except Exception as exc:
        raise ExcelProcessError("清空输出计算区域失败。") from exc


def write_headers(worksheet) -> None:
    """写入左侧和右侧表头。"""
    try:
        for column, title in LEFT_HEADERS.items():
            worksheet.Cells(1, column).Value = title

        for column, title in RIGHT_HEADERS.items():
            worksheet.Cells(2, column).Value = title

        worksheet.Range(
            worksheet.Cells(1, COL_TIME),
            worksheet.Cells(1, COL_AVG_GROUP_2),
        ).Font.Bold = True
        worksheet.Range(
            worksheet.Cells(2, COL_RESULT_NO),
            worksheet.Cells(2, COL_RESULT_PERMEABILITY),
        ).Font.Bold = True
    except Exception as exc:
        raise ExcelProcessError("写入表头失败。") from exc


def validate_group_size(group_size: int, name: str) -> None:
    """校验分组大小常量。"""
    if not isinstance(group_size, int) or group_size <= 0:
        raise ExcelProcessError(f"{name} 必须是大于 0 的整数，请检查脚本顶部配置。")


def write_diff_formulas(worksheet: object, last_output_data_row: int) -> None:
    """写入 C 列液滴质量差公式。"""
    if last_output_data_row < 3:
        return

    try:
        target_range = worksheet.Range(
            worksheet.Cells(3, COL_DIFF),
            worksheet.Cells(last_output_data_row, COL_DIFF),
        )
        # R1C1：当前行 B 列 - 上一行 B 列。
        target_range.FormulaR1C1 = "=RC[-1]-R[-1]C[-1]"
    except Exception as exc:
        raise ExcelProcessError("写入 C 列液滴质量差公式失败。") from exc


def write_average_formulas(
    worksheet: object,
    target_column: int,
    group_size: int,
    last_output_data_row: int,
) -> List[int]:
    """
    按分组大小写入平均值公式。

    返回每个平均值所在的行号，例如 AVG_GROUP_1=5 时返回 [7, 12, 17, ...]。
    """
    result_rows: List[int] = []
    start_row = 3

    while start_row + group_size - 1 <= last_output_data_row:
        end_row = start_row + group_size - 1
        try:
            worksheet.Cells(end_row, target_column).Formula = f"=AVERAGE(C{start_row}:C{end_row})"
        except Exception as exc:
            raise ExcelProcessError(
                f"写入第 {end_row} 行平均值公式失败。"
            ) from exc
        result_rows.append(end_row)
        start_row = end_row + 1

    return result_rows


def write_result_base_area(
    worksheet,
    average_rows: Sequence[int],
    runtime_settings: RuntimeSettings,
) -> None:
    """Populate the shared pure-water result base area G:J."""
    for index, source_row in enumerate(average_rows, start=1):
        output_row = index + 2
        try:
            worksheet.Cells(output_row, COL_RESULT_NO).Value = index
            worksheet.Cells(output_row, COL_RESULT_WEIGHT).Formula = f"=D{source_row}"
            worksheet.Cells(output_row, COL_RESULT_TIME).Value = runtime_settings.time_seconds
            worksheet.Cells(output_row, COL_RESULT_AREA).Value = runtime_settings.membrane_area_cm2
        except Exception as exc:
            raise ExcelProcessError(f"写入右侧结果区第 {output_row} 行基础数据失败。") from exc

def parse_result_h_value(value, result_index: int) -> float:
    """解析右侧结果区 H 列重量值，用于压力切换点检测。"""
    return _parse_float_cell_value(value, f"第 {result_index} 个 H 值")


def read_result_h_values(worksheet, result_count: int) -> List[float]:
    """从右侧结果区 H 列读取实际生成的重量g值。"""
    h_values: List[float] = []

    for result_index in range(1, result_count + 1):
        output_row = result_index + 2
        try:
            value = worksheet.Cells(output_row, COL_RESULT_WEIGHT).Value
        except Exception as exc:
            raise ExcelProcessError(f"读取右侧结果区第 {result_index} 个 H 值失败。") from exc

        if is_blank_value(value):
            raise ExcelProcessError(f"右侧结果区第 {result_index} 个 H 值为空，无法判断压力切换点。")

        h_values.append(parse_result_h_value(value, result_index))

    return h_values


def detect_pressure_switch_index(h_values: Sequence[float]) -> Optional[int]:
    """
    自动识别压力切换开始编号。

    严格规则：
        当前三项平均值 = mean(H[i], H[i+1], H[i+2])
        前面所有项平均值 = mean(H[1], H[2], ..., H[i-1])
        若 当前三项平均值 < 0.85 * 前面所有项平均值，则从第 i 个编号开始切换。

    i 使用右侧结果区编号，从 1 开始计数。至少前面已有 5 个 H 值
    才允许判断，所以 i 最小为 6；同时 H[i]、H[i+1]、H[i+2]
    必须都存在，因此最少需要 8 个有效 H 值。

    如果存在多个候选切换点，默认取第一个满足条件的位置，因为它代表
    最早出现稳定下降的分段起点。
    """
    required_count = PRESSURE_SWITCH_MIN_PREVIOUS_COUNT + PRESSURE_SWITCH_CURRENT_COUNT
    if len(h_values) < required_count:
        return None

    # 0 基索引 5 对应右侧结果区第 6 个编号。
    first_current_index = PRESSURE_SWITCH_MIN_PREVIOUS_COUNT
    last_current_index = len(h_values) - PRESSURE_SWITCH_CURRENT_COUNT

    for current_index in range(first_current_index, last_current_index + 1):
        # “前面所有 H 值”是从第 1 个 H 值一直到 H[i-1]，不是固定窗口。
        previous_values = h_values[:current_index]
        current_values = h_values[
            current_index:current_index + PRESSURE_SWITCH_CURRENT_COUNT
        ]

        previous_mean = sum(previous_values) / len(previous_values)
        current_mean = sum(current_values) / PRESSURE_SWITCH_CURRENT_COUNT

        if current_mean < PRESSURE_SWITCH_RATIO * previous_mean:
            return current_index + 1

    return None


def format_result_weight_table(h_values: Sequence[float]) -> str:
    """把实际生成的 G/H 数据格式化为便于人工判断的文本表。"""
    lines = ["编号    重量g", "----    ----------------"]
    for index, value in enumerate(h_values, start=1):
        lines.append(f"{index:<6}  {value:.6f}")
    return "\n".join(lines)


def parse_manual_switch_index(text: str, max_index: int) -> int:
    """校验用户手动输入的压力切换开始编号。"""
    normalized = normalize_text(text)
    try:
        value = int(normalized)
    except ValueError as exc:
        raise ExcelProcessError("切换编号必须是整数。") from exc

    if value < 1:
        raise ExcelProcessError("切换编号不能小于 1。")
    if value > max_index:
        raise ExcelProcessError(f"切换编号不能大于最后一个编号 {max_index}。")

    return value


def prompt_manual_switch_index_cli(h_values: Sequence[float]) -> int:
    """命令行环境下展示 G/H 并让用户输入切换编号。"""
    table_text = format_result_weight_table(h_values)
    print("\n未能自动识别压力切换点，请根据右侧结果区 G/H 数据手动判断：")
    print(table_text)

    while True:
        text = input("请输入从第几个编号开始切换压力，直接回车取消：").strip()
        if not text:
            raise UserCancelledError("已取消输入压力切换编号，程序已终止。")
        try:
            return parse_manual_switch_index(text, len(h_values))
        except ExcelProcessError as exc:
            print(f"输入无效：{exc}")


def prompt_manual_switch_index_gui(h_values: Sequence[float]) -> int:
    """GUI 环境下用可滚动文本框展示 G/H，并输入切换编号。"""
    global _TK_ROOT

    root = get_tk_root()
    if root is None or tk is None:
        return prompt_manual_switch_index_cli(h_values)

    table_text = format_result_weight_table(h_values)
    selected = {"value": None, "cancelled": False}

    # 这里直接复用 Tk 主窗口，不再创建依附于隐藏 root 的 Toplevel。
    # Windows 下“隐藏父窗口 + 模态 Toplevel”偶尔会出现任务栏可见但窗口点不开。
    for child in root.winfo_children():
        child.destroy()

    root.title("手动输入压力切换编号")
    root.geometry("560x620")
    root.minsize(520, 520)
    root.resizable(True, True)
    root.deiconify()
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass

    prompt = (
        "未能自动识别压力切换点。\n"
        "请根据下方实际生成的 G/H 数据，输入“从第几个编号开始切换压力”。\n"
        "输入 1 表示全部使用后半段压力。"
    )
    label = tk.Label(root, text=prompt, justify="left", anchor="w")
    label.pack(fill="x", padx=12, pady=(12, 8))

    text_frame = tk.Frame(root)
    text_frame.pack(fill="both", expand=True, padx=12, pady=4)

    scrollbar = tk.Scrollbar(text_frame)
    scrollbar.pack(side="right", fill="y")

    text_widget = tk.Text(
        text_frame,
        height=20,
        wrap="none",
        yscrollcommand=scrollbar.set,
    )
    text_widget.insert("1.0", table_text)
    text_widget.configure(state="disabled")
    text_widget.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=text_widget.yview)

    entry_frame = tk.Frame(root)
    entry_frame.pack(fill="x", padx=12, pady=8)
    tk.Label(entry_frame, text="切换开始编号：").pack(side="left")
    entry = tk.Entry(entry_frame)
    entry.pack(side="left", fill="x", expand=True)

    button_frame = tk.Frame(root)
    button_frame.pack(fill="x", padx=12, pady=(0, 12))

    def submit() -> None:
        try:
            selected["value"] = parse_manual_switch_index(entry.get(), len(h_values))
        except ExcelProcessError as exc:
            if messagebox is not None:
                messagebox.showwarning("输入无效", str(exc), parent=root)
            return

        root.quit()

    def cancel() -> None:
        selected["cancelled"] = True
        root.quit()

    tk.Button(button_frame, text="确定", command=submit).pack(side="right", padx=(8, 0))
    tk.Button(button_frame, text="取消", command=cancel).pack(side="right")
    root.protocol("WM_DELETE_WINDOW", cancel)
    entry.bind("<Return>", lambda _event: submit())

    _center_and_focus_window(root, 560, 620)

    entry.focus_set()
    root.mainloop()

    try:
        root.attributes("-topmost", False)
        root.withdraw()
    except Exception:
        pass

    if selected["cancelled"] or selected["value"] is None:
        raise UserCancelledError("已取消输入压力切换编号，程序已终止。")

    return int(selected["value"])


def prompt_manual_switch_index(h_values: Sequence[float]) -> int:
    """自动识别失败后，展示 G/H 数据并让用户输入切换编号。"""
    if not h_values:
        raise ExcelProcessError("右侧结果区没有生成任何 G/H 数据，无法进行压力分段。")

    root = get_tk_root()
    if root is not None and tk is not None:
        return prompt_manual_switch_index_gui(h_values)

    return prompt_manual_switch_index_cli(h_values)


def choose_pressure_switch_index(
    h_values: Sequence[float],
    injected_switch_index: Optional[int] = None,
) -> Tuple[int, bool]:
    """????????????????????????????"""
    switch_index = detect_pressure_switch_index(h_values)
    if switch_index is not None:
        debug_log_event(
            "pressure_switch_auto_detected",
            h_count=len(h_values),
            switch_index=switch_index,
        )
        return switch_index, True

    if injected_switch_index is not None:
        parsed_index = parse_manual_switch_index(str(injected_switch_index), len(h_values))
        debug_log_event(
            "pressure_switch_injected_fallback_used",
            h_count=len(h_values),
            switch_index=parsed_index,
        )
        return parsed_index, False

    debug_log_event(
        "pressure_switch_auto_detect_failed_waiting_manual_input",
        h_count=len(h_values),
        first_h=h_values[0] if h_values else None,
        last_h=h_values[-1] if h_values else None,
    )
    manual_index = prompt_manual_switch_index(h_values)
    debug_log_event(
        "pressure_switch_manual_input_used",
        h_count=len(h_values),
        switch_index=manual_index,
    )
    return manual_index, False


def write_segmented_pressure_and_formulas(
    worksheet,
    result_count: int,
    pressure_settings: PressureSettings,
    switch_index: int,
) -> None:
    """按切换编号分段写入 K 列压力，并写入 L/M 公式。"""
    if result_count <= 0:
        raise ExcelProcessError("右侧结果区没有生成数据，无法写入分段压力。")
    if switch_index < 1 or switch_index > result_count:
        raise ExcelProcessError(
            f"压力切换编号 {switch_index} 超出实际编号范围 1~{result_count}。"
        )

    for index in range(1, result_count + 1):
        output_row = index + 2
        pressure = (
            pressure_settings.second_segment_pressure
            if index >= switch_index
            else pressure_settings.first_segment_pressure
        )

        try:
            worksheet.Cells(output_row, COL_RESULT_PRESSURE).Value = pressure

            # L = H * 10 * 3600 / I / J。L 不直接除压力，保持原通量定义。
            worksheet.Cells(output_row, COL_RESULT_FLUX).FormulaR1C1 = "=RC[-4]*10*3600/RC[-3]/RC[-2]"

            # M = L / K，K 已按切换编号写成对应压力。
            worksheet.Cells(output_row, COL_RESULT_PERMEABILITY).FormulaR1C1 = "=RC[-1]/RC[-2]"
        except Exception as exc:
            raise ExcelProcessError(f"写入右侧结果区第 {output_row} 行分段压力失败。") from exc


def write_single_pressure_and_formulas(
    worksheet,
    result_count: int,
    pressure_bar: float,
) -> None:
    """
    污染物模式：整段结果使用同一个运行压力。

    本轮污染物分支不做压力切换识别、不做手动切换编号、不做最终异常回修，
    只按当前单一压力写 K，并按既有公式写 L/M。
    """
    if result_count <= 0:
        raise ExcelProcessError("右侧结果区没有生成数据，无法写入运行压力。")

    for index in range(1, result_count + 1):
        output_row = index + 2
        try:
            worksheet.Cells(output_row, COL_RESULT_PRESSURE).Value = pressure_bar

            # L = H * 10 * 3600 / I / J。
            worksheet.Cells(output_row, COL_RESULT_FLUX).FormulaR1C1 = "=RC[-4]*10*3600/RC[-3]/RC[-2]"

            # M = L / K；污染物模式下 K 为整段统一压力。
            worksheet.Cells(output_row, COL_RESULT_PERMEABILITY).FormulaR1C1 = "=RC[-1]/RC[-2]"
        except Exception as exc:
            raise ExcelProcessError(f"写入污染物模式第 {output_row} 行压力和公式失败。") from exc


def write_pollutant_result_base_area(
    worksheet,
    last_output_data_row: int,
    runtime_settings: RuntimeSettings,
) -> int:
    """Populate pollutant mode result area G:M with H mapped from column C."""
    result_count = max(last_output_data_row - 2, 0)
    if result_count <= 0:
        raise ExcelProcessError("有效 C 列液滴质量差不足，无法生成污染物模式结果区。")

    try:
        worksheet.Range(
            worksheet.Cells(3, COL_RESULT_NO),
            worksheet.Cells(max(3, last_output_data_row), COL_RESULT_PERMEABILITY),
        ).ClearContents()
    except Exception as exc:
        raise ExcelProcessError("清空污染物模式 G:M 结果区失败。") from exc

    for index in range(1, result_count + 1):
        output_row = index + 2
        try:
            worksheet.Cells(output_row, COL_RESULT_NO).Value = index
            worksheet.Cells(output_row, COL_RESULT_WEIGHT).Formula = f"=C{output_row}"
            worksheet.Cells(output_row, COL_RESULT_TIME).Value = runtime_settings.time_seconds
            worksheet.Cells(output_row, COL_RESULT_AREA).Value = runtime_settings.membrane_area_cm2
        except Exception as exc:
            raise ExcelProcessError(f"写入污染物模式 G:M 基础区第 {output_row} 行失败。") from exc

    return result_count

def pollutant_permeability_formula(
    c_start_row: int,
    c_end_row: int,
    pressure_bar: float,
    runtime_settings: RuntimeSettings,
) -> str:
    """Build the pollutant permeability formula using runtime time and area."""
    return (
        f"=AVERAGE(C{c_start_row}:C{c_end_row})"
        f"*10*3600/{runtime_settings.time_seconds:g}"
        f"/{runtime_settings.membrane_area_cm2:g}/{pressure_bar:g}"
    )

def split_count_evenly(total_count: int, target_group_count: int) -> List[int]:
    """
    将 total_count 个点尽量平均分为 target_group_count 组。

    如果不能整除，前面的组多 1 个点，后面的组少 1 个点。
    如果数据点不足目标组数，则实际组数 = min(目标组数, 有效 C 点数)。
    """
    if total_count <= 0:
        return []

    actual_group_count = min(target_group_count, total_count)
    base_size = total_count // actual_group_count
    extra_count = total_count % actual_group_count

    return [
        base_size + (1 if index < extra_count else 0)
        for index in range(actual_group_count)
    ]


def clear_pollutant_analysis_area(worksheet, last_output_data_row: int) -> None:
    """清空污染物模式 O:V 分析区，避免旧内容残留。"""
    clear_to_row = max(last_output_data_row + 5, 40)
    try:
        worksheet.Range(
            worksheet.Cells(2, COL_POLLUTANT_5_POINT_PERMEABILITY),
            worksheet.Cells(clear_to_row, COL_POLLUTANT_5_GROUP_NORMALIZED),
        ).ClearContents()
    except Exception as exc:
        raise ExcelProcessError("清空污染物模式 O:V 分析区失败。") from exc


def write_pollutant_analysis_headers(worksheet) -> None:
    """写入污染物模式 O:V 分析区表头。"""
    try:
        for column, title in POLLUTANT_ANALYSIS_HEADERS.items():
            worksheet.Cells(2, column).Value = title

        worksheet.Range(
            worksheet.Cells(2, COL_POLLUTANT_5_POINT_PERMEABILITY),
            worksheet.Cells(2, COL_POLLUTANT_5_GROUP_NORMALIZED),
        ).Font.Bold = True
    except Exception as exc:
        raise ExcelProcessError("写入污染物模式 O:V 表头失败。") from exc


def write_pollutant_baseline_row(worksheet, pure_water_flux: float) -> None:
    """
    写入污染物 O/Q/S/U 的基准纯水通量，以及 P/R/T/V 的归一化起点。

    P3/R3/T3/V3 都是对应基准值除以纯水通量，因此都等于 1。
    """
    pairs = (
        (COL_POLLUTANT_5_POINT_PERMEABILITY, COL_POLLUTANT_5_POINT_NORMALIZED),
        (COL_POLLUTANT_15_GROUP_PERMEABILITY, COL_POLLUTANT_15_GROUP_NORMALIZED),
        (COL_POLLUTANT_10_GROUP_PERMEABILITY, COL_POLLUTANT_10_GROUP_NORMALIZED),
        (COL_POLLUTANT_5_GROUP_PERMEABILITY, COL_POLLUTANT_5_GROUP_NORMALIZED),
    )

    try:
        for permeability_column, normalized_column in pairs:
            worksheet.Cells(3, permeability_column).Value = pure_water_flux
            worksheet.Cells(3, normalized_column).Formula = (
                pollutant_normalized_formula(permeability_column, 3)
            )
    except Exception as exc:
        raise ExcelProcessError("写入污染物模式基准纯水通量失败。") from exc


def pollutant_normalized_formula(permeability_column: int, output_row: int) -> str:
    """
    污染物模式归一化公式。

    P/R/T/V 不再除以 Python 写死数值，而是引用各自渗透性列第 3 行
    的基准纯水通量单元格，例如 P4 = O4/$O$3。
    """
    source_column_letter = column_letter(permeability_column)
    return f"={source_column_letter}{output_row}/${source_column_letter}$3"


def column_letter(column_number: int) -> str:
    """把 Excel 列号转换为列字母，例如 15 -> O。"""
    letters = ""
    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def write_pollutant_fixed_size_groups(
    worksheet,
    c_start_row: int,
    c_count: int,
    group_size: int,
    permeability_column: int,
    normalized_column: int,
    pressure_bar: float,
    pure_water_flux: float,
    runtime_settings: RuntimeSettings,
) -> int:
    """O \u5217\u89c4\u5219\uff1aC \u5217\u6bcf 5 \u4e2a\u70b9\u4e00\u7ec4\uff0c\u672b\u7ec4\u4e0d\u8db3 5 \u4e2a\u70b9\u65f6\u76f4\u63a5\u4e22\u5f03\u3002"""
    group_count = c_count // group_size

    for group_index in range(group_count):
        source_start_row = c_start_row + group_index * group_size
        source_end_row = source_start_row + group_size - 1
        output_row = 4 + group_index

        try:
            worksheet.Cells(output_row, permeability_column).Formula = pollutant_permeability_formula(
                source_start_row,
                source_end_row,
                pressure_bar,
                runtime_settings,
            )
            worksheet.Cells(output_row, normalized_column).Formula = pollutant_normalized_formula(
                permeability_column,
                output_row,
            )
        except Exception as exc:
            raise ExcelProcessError(f"\u5199\u5165\u6c61\u67d3\u7269\u6a21\u5f0f\u7b2c {output_row} \u884c 5 \u70b9\u5206\u7ec4\u7ed3\u679c\u5931\u8d25\u3002") from exc

    return group_count


def write_pollutant_even_groups(
    worksheet,
    c_start_row: int,
    c_count: int,
    target_group_count: int,
    permeability_column: int,
    normalized_column: int,
    pressure_bar: float,
    pure_water_flux: float,
    runtime_settings: RuntimeSettings,
) -> int:
    """Q/S/U \u89c4\u5219\uff1a\u76f4\u63a5\u57fa\u4e8e\u5168\u90e8\u6709\u6548 C \u70b9\u5c3d\u91cf\u5e73\u5747\u5206\u7ec4\u3002"""
    group_sizes = split_count_evenly(c_count, target_group_count)
    source_start_row = c_start_row

    for group_index, group_size in enumerate(group_sizes):
        source_end_row = source_start_row + group_size - 1
        output_row = 4 + group_index

        try:
            worksheet.Cells(output_row, permeability_column).Formula = pollutant_permeability_formula(
                source_start_row,
                source_end_row,
                pressure_bar,
                runtime_settings,
            )
            worksheet.Cells(output_row, normalized_column).Formula = pollutant_normalized_formula(
                permeability_column,
                output_row,
            )
        except Exception as exc:
            raise ExcelProcessError(f"\u5199\u5165\u6c61\u67d3\u7269\u6a21\u5f0f\u7b2c {output_row} \u884c\u5e73\u5747\u5206\u7ec4\u7ed3\u679c\u5931\u8d25\u3002") from exc

        source_start_row = source_end_row + 1

    return len(group_sizes)


def write_pollutant_analysis_area(
    worksheet,
    last_output_data_row: int,
    pressure_bar: float,
    pure_water_flux: float,
    runtime_settings: RuntimeSettings,
) -> PollutantAnalysisResult:
    """\u6309\u6c61\u67d3\u7269\u6a21\u5f0f\u89c4\u5219\u751f\u6210 O:V \u5206\u6790\u533a\u3002"""
    c_start_row = 3
    c_count = max(last_output_data_row - 2, 0)
    if c_count <= 0:
        raise ExcelProcessError("\u6709\u6548 C \u5217\u6db2\u6ef4\u8d28\u91cf\u5dee\u4e0d\u8db3\uff0c\u65e0\u6cd5\u751f\u6210\u6c61\u67d3\u7269\u6a21\u5f0f O:V \u5206\u6790\u533a\u3002")

    clear_pollutant_analysis_area(worksheet, last_output_data_row)
    write_pollutant_analysis_headers(worksheet)
    write_pollutant_baseline_row(worksheet, pure_water_flux)

    five_point_count = write_pollutant_fixed_size_groups(
        worksheet,
        c_start_row,
        c_count,
        5,
        COL_POLLUTANT_5_POINT_PERMEABILITY,
        COL_POLLUTANT_5_POINT_NORMALIZED,
        pressure_bar,
        pure_water_flux,
        runtime_settings,
    )
    even_group_counts = {}
    configs = [
        (15, COL_POLLUTANT_15_GROUP_PERMEABILITY, COL_POLLUTANT_15_GROUP_NORMALIZED),
        (10, COL_POLLUTANT_10_GROUP_PERMEABILITY, COL_POLLUTANT_10_GROUP_NORMALIZED),
        (5, COL_POLLUTANT_5_GROUP_PERMEABILITY, COL_POLLUTANT_5_GROUP_NORMALIZED),
    ]
    for target_group_count, permeability_column, normalized_column in configs:
        even_group_counts[target_group_count] = write_pollutant_even_groups(
            worksheet,
            c_start_row,
            c_count,
            target_group_count,
            permeability_column,
            normalized_column,
            pressure_bar,
            pure_water_flux,
            runtime_settings,
        )

    return PollutantAnalysisResult(
        five_point_count=five_point_count,
        fifteen_group_count=even_group_counts[15],
        ten_group_count=even_group_counts[10],
        five_group_count=even_group_counts[5],
    )


def clear_pollutant_standard_10min_area(worksheet, last_output_data_row: int) -> None:
    """清空污染物模式 X:Y 标准十分钟结果区，避免旧内容残留。"""
    clear_to_row = max(last_output_data_row + 5, STANDARD_10MIN_GROUP_COUNT + 3, 40)
    try:
        worksheet.Range(
            worksheet.Cells(2, COL_STANDARD_10MIN_PERMEABILITY),
            worksheet.Cells(clear_to_row, COL_STANDARD_10MIN_NORMALIZED),
        ).ClearContents()
    except Exception as exc:
        raise ExcelProcessError("清空污染物模式 X:Y 标准十分钟结果区失败。") from exc


def write_pollutant_standard_10min_area(
    worksheet,
    result_count: int,
    pure_water_flux: float,
    last_output_data_row: int,
) -> int:
    """
    生成污染物模式 X:Y 标准十分钟结果区。

    X 列直接基于污染物模式已经生成的 M 列渗透性结果：
    M3:M152 对应 G=1~150，按每 30 个 M 值一组取平均，共 5 组。
    """
    if result_count < STANDARD_10MIN_REQUIRED_COUNT:
        raise ExcelProcessError(
            "标准十分钟计算需要至少150个污染物模式 M 列渗透性数据，请检查原始数据长度。"
        )

    clear_pollutant_standard_10min_area(worksheet, last_output_data_row)

    try:
        for column, title in POLLUTANT_STANDARD_10MIN_HEADERS.items():
            worksheet.Cells(2, column).Value = title

        worksheet.Range(
            worksheet.Cells(2, COL_STANDARD_10MIN_PERMEABILITY),
            worksheet.Cells(2, COL_STANDARD_10MIN_NORMALIZED),
        ).Font.Bold = True

        # 第 3 行写基准纯水通量；Y3 通过引用 X3 归一化，结果为 1。
        worksheet.Cells(3, COL_STANDARD_10MIN_PERMEABILITY).Value = pure_water_flux
        worksheet.Cells(3, COL_STANDARD_10MIN_NORMALIZED).Formula = (
            pollutant_normalized_formula(COL_STANDARD_10MIN_PERMEABILITY, 3)
        )

        for group_index in range(STANDARD_10MIN_GROUP_COUNT):
            output_row = 4 + group_index
            m_start_row = 3 + group_index * STANDARD_10MIN_GROUP_SIZE
            m_end_row = m_start_row + STANDARD_10MIN_GROUP_SIZE - 1

            worksheet.Cells(output_row, COL_STANDARD_10MIN_PERMEABILITY).Formula = (
                f"=AVERAGE(M{m_start_row}:M{m_end_row})"
            )
            worksheet.Cells(output_row, COL_STANDARD_10MIN_NORMALIZED).Formula = (
                pollutant_normalized_formula(COL_STANDARD_10MIN_PERMEABILITY, output_row)
            )

        worksheet.Range(
            worksheet.Cells(3, COL_STANDARD_10MIN_PERMEABILITY),
            worksheet.Cells(STANDARD_10MIN_GROUP_COUNT + 3, COL_STANDARD_10MIN_NORMALIZED),
        ).NumberFormat = "0.000000"
        worksheet.Columns("X:Y").AutoFit()
    except Exception as exc:
        raise ExcelProcessError("写入污染物模式 X:Y 标准十分钟结果区失败。") from exc

    return STANDARD_10MIN_GROUP_COUNT + 1


def clear_pollutant_custom_average_area(worksheet, point_count: int, last_output_data_row: int) -> None:
    """清空污染物模式 AA:AB 用户指定平均点区域，Z 列保留为空白间隔。"""
    clear_to_row = max(last_output_data_row + 5, point_count + 3, 40)
    try:
        worksheet.Range(
            worksheet.Cells(2, COL_CUSTOM_AVERAGE_PERMEABILITY),
            worksheet.Cells(clear_to_row, COL_CUSTOM_AVERAGE_NORMALIZED),
        ).ClearContents()
    except Exception as exc:
        raise ExcelProcessError("清空污染物模式 AA:AB 散点分析区失败。") from exc


def write_pollutant_custom_average_area(
    worksheet,
    last_output_data_row: int,
    pressure_bar: float,
    pure_water_flux: float,
    runtime_settings: RuntimeSettings,
    custom_settings: CustomAverageSettings,
) -> int:
    """
    生成污染物模式 AA:AB 用户指定平均点结果区。

    数据源固定为 C3 开始的液滴质量差，每组点数由用户输入分钟数和采样时间换算。
    """
    c_start_row = 3
    c_count = max(last_output_data_row - 2, 0)
    required_count = custom_settings.points_per_group * custom_settings.point_count
    if c_count < required_count:
        raise ExcelProcessError("当前C列有效数据不足，无法生成指定数量的平均点，请检查原始数据长度。")

    clear_pollutant_custom_average_area(
        worksheet,
        custom_settings.point_count,
        last_output_data_row,
    )

    try:
        worksheet.Cells(2, COL_CUSTOM_AVERAGE_PERMEABILITY).Value = (
            f"{custom_settings.name_prefix}-渗透性"
        )
        worksheet.Cells(2, COL_CUSTOM_AVERAGE_NORMALIZED).Value = (
            f"{custom_settings.name_prefix}-归一化通量"
        )
        worksheet.Range(
            worksheet.Cells(2, COL_CUSTOM_AVERAGE_PERMEABILITY),
            worksheet.Cells(2, COL_CUSTOM_AVERAGE_NORMALIZED),
        ).Font.Bold = True

        # 第 3 行写入纯水通量基准，AB3 通过引用 AA3 得到归一化起点 1。
        worksheet.Cells(3, COL_CUSTOM_AVERAGE_PERMEABILITY).Value = pure_water_flux
        worksheet.Cells(3, COL_CUSTOM_AVERAGE_NORMALIZED).Formula = (
            pollutant_normalized_formula(COL_CUSTOM_AVERAGE_PERMEABILITY, 3)
        )

        for group_index in range(custom_settings.point_count):
            source_start_row = c_start_row + group_index * custom_settings.points_per_group
            source_end_row = source_start_row + custom_settings.points_per_group - 1
            output_row = 4 + group_index

            worksheet.Cells(output_row, COL_CUSTOM_AVERAGE_PERMEABILITY).Formula = (
                pollutant_permeability_formula(
                    source_start_row,
                    source_end_row,
                    pressure_bar,
                    runtime_settings,
                )
            )
            worksheet.Cells(output_row, COL_CUSTOM_AVERAGE_NORMALIZED).Formula = (
                pollutant_normalized_formula(COL_CUSTOM_AVERAGE_PERMEABILITY, output_row)
            )

        worksheet.Range(
            worksheet.Cells(3, COL_CUSTOM_AVERAGE_PERMEABILITY),
            worksheet.Cells(custom_settings.point_count + 3, COL_CUSTOM_AVERAGE_NORMALIZED),
        ).NumberFormat = "0.000000"
        worksheet.Columns("AA:AB").AutoFit()
    except Exception as exc:
        raise ExcelProcessError("写入污染物模式 AA:AB 散点分析区失败。") from exc

    return custom_settings.point_count + 1


def get_or_create_custom_average_scatter_sheet(workbook, sheet_name: str):
    """获取或新建用户指定平均点散点图工作表。"""
    try:
        for index in range(1, int(workbook.Worksheets.Count) + 1):
            candidate = workbook.Worksheets(index)
            if str(candidate.Name) == sheet_name:
                return candidate

        chart_sheet = workbook.Worksheets.Add(
            After=workbook.Worksheets(int(workbook.Worksheets.Count))
        )
        chart_sheet.Name = sheet_name
        return chart_sheet
    except Exception as exc:
        raise ExcelProcessError(f"创建散点图工作表“{sheet_name}”失败。") from exc


def create_custom_average_scatter_sheet(
    source_worksheet,
    custom_settings: CustomAverageSettings,
) -> None:
    """生成单独的用户指定平均点散点图工作表。"""
    sheet_name = f"{custom_settings.name_prefix}图"
    workbook = source_worksheet.Parent
    chart_sheet = get_or_create_custom_average_scatter_sheet(workbook, sheet_name)
    clear_scatter_chart_sheet(chart_sheet)

    x_range, y_range = write_scatter_chart_source_data(
        source_worksheet,
        chart_sheet,
        1,
        COL_CUSTOM_AVERAGE_NORMALIZED,
        custom_settings.point_count + 1,
        custom_settings.name_prefix,
    )
    add_scatter_chart(chart_sheet, x_range, y_range, custom_settings.name_prefix, 20, 20)

    try:
        chart_sheet.Columns("A:B").AutoFit()
    except Exception:
        pass


def quote_excel_sheet_name(sheet_name: str) -> str:
    """生成 Excel 公式里安全可用的工作表引用名。"""
    return "'" + sheet_name.replace("'", "''") + "'"


def get_or_create_scatter_chart_sheet(workbook):
    """获取或新建污染物模式散点图工作表。"""
    chart_sheet_name = "散点图工作表"

    try:
        for index in range(1, int(workbook.Worksheets.Count) + 1):
            candidate = workbook.Worksheets(index)
            if str(candidate.Name) == chart_sheet_name:
                return candidate

        chart_sheet = workbook.Worksheets.Add(
            After=workbook.Worksheets(int(workbook.Worksheets.Count))
        )
        chart_sheet.Name = chart_sheet_name
        return chart_sheet
    except Exception as exc:
        raise ExcelProcessError("创建散点图工作表失败。") from exc


def clear_scatter_chart_sheet(chart_sheet) -> None:
    """清空散点图工作表中的辅助数据和旧图表。"""
    try:
        chart_sheet.Cells.Clear()
        chart_objects = chart_sheet.ChartObjects()
        for index in range(int(chart_objects.Count), 0, -1):
            chart_objects(index).Delete()
    except Exception as exc:
        raise ExcelProcessError("清空散点图工作表失败。") from exc


def write_scatter_chart_source_data(
    source_worksheet,
    chart_sheet,
    start_column: int,
    normalized_column: int,
    point_count: int,
    title: str,
) -> Tuple[object, object]:
    """
    在散点图工作表写入某张图的 X/Y 辅助数据。

    X 轴固定为 1, 2, 3, ...；
    Y 轴引用对应归一化通量列，从第 3 行的基准点 1 开始。
    """
    source_sheet_ref = quote_excel_sheet_name(str(source_worksheet.Name))
    y_column_letter = column_letter(normalized_column)

    try:
        chart_sheet.Cells(1, start_column).Value = "序号"
        chart_sheet.Cells(1, start_column + 1).Value = title

        for point_index in range(1, point_count + 1):
            output_row = point_index + 1
            source_row = point_index + 2
            chart_sheet.Cells(output_row, start_column).Value = point_index
            chart_sheet.Cells(output_row, start_column + 1).Formula = (
                f"={source_sheet_ref}!{y_column_letter}{source_row}"
            )

        x_range = chart_sheet.Range(
            chart_sheet.Cells(2, start_column),
            chart_sheet.Cells(point_count + 1, start_column),
        )
        y_range = chart_sheet.Range(
            chart_sheet.Cells(2, start_column + 1),
            chart_sheet.Cells(point_count + 1, start_column + 1),
        )
        return x_range, y_range
    except Exception as exc:
        raise ExcelProcessError(f"写入散点图“{title}”辅助数据失败。") from exc


def add_scatter_chart(chart_sheet, x_range, y_range, title: str, left: int, top: int) -> None:
    """在散点图工作表新增一张独立散点图。"""
    try:
        chart_object = chart_sheet.ChartObjects().Add(left, top, 420, 260)
        chart = chart_object.Chart
        chart.ChartType = XL_XY_SCATTER

        series = chart.SeriesCollection().NewSeries()
        series.XValues = x_range
        series.Values = y_range
        series.Name = title

        chart.HasTitle = True
        chart.ChartTitle.Text = title

        try:
            chart.Axes(XL_CATEGORY_AXIS).HasTitle = True
            chart.Axes(XL_CATEGORY_AXIS).AxisTitle.Text = "序号"
            chart.Axes(XL_VALUE_AXIS).HasTitle = True
            chart.Axes(XL_VALUE_AXIS).AxisTitle.Text = "归一化通量"
        except Exception:
            pass
    except Exception as exc:
        raise ExcelProcessError(f"生成散点图“{title}”失败。") from exc


def add_standard_10min_scatter_chart(source_worksheet, chart_sheet) -> None:
    """
    在散点图工作表中新增“标准十分钟”散点图。

    数据来自污染物模式 Y3:Y8，第一个点 Y3 为归一化基准 1。
    """
    x_range, y_range = write_scatter_chart_source_data(
        source_worksheet,
        chart_sheet,
        13,
        COL_STANDARD_10MIN_NORMALIZED,
        STANDARD_10MIN_GROUP_COUNT + 1,
        "标准十分钟",
    )
    add_scatter_chart(chart_sheet, x_range, y_range, "标准十分钟", 20, 620)


def create_pollutant_scatter_charts(
    worksheet,
    analysis_result: PollutantAnalysisResult,
) -> None:
    """
    污染物模式：生成 5 张独立散点图。

    每张图的第一个点都来自第 3 行归一化起点，因此都是 (1, 1)。
    """
    workbook = worksheet.Parent
    chart_sheet = get_or_create_scatter_chart_sheet(workbook)
    clear_scatter_chart_sheet(chart_sheet)

    chart_specs = (
        (
            "P列归一化通量散点图",
            COL_POLLUTANT_5_POINT_NORMALIZED,
            analysis_result.five_point_count + 1,
            1,
            20,
            20,
        ),
        (
            "R列归一化通量散点图",
            COL_POLLUTANT_15_GROUP_NORMALIZED,
            analysis_result.fifteen_group_count + 1,
            4,
            470,
            20,
        ),
        (
            "T列归一化通量散点图",
            COL_POLLUTANT_10_GROUP_NORMALIZED,
            analysis_result.ten_group_count + 1,
            7,
            20,
            320,
        ),
        (
            "V列归一化通量散点图",
            COL_POLLUTANT_5_GROUP_NORMALIZED,
            analysis_result.five_group_count + 1,
            10,
            470,
            320,
        ),
    )

    for title, normalized_column, point_count, data_start_column, left, top in chart_specs:
        x_range, y_range = write_scatter_chart_source_data(
            worksheet,
            chart_sheet,
            data_start_column,
            normalized_column,
            point_count,
            title,
        )
        add_scatter_chart(chart_sheet, x_range, y_range, title, left, top)

    add_standard_10min_scatter_chart(worksheet, chart_sheet)

    try:
        chart_sheet.Columns("A:N").AutoFit()
    except Exception:
        pass


def parse_result_m_value(value, result_index: int) -> float:
    """解析右侧结果区 M 列渗透性值，用于后半段平均值统计。"""
    return _parse_float_cell_value(value, f"第 {result_index} 个 M 值")


def parse_result_pressure_value(value, result_index: int) -> float:
    """解析右侧结果区 K 列运行压力值，用于最终校验。"""
    return _parse_float_cell_value(value, f"第 {result_index} 个 K 值")


def pressures_are_equal(left: float, right: float) -> bool:
    """
    判断两个压力是否相等。

    Excel COM 读写浮点数时可能出现极小误差，因此这里用很小的容差比较；
    业务含义仍然是“当前 K 列等于后半段运行压力”。
    """
    return math.isclose(float(left), float(right), rel_tol=1e-9, abs_tol=1e-9)


def read_result_pressure_value(worksheet, result_index: int) -> float:
    """读取某个结果编号对应的 K 列压力。"""
    output_row = result_index + 2
    try:
        value = worksheet.Cells(output_row, COL_RESULT_PRESSURE).Value
    except Exception as exc:
        raise ExcelProcessError(f"读取右侧结果区第 {result_index} 个 K 值失败。") from exc

    if is_blank_value(value):
        raise ExcelProcessError(f"右侧结果区第 {result_index} 个 K 值为空，无法执行最终校验。")

    return parse_result_pressure_value(value, result_index)


def read_valid_result_m_values(worksheet, result_count: int) -> List[Tuple[int, float]]:
    """
    读取所有有效结果行中的 M 列渗透性值。

    注意：
        这里只读取 result_count 对应的有效结果行，不读取 M 列最后额外输出的
        “后半段渗透性平均值”单元格。
    """
    values: List[Tuple[int, float]] = []
    for result_index in range(1, result_count + 1):
        output_row = result_index + 2
        try:
            value = worksheet.Cells(output_row, COL_RESULT_PERMEABILITY).Value
        except Exception as exc:
            raise ExcelProcessError(f"读取右侧结果区第 {result_index} 个 M 值失败。") from exc

        if is_blank_value(value):
            continue

        values.append((result_index, parse_result_m_value(value, result_index)))

    return values


def rewrite_result_rows_to_first_pressure(
    worksheet,
    result_indices: Sequence[int],
    pressure_settings: PressureSettings,
) -> None:
    """
    将指定结果编号对应行的 K 列改为前半段压力，并重写 L/M 公式。

    同一轮最终校验中发现的多个异常行会一起传入本函数，统一改压后再计算。
    """
    for result_index in result_indices:
        output_row = result_index + 2
        try:
            worksheet.Cells(output_row, COL_RESULT_PRESSURE).Value = (
                pressure_settings.first_segment_pressure
            )

            # L = H * 10 * 3600 / I / J。
            worksheet.Cells(output_row, COL_RESULT_FLUX).FormulaR1C1 = "=RC[-4]*10*3600/RC[-3]/RC[-2]"

            # M = L / K；K 已改为前半段压力。
            worksheet.Cells(output_row, COL_RESULT_PERMEABILITY).FormulaR1C1 = "=RC[-1]/RC[-2]"
        except Exception as exc:
            raise ExcelProcessError(f"回修右侧结果区第 {output_row} 行压力失败。") from exc


def calculate_worksheet_for_final_check(worksheet) -> None:
    """最终校验阶段需要读取公式结果，因此计算失败时应明确报错。"""
    try:
        worksheet.Calculate()
    except Exception as exc:
        raise ExcelProcessError("最终校验阶段重新计算 L/M 公式失败。") from exc


def calculate_workbook_before_save(worksheet) -> None:
    """保存前统一强制刷新工作簿公式，确保结果区值已经稳定落盘。"""
    try:
        workbook = getattr(worksheet, "Parent", None)
        application = getattr(worksheet, "Application", None)

        if application is not None:
            try:
                application.CalculateBeforeSave = True
            except Exception:
                pass

            for method_name in ("CalculateFullRebuild", "CalculateFull", "Calculate"):
                method = getattr(application, method_name, None)
                if callable(method):
                    try:
                        method()
                    except Exception:
                        continue

            async_calculate = getattr(application, "CalculateUntilAsyncQueriesDone", None)
            if callable(async_calculate):
                try:
                    async_calculate()
                except Exception:
                    pass

        if workbook is not None:
            try:
                workbook.ForceFullCalculation = True
            except Exception:
                pass

            try:
                workbook.Calculate()
            except Exception:
                pass

        worksheet.Calculate()
    except Exception as exc:
        raise ExcelProcessError("保存前重新计算工作簿失败。") from exc


def ensure_pure_water_result_area_ready(worksheet, result_count: int) -> None:
    """在纯水模式保存前确认 H/M 结果区已有可读实值。"""
    if result_count <= 0:
        raise ExcelProcessError("右侧结果区没有生成数据，无法校验纯水模式结果。")

    h_values = read_result_h_values(worksheet, result_count)
    if len(h_values) != result_count:
        raise ExcelProcessError("纯水模式结果区生成失败，H 列结果数量异常。")

    valid_m_values = read_valid_result_m_values(worksheet, result_count)
    if len(valid_m_values) != result_count:
        raise ExcelProcessError("纯水模式结果区生成失败，M 列存在空值，请检查 K/L/M 公式计算。")


def final_validate_and_repair_pressure_assignment(
    worksheet,
    result_count: int,
    pressure_settings: PressureSettings,
) -> int:
    """
    最终校验并回修压力分配。

    循环规则：
        1. 读取当前所有有效结果行的 M 值，并计算整体平均值。
        2. threshold = mean(all_valid_M_values) * 1.15。
        3. 只检查当前 K 列等于后半段压力的那些行。
        4. 若某行 M > threshold，则该行判定为异常。
        5. 同一轮发现的所有异常行全部改为前半段压力，并统一重算 L/M。
        6. 重复上述流程，直到没有异常值。

    返回：
        被回修为前半段压力的结果行数量。
    """
    if result_count <= 0:
        raise ExcelProcessError("右侧结果区没有生成数据，无法执行最终校验。")

    repaired_count = 0
    max_rounds = result_count + 1

    for _round_number in range(1, max_rounds + 1):
        valid_m_values = read_valid_result_m_values(worksheet, result_count)
        if not valid_m_values:
            raise ExcelProcessError("没有可用于最终校验的有效 M 列渗透性数值。")

        all_m_average = sum(value for _index, value in valid_m_values) / len(valid_m_values)
        threshold = all_m_average * FINAL_PERMEABILITY_OUTLIER_RATIO

        outlier_indices: List[int] = []
        for result_index, m_value in valid_m_values:
            pressure = read_result_pressure_value(worksheet, result_index)

            # 只检查当前 K 列等于后半段压力的行；阈值平均值仍来自全部有效 M 值。
            if (
                pressures_are_equal(pressure, pressure_settings.second_segment_pressure)
                and m_value > threshold
            ):
                outlier_indices.append(result_index)

        if not outlier_indices:
            return repaired_count

        rewrite_result_rows_to_first_pressure(
            worksheet,
            outlier_indices,
            pressure_settings,
        )
        repaired_count += len(outlier_indices)
        calculate_worksheet_for_final_check(worksheet)

    raise ExcelProcessError("最终校验循环次数异常，请检查 M/K 列数据。")


def write_second_segment_permeability_average(
    worksheet,
    result_count: int,
    pressure_settings: PressureSettings,
) -> Optional[float]:
    """
    计算并写入后半段压力对应的渗透性平均值。

    统计范围明确为：
        当前 K 列仍然等于后半段运行压力的所有结果行对应的 M 值。

    这意味着最终校验把某些异常行回修为前半段压力后，这些行不会再计入
    后半段渗透性平均值。平均值写入 M 列最后一个有效结果值的下一行。

    如果最终没有任何 K 列等于后半段压力的结果行，本函数采用“留空”策略：
    清空 M 列最后一个有效值下一行，避免输出误导性的平均值。
    """
    if result_count <= 0:
        raise ExcelProcessError("右侧结果区没有生成数据，无法计算后半段渗透性平均值。")

    average_output_row = result_count + 3
    try:
        worksheet.Cells(average_output_row, COL_RESULT_PERMEABILITY).ClearContents()
    except Exception as exc:
        raise ExcelProcessError("清空后半段渗透性平均值输出单元格失败。") from exc

    values: List[float] = []
    for result_index, m_value in read_valid_result_m_values(worksheet, result_count):
        pressure = read_result_pressure_value(worksheet, result_index)
        if pressures_are_equal(pressure, pressure_settings.second_segment_pressure):
            values.append(m_value)

    if not values:
        return None

    average_value = sum(values) / len(values)

    try:
        # 只把平均值数值写入 M 列最后一个有效值的下一行，不写公式，不覆盖有效数据。
        worksheet.Cells(average_output_row, COL_RESULT_PERMEABILITY).Value = average_value
        worksheet.Cells(average_output_row, COL_RESULT_PERMEABILITY).NumberFormat = "0.000000"
    except Exception as exc:
        raise ExcelProcessError("写入后半段渗透性平均值失败。") from exc

    return average_value


def write_all_result_m_average(worksheet, result_count: int) -> float:
    """
    纯水单一压力模式：计算全部有效 M 值平均值并写入额外平均值单元格。

    当纯水模式下前半段压力与后半段压力相同，本次实验视为全程单一压力。
    此时 M 列最后额外平均值不再表示“后半段平均值”，而是 M3 到最后
    一个有效结果行的全部有效 M 值平均值。
    """
    if result_count <= 0:
        raise ExcelProcessError("右侧结果区没有生成数据，无法计算全部 M 列平均值。")

    average_output_row = result_count + 3
    try:
        worksheet.Cells(average_output_row, COL_RESULT_PERMEABILITY).ClearContents()
    except Exception as exc:
        raise ExcelProcessError("清空全部 M 列平均值输出单元格失败。") from exc

    valid_m_values = read_valid_result_m_values(worksheet, result_count)
    if not valid_m_values:
        raise ExcelProcessError("没有可用于求平均的有效 M 列渗透性数值。")

    average_value = sum(value for _index, value in valid_m_values) / len(valid_m_values)

    try:
        # 只写入数值，位置仍是 M 列最后一个有效值的下一行。
        worksheet.Cells(average_output_row, COL_RESULT_PERMEABILITY).Value = average_value
        worksheet.Cells(average_output_row, COL_RESULT_PERMEABILITY).NumberFormat = "0.000000"
    except Exception as exc:
        raise ExcelProcessError("写入全部 M 列平均值失败。") from exc

    return average_value


def format_output_sheet(worksheet: object, last_output_data_row: int, result_count: int) -> None:
    """做少量格式整理，重点是稳定可读，不做复杂美化。"""
    try:
        if last_output_data_row >= 3:
            worksheet.Range(
                worksheet.Cells(3, COL_DIFF),
                worksheet.Cells(last_output_data_row, COL_AVG_GROUP_2),
            ).NumberFormat = "0.000000"

        if result_count > 0:
            last_result_row = result_count + 2
            worksheet.Range(
                worksheet.Cells(3, COL_RESULT_WEIGHT),
                worksheet.Cells(last_result_row, COL_RESULT_WEIGHT),
            ).NumberFormat = "0.000000"
            worksheet.Range(
                worksheet.Cells(3, COL_RESULT_FLUX),
                worksheet.Cells(last_result_row, COL_RESULT_PERMEABILITY),
            ).NumberFormat = "0.000000"
            worksheet.Range(
                worksheet.Cells(3, COL_RESULT_TIME),
                worksheet.Cells(last_result_row, COL_RESULT_TIME),
            ).NumberFormat = "0"
            worksheet.Range(
                worksheet.Cells(3, COL_RESULT_AREA),
                worksheet.Cells(last_result_row, COL_RESULT_PRESSURE),
            ).NumberFormat = "0.00"

        worksheet.Columns("A:M").AutoFit()
    except Exception:
        # 格式化失败不影响计算结果，不中断保存。
        pass


def build_common_output_layout(
    worksheet,
    source_data_info: SourceDataInfo,
    runtime_settings: RuntimeSettings,
    write_pure_result_area: bool = True,
) -> Tuple[int, int]:
    """
    生成纯水/污染物模式共用的基础布局。

    共用内容包括：
        起点裁剪、插入表头、C 列质量差、D/E 平均值。
        纯水模式还会生成 G:J 结果区基础列；污染物模式会在后续
        单独用 C 列逐行映射生成 G:M。
    返回：
        (左侧输出数据最后一行, 右侧结果区实际数据行数)
    """
    validate_group_size(AVG_GROUP_1, "AVG_GROUP_1")
    validate_group_size(AVG_GROUP_2, "AVG_GROUP_2")

    source_data_row_count = (
        source_data_info.last_b_row - source_data_info.calculation_start_row + 1
    )
    if source_data_row_count < 1:
        raise ExcelProcessError("正式计算起点超出有效数据范围，请检查原始 B 列数据。")

    trim_rows_before_calculation_start(
        worksheet,
        source_data_info.calculation_start_row,
    )

    try:
        # 在顶部插入 1 行后，正式计算起点数据会出现在 A2:B2。
        worksheet.Rows(1).Insert(Shift=XL_SHIFT_DOWN)
    except Exception as exc:
        raise ExcelProcessError("无法在工作表顶部插入表头行。") from exc

    last_output_data_row = source_data_row_count + 1

    clear_generated_regions(worksheet, last_output_data_row)
    write_headers(worksheet)
    write_diff_formulas(worksheet, last_output_data_row)

    avg_group_1_rows = write_average_formulas(
        worksheet,
        COL_AVG_GROUP_1,
        AVG_GROUP_1,
        last_output_data_row,
    )
    write_average_formulas(
        worksheet,
        COL_AVG_GROUP_2,
        AVG_GROUP_2,
        last_output_data_row,
    )

    result_count = len(avg_group_1_rows)

    debug_log_event(
        "build_common_output_layout_after_cde",
        last_output_data_row=last_output_data_row,
        result_count=result_count,
        write_pure_result_area=write_pure_result_area,
        c3=get_cell_debug_state(worksheet, 3, COL_DIFF),
        d7=get_cell_debug_state(worksheet, 7, COL_AVG_GROUP_1),
        e12=get_cell_debug_state(worksheet, 12, COL_AVG_GROUP_2),
    )

    if write_pure_result_area:
        write_result_base_area(worksheet, avg_group_1_rows, runtime_settings)
        debug_log_event(
            "build_common_output_layout_after_gj",
            result_count=result_count,
            time_seconds=runtime_settings.time_seconds,
            membrane_area_cm2=runtime_settings.membrane_area_cm2,
            g3=get_cell_debug_state(worksheet, 3, COL_RESULT_NO),
            h3=get_cell_debug_state(worksheet, 3, COL_RESULT_WEIGHT),
            i3=get_cell_debug_state(worksheet, 3, COL_RESULT_TIME),
            j3=get_cell_debug_state(worksheet, 3, COL_RESULT_AREA),
        )

    try:
        worksheet.Calculate()
    except Exception:
        # 先尝试计算 H 列公式，后续纯水自动识别或污染物公式都依赖这些结果。
        pass

    return last_output_data_row, result_count


def apply_output_layout(
    worksheet,
    source_data_info: SourceDataInfo,
    pressure_settings: PressureSettings,
    runtime_settings: RuntimeSettings,
    injected_switch_index: Optional[int] = None,
) -> OutputLayoutResult:
    """Build the pure-water segmented-pressure output layout."""
    last_output_data_row, result_count = build_common_output_layout(
        worksheet,
        source_data_info,
        runtime_settings,
    )

    if result_count == 0:
        raise ExcelProcessError("??????????????????????????????????")

    h_values = read_result_h_values(worksheet, result_count)
    switch_index, auto_detected = choose_pressure_switch_index(h_values, injected_switch_index)
    write_segmented_pressure_and_formulas(worksheet, result_count, pressure_settings, switch_index)
    debug_log_event(
        "pure_water_segmented_after_klm",
        last_output_data_row=last_output_data_row,
        result_count=result_count,
        switch_index=switch_index,
        switch_auto_detected=auto_detected,
        k3=get_cell_debug_state(worksheet, 3, COL_RESULT_PRESSURE),
        l3=get_cell_debug_state(worksheet, 3, COL_RESULT_FLUX),
        m3=get_cell_debug_state(worksheet, 3, COL_RESULT_PERMEABILITY),
    )

    try:
        worksheet.Calculate()
    except Exception:
        pass

    repaired_count = final_validate_and_repair_pressure_assignment(worksheet, result_count, pressure_settings)
    average_value = write_second_segment_permeability_average(worksheet, result_count, pressure_settings)
    debug_log_event(
        "pure_water_segmented_after_final_check",
        result_count=result_count,
        repaired_count=repaired_count,
        second_segment_average=average_value,
        k3=get_cell_debug_state(worksheet, 3, COL_RESULT_PRESSURE),
        l3=get_cell_debug_state(worksheet, 3, COL_RESULT_FLUX),
        m3=get_cell_debug_state(worksheet, 3, COL_RESULT_PERMEABILITY),
    )
    format_output_sheet(worksheet, last_output_data_row, result_count)

    return OutputLayoutResult(
        result_count=result_count,
        pressure_switch_index=switch_index,
        pressure_switch_auto_detected=auto_detected,
    )


def apply_single_pressure_pure_water_output_layout(
    worksheet,
    source_data_info: SourceDataInfo,
    pressure_bar: float,
    runtime_settings: RuntimeSettings,
) -> OutputLayoutResult:
    """Build the pure-water single-pressure output layout."""
    last_output_data_row, result_count = build_common_output_layout(
        worksheet,
        source_data_info,
        runtime_settings,
    )

    if result_count == 0:
        raise ExcelProcessError("有效液滴质量差数量不足，无法生成每五个取平均结果，请检查原始数据长度")

    write_single_pressure_and_formulas(worksheet, result_count, pressure_bar)
    debug_log_event(
        "pure_water_single_pressure_after_klm",
        last_output_data_row=last_output_data_row,
        result_count=result_count,
        pressure_bar=pressure_bar,
        k3=get_cell_debug_state(worksheet, 3, COL_RESULT_PRESSURE),
        l3=get_cell_debug_state(worksheet, 3, COL_RESULT_FLUX),
        m3=get_cell_debug_state(worksheet, 3, COL_RESULT_PERMEABILITY),
    )

    try:
        worksheet.Calculate()
    except Exception:
        pass

    average_value = write_all_result_m_average(worksheet, result_count)
    debug_log_event(
        "pure_water_single_pressure_average_written",
        result_count=result_count,
        all_m_average=average_value,
        m_average_cell=get_cell_debug_state(worksheet, result_count + 3, COL_RESULT_PERMEABILITY),
    )
    format_output_sheet(worksheet, last_output_data_row, result_count)

    return OutputLayoutResult(
        result_count=result_count,
        pressure_switch_index=0,
        pressure_switch_auto_detected=False,
    )

def apply_pollutant_output_layout(
    worksheet,
    source_data_info: SourceDataInfo,
    pressure_bar: float,
    pure_water_flux: float,
    runtime_settings: RuntimeSettings,
    custom_average_settings: Optional[CustomAverageSettings] = None,
) -> OutputLayoutResult:
    """\u6c61\u67d3\u7269\u6a21\u5f0f\uff1a\u751f\u6210\u57fa\u7840\u5e03\u5c40\u540e\uff0c\u6574\u6bb5\u7ed3\u679c\u4f7f\u7528\u5355\u4e00\u8fd0\u884c\u538b\u529b\u3002"""
    last_output_data_row, _pure_result_count = build_common_output_layout(
        worksheet,
        source_data_info,
        runtime_settings,
        write_pure_result_area=False,
    )
    result_count = write_pollutant_result_base_area(
        worksheet,
        last_output_data_row,
        runtime_settings,
    )
    debug_log_event(
        "pollutant_after_gj",
        last_output_data_row=last_output_data_row,
        result_count=result_count,
        g3=get_cell_debug_state(worksheet, 3, COL_RESULT_NO),
        h3=get_cell_debug_state(worksheet, 3, COL_RESULT_WEIGHT),
        i3=get_cell_debug_state(worksheet, 3, COL_RESULT_TIME),
        j3=get_cell_debug_state(worksheet, 3, COL_RESULT_AREA),
    )
    write_single_pressure_and_formulas(worksheet, result_count, pressure_bar)
    debug_log_event(
        "pollutant_after_klm",
        result_count=result_count,
        pressure_bar=pressure_bar,
        k3=get_cell_debug_state(worksheet, 3, COL_RESULT_PRESSURE),
        l3=get_cell_debug_state(worksheet, 3, COL_RESULT_FLUX),
        m3=get_cell_debug_state(worksheet, 3, COL_RESULT_PERMEABILITY),
    )
    analysis_result = write_pollutant_analysis_area(
        worksheet,
        last_output_data_row,
        pressure_bar,
        pure_water_flux,
        runtime_settings,
    )
    standard_10min_point_count = write_pollutant_standard_10min_area(
        worksheet,
        result_count,
        pure_water_flux,
        last_output_data_row,
    )
    debug_log_event(
        "pollutant_standard_10min_area_done",
        result_count=result_count,
        point_count=standard_10min_point_count,
        x3=get_cell_debug_state(worksheet, 3, COL_STANDARD_10MIN_PERMEABILITY),
        y3=get_cell_debug_state(worksheet, 3, COL_STANDARD_10MIN_NORMALIZED),
        x4=get_cell_debug_state(worksheet, 4, COL_STANDARD_10MIN_PERMEABILITY),
        y4=get_cell_debug_state(worksheet, 4, COL_STANDARD_10MIN_NORMALIZED),
        x8=get_cell_debug_state(worksheet, 8, COL_STANDARD_10MIN_PERMEABILITY),
        y8=get_cell_debug_state(worksheet, 8, COL_STANDARD_10MIN_NORMALIZED),
    )

    try:
        worksheet.Calculate()
    except Exception:
        pass

    create_pollutant_scatter_charts(worksheet, analysis_result)

    if custom_average_settings is not None:
        custom_point_count = write_pollutant_custom_average_area(
            worksheet,
            last_output_data_row,
            pressure_bar,
            pure_water_flux,
            runtime_settings,
            custom_average_settings,
        )
        debug_log_event(
            "pollutant_custom_average_area_done",
            point_count=custom_point_count,
            points_per_group=custom_average_settings.points_per_group,
            name_prefix=custom_average_settings.name_prefix,
            aa3=get_cell_debug_state(worksheet, 3, COL_CUSTOM_AVERAGE_PERMEABILITY),
            ab3=get_cell_debug_state(worksheet, 3, COL_CUSTOM_AVERAGE_NORMALIZED),
            aa4=get_cell_debug_state(worksheet, 4, COL_CUSTOM_AVERAGE_PERMEABILITY),
            ab4=get_cell_debug_state(worksheet, 4, COL_CUSTOM_AVERAGE_NORMALIZED),
        )
        try:
            worksheet.Calculate()
        except Exception:
            pass
        create_custom_average_scatter_sheet(worksheet, custom_average_settings)

    format_output_sheet(worksheet, last_output_data_row, result_count)

    return OutputLayoutResult(
        result_count=result_count,
        pressure_switch_index=0,
        pressure_switch_auto_detected=False,
    )

def save_workbook_as_xlsm(workbook, output_path: Path) -> None:
    """将工作簿另存为 .xlsm。"""
    try:
        workbook.SaveAs(str(output_path), FileFormat=XL_FILE_FORMAT_XLSM)
    except Exception as exc:
        raise ExcelProcessError(
            f"输出文件无法保存：{output_path}\n"
            "可能原因：输出文件被占用、目录无写入权限，或 Excel 不允许写入该路径。"
        ) from exc

    if not output_path.exists():
        raise ExcelProcessError(f"输出文件保存后不存在：{output_path}")

    try:
        output_size = output_path.stat().st_size
    except OSError as exc:
        raise ExcelProcessError(f"无法读取输出文件大小：{output_path}") from exc

    if output_size <= 0:
        raise ExcelProcessError(f"输出文件保存后仍为空文件：{output_path}")


def process_workbook_file(
    input_path: Path,
    output_name_suffix: str,
    layout_builder: Callable,
    output_directory: Optional[Path] = None,
    debug_context: Optional[dict] = None,
) -> Tuple[Path, OutputLayoutResult]:
    """?? Excel COM ?????"""
    ensure_pywin32_available()

    excel = None
    workbook = None
    output_path: Optional[Path] = None
    com_initialized = False
    debug_context = debug_context or {}

    debug_log_event(
        "process_workbook_file_start",
        input_path=input_path,
        output_name_suffix=output_name_suffix,
        output_directory=output_directory,
        **debug_context,
    )

    try:
        pythoncom.CoInitialize()
        com_initialized = True

        excel = start_excel_application()
        workbook = open_workbook_readonly(excel, input_path)
        worksheet = get_worksheet_by_name(workbook, SHEET_NAME)

        source_data_info = validate_source_data(worksheet)
        debug_log_event(
            "validate_source_data_done",
            input_path=input_path,
            last_b_row=source_data_info.last_b_row,
            calculation_start_row=source_data_info.calculation_start_row,
            **debug_context,
        )

        output_path = reserve_unique_output_path(
            input_path,
            output_name_suffix,
            output_directory=output_directory,
        )
        debug_log_event(
            "output_path_reserved",
            input_path=input_path,
            output_path=output_path,
            output_exists=output_path.exists(),
            output_size=get_file_size_text(output_path),
            **debug_context,
        )

        layout_result = layout_builder(worksheet, source_data_info)
        debug_log_event(
            "layout_builder_done",
            input_path=input_path,
            output_path=output_path,
            result_count=layout_result.result_count,
            pressure_switch_index=layout_result.pressure_switch_index,
            pressure_switch_auto_detected=layout_result.pressure_switch_auto_detected,
            c3=get_cell_debug_state(worksheet, 3, COL_DIFF),
            d7=get_cell_debug_state(worksheet, 7, COL_AVG_GROUP_1),
            e12=get_cell_debug_state(worksheet, 12, COL_AVG_GROUP_2),
            g3=get_cell_debug_state(worksheet, 3, COL_RESULT_NO),
            h3=get_cell_debug_state(worksheet, 3, COL_RESULT_WEIGHT),
            i3=get_cell_debug_state(worksheet, 3, COL_RESULT_TIME),
            j3=get_cell_debug_state(worksheet, 3, COL_RESULT_AREA),
            k3=get_cell_debug_state(worksheet, 3, COL_RESULT_PRESSURE),
            l3=get_cell_debug_state(worksheet, 3, COL_RESULT_FLUX),
            m3=get_cell_debug_state(worksheet, 3, COL_RESULT_PERMEABILITY),
            **debug_context,
        )

        calculate_workbook_before_save(worksheet)
        debug_log_event(
            "calculate_workbook_before_save_done",
            input_path=input_path,
            output_path=output_path,
            c3=get_cell_debug_state(worksheet, 3, COL_DIFF),
            d7=get_cell_debug_state(worksheet, 7, COL_AVG_GROUP_1),
            h3=get_cell_debug_state(worksheet, 3, COL_RESULT_WEIGHT),
            l3=get_cell_debug_state(worksheet, 3, COL_RESULT_FLUX),
            m3=get_cell_debug_state(worksheet, 3, COL_RESULT_PERMEABILITY),
            **debug_context,
        )

        if output_name_suffix == PURE_WATER_OUTPUT_NAME_SUFFIX:
            ensure_pure_water_result_area_ready(worksheet, layout_result.result_count)

        debug_log_event(
            "before_save_workbook",
            output_path=output_path,
            output_exists=output_path.exists(),
            output_size=get_file_size_text(output_path),
            **debug_context,
        )
        prepare_output_path_for_excel_save(output_path)
        save_workbook_as_xlsm(workbook, output_path)
        debug_log_event(
            "after_save_workbook",
            output_path=output_path,
            output_exists=output_path.exists(),
            output_size=get_file_size_text(output_path),
            **debug_context,
        )

        return output_path, layout_result
    except ExcelProcessError:
        if output_path is not None:
            remove_reserved_output_file(output_path)
        raise
    except Exception as exc:
        if output_path is not None:
            remove_reserved_output_file(output_path)

        if is_com_error(exc):
            raise ExcelProcessError(f"Excel COM ?????{exc}") from exc
        raise
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass

        if excel is not None:
            try:
                excel.DisplayAlerts = False
                excel.Quit()
            except Exception:
                pass

        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def process_pure_water(
    input_path: Path,
    pressure_settings: PressureSettings,
    runtime_settings: RuntimeSettings,
    output_directory: Optional[Path] = None,
    manual_switch_index: Optional[int] = None,
) -> Tuple[Path, OutputLayoutResult]:
    """???????"""
    single_pressure_branch = pressures_are_equal(
        pressure_settings.first_segment_pressure,
        pressure_settings.second_segment_pressure,
    )
    branch_name = "single_pressure" if single_pressure_branch else "segmented_pressure"
    debug_context = {
        "experiment_mode": MODE_PURE_WATER,
        "time_seconds": runtime_settings.time_seconds,
        "membrane_area_cm2": runtime_settings.membrane_area_cm2,
        "first_segment_pressure": pressure_settings.first_segment_pressure,
        "second_segment_pressure": pressure_settings.second_segment_pressure,
        "is_single_pressure_branch": single_pressure_branch,
        "is_segmented_pressure_branch": not single_pressure_branch,
        "manual_switch_index": manual_switch_index,
    }

    debug_log_event(
        "process_pure_water_start",
        input_path=input_path,
        output_directory=output_directory,
        branch_name=branch_name,
        **debug_context,
    )

    if single_pressure_branch:
        return process_workbook_file(
            input_path,
            PURE_WATER_OUTPUT_NAME_SUFFIX,
            lambda worksheet, source_data_info: apply_single_pressure_pure_water_output_layout(
                worksheet,
                source_data_info,
                pressure_settings.first_segment_pressure,
                runtime_settings,
            ),
            output_directory=output_directory,
            debug_context=debug_context,
        )

    return process_workbook_file(
        input_path,
        PURE_WATER_OUTPUT_NAME_SUFFIX,
        lambda worksheet, source_data_info: apply_output_layout(
            worksheet,
            source_data_info,
            pressure_settings,
            runtime_settings,
            injected_switch_index=manual_switch_index,
        ),
        output_directory=output_directory,
        debug_context=debug_context,
    )


def process_pollutant(
    input_path: Path,
    pressure_bar: float,
    pure_water_flux: float,
    runtime_settings: RuntimeSettings,
    output_directory: Optional[Path] = None,
    custom_average_settings: Optional[CustomAverageSettings] = None,
) -> Tuple[Path, OutputLayoutResult]:
    """?????????????????????????"""
    debug_context = {
        "experiment_mode": MODE_POLLUTANT,
        "time_seconds": runtime_settings.time_seconds,
        "membrane_area_cm2": runtime_settings.membrane_area_cm2,
        "pressure_bar": pressure_bar,
        "pure_water_flux": pure_water_flux,
        "custom_average_settings": custom_average_settings,
    }
    debug_log_event(
        "process_pollutant_start",
        input_path=input_path,
        output_directory=output_directory,
        **debug_context,
    )
    return process_workbook_file(
        input_path,
        POLLUTANT_OUTPUT_NAME_SUFFIX,
        lambda worksheet, source_data_info: apply_pollutant_output_layout(
            worksheet,
            source_data_info,
            pressure_bar,
            pure_water_flux,
            runtime_settings,
            custom_average_settings,
        ),
        output_directory=output_directory,
        debug_context=debug_context,
    )


def process_excel_file(
    input_path: Path,
    pressure_settings: PressureSettings,
    runtime_settings: Optional[RuntimeSettings] = None,
    output_directory: Optional[Path] = None,
) -> Tuple[Path, OutputLayoutResult]:
    """?????????????????"""
    if runtime_settings is None:
        runtime_settings = RuntimeSettings(
            time_seconds=DEFAULT_TIME_SECONDS,
            membrane_area_cm2=DEFAULT_MEMBRANE_AREA_CM2,
        )

    return process_pure_water(
        input_path,
        pressure_settings,
        runtime_settings,
        output_directory=output_directory,
    )


def process_with_test_mode(
    input_path: Path,
    test_mode_options: TestModeOptions,
) -> Tuple[Path, OutputLayoutResult, str]:
    """???????? main ???????????????"""
    debug_log_event(
        "process_with_test_mode_start",
        input_path=input_path,
        experiment_mode=test_mode_options.experiment_mode,
        time_seconds=test_mode_options.runtime_settings.time_seconds,
        membrane_area_cm2=test_mode_options.runtime_settings.membrane_area_cm2,
        pressure_settings=test_mode_options.pressure_settings,
        pressure_bar=test_mode_options.pressure_bar,
        pure_water_flux=test_mode_options.pure_water_flux,
        custom_average_settings=test_mode_options.custom_average_settings,
        manual_switch_index=test_mode_options.manual_switch_index,
        output_directory=test_mode_options.output_directory,
        allow_xlsm_input=test_mode_options.allow_xlsm_input,
    )
    validated_input_path = validate_input_file(
        input_path,
        allow_xlsm_input=test_mode_options.allow_xlsm_input,
    )
    runtime_settings = test_mode_options.runtime_settings

    if test_mode_options.experiment_mode == MODE_PURE_WATER:
        if test_mode_options.pressure_settings is None:
            raise ExcelProcessError("纯水模式缺少前半段/后半段压力参数。")

        output_path, layout_result = process_pure_water(
            validated_input_path,
            test_mode_options.pressure_settings,
            runtime_settings,
            output_directory=test_mode_options.output_directory,
            manual_switch_index=test_mode_options.manual_switch_index,
        )
        mode_detail = _build_mode_detail(
            MODE_PURE_WATER,
            runtime_settings,
            layout_result,
        )
        return output_path, layout_result, mode_detail

    if test_mode_options.experiment_mode == MODE_POLLUTANT:
        if test_mode_options.pressure_bar is None:
            raise ExcelProcessError("污染物模式缺少运行压力参数。")
        if test_mode_options.pure_water_flux is None:
            raise ExcelProcessError("污染物模式缺少纯水通量参数。")

        output_path, layout_result = process_pollutant(
            validated_input_path,
            test_mode_options.pressure_bar,
            test_mode_options.pure_water_flux,
            runtime_settings,
            output_directory=test_mode_options.output_directory,
            custom_average_settings=test_mode_options.custom_average_settings,
        )
        mode_detail = _build_mode_detail(
            MODE_POLLUTANT,
            runtime_settings,
            layout_result,
            pressure_bar=test_mode_options.pressure_bar,
            pure_water_flux=test_mode_options.pure_water_flux,
        )
        return output_path, layout_result, mode_detail

    raise ExcelProcessError(f"未知实验模式：{test_mode_options.experiment_mode}")


def main(
    argv: Optional[Sequence[str]] = None,
    test_mode_options: Optional[TestModeOptions] = None,
) -> int:
    """?????"""
    if argv is None:
        argv = sys.argv

    should_show_messages = test_mode_options is None or test_mode_options.show_messages

    try:
        input_path = parse_input_args(argv)
        debug_log_event(
            "main_start",
            argv=list(argv),
            input_path_from_args=input_path,
            has_test_mode_options=test_mode_options is not None,
        )

        if test_mode_options is not None:
            if input_path is None:
                raise ExcelProcessError("测试模式必须通过命令行参数提供输入文件路径。")

            debug_log_event(
                "main_test_mode_path",
                input_path=input_path,
                experiment_mode=test_mode_options.experiment_mode,
                time_seconds=test_mode_options.runtime_settings.time_seconds,
                membrane_area_cm2=test_mode_options.runtime_settings.membrane_area_cm2,
                pressure_settings=test_mode_options.pressure_settings,
                pressure_bar=test_mode_options.pressure_bar,
                pure_water_flux=test_mode_options.pure_water_flux,
                custom_average_settings=test_mode_options.custom_average_settings,
                manual_switch_index=test_mode_options.manual_switch_index,
                output_directory=test_mode_options.output_directory,
            )
            output_path, layout_result, mode_detail = process_with_test_mode(
                input_path,
                test_mode_options,
            )
        else:
            experiment_mode = select_experiment_mode()
            runtime_settings = get_time_and_area_inputs()
            debug_log_event(
                "main_manual_after_mode_and_runtime",
                experiment_mode=experiment_mode,
                time_seconds=runtime_settings.time_seconds,
                membrane_area_cm2=runtime_settings.membrane_area_cm2,
            )

            if input_path is None:
                input_path = select_input_file()

            input_path = validate_input_file(input_path)
            debug_log_event(
                "main_manual_input_file_validated",
                experiment_mode=experiment_mode,
                input_path=input_path,
                time_seconds=runtime_settings.time_seconds,
                membrane_area_cm2=runtime_settings.membrane_area_cm2,
            )

            if experiment_mode == MODE_PURE_WATER:
                pressure_settings = get_pressure_inputs()
                single_pressure_branch = pressures_are_equal(
                    pressure_settings.first_segment_pressure,
                    pressure_settings.second_segment_pressure,
                )
                debug_log_event(
                    "main_manual_pure_water_before_core",
                    input_path=input_path,
                    time_seconds=runtime_settings.time_seconds,
                    membrane_area_cm2=runtime_settings.membrane_area_cm2,
                    first_segment_pressure=pressure_settings.first_segment_pressure,
                    second_segment_pressure=pressure_settings.second_segment_pressure,
                    is_single_pressure_branch=single_pressure_branch,
                    is_segmented_pressure_branch=not single_pressure_branch,
                )
                output_path, layout_result = process_pure_water(
                    input_path,
                    pressure_settings,
                    runtime_settings,
                )
                mode_detail = _build_mode_detail(
                    MODE_PURE_WATER,
                    runtime_settings,
                    layout_result,
                )
            elif experiment_mode == MODE_POLLUTANT:
                pressure_bar = get_pressure_input("运行压力")
                pure_water_flux = get_pure_water_flux_input()
                custom_average_settings = get_pollutant_custom_average_settings(runtime_settings)
                debug_log_event(
                    "main_manual_pollutant_before_core",
                    input_path=input_path,
                    time_seconds=runtime_settings.time_seconds,
                    membrane_area_cm2=runtime_settings.membrane_area_cm2,
                    pressure_bar=pressure_bar,
                    pure_water_flux=pure_water_flux,
                    custom_average_settings=custom_average_settings,
                )
                output_path, layout_result = process_pollutant(
                    input_path,
                    pressure_bar,
                    pure_water_flux,
                    runtime_settings,
                    custom_average_settings=custom_average_settings,
                )
                mode_detail = _build_mode_detail(
                    MODE_POLLUTANT,
                    runtime_settings,
                    layout_result,
                    pressure_bar=pressure_bar,
                    pure_water_flux=pure_water_flux,
                )
            else:
                raise ExcelProcessError(f"未知实验模式：{experiment_mode}")

        debug_log_event(
            "main_done",
            input_path=input_path,
            output_path=output_path,
            output_exists=output_path.exists(),
            output_size=get_file_size_text(output_path),
            result_count=layout_result.result_count,
            pressure_switch_index=layout_result.pressure_switch_index,
            pressure_switch_auto_detected=layout_result.pressure_switch_auto_detected,
        )

        if should_show_messages:
            show_user_message(
                "完成",
                "Excel 数据处理完成。\n"
                f"输入文件：{input_path}\n"
                f"输出文件：{output_path}\n"
                f"{mode_detail}",
                is_error=False,
            )
        return 0

    except UserCancelledError as exc:
        debug_log_event("main_user_cancelled", message=exc)
        if should_show_messages:
            show_user_message("已取消", str(exc), is_error=False)
        return 1
    except ExcelProcessError as exc:
        debug_log_event("main_excel_process_error", message=exc)
        if should_show_messages:
            show_user_message("错误", str(exc), is_error=True)
        return 2
    except Exception as exc:
        detail = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
        debug_log_event("main_unexpected_error", message=exc, traceback=detail)
        if should_show_messages:
            show_user_message(
                "未知错误",
                f"{exc}\n\n详细信息：\n{detail}",
                is_error=True,
            )
        return 3
    finally:
        global _TK_ROOT
        if _TK_ROOT is not None:
            try:
                for child in _TK_ROOT.winfo_children():
                    child.destroy()
                _TK_ROOT.destroy()
            except Exception:
                pass
            _TK_ROOT = None


if __name__ == "__main__":
    sys.exit(main())
