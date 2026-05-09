# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
import importlib.util
import os
import sys
import traceback
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Dict, List

PROJECT_DIR = Path(__file__).resolve().parent
MAIN_SCRIPT_PATH = PROJECT_DIR / "process_excel-1.py"
TEST_INPUT_DIR = PROJECT_DIR / "test_excels"
TEST_INPUT_FILE = PROJECT_DIR / "test_excels.xls"
TEST_OUTPUT_DIR = PROJECT_DIR / "batch_test_outputs"
REPORT_TXT_PATH = PROJECT_DIR / "batch_test_report.txt"
REPORT_CSV_PATH = PROJECT_DIR / "batch_test_report.csv"
ALLOWED_SUFFIXES = {".xls", ".xlsm"}


def load_main_module():
    spec = importlib.util.spec_from_file_location("process_excel_1_module", MAIN_SCRIPT_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"\u65e0\u6cd5\u52a0\u8f7d\u4e3b\u7a0b\u5e8f\u811a\u672c\uff1a{MAIN_SCRIPT_PATH}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    if hasattr(module, "debug_log_event"):
        module.debug_log_event(
            "batch_self_check_loaded_main_module",
            batch_file=Path(__file__).resolve(),
            batch_sys_executable=sys.executable,
            batch_cwd=os.getcwd(),
            main_script_path=MAIN_SCRIPT_PATH,
            loaded_module_file=getattr(module, "__file__", ""),
        )
    return module


def iter_test_files() -> List[Path]:
    if TEST_INPUT_DIR.exists() and TEST_INPUT_DIR.is_dir():
        files = [
            path
            for path in TEST_INPUT_DIR.rglob("*")
            if path.is_file() and path.suffix.lower() in ALLOWED_SUFFIXES
        ]
        return sorted(files, key=lambda item: str(item.relative_to(TEST_INPUT_DIR)).lower())

    if TEST_INPUT_FILE.exists() and TEST_INPUT_FILE.is_file() and TEST_INPUT_FILE.suffix.lower() in ALLOWED_SUFFIXES:
        return [TEST_INPUT_FILE]

    return []


def build_runtime_settings(module):
    return module.RuntimeSettings(
        time_seconds=module.DEFAULT_TIME_SECONDS,
        membrane_area_cm2=module.DEFAULT_MEMBRANE_AREA_CM2,
    )


def build_test_options(module, mode: str):
    runtime_settings = build_runtime_settings(module)

    if mode == module.MODE_PURE_WATER:
        return module.TestModeOptions(
            experiment_mode=module.MODE_PURE_WATER,
            runtime_settings=runtime_settings,
            pressure_settings=module.PressureSettings(1.5, 1.0),
            manual_switch_index=1,
            output_directory=TEST_OUTPUT_DIR / module.MODE_PURE_WATER,
            allow_xlsm_input=True,
            show_messages=False,
        )

    if mode == module.MODE_POLLUTANT:
        return module.TestModeOptions(
            experiment_mode=module.MODE_POLLUTANT,
            runtime_settings=runtime_settings,
            pressure_bar=1.5,
            pure_water_flux=1.0,
            output_directory=TEST_OUTPUT_DIR / module.MODE_POLLUTANT,
            allow_xlsm_input=True,
            show_messages=False,
        )

    raise ValueError(f"\u672a\u77e5\u6d4b\u8bd5\u6a21\u5f0f\uff1a{mode}")


def format_failure_reason(error_text: str) -> str:
    first_line = (error_text or "").strip().splitlines()
    if not first_line:
        return "\u672a\u77e5\u9519\u8bef"
    return first_line[0].strip()


def run_single_case(module, input_path: Path, mode: str) -> Dict[str, str]:
    options = build_test_options(module, mode)
    started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    record: Dict[str, str] = {
        "file_name": input_path.name,
        "relative_path": str(input_path.relative_to(TEST_INPUT_DIR)) if TEST_INPUT_DIR.exists() and TEST_INPUT_DIR.is_dir() else input_path.name,
        "mode": mode,
        "started_at": started_at,
        "success": "False",
        "output_path": "",
        "result_count": "",
        "error_message": "",
        "failure_reason": "",
        "traceback": "",
    }

    try:
        if hasattr(module, "debug_log_event"):
            module.debug_log_event(
                "batch_self_check_case_start",
                batch_file=Path(__file__).resolve(),
                batch_sys_executable=sys.executable,
                batch_cwd=os.getcwd(),
                input_path=input_path.resolve(strict=False),
                mode=mode,
                time_seconds=options.runtime_settings.time_seconds,
                membrane_area_cm2=options.runtime_settings.membrane_area_cm2,
                pressure_settings=options.pressure_settings,
                pressure_bar=options.pressure_bar,
                pure_water_flux=options.pure_water_flux,
                manual_switch_index=options.manual_switch_index,
                output_directory=options.output_directory,
                allow_xlsm_input=options.allow_xlsm_input,
            )
        output_path, layout_result, _mode_detail = module.process_with_test_mode(input_path, options)
        record["success"] = "True"
        record["output_path"] = str(output_path)
        record["result_count"] = str(layout_result.result_count)
        if hasattr(module, "debug_log_event"):
            module.debug_log_event(
                "batch_self_check_case_done",
                input_path=input_path.resolve(strict=False),
                output_path=output_path,
                mode=mode,
                result_count=layout_result.result_count,
                pressure_switch_index=layout_result.pressure_switch_index,
                pressure_switch_auto_detected=layout_result.pressure_switch_auto_detected,
                output_exists=Path(output_path).exists(),
                output_size=module.get_file_size_text(Path(output_path)) if hasattr(module, "get_file_size_text") else "",
            )
    except Exception as exc:
        error_text = str(exc)
        record["error_message"] = error_text
        record["failure_reason"] = format_failure_reason(error_text)
        record["traceback"] = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
        if hasattr(module, "debug_log_event"):
            module.debug_log_event(
                "batch_self_check_case_error",
                input_path=input_path.resolve(strict=False),
                mode=mode,
                error_message=error_text,
                traceback=record["traceback"],
            )

    return record


def write_csv_report(records: List[Dict[str, str]]) -> None:
    REPORT_CSV_PATH.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "file_name",
        "relative_path",
        "mode",
        "started_at",
        "success",
        "output_path",
        "result_count",
        "failure_reason",
        "error_message",
        "traceback",
    ]
    with REPORT_CSV_PATH.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(records)


def write_text_report(records: List[Dict[str, str]], test_files: List[Path]) -> None:
    total_cases = len(records)
    success_count = sum(1 for record in records if record["success"] == "True")
    failure_records = [record for record in records if record["success"] != "True"]
    failure_counter = Counter(record["failure_reason"] for record in failure_records if record["failure_reason"])

    lines: List[str] = []
    lines.append("\u6279\u91cf\u81ea\u67e5\u62a5\u544a")
    lines.append(f"\u751f\u6210\u65f6\u95f4\uff1a{datetime.now():%Y-%m-%d %H:%M:%S}")
    input_source = TEST_INPUT_DIR if TEST_INPUT_DIR.exists() and TEST_INPUT_DIR.is_dir() else TEST_INPUT_FILE
    lines.append(f"\u6d4b\u8bd5\u8f93\u5165\u6e90\uff1a{input_source}")
    lines.append(f"\u6d4b\u8bd5\u6587\u4ef6\u6570\u91cf\uff1a{len(test_files)}")
    lines.append(f"\u6d4b\u8bd5\u7528\u4f8b\u6570\u91cf\uff1a{total_cases}")
    lines.append(f"\u6210\u529f\u6570\u91cf\uff1a{success_count}")
    lines.append(f"\u5931\u8d25\u6570\u91cf\uff1a{len(failure_records)}")
    lines.append("")

    if not test_files:
        lines.append("\u5f53\u524d\u672a\u627e\u5230 test_excels \u76ee\u5f55\uff0c\u4e5f\u672a\u627e\u5230 test_excels.xls \u6d4b\u8bd5\u6587\u4ef6\u3002")
        lines.append("")

    if failure_counter:
        lines.append("\u6700\u5e38\u89c1\u5931\u8d25\u539f\u56e0\uff1a")
        for reason, count in failure_counter.most_common():
            lines.append(f"- {reason}\uff1a{count} \u6b21")
        lines.append("")
    elif total_cases > 0:
        lines.append("\u672c\u8f6e\u6ca1\u6709\u5931\u8d25\u7528\u4f8b\u3002")
        lines.append("")

    lines.append("\u9010\u9879\u7ed3\u679c\uff1a")
    if not records:
        lines.append("- \u65e0\u53ef\u6267\u884c\u6d4b\u8bd5\u7528\u4f8b")
    else:
        for index, record in enumerate(records, start=1):
            status_text = "\u6210\u529f" if record["success"] == "True" else "\u5931\u8d25"
            lines.append(f"{index}. \u6587\u4ef6\uff1a{record['relative_path']}")
            lines.append(f"   \u6a21\u5f0f\uff1a{record['mode']}")
            lines.append(f"   \u72b6\u6001\uff1a{status_text}")
            lines.append(f"   \u8f93\u51fa\uff1a{record['output_path'] or '\u65e0'}")
            if record["result_count"]:
                lines.append(f"   \u7ed3\u679c\u884c\u6570\uff1a{record['result_count']}")
            if record["failure_reason"]:
                lines.append(f"   \u5931\u8d25\u539f\u56e0\uff1a{record['failure_reason']}")
            if record["traceback"]:
                lines.append("   \u5806\u6808\uff1a")
                for stack_line in record["traceback"].rstrip().splitlines():
                    lines.append(f"     {stack_line}")
            lines.append("")

    REPORT_TXT_PATH.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    module = load_main_module()
    TEST_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    test_files = iter_test_files()
    records: List[Dict[str, str]] = []

    for input_path in test_files:
        for mode in (module.MODE_PURE_WATER, module.MODE_POLLUTANT):
            print(f"\u6b63\u5728\u6d4b\u8bd5\uff1a{input_path.name} [{mode}]")
            records.append(run_single_case(module, input_path, mode))

    write_csv_report(records)
    write_text_report(records, test_files)

    failure_records = [record for record in records if record["success"] != "True"]
    print(f"\u6d4b\u8bd5\u6587\u4ef6\u6570\u91cf\uff1a{len(test_files)}")
    print(f"\u6d4b\u8bd5\u7528\u4f8b\u6570\u91cf\uff1a{len(records)}")
    print(f"\u5931\u8d25\u6570\u91cf\uff1a{len(failure_records)}")
    print(f"CSV \u62a5\u544a\uff1a{REPORT_CSV_PATH}")
    print(f"TXT \u62a5\u544a\uff1a{REPORT_TXT_PATH}")

    return 1 if failure_records else 0


if __name__ == "__main__":
    raise SystemExit(main())
