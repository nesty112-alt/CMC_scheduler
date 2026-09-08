"""Create sample files and verify scheduler execution paths."""
import os
import sqlite3
import subprocess
import sys
import time
import traceback
import uuid
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "samples"
OUTPUT = SAMPLES / "output"
VALIDATION = SAMPLES / "validation"
XLSM_PATH = SAMPLES / "sample_macro.xlsm"
PYTHON_JOB = SAMPLES / "sample_job.py"
PYTHON_RESULT = OUTPUT / "python_result.txt"
VBA_RESULT = OUTPUT / "vba_result.txt"
SCHEDULER_RESULT = OUTPUT / "scheduler_fired.txt"
RULES_XLSX = VALIDATION / "rules.xlsx"
MAIN_CSV = VALIDATION / "main.csv"

sys.path.insert(0, str(ROOT))
os.chdir(ROOT)

RESULTS = []


def record(name: str, ok: bool, detail: str = "") -> None:
    RESULTS.append((name, ok, detail))
    status = "PASS" if ok else "FAIL"
    print(f"[{status}] {name}" + (f" - {detail}" if detail else ""))


def reset_outputs() -> None:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    for path in (PYTHON_RESULT, VBA_RESULT, SCHEDULER_RESULT):
        if path.exists():
            path.unlink()


def create_validation_rules() -> None:
    import pandas as pd

    df = pd.DataFrame(
        [
            {
                "number": 1,
                "Rule_Name": "금액이 0 이하",
                "SQL_Query": "SELECT * FROM main WHERE amount <= 0",
            },
            {
                "number": 2,
                "Rule_Name": "상태값이 ERR",
                "SQL_Query": "SELECT * FROM main WHERE status = 'ERR'",
            },
            {
                "number": 3,
                "Rule_Name": "정상 건수 확인",
                "SQL_Query": "SELECT * FROM main WHERE amount < -99999",
            },
        ]
    )
    VALIDATION.mkdir(parents=True, exist_ok=True)
    df.to_excel(RULES_XLSX, index=False)


def _excel_accessvbom_key() -> str:
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    try:
        major = str(excel.Version).split(".")[0]
    finally:
        excel.Quit()
    return rf"Software\Microsoft\Office\{major}.0\Excel\Security"


def enable_accessvbom_temporarily():
    import winreg

    key_path = _excel_accessvbom_key()
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
    try:
        try:
            previous = winreg.QueryValueEx(key, "AccessVBOM")[0]
            had_value = True
        except FileNotFoundError:
            previous = None
            had_value = False
        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
        return key_path, had_value, previous
    finally:
        winreg.CloseKey(key)


def restore_accessvbom(state) -> None:
    import winreg

    key_path, had_value, previous = state
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
    try:
        if had_value:
            winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, int(previous))
        else:
            winreg.DeleteValue(key, "AccessVBOM")
    except FileNotFoundError:
        pass
    finally:
        winreg.CloseKey(key)


def create_excel_macro_workbook() -> str:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Add()
        try:
            project = wb.VBProject
        except Exception as exc:
            raise RuntimeError(
                "Excel에서 'VBA 프로젝트 개체 모델에 액세스할 수 있음'이 꺼져 있습니다. "
                f"({exc})"
            ) from exc

        component = project.VBComponents.Add(1)  # vbext_ct_StdModule
        component.Name = "SampleModule"
        result_path = str(VBA_RESULT).replace("\\", "\\\\")
        component.CodeModule.AddFromString(
            "\n".join(
                [
                    "Public Sub WriteSampleResult()",
                    "    Dim fso As Object",
                    "    Dim ts As Object",
                    f'    Set fso = CreateObject("Scripting.FileSystemObject")',
                    f'    Set ts = fso.CreateTextFile("{result_path}", True)',
                    '    ts.WriteLine "VBA_SAMPLE_OK " & Format(Now, "yyyy-mm-dd hh:nn:ss")',
                    "    ts.Close",
                    "End Sub",
                    "",
                    "Public Function SampleHelper() As String",
                    '    SampleHelper = "helper"',
                    "End Function",
                ]
            )
        )
        if XLSM_PATH.exists():
            XLSM_PATH.unlink()
        wb.SaveAs(str(XLSM_PATH), FileFormat=52)  # xlOpenXMLWorkbookMacroEnabled
        return str(XLSM_PATH)
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def test_python_job() -> None:
    from scheduler_app import list_excel_macros  # noqa: F401  ensure import works

    flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    result = subprocess.run(
        [sys.executable, str(PYTHON_JOB)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=True,
        cwd=str(SAMPLES),
        creationflags=flags,
    )
    ok = PYTHON_RESULT.exists() and "CMC_SAMPLE_OK" in PYTHON_RESULT.read_text(encoding="utf-8")
    record("Python 샘플 실행", ok, result.stdout.strip())


def test_scheduler_trigger() -> None:
    from apscheduler.schedulers.background import BackgroundScheduler
    from apscheduler.triggers.date import DateTrigger
    from tzlocal import get_localzone

    fired = {"ok": False}

    def job():
        OUTPUT.mkdir(parents=True, exist_ok=True)
        SCHEDULER_RESULT.write_text("SCHEDULER_OK\n", encoding="utf-8")
        fired["ok"] = True

    tz = get_localzone()
    scheduler = BackgroundScheduler(timezone=tz)
    scheduler.start()
    run_at = datetime.now() + timedelta(seconds=2)
    scheduler.add_job(job, trigger=DateTrigger(run_date=run_at, timezone=tz))
    deadline = time.time() + 8
    while time.time() < deadline and not fired["ok"]:
        time.sleep(0.2)
    scheduler.shutdown(wait=False)
    ok = fired["ok"] and SCHEDULER_RESULT.exists()
    record("APScheduler 1회 트리거", ok, f"실행시각 {run_at.strftime('%H:%M:%S')}")


def test_macro_list_and_run() -> None:
    from scheduler_app import list_excel_macros

    access_state = None
    try:
        access_state = enable_accessvbom_temporarily()
        create_excel_macro_workbook()
    except Exception as exc:
        record("Excel 매크로 샘플 생성", False, str(exc))
        return
    finally:
        if access_state is not None:
            try:
                restore_accessvbom(access_state)
            except Exception:
                pass

    record("Excel 매크로 샘플 생성", True, str(XLSM_PATH.name))

    macros = list_excel_macros(str(XLSM_PATH))
    has_target = any(name.lower().endswith("writesampleresult") for name in macros)
    record("VBA 매크로 목록 읽기", has_target, ", ".join(macros) if macros else "목록 없음")
    if not has_target:
        return

    target = next(name for name in macros if name.lower().endswith("writesampleresult"))
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        abs_path = os.path.abspath(str(XLSM_PATH))
        wb = excel.Workbooks.Open(abs_path)
        excel.Application.Run(f"'{abs_path}'!{target}")
        wb.Save()
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    ok = VBA_RESULT.exists() and "VBA_SAMPLE_OK" in VBA_RESULT.read_text(encoding="utf-8")
    record("VBA 매크로 실행", ok, VBA_RESULT.read_text(encoding="utf-8").strip() if ok else "결과 파일 없음")


def test_validation_sql() -> None:
    import pandas as pd

    create_validation_rules()
    conn = sqlite3.connect(":memory:")
    try:
        df_data = pd.read_csv(MAIN_CSV, encoding="utf-8")
        df_data.to_sql("main", conn, if_exists="replace", index=False)
        df_rules = pd.read_excel(RULES_XLSX)
        counts = []
        for _, row in df_rules.iterrows():
            result_df = pd.read_sql_query(row["SQL_Query"], conn)
            counts.append((row["Rule_Name"], len(result_df)))
        expected = {
            "금액이 0 이하": 3,
            "상태값이 ERR": 2,
            "정상 건수 확인": 0,
        }
        ok = all(expected[name] == cnt for name, cnt in counts)
        detail = ", ".join(f"{name}={cnt}건" for name, cnt in counts)
        record("검증 SQL 샘플", ok, detail)
    finally:
        conn.close()


def main() -> int:
    print("=== CMC 스케줄러 샘플 검증 ===")
    reset_outputs()
    try:
        test_python_job()
    except Exception:
        record("Python 샘플 실행", False, traceback.format_exc())
    try:
        test_scheduler_trigger()
    except Exception:
        record("APScheduler 1회 트리거", False, traceback.format_exc())
    try:
        test_macro_list_and_run()
    except Exception:
        record("VBA 매크로 검증", False, traceback.format_exc())
    try:
        test_validation_sql()
    except Exception:
        record("검증 SQL 샘플", False, traceback.format_exc())

    print("\n=== 결과 요약 ===")
    passed = sum(1 for _, ok, _ in RESULTS if ok)
    failed = sum(1 for _, ok, _ in RESULTS if not ok)
    for name, ok, detail in RESULTS:
        print(f" - {name}: {'성공' if ok else '실패'}{(' (' + detail + ')') if detail else ''}")
    print(f"총 {len(RESULTS)}건 중 성공 {passed}, 실패 {failed}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
