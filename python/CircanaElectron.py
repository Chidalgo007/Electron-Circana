import gc
import os
import sys
import time
import platform
import ctypes
import pythoncom
import win32com.client as win32
import win32process
import io

# === UTF-8 output for Electron logs ===
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# === Prevent sleep constants ===
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002 # Optional: keep screen awake too

excel_pid = None

# === Helper functions ===
def force_console_output(message):
    """Print and flush immediately (so Electron sees it)."""
    try:
        print(message)
    except UnicodeEncodeError:
        print(message.encode("utf-8", errors="replace").decode("utf-8"))
    sys.stdout.flush()
    sys.stderr.flush()

def prevent_sleep():
    if platform.system() == "Windows":
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
        )

def allow_sleep():
    if platform.system() == "Windows":
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

def start_excel():
    """Start Excel and track PID for external kill."""
    global excel_pid
    excel_app = win32.DispatchEx("Excel.Application")
    hwnd = excel_app.Hwnd
    _, excel_pid = win32process.GetWindowThreadProcessId(hwnd)
    force_console_output(f"✅ Excel started with PID:{excel_pid}")
    return excel_app

# === Main automation function ===
def automate_excel_process(file_path):
    global excel_pid
    try:
        pythoncom.CoInitialize()

        excel = start_excel()
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"[Circana] File not found: {file_path}")

        force_console_output(f"[Circana] 🔄 Opening workbook: {os.path.basename(file_path)}")
        wb = excel.Workbooks.Open(file_path, UpdateLinks=0, ReadOnly=False)
        force_console_output("[Circana] ✅ Workbook opened successfully")

        # === 1. Refresh connections ===
        for idx, conn in enumerate(wb.Connections, start=1):
            try:
                force_console_output(f"[Circana] 🔄 Refreshing connection {idx}: {conn.Name}")
                conn.Refresh()
                time.sleep(2)
                force_console_output(f"[Circana] ✅ Connection {idx} refreshed")
            except Exception as e:
                force_console_output(f"[Circana] ⚠️ Connection {idx} error: {e}")

        # === 2. Update Date Pivot ===
        try:
            force_console_output("[Circana] 🔄 Updating Dates")
            dates_sheet = wb.Sheets("Dates")
            pivot = dates_sheet.PivotTables("PivotTable2")
            field_name = "[TSM].[Date].[Date]"
            pivot_field = pivot.PivotFields(field_name)
            pivot_field.ClearAllFilters()
            items = [item.Name for item in pivot_field.PivotItems()]
            non_blank_items = [d for d in items if d and not d.strip().endswith(".&") and d.strip() != ""]
            if not non_blank_items:
                raise Exception("No non-blank dates found")
            pivot_field.VisibleItemsList = non_blank_items
            force_console_output("[Circana] ✅ Dates updated")
        except Exception as e:
            force_console_output(f"[Circana] ❌ OLAP pivot failed: {e}")

        # === 3. Update Nespresso Pivot ===
        try:
            force_console_output("[Circana] 🔄 Refreshing Nespresso pivot")
            pivot_sheet = wb.Sheets("Nespresso")
            pivot_table = pivot_sheet.PivotTables("PivotTable1")
            pivot_table.RefreshTable()
            excel.CalculateUntilAsyncQueriesDone()
            force_console_output("[Circana] ✅ Nespresso pivot refreshed")
        except Exception as e:
            force_console_output(f"[Circana] ❌ Nespresso pivot failed: {e}")

        # === 4. Finalize ===
        try:
            excel.Calculation = -4105
            excel.Calculate()
            force_console_output("[Circana] ✅ Calculations completed")
        except Exception as e:
            try:
                wb.Calculate()
                force_console_output("[Circana] ✅ Manual calculation completed")
            except:
                force_console_output(f"[Circana] ❌ Manual calculation failed: {e}")

        wb.Save()
        wb.Close()
        del wb
        gc.collect()
        force_console_output("[Circana] ✅ Workbook saved and closed")
        return True

    except Exception as e:
        force_console_output(f"[Circana] ❌ Critical error: {e}")
        return False

    finally:
        pythoncom.CoUninitialize()
        force_console_output("[Circana] ✅ COM uninitialized")


# === Main entry ===
if __name__ == "__main__":
    if len(sys.argv) < 2:
        force_console_output("[Circana] ❌ Missing file argument")
        sys.exit(1)

    file_path = sys.argv[1]
    prevent_sleep()
    start_time = time.time()
    success = automate_excel_process(file_path)
    allow_sleep()
    elapsed = int(time.time() - start_time)
    mins, secs = divmod(elapsed, 60)
    force_console_output(f"⏱️ Total time: {mins}m {secs}s")
    force_console_output("🎉 SUCCESS!" if success else "❌ FAILED")
    sys.exit(0 if success else 1)
