import gc
import win32com.client as win32
import win32process
import os
import pythoncom
from datetime import datetime
import time
import ctypes
import io
import sys

# === UTF-8 output for Electron logs ===
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


# === Helper functions ===
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002  # Optional: keep screen awake too

npdPID = None # Global to track Excel PID, to kill if needed

def prevent_sleep():
    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
    )

def allow_sleep():
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)


def force_console_output(message):
    """Print and flush immediately (so Electron sees it)."""
    try:
        print(message)
    except UnicodeEncodeError:
        print(message.encode("utf-8", errors="replace").decode("utf-8"))
    sys.stdout.flush()
    sys.stderr.flush()
    
def start_excel():
    """Start Excel and track PID for external kill."""
    global npdPID
    excel_app = win32.DispatchEx("Excel.Application")
    hwnd = excel_app.Hwnd
    _, npdPID = win32process.GetWindowThreadProcessId(hwnd)
    force_console_output(f"✅ Excel started with PID:{npdPID}")
    return excel_app

# =========================================
# Main function to automate Excel process |
# =========================================

def automate_excel_process(file_path):
    wb = None
    excel = None
    global npdPID
    
    try:
        pythoncom.CoInitialize()
        excel = start_excel()

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"[NPD] File not found: {file_path}")

        wb = excel.Workbooks.Open(file_path, UpdateLinks=0, ReadOnly=False)
        force_console_output(f"[NPD] Opening workbook: {file_path}")
        
         # --- Refresh pivots safely ---
        force_console_output("\n\n[NPD] == 1. Refreshing Pivot Tables ==\n")
        try:
            sheet = wb.Sheets("Actuals")
            pivot = sheet.PivotTables("PivotTable1")
            pivot.PivotCache().Refresh()
            excel.CalculateUntilAsyncQueriesDone()
            force_console_output("[NPD] ✅ Pivot table refreshed successfully.")
        except Exception as e:
            force_console_output(f"[NPD] ⚠️ Pivot refreshing failed: {e}")

        # --- Update date filters ---
        try:
            force_console_output("\n[NPD] == 2. Updating Date Filters ==")
            pivot_field = pivot.PivotFields("[Table1].[Weeks].[Weeks]")

            # select all items available
            items = [item.Name for item in pivot_field.PivotItems()]
            force_console_output("[NPD] 📅 Dates Found")
            # filter and Keep all non-blank values
            non_blank_items = [
                d for d in items 
                if d and not d.strip().endswith(".&") and d.strip() != ""
            ]
            if not non_blank_items:
                raise Exception("No non-blank dates found.")
            
            # display the selected (filter items) only non-blank values
            pivot_field.VisibleItemsList = non_blank_items

            force_console_output("[NPD] ✅ Dates updated successfully.")
        except Exception as e:
            force_console_output(f"[NPD] ❌ Date filter update failed: {e}")

        # --- Refresh Power BI pivots ---
        force_console_output("\n[NPD] == 3. Refresh Power BI Pivot Table ==")
        try:
            sheet = wb.Sheets("For Power BI") # select sheet For Power BI
            pivot = sheet.PivotTables("PivotTable2")
            pivot.PivotCache().Refresh()
            pivot.RefreshTable()
            force_console_output("[NPD] ✅ Power BI Table refreshed.")

        except Exception as e:
            force_console_output(f"[NPD] ⚠️ Power BI Pivots failed: {e}")

        force_console_output("\n[NPD] == Finalizing & Saving ==")
        
        # --- Final save and close ---

        wb.Save()
        wb.Close()
        del wb
        gc.collect()
        force_console_output("[NPD] ✅ Workbook saved and closed")
        return True

    except Exception as e:
        force_console_output(f"[NPD] ❌ Critical error: {e}")
        return False

    finally:
        pythoncom.CoUninitialize()
        force_console_output("[NPD] ✅ COM uninitialized")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        force_console_output("[NPD] ❌ Missing file argument")
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
        
        