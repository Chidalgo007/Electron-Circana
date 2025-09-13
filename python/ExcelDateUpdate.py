import os
import win32com.client as win32
import pythoncom
import sys
import io

# === UTF-8 output for Electron logs ===
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# === Helper functions ===
def force_console_output(message):
    """Print and flush immediately (so Electron sees it)."""
    try:
        print(message)
    except UnicodeEncodeError:
        print(message.encode("utf-8", errors="replace").decode("utf-8"))
    sys.stdout.flush()
    sys.stderr.flush()

# ========== Main Automation Logic ========== #
def run_excel_updates(destination_folder):
    pythoncom.CoInitialize()

    try:
        try:
            excel = win32.GetActiveObject("Excel.Application")
            force_console_output("Connected to running Excel instance")
        except Exception:
            excel = win32.Dispatch("Excel.Application")
            force_console_output("Created new Excel instance")

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False

        # === Calendar Sheet Update ===
        force_console_output("\n\n=== Updating Calendar Sheet ===")
        calendar_file_path = os.path.join(destination_folder, "Calendar.xlsx")
        if not os.path.exists(calendar_file_path):
            raise FileNotFoundError(f"❌  File not found: {calendar_file_path}")
        
        wb = excel.Workbooks.Open(calendar_file_path, UpdateLinks=0, ReadOnly=False)
        sheet = wb.Sheets("Sheet1")
        date_value = sheet.Range("J5").Value
        sheet.Range("J1").Value = date_value
        force_console_output(f"📅  Calendar.xlsx: J1 updated to {date_value}")
        wb.Save()
        wb.Close(SaveChanges=True)

        # === Week Sheet Update ===
        force_console_output("\n\n=== Updating Week Sheet ===")
        week_file_path = os.path.join(destination_folder, "Weeks.xlsx")
        
        if not os.path.exists(week_file_path):
            raise FileNotFoundError(f"❌  File not found: {week_file_path}")
        wb = excel.Workbooks.Open(week_file_path, UpdateLinks=0, ReadOnly=False)
        sheet = wb.Sheets("Sheet1")
        date_value = sheet.Range("P2").Value
        sheet.Range("A2").Value = date_value
        force_console_output(f"📅  Weeks.xlsx: A2 updated to {date_value}")
        wb.Save()
        wb.Close(SaveChanges=True)

    finally:
        excel.Quit()
        

if __name__ == "__main__":
    file_path = sys.argv[1]
    run_excel_updates(file_path)