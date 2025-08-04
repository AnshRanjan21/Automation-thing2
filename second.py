import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def select_file(title="Select a file"):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx")])
    return file_path

def update_report(report_path, dump_path):
    try:
        # Load data
        df_report = pd.read_excel(report_path, sheet_name="Data")
        df_dump = pd.read_excel(dump_path)

        # Parse dates
        df_report["Created On"] = pd.to_datetime(df_report["Created On"])
        df_dump["Created On"] = pd.to_datetime(df_dump["Created On"])

        # Normalize keys
        df_report["ParentID"] = df_report["ParentID"].fillna(0).astype(int).astype(str)
        df_dump["ParentID"] = df_dump["ParentID"].fillna(0).astype(int).astype(str)

        # 1. Update statuses
        if "Status" in df_report.columns and "Status" in df_dump.columns:
            dump_status_map = df_dump.set_index("ParentID")["Status"].astype(str).to_dict()
            df_report["Status"] = df_report.apply(
                lambda row: dump_status_map[row["ParentID"]]
                if row["ParentID"] in dump_status_map and dump_status_map[row["ParentID"]] != str(row["Status"])
                else row["Status"],
                axis=1
            )

        # 2. Append new rows
        last_created_on = df_report["Created On"].max()
        new_rows = df_dump[df_dump["Created On"] > last_created_on]
        combined = pd.concat([df_report, new_rows], ignore_index=True)

        # 3. Overwrite original file
        with pd.ExcelWriter(report_path, engine="openpyxl", datetime_format='mm/dd/yyyy hh:mm:ss') as writer:
            combined.to_excel(writer, index=False, sheet_name="Data")

        messagebox.showinfo("Success", f"✅ Report updated successfully!\n\nNew rows added: {len(new_rows)}")
    
    except Exception as e:
        messagebox.showerror("Error", f"❌ Failed to update report:\n{e}")

if __name__ == "__main__":
    print("📄 Select Original Report Excel File")
    report_file = select_file("Select the original Report Excel file (with 'Data' sheet)")

    if not report_file:
        print("Report file not selected. Exiting.")
        exit()

    print("📄 Select New Dump Excel File")
    dump_file = select_file("Select the new Dump Excel file")

    if not dump_file:
        print("Dump file not selected. Exiting.")
        exit()

    print(f"Processing...\nReport: {report_file}\nDump: {dump_file}")
    update_report(report_file, dump_file)
