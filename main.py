import streamlit as st
import pandas as pd
import io

def display_safe_dataframe(df, title=None):
    if title:
        st.subheader(title)
    df_safe = df.copy()
    if 'ParentID' in df_safe.columns:
        df_safe['ParentID'] = df_safe['ParentID'].fillna(0).astype(int).astype(str)
    st.dataframe(df_safe)

def upload_csv_files():
    st.title("📊 TRANSFORMER")
    st.write("Upload report and dump Excel files to get started.")

    file1 = st.file_uploader("Upload Report Excel File", type="xlsx", key="file1")
    file2 = st.file_uploader("Upload Dump Excel File", type="xlsx", key="file2")

    df1, df2 = None, None

    if file1:
        df1 = pd.read_excel(file1, sheet_name="Data")

    if file2:
        df2 = pd.read_excel(file2)

    return df1, df2

def clean_and_filter_data(df_report, df_dump):
    try:
        df_report["Created On"] = pd.to_datetime(df_report["Created On"], format="%m/%d/%Y %H:%M:%S")
        df_dump["Created On"] = pd.to_datetime(df_dump["Created On"], format="%m/%d/%Y %H:%M:%S")

        last_created_on = df_report["Created On"].max()

        if "ParentID" not in df_report.columns or "ParentID" not in df_dump.columns:
            st.error("❌ 'ParentID' column not found in one or both files.")
            return

        before_df = df_dump[df_dump["Created On"] <= last_created_on]
        after_df = df_dump[df_dump["Created On"] > last_created_on]

        report_ids = set(df_report["ParentID"].dropna().astype(str))
        before_valid = before_df.dropna(subset=["ParentID"]).copy()
        before_invalid = before_valid[~before_valid["ParentID"].astype(str).isin(report_ids)]

        before_cleaned = before_df.drop(index=before_invalid.index)
        df_cleaned = pd.concat([before_cleaned, after_df], ignore_index=True)

        removed_change_rows = pd.DataFrame()
        if "Record Type" in df_cleaned.columns and "Record Type" in df_report.columns:
            df_changes = df_cleaned[df_cleaned["Record Type"].str.lower() == "change"].copy()
            report_keys = set(
                zip(df_report["Record Type"].str.lower(), df_report["Created On"])
            )
            df_changes["key"] = list(zip(df_changes["Record Type"].str.lower(), df_changes["Created On"]))
            unmatched_changes = df_changes[~df_changes["key"].isin(report_keys)]
            removed_change_rows = unmatched_changes
            df_cleaned = df_cleaned.drop(index=unmatched_changes.index)
        else:
            unmatched_changes = pd.DataFrame()

        changed_status_df = pd.DataFrame()
        if "Status" in df_report.columns and "Status" in df_cleaned.columns:
            rep_status = df_report[["ParentID", "Status"]].dropna()
            dump_status = df_cleaned[["ParentID", "Status"]].dropna()
            rep_status["ParentID"] = rep_status["ParentID"].astype(str)
            dump_status["ParentID"] = dump_status["ParentID"].astype(str)
            rep_status["Status"] = rep_status["Status"].astype(str)
            dump_status["Status"] = dump_status["Status"].astype(str)

            merged = pd.merge(rep_status, dump_status, on="ParentID", how="inner", suffixes=("_Report", "_Dump"))
            changed_status = merged[merged["Status_Report"] != merged["Status_Dump"]]

            if not changed_status.empty:
                parent_ids_with_changes = changed_status["ParentID"].tolist()
                changed_status_df = df_cleaned[df_cleaned["ParentID"].astype(str).isin(parent_ids_with_changes)]

        # KPIs
        col1, col2, col3 = st.columns(3)
        col1.metric("Removed Unmatched ParentIDs", len(before_invalid))
        col2.metric("Removed Unmatched 'Change' Rows", len(removed_change_rows))
        col3.metric("Status Changes Detected", len(changed_status_df))

        if not changed_status_df.empty:
            display_safe_dataframe(changed_status_df, "🔍 Full Rows with Status Changes")
        st.info(f"🆕 New Entries in Dump after Report ends: {len(after_df)}")

        display_safe_dataframe(df_cleaned, "📄 Cleaned Dump DataFrame")

        return df_cleaned, changed_status_df, after_df

    except Exception as e:
        st.error(f"❌ Error during cleaning/filtering: {e}")

def download_csv(df, filename, label, sheet_name="Sheet1"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='mm/dd/yyyy hh:mm:ss') as writer:
        df_to_export = df.copy()
        df_to_export.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        date_format = workbook.add_format({'num_format': 'mm/dd/yyyy hh:mm:ss'})
        for idx, col in enumerate(df_to_export.columns):
            if pd.api.types.is_datetime64_any_dtype(df_to_export[col]):
                worksheet.set_column(idx, idx, 20, date_format)
            else:
                worksheet.set_column(idx, idx, 20)

    output.seek(0)
    st.download_button(
        label=label,
        data=output,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def main():
    st.set_page_config(page_title="CSV Transformer", layout="centered")
    df_report, df_dump = upload_csv_files()

    if st.button("🚀 Clean & Filter Dump Data"):
        if df_report is None or df_dump is None:
            st.warning("⚠️ Please upload both Report and Dump Excel files before proceeding.")
        else:
            result = clean_and_filter_data(df_report, df_dump)
            if result is not None:
                df_cleaned, df_status_changes, df_new_entries = result
                st.session_state["df_cleaned"] = df_cleaned
                st.session_state["df_status_changes"] = df_status_changes
                st.session_state["df_new_entries"] = df_new_entries

    if "df_cleaned" in st.session_state:
        col1, col2 = st.columns(2)
        with col1:
            download_csv(
                st.session_state["df_cleaned"],
                "cleaned_dump.xlsx",
                "📥 Download Cleaned Dump"
            )
        with col2:
            updates_dict = {
                "Status_Changes": st.session_state["df_status_changes"],
                "New_Entries": st.session_state["df_new_entries"]
            }
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='mm/dd/yyyy hh:mm:ss') as writer:
                for sheet_name, df in updates_dict.items():
                    df_copy = df.copy()
                    df_copy.to_excel(writer, index=False, sheet_name=sheet_name)
                    workbook = writer.book
                    worksheet = writer.sheets[sheet_name]

                    date_format = workbook.add_format({'num_format': 'mm/dd/yyyy hh:mm:ss'})
                    for idx, col in enumerate(df_copy.columns):
                        if pd.api.types.is_datetime64_any_dtype(df_copy[col]):
                            worksheet.set_column(idx, idx, 20, date_format)
                        else:
                            worksheet.set_column(idx, idx, 20)

            output.seek(0)
            st.download_button(
                label="📤 Download Updates (Changes + New)",
                data=output,
                file_name="dump_updates.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
