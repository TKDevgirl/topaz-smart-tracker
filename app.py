import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime
from io import BytesIO
import pandas as pd

tracking_sheets = ["RFA", "RFI"]
takenaka_sheets = ["MAT_ICT", "DWG_ICT", "MTS_ICT"]

def base_doc_no(doc_no):
    doc_no = str(doc_no).strip()
    parts = doc_no.split("-")
    if parts[-1].isdigit() and len(parts[-1]) == 2:
        return "-".join(parts[:-1])
    return doc_no

def generate_report(tracking_file, takenaka_file):
    takenaka_wb = load_workbook(takenaka_file, data_only=True)
    takenaka_map = {}

    for sheet in takenaka_sheets:
        if sheet not in takenaka_wb.sheetnames:
            continue

        ws = takenaka_wb[sheet]
        for row in range(11, ws.max_row + 1):
            doc_no = ws[f"E{row}"].value
            if not doc_no or "DETH-NSC" not in str(doc_no):
                continue

            takenaka_map[base_doc_no(doc_no)] = {
                "sheet": sheet,
                "doc_no": doc_no,
                "status1": ws[f"AA{row}"].value,
                "status2": ws[f"AB{row}"].value,
                "status3": ws[f"AC{row}"].value,
            }

    out_wb = load_workbook(tracking_file)
    report_sheet = "Open_On_Process_Compare"

    if report_sheet in out_wb.sheetnames:
        del out_wb[report_sheet]

    report_ws = out_wb.create_sheet(report_sheet)

    headers = [
        "Tracking Sheet", "Document No", "Document Name", "Tracking Status",
        "Takenaka Sheet", "Takenaka Doc No",
        "Takenaka Status 1", "Takenaka Status 2", "Takenaka Status 3",
        "Action", "Checked Time"
    ]
    report_ws.append(headers)

    for cell in report_ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

    green = PatternFill(fill_type="solid", fgColor="C6EFCE")
    yellow = PatternFill(fill_type="solid", fgColor="FFEB9C")
    red = PatternFill(fill_type="solid", fgColor="FFC7CE")
    blue = PatternFill(fill_type="solid", fgColor="BDD7EE")

    rows = []
    total_docs = 0
    open_count = 0
    action_count = {}

    for sheet in tracking_sheets:
        if sheet not in out_wb.sheetnames:
            continue

        ws = out_wb[sheet]
        for row in range(2, ws.max_row + 1):
            doc_no = ws[f"B{row}"].value
            doc_name = ws[f"D{row}"].value
            status = ws[f"F{row}"].value

            if not doc_no:
                continue

            total_docs += 1

            if str(status or "").strip().upper() != "OPEN":
                continue

            open_count += 1
            key = base_doc_no(doc_no)
            checked_time = datetime.now().strftime("%d-%b-%Y %H:%M:%S")

            if key in takenaka_map:
                src = takenaka_map[key]
                s1 = str(src["status1"] or "").strip().upper()
                s2 = str(src["status2"] or "").strip().upper()
                s3 = str(src["status3"] or "").strip().upper()

                if s1 == "CLOSED":
                    action = "UPDATE TRACKING TO CLOSED"
                    fill = green
                elif s1 == "RETURNED":
                    action = "RETURNED BY NV5 / NEED RESUBMIT"
                    fill = red
                elif s3 == "OVERDUE":
                    action = "OVERDUE / FOLLOW UP"
                    fill = red
                elif s1 == "OPEN" and s2 == "ON PROCESS":
                    action = "OPEN & ON PROCESS"
                    fill = yellow
                elif s1 == "OPEN":
                    action = "OPEN"
                    fill = blue
                else:
                    action = "CHECK"
                    fill = blue

                new_row = [
                    sheet, doc_no, doc_name, status,
                    src["sheet"], src["doc_no"],
                    src["status1"], src["status2"], src["status3"],
                    action, checked_time
                ]
            else:
                action = "NOT FOUND IN TAKENAKA SOURCE"
                fill = red
                new_row = [
                    sheet, doc_no, doc_name, status,
                    "", "", "", "", "",
                    action, checked_time
                ]

            report_ws.append(new_row)
            report_ws[f"J{report_ws.max_row}"].fill = fill

            rows.append({
                "Tracking Sheet": new_row[0],
                "Document No": new_row[1],
                "Document Name": new_row[2],
                "Tracking Status": new_row[3],
                "Takenaka Status 1": new_row[6],
                "Takenaka Status 2": new_row[7],
                "Takenaka Status 3": new_row[8],
                "Action": new_row[9],
            })

            action_count[action] = action_count.get(action, 0) + 1

    for col in report_ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        report_ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

    output = BytesIO()
    out_wb.save(output)
    output.seek(0)

    return output, total_docs, open_count, action_count, rows

st.set_page_config(
    page_title="Topaz Smart Document Tracker",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Topaz Smart Document Tracker")
st.caption("Compare OPEN documents from Tracking Document with Takenaka Status.")

tracking_file = st.file_uploader("1) Upload Tracking_document.xlsx", type=["xlsx"])
takenaka_file = st.file_uploader("2) Upload Takenaka Summary.xlsx", type=["xlsx"])

if tracking_file and takenaka_file:
    if st.button("Generate Report", type="primary"):
        report, total_docs, open_count, action_count, rows = generate_report(
            tracking_file, takenaka_file
        )

        st.success("Report generated successfully ka ✅")

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Documents", total_docs)
        c2.metric("Open in Tracking", open_count)
        c3.metric("Open & On Process", action_count.get("OPEN & ON PROCESS", 0))
        c4.metric("Need Update", action_count.get("UPDATE TRACKING TO CLOSED", 0))
        c5.metric("Overdue / Follow Up", action_count.get("OVERDUE / FOLLOW UP", 0))

        st.subheader("📊 Action Summary")
        summary_df = pd.DataFrame(
            [{"Action": k, "Count": v} for k, v in action_count.items()]
        )
        st.dataframe(summary_df, use_container_width=True)

        st.subheader("📋 Action List")
        result_df = pd.DataFrame(rows)

        selected_action = st.selectbox(
            "Filter by Action",
            ["All"] + sorted(result_df["Action"].unique().tolist())
        )

        if selected_action != "All":
            result_df = result_df[result_df["Action"] == selected_action]

        search = st.text_input("Search Document No / Document Name")
        if search:
            result_df = result_df[
                result_df.astype(str).apply(
                    lambda x: x.str.contains(search, case=False, na=False)
                ).any(axis=1)
            ]

        st.dataframe(result_df, use_container_width=True)

        st.download_button(
            label="⬇️ Download Excel Report",
            data=report,
            file_name="Open_On_Process_Compare.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Please upload both Excel files to generate the report.")
