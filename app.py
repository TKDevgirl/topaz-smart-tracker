import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime
from io import BytesIO

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

                report_ws.append([
                    sheet, doc_no, doc_name, status,
                    src["sheet"], src["doc_no"],
                    src["status1"], src["status2"], src["status3"],
                    action, checked_time
                ])
            else:
                action = "NOT FOUND IN TAKENAKA SOURCE"
                fill = red
                report_ws.append([
                    sheet, doc_no, doc_name, status,
                    "", "", "", "", "",
                    action, checked_time
                ])

            report_ws[f"J{report_ws.max_row}"].fill = fill
            action_count[action] = action_count.get(action, 0) + 1

    summary_row = report_ws.max_row + 3
    report_ws[f"A{summary_row}"] = "SUMMARY"
    report_ws[f"A{summary_row}"].font = Font(bold=True)

    r = summary_row + 1
    report_ws[f"A{r}"] = "OPEN DOCUMENTS IN TRACKING"
    report_ws[f"B{r}"] = open_count

    for action, count in action_count.items():
        r += 1
        report_ws[f"A{r}"] = action
        report_ws[f"B{r}"] = count

    r += 1
    report_ws[f"A{r}"] = "GENERATED TIME"
    report_ws[f"B{r}"] = datetime.now().strftime("%d-%b-%Y %H:%M:%S")

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
    return output, open_count, action_count

st.set_page_config(page_title="Topaz Smart Document Tracker", page_icon="📄", layout="centered")

st.title("📄 Topaz Smart Document Tracker")
st.caption("Upload Tracking Document and Takenaka Summary, then download the comparison report.")

tracking_file = st.file_uploader("1) Upload Tracking_document.xlsx", type=["xlsx"])
takenaka_file = st.file_uploader("2) Upload Takenaka Summary.xlsx", type=["xlsx"])

if tracking_file and takenaka_file:
    if st.button("Generate Report", type="primary"):
        report, open_count, action_count = generate_report(tracking_file, takenaka_file)

        st.success(f"Report generated. Open documents checked: {open_count}")

        col1, col2 = st.columns(2)
        col1.metric("Open documents", open_count)
        col2.metric("Actions found", sum(action_count.values()))

        st.write("### Summary")
        st.dataframe(
            [{"Action": k, "Count": v} for k, v in action_count.items()],
            use_container_width=True
        )

        st.download_button(
            label="⬇️ Download Excel Report",
            data=report,
            file_name="Open_On_Process_Compare.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload both Excel files to generate the report.")