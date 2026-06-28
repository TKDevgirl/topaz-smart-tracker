# Excel column mapping.
# If the Excel format changes, edit this file instead of business logic.

TRACKING_COLUMNS = {
    "doc_no": "B",
    "category": "C",
    "doc_name": "D",
    "revision": "E",
    "status": "F",
    "info": "G",
}

TAKENAKA_COLUMNS = {
    "start_row": 11,
    "doc_no": "E",
    "status_1": "AA",
    "status_2": "AB",
    "status_3": "AC",
}

STANDARD_CATEGORIES = ["MAT", "MCR", "MTS", "CVI", "DWG"]
STANDARD_STATUSES = ["Open", "On Progress", "Approved"]
