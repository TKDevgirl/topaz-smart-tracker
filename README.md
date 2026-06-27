# Topaz Smart Document Tracker

## How to run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## How to deploy on Streamlit Community Cloud
1. Create a GitHub repository.
2. Upload `app.py` and `requirements.txt`.
3. Go to Streamlit Community Cloud.
4. Click New app.
5. Select the repository and `app.py`.
6. Click Deploy.

## User flow
1. Upload `Tracking_document.xlsx`.
2. Upload Takenaka Summary Excel.
3. Click Generate Report.
4. Download `Open_On_Process_Compare.xlsx`.