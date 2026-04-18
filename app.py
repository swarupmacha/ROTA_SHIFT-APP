import streamlit as st
import pandas as pd
import win32com.client as win32

st.set_page_config(page_title="ROTA Generator", layout="wide")
st.title("📊 ROTA Email Generator (Outlook Automation)")

# Upload file
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:

    xls = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("Select Sheet", xls.sheet_names)

    col1, col2 = st.columns(2)

    with col1:
        start_day = st.number_input("Start Day", 1, 31, 19)
    with col2:
        end_day = st.number_input("End Day", 1, 31, 23)

    if st.button("Generate & Open Outlook"):

        # ==============================
        # READ + CLEAN DATA
        # ==============================
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=1)

        df = df[df["Name"].notna()]
        df = df[df["Name"] != "D&T"]

        valid_cols = [col for col in df.columns if col == "Name" or str(col).isdigit()]
        df = df[valid_cols]

        week_dates = list(range(start_day, end_day + 1))
        cols = ["Name"] + [col for col in df.columns if str(col) in map(str, week_dates)]

        week_df = df[cols]
        week_df = week_df.dropna(how='all', subset=cols[1:])
        week_df = week_df.fillna("")
        week_df.columns = ["Name"] + [str(d) for d in week_dates]
        week_df = week_df.astype(str).reset_index(drop=True)

        # ==============================
        # PREVIEW
        # ==============================
        st.subheader("Preview")
        st.dataframe(week_df)

        # ==============================
        # CREATE CLEAN HTML TABLE
        # ==============================
        html_table = "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>"

        # Header
        html_table += "<tr style='background-color:#2E7D32; color:white;'>"
        html_table += "<th>Name</th>"
        for d in week_dates:
            html_table += f"<th>{d}</th>"
        html_table += "</tr>"

        # Rows
        for _, row in week_df.iterrows():
            html_table += "<tr>"
            html_table += f"<td>{row['Name']}</td>"

            for d in week_dates:
                val = row[str(d)]

                if val == "N":
                    color = "#00BFFF"
                elif val == "1":
                    color = "#FFA500"
                elif val == "2":
                    color = "#FFD700"
                else:
                    color = ""

                if color:
                    html_table += f"<td style='background-color:{color}; text-align:center;'>{val}</td>"
                else:
                    html_table += f"<td style='text-align:center;'>{val}</td>"

            html_table += "</tr>"

        html_table += "</table>"

        # ==============================
        # EMAIL BODY
        # ==============================
        email_body = f"""
        <html>
        <body>

        <p>Hi All,</p>

        <p>Please find below your shifts for upcoming week.</p>

        {html_table}

        <p>Thanks & Regards,<br>
        Your Name</p>

        </body>
        </html>
        """

        # ==============================
        # EXTRACT EMAILS
        # ==============================
        names = week_df["Name"].dropna().unique()

        # 👉 change domain if needed
        email_list = [name.strip() + "@accenture.com" for name in names]
        to_emails = ";".join(email_list)

        # ==============================
        # OPEN OUTLOOK
        # ==============================
        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            mail.To = to_emails
            mail.Subject = "24x7 Monitoring Shifts - Reminder"
            mail.HTMLBody = email_body

            mail.Display()

            st.success("✅ Outlook opened with formatted table!")

        except Exception as e:
            st.error(f"Error opening Outlook: {e}")