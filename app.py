import streamlit as st
import pandas as pd
import urllib.parse

st.set_page_config(page_title="ROTA Generator", layout="wide")
st.title("📊 ROTA Email Generator")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:

    xls = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("Select Sheet", xls.sheet_names)

    col1, col2 = st.columns(2)

    with col1:
        start_day = st.number_input("Start Day", 1, 31, 19)
    with col2:
        end_day = st.number_input("End Day", 1, 31, 23)

    if st.button("Generate Draft Mail"):

        # ==============================
        # READ DATA
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

        st.subheader("Preview")
        st.dataframe(week_df)

        # ==============================
        # CREATE TAB TABLE (BEST FOR BODY)
        # ==============================
        table_lines = []
        header = ["Name"] + [str(d) for d in week_dates]
        table_lines.append("\t".join(header))

        for _, row in week_df.iterrows():
            row_data = [row["Name"]]
            for d in week_dates:
                row_data.append(row[str(d)])
            table_lines.append("\t".join(row_data))

        table_text = "\n".join(table_lines)

        # ==============================
        # EMAIL BODY
        # ==============================
        email_body = f"""Hi All,

Please find below your shifts for upcoming week.

{table_text}

Thanks & Regards,
Your Name
"""

        # ==============================
        # EMAIL IDS
        # ==============================
        names = week_df["Name"].dropna().unique()
        emails = [name.strip() + "@accenture.com" for name in names]
        to_emails = ";".join(emails)

        # ==============================
        # CREATE OUTLOOK WEB LINK
        # ==============================
        subject = "24x7 Monitoring Shifts - Reminder"

        base_url = "https://outlook.office.com/mail/deeplink/compose"

        params = {
            "to": to_emails,
            "subject": subject,
            "body": email_body
        }

        query_string = urllib.parse.urlencode(params)

        outlook_url = f"{base_url}?{query_string}"

        # ==============================
        # BUTTON TO OPEN DRAFT
        # ==============================
        st.markdown(
            f'<a href="{outlook_url}" target="_blank">'
            f'<button style="padding:10px 20px;font-size:16px;">📧 Open Draft in Outlook</button>'
            f'</a>',
            unsafe_allow_html=True
        )

        st.success("✅ Click button → Draft mail opens in Outlook Web!")