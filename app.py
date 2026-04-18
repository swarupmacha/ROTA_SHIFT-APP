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

        st.subheader("Preview")
        st.dataframe(week_df)

        # ==============================
        # CREATE CLEAN PIPE TABLE
        # ==============================
        name_width = 25
        col_width = 4

        lines = []

        # Header
        header = "Name".ljust(name_width) + " |"
        for d in week_dates:
            header += f" {str(d).rjust(2)} |"
        lines.append(header)

        # Separator
        lines.append("-" * len(header))

        # Rows
        for _, row in week_df.iterrows():
            line = str(row["Name"]).ljust(name_width) + " |"

            for d in week_dates:
                val = row[str(d)]
                if str(val) == "nan":
                    val = ""
                line += f" {str(val).rjust(2)} |"

            lines.append(line)

        table_text = "\n".join(lines)

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
        # OPEN DRAFT BUTTON
        # ==============================
        st.markdown(
            f'<a href="{outlook_url}" target="_blank">'
            f'<button style="padding:10px 20px;font-size:16px;">📧 Open Draft in Outlook</button>'
            f'</a>',
            unsafe_allow_html=True
        )

        # ==============================
        # OPTIONAL VIEW
        # ==============================
        st.subheader("Email Preview (Copy if needed)")
        st.text_area("", email_body, height=300)

        st.success("✅ Click button → Draft mail opens with formatted table!")