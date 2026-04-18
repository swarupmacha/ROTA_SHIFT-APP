import streamlit as st
import pandas as pd
import urllib.parse

st.set_page_config(page_title="ROTA Generator", layout="wide")
st.title("📊 ROTA Email Generator")

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

    if st.button("Generate Email"):

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
        # PREVIEW IN APP
        # ==============================
        st.subheader("Preview")
        st.dataframe(week_df)

        # ==============================
        # CLEAN ALIGNED TABLE TEXT
        # ==============================
        name_width = 25
        col_width = 5

        # Header
        header = f"{'Name'.ljust(name_width)}"
        for d in week_dates:
            header += str(d).rjust(col_width)

        lines = [header]

        # Rows
        for _, row in week_df.iterrows():
            line = row["Name"].ljust(name_width)
            for d in week_dates:
                val = row[str(d)]
                line += str(val).rjust(col_width)
            lines.append(line)

        table_text = "\n".join(lines)

        # ==============================
        # EMAIL BODY
        # ==============================
        email_body = f"""
Hi All,

Please find below your shifts for upcoming week.

{table_text}

Thanks & Regards,
Your Name
"""

        # ==============================
        # EXTRACT EMAILS
        # ==============================
        names = week_df["Name"].dropna().unique()

        # 👉 Change domain if needed
        email_list = [name.strip() + "@gmail.com" for name in names]

        # Outlook uses ;
        to_emails = ";".join(email_list)

        # ==============================
        # CREATE MAILTO LINK
        # ==============================
        subject = "24x7 Monitoring Shifts - Reminder"

        encoded_subject = urllib.parse.quote(subject)
        encoded_body = urllib.parse.quote(email_body)
        encoded_to = urllib.parse.quote(to_emails)

        mailto_link = (
            f"mailto:{encoded_to}"
            f"?subject={encoded_subject}"
            f"&body={encoded_body}"
        )

        # ==============================
        # OUTLOOK BUTTON
        # ==============================
        st.markdown(
            f'<a href="{mailto_link}">'
            f'<button style="padding:10px 20px;font-size:16px;">📧 Open Outlook</button>'
            f'</a>',
            unsafe_allow_html=True
        )

        # ==============================
        # OPTIONAL VIEW
        # ==============================
        show_body = st.checkbox("Show Email Body")

        if show_body:
            st.text_area("Email Content", email_body, height=300)

        st.success("✅ Email Generated Successfully!")