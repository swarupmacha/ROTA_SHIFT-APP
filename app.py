import streamlit as st
import pandas as pd
import urllib.parse

st.set_page_config(page_title="ROTA Generator", layout="wide")
st.title("📊 ROTA Email Generator")

# ==============================
# DATE SUFFIX FUNCTION
# ==============================
def get_day_suffix(day):
    if 11 <= day <= 13:
        return f"{day}th"
    last = day % 10
    if last == 1:
        return f"{day}st"
    elif last == 2:
        return f"{day}nd"
    elif last == 3:
        return f"{day}rd"
    else:
        return f"{day}th"

# ==============================
# MONTH MAP
# ==============================
month_map = {
    "Jan": "Jan", "Feb": "Feb", "Mar": "Mar", "Apr": "Apr",
    "May": "May", "Jun": "Jun", "Jul": "Jul", "Aug": "Aug",
    "Sep": "Sep", "Oct": "Oct", "Nov": "Nov", "Dec": "Dec"
}

# ==============================
# FILE UPLOAD
# ==============================
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:

    xls = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("Select Sheet", xls.sheet_names)

    col1, col2 = st.columns(2)

    with col1:
        start_day = st.number_input("Start Day", 1, 31, 13)
    with col2:
        end_day = st.number_input("End Day", 1, 31, 19)

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
        # PREVIEW
        # ==============================
        st.subheader("Preview")
        st.dataframe(week_df)

        # ==============================
        # TABLE TEXT
        # ==============================
        table_text = week_df.to_string(index=False)

        # ==============================
        # DYNAMIC SUBJECT
        # ==============================
        start_str = get_day_suffix(start_day)
        end_str = get_day_suffix(end_day)

        month = month_map.get(sheet_name[:3], sheet_name[:3])

        subject = f"24x7 Monitoring Shifts - {start_str} {month} to {end_str} {month} 2026"

        # ==============================
        # EMAIL BODY
        # ==============================
        email_body = f"""
Hi All,

Please find below your shifts for upcoming week.

Note: You all will be doing complete 9 hours shifts 

Please ensure to come in shifts on time and follow the escalation matrix as below in case of any issues:
Respective Supervisors
Segment wise ROTA POCs:
Portfolio Leads

Segment wise ROTA POCs are:
Corporate Banking         --> Stubbs
Retail Banking           --> Remhana Dekaot
Tech & Data Capabilities --> KL Rahul
Insurance                --> Raisk Salam

Below are the MIM POCs for next week, in case of any MIM situation please reach out to them:
- Rajesh Chawal (Banking)
- Naleem Baking (Banking)
- Shiavam Shendhe (Insurance)

The shift timings are as follows:

1st Shift   5:30 am to 2:30 pm (IST)
2nd Shift   1:30 pm to 10:30 pm (IST)
Night Shift 9:30 pm to 6:30 am (IST)


{table_text}

Thanks & Regards,
Your Name
"""

        # ==============================
        # EMAIL IDS
        # ==============================
        names = week_df["Name"].dropna().unique()
        email_list = [name.strip() + "@accenture.com" for name in names]
        to_emails = ";".join(email_list)

        # ==============================
        # MAILTO LINK
        # ==============================
        encoded_subject = urllib.parse.quote(subject)
        encoded_body = urllib.parse.quote(email_body)
        encoded_to = urllib.parse.quote(to_emails)

        mailto_link = (
            f"mailto:{encoded_to}"
            f"?subject={encoded_subject}"
            f"&body={encoded_body}"
        )

        # ==============================
        # BUTTON
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