import streamlit as st
import pandas as pd

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

        def highlight_shift(val):
            if val == "N":
                return "background-color:#00BFFF; color:white;"
            elif val == "1":
                return "background-color:#FFA500; color:white;"
            elif val == "2":
                return "background-color:#FFD700; color:black;"
            return ""

        styled_df = (
    week_df.style
    .map(highlight_shift)
    .hide(axis="index")
    .set_properties(**{
        'text-align': 'center',
        'border': '1px solid black',
        'color': 'white'   # 🔥 THIS MAKES TEXT WHITE
    })
    .set_table_styles([
        {
            'selector': 'th',
            'props': [
                ('background-color', '#2E7D32'),
                ('color', 'white'),   # header text white
                ('font-weight', 'bold'),
                ('border', '1px solid black'),
                ('text-align', 'center')
            ]
        }
    ])
)

        html_table = styled_df.to_html()

        email_body = f"""
        <h4>Hi All,</h4>
        <p>Please find below your shifts for upcoming week.</p>
        {html_table}
        <br><br>
        Thanks & Regards
        """
        

        st.success("Email Generated!")

        st.components.v1.html(email_body, height=600, scrolling=True)

        st.code(email_body, language="html")