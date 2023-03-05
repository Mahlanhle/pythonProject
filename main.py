import streamlit as st
import pandas as pd
import plotly.express as px
import base64
from io import StringIO, BytesIO

def generate_excel_download_link(df):
    # Credit Excel: https://discuss.streamlit.io/t/how-to-add-a-download-excel-csv-function-to-a-button/4474/5
    towrite = BytesIO()
    df.to_excel(towrite, encoding="utf-8", index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data_download.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

st.set_page_config(page_title='Excel plotter')
st.title('Doppio Excel plotter')
st.subheader('Feed me with your Excel file')

uploaded_file = st.file_uploader('Choose an XLSX file ', type='xlsx')
if uploaded_file:
    st.markdown('---')
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    st.dataframe(df)
    groupby_column = st.selectbox(
        'what would you like to analyse ?',
        ('Ship Mode', 'Segment', 'Category', 'Sub-Category'),
     )

    # --Group Dataframe
    output_columns = ['Sales', 'Profit']
    df_grouped = df.groupby(by=[groupby_column], as_index=False)[output_columns].sum()

    #--plot Dataframe
    fig = px.bar(
        df_grouped,
        x=groupby_column,
        y='Sales',
        color='Profit',
        color_continuous_scale=['red', 'yellow', 'green'],
        template='plotly_white',
        title=f'<b>Scales and profit by {groupby_column}</b>'
    )
    st.plotly_chart(fig)
    st.subheader('Download:')
    generate_excel_download_link(df_grouped)




