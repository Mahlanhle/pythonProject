import pandas as pd
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title='Doppio Zero')
st.title('Excel Plotter')
st.subheader('Feed me with excel files')

uploaded_file = st.file_uploader('Choose an XLSX file ',type = 'xlsx')
if uploaded_file:
    st.markdown('---')
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    st.dataframe(df)
    groupby_column = st.selectbox(
        'What would you like to analyse?',
        ('Ship Mode', 'Segment', 'Category', 'Sub-Category'),
    )

    #---Dataframe
    output_columns = ['Sales', 'Profit']
    df_grouped = df.groupby(by=[groupby_column], as_index=False)[output_columns].sum()
    st.dataframe(df_grouped)






