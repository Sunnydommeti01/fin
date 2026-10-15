
import streamlit as st
import pandas as pd

st.set_page_config(page_title="FIN Dashboard", layout="wide")

@st.cache_data(ttl=60)
def load_data():
    return pd.read_excel("Rengy_Streamlit_Sample_Data.xlsx")

df = load_data()

st.title("FIN - Excel Dashboard")

col1, col2, col3 = st.columns(3)

col1.metric("Total Projects", len(df))
col2.metric("Total Project Value", f"₹{df['Project Value'].sum():,.0f}")
col3.metric("Pending Amount", f"₹{df['Pending Amount'].sum():,.0f}")

st.subheader("Excel Data")
st.dataframe(df, use_container_width=True)
