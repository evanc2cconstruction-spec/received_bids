import pandas as pd
from datetime import datetime
import streamlit as st
import openpyxl

sheet_name = 'Cottage Inn - Toledo, OH'

df = pd.read_excel('public_bids_received_2025-09-19.xlsx', sheet_name=sheet_name)

st.title(sheet_name)
st.dataframe(df)

# It turns out, I won't need to use openpyxl
# Basically, I just need to get every single sheet into a pandas df
# Then, I can get all of these posted