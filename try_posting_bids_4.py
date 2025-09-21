import pandas as pd
from datetime import datetime
import streamlit as st
import openpyxl

excel_file = 'public_bids_received_2025-09-19.xlsx'
sheet_names = excel_file.sheet_names

for sheet_name in sheet_names:
	df = pd.read_excel(excel_file, sheet_name=sheet_name)

	st.title(sheet_name)
	st.dataframe(df)


# It turns out, I won't need to use openpyxl
# Basically, I just need to get every single sheet into a pandas df
# Then, I can get all of these posted