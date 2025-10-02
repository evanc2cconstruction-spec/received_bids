import pandas as pd
from datetime import datetime
import streamlit as st
import openpyxl
import glob
import os

# excel_file = pd.ExcelFile('public_bids_received_2025-09-19.xlsx')
files = glob.glob("public_bids_received_*.xlsx")

if files:
	latest_file = max(files)
	excel_file = pd.ExcelFile(latest_file)

	sheet_names = excel_file.sheet_names

	for sheet_name in sheet_names:
		df = pd.read_excel(excel_file, sheet_name=sheet_name)

		st.title(sheet_name)
		st.dataframe(df)

else:
	st.error("No report found")


# It turns out, I won't need to use openpyxl
# Basically, I just need to get every single sheet into a pandas df
# Then, I can get all of these posted