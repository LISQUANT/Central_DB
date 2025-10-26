from numpy import e
from openpyxl import workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook.workbook import Workbook
import pandas as pd
import openpyxl
import os
from pandas.core.arrays import period
import yfinance as yf

INPUT_DIR = "Valuation_Models_Previous"
OUTPUT_FILE = "Valuation_Models_Results/Consolidated_Targets.csv"

data_records = []

for file in os.listdir(INPUT_DIR):
    if file.endswith(".xslx"):
        file_path = os.path.join(INPUT_DIR, file)
        workbook = load_workbook(filename=file_path, data_only=True)
        sheet = workbook.active  # may change if the data isnt in the first sheet

        try:
            # get data
            ticker = sheet["#cell"].value
            ticker_yf = yf.Ticker(ticker)
            value_amd = sheet["#cell"].value  # value at making date
            value_acd = ticker_yf.history(period="1d")  # value at current date
            rec = sheet["#cell"].value  # recommendation
            pct_chg = ((value_acd / value_amd) - 1) * 100 + "%"

            # export data
            data_records.append(
                {
                    "Ticker": ticker,
                    "Value at Making Date": value_amd,
                    "Value at current date": value_acd,
                    "Change in Price": pct_chg,
                    "Recommendation": rec,
                    "Source File": file,
                }
            )
        except Exception as e:
            print(f"Error reading file {file}: {e}")

df = pd.DataFrame(data_records)

df.to_csv(OUTPUT_FILE, index=False)

print(f"The transfer o files was a success in {OUTPUT_FILE}")
