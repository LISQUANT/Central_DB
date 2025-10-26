from openpyxl import load_workbook
import pandas as pd
import os
import yfinance as yf

# Cell mappings for each dir
CELL_MAPPINGS = {
    "Valuation_Models_Previous/23-24-1": {
        "ticker_cell": "A1",
        "date_cell": "B2",
        "rec_cell": "C1",
    },
    "Valuation_Models_Previous/23-24-2": {
        "ticker_cell": "A1",
        "date_cell": "B2",
        "rec_cell": "C1",
    },
    "Valuation_Models_Previous/24-25-1-4": {
        "ticker_cell": "A1",
        "date_cell": "B2",
        "rec_cell": "C1",
    },
    "Valuation_Models_Previous/24-25-5": {
        "ticker_cell": "A1",
        "date_cell": "B2",
        "rec_cell": "C1",
    },
    "Valuation_Models_Previous/25-26_1": {
        "ticker_cell": "A1",
        "date_cell": "B2",
        "rec_cell": "C1",
    },
}

OUTPUT_FILE = "Valuation_Models_Results/Consolidated_Targets.csv"

data_records = []

# Loop through each directory and its cell mapping
for input_dir, cell_map in CELL_MAPPINGS.items():
    print(f"Processing directory: {input_dir}")

    # Loop through each file in the directory
    for file in os.listdir(input_dir):
        if file.endswith(".xlsx") or file.endswith(".xls"):
            file_path = os.path.join(input_dir, file)
            try:
                workbook = load_workbook(filename=file_path, data_only=True)
                sheet = workbook.active

                ticker = sheet[cell_map["ticker_cell"]].value
                making_date = sheet[cell_map["date_cell"]].value
                rec = sheet[cell_map["rec_cell"]].value

                # Skip if ticker is None or empty
                if not ticker:
                    print(f"Skipping {file}: No ticker found")
                    continue

                if not making_date:
                    print(f"Skipping {file}: No date found")
                    continue

                # Get yfinance ticker object
                ticker_yf = yf.Ticker(ticker)
                # Get price at making date
                if isinstance(making_date, pd.Timestamp) or hasattr(
                    making_date, "strftime"
                ):
                    date_str = making_date.strftime("%Y-%m-%d")
                else:
                    date_str = str(making_date)

                history_making = ticker_yf.history(
                    start=date_str, end=pd.to_datetime(date_str) + pd.Timedelta(days=5)
                )

                if history_making.empty:
                    print(
                        f"Warning: No price data for {ticker} at date {date_str} in {file}"
                    )
                    value_amd = None
                else:
                    value_amd = history_making["Close"].iloc[0]

                history_current = ticker_yf.history(period="1d")

                if history_current.empty:
                    print(f"Warning: No current price data for {ticker} in {file}")
                    value_acd = None
                    pct_chg = None
                else:
                    value_acd = history_current["Close"].iloc[-1]
                    if value_amd:
                        pct_chg = ((value_acd / value_amd) - 1) * 100
                    else:
                        pct_chg = None

                # Export data
                data_records.append(
                    {
                        "Ticker": ticker,
                        "Making Date": making_date,
                        "Value at Making Date": value_amd,
                        "Value at Current Date": value_acd,
                        "Change in Price (%)": pct_chg,
                        "Recommendation": rec,
                        "Source Folder": input_dir,
                        "Source File": file,
                    }
                )
                print(f"Successfully processed: {file}")

            except Exception as e:
                print(f"Error reading file {file}: {e}")

df = pd.DataFrame(data_records)
df.to_csv(OUTPUT_FILE, index=False)

print(f"\nThe transfer of files was a success! Saved to {OUTPUT_FILE}")
print(f"Total records processed: {len(data_records)}")
