import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta

# --- CONFIG ---
ticker = "^BSESN"
interval = "5m"
max_days_to_check = 7  # Check up to 7 days back for last trading day

# --- FUNCTION: Get last valid intraday data ---
def get_last_trading_day():
    for i in range(max_days_to_check):
        date_to_check = datetime.now() - timedelta(days=i)
        start_str = date_to_check.strftime('%Y-%m-%d')
        end_str = (date_to_check + timedelta(days=1)).strftime('%Y-%m-%d')

        df = yf.download(
            ticker,
            start=start_str,
            end=end_str,
            interval=interval,
            progress=False,
            auto_adjust=False
        )

        if not df.empty:
            return df, start_str
    return None, None

# --- MAIN EXECUTION ---
df, trading_date = get_last_trading_day()

if df is not None:
    df.reset_index(inplace=True)

    # ✅ Convert timezone from UTC to IST
    df['Datetime'] = df['Datetime'].dt.tz_convert('Asia/Kolkata')

    # ✅ Filter market hours (09:00 to 15:30 IST)
    df = df[
        (df['Datetime'].dt.time >= datetime.strptime("09:00", "%H:%M").time()) &
        (df['Datetime'].dt.time <= datetime.strptime("15:30", "%H:%M").time())
    ]

    # ✅ Add Time column and extract required columns
    df['Time'] = df['Datetime'].dt.strftime("%H:%M")
    final_df = df[['Time', 'Close', 'Open', 'High', 'Low']]

    # ✅ Flatten MultiIndex columns if present (to avoid Excel export error)
    if isinstance(final_df.columns, pd.MultiIndex):
        final_df.columns = ['_'.join(col).strip() if isinstance(col, tuple) else col for col in final_df.columns]

    # ✅ Save to Excel
    output_file = f"Sensex_Intraday_data_{trading_date}.xlsx"
    final_df.to_excel(output_file, index=False)
    print(f"✅ Clean intraday data saved to '{output_file}'")

else:
    print("⚠️ No intraday data found in the last few days. Market may have been closed.")
