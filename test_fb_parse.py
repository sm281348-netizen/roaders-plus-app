import pandas as pd
import re
import datetime
from app import conn

def test_parse_fb():
    try:
        df_raw = conn.read(worksheet="f&b_data", ttl=0)
        if df_raw is None or df_raw.empty:
            print("f&b_data is empty or not found.")
            return

        print(f"Loaded f&b_data shape: {df_raw.shape}")
        
        parsed_rows = []
        current_year = datetime.datetime.now().year # Default year

        # Variables to keep track of the most recent summary metrics
        current_month_rev = 0
        current_month_avg_spent = 0
        
        # When we hit a summary block, it applies to the month of the dates we just processed
        # Wait, if we read top-down, we process dates, then hit the summary block.
        # So we should assign the summary block values to all dates of that month!
        # Let's keep track of dates in the current month block
        current_month_dates = []

        for i, row in df_raw.iterrows():
            col_a = str(row.iloc[0]).strip()
            if not col_a or col_a == 'nan':
                continue
                
            # Check for summary rows
            if col_a == "已結算營收":
                val = str(row.iloc[1]).replace("NT$", "").replace(",", "").strip()
                current_month_rev = int(float(val)) if val and val != 'nan' else 0
                continue
            elif col_a == "平均客單價":
                val = str(row.iloc[1]).replace("NT$", "").replace(",", "").strip()
                current_month_avg_spent = int(float(val)) if val and val != 'nan' else 0
                
                # Apply the parsed monthly metrics to the current month's dates
                for r in current_month_dates:
                    r['rest_month_rev'] = current_month_rev
                    r['rest_avg_spent'] = current_month_avg_spent
                # Clear the block
                current_month_dates = []
                continue
                
            # Date row like "1/1週四"
            m_date = re.match(r'^(\d{1,2})/(\d{1,2})', col_a)
            if m_date:
                m, d = int(m_date.group(1)), int(m_date.group(2))
                date_str = f"{current_year}-{m:02d}-{d:02d}"
                
                # Parse daily columns (0-indexed)
                # G (index 6): Breakfast Actual Total
                rest_breakfast = pd.to_numeric(row.iloc[6], errors='coerce') if len(row) > 6 else 0
                # M (index 12): Afternoon Tea Actual Total
                rest_day_guests = pd.to_numeric(row.iloc[12], errors='coerce') if len(row) > 12 else 0
                # N (index 13): HH
                rest_hh_guests = pd.to_numeric(row.iloc[13], errors='coerce') if len(row) > 13 else 0
                
                row_dict = {
                    'date': date_str,
                    'rest_breakfast': rest_breakfast if pd.notna(rest_breakfast) else 0,
                    'rest_day_guests': rest_day_guests if pd.notna(rest_day_guests) else 0,
                    'rest_hh_guests': rest_hh_guests if pd.notna(rest_hh_guests) else 0,
                    'rest_month_rev': 0, # Will be filled later when we hit summary
                    'rest_avg_spent': 0  # Will be filled later when we hit summary
                }
                parsed_rows.append(row_dict)
                current_month_dates.append(row_dict)

        df_final = pd.DataFrame(parsed_rows)
        print("Parsed DataFrame:")
        print(df_final.head(10))
        print("...")
        print(df_final.tail(5))
        
    except Exception as e:
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_parse_fb()
