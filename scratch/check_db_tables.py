import sqlite3

try:
    conn = sqlite3.connect("roaders_plus.db")
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    print("Tables in roaders_plus.db:", tables)
    for table in tables:
        t_name = table[0]
        cursor.execute(f"PRAGMA table_info({t_name});")
        columns = cursor.fetchall()
        print(f"\nTable {t_name} columns:")
        for col in columns:
            print(f"  {col[1]} ({col[2]})")
        
        # Select first few rows
        cursor.execute(f"SELECT * FROM {t_name} LIMIT 3;")
        rows = cursor.fetchall()
        print(f"Table {t_name} preview (first 3 rows):")
        for row in rows:
            print(f"  {row}")
except Exception as e:
    print("Error:", e)
