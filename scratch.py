import pandas as pd

def add_employee(e_id, name, dept, pos, salary):
    try:
        df = pd.DataFrame()
        required_cols = ["employee_id", "name", "dept", "position", "salary"]
        if df is None or df.empty or not all(c in df.columns for c in required_cols):
            if df is None or df.empty:
                df = pd.DataFrame(columns=required_cols)
            else:
                for c in required_cols:
                    if c not in df.columns:
                        df[c] = ""
        
        if str(e_id) in df['employee_id'].astype(str).values:
            return "ID_EXISTS"
            
        new_emp = pd.DataFrame([{"employee_id": str(e_id), "name": name, "dept": dept, "position": pos, "salary": salary}])
        df = pd.concat([df, new_emp], ignore_index=True)
        return df
    except Exception as e:
        return str(e)

print(add_employee("E001", "John", "IT", "Dev", 50000))
