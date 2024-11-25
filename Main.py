import pandas as pd
from sqlalchemy import create_engine

# Replace with your actual database connection details
server = ""
database = ""
username = ""
password = ""
driver = "ODBC Driver 17 for SQL Server"  # Ensure the correct driver is installed

# Create a SQLAlchemy engine
engine = create_engine(f"mssql+pyodbc://{username}:{password}@{server}/{database}?driver={driver}")

# Path to your Excel file
excel_file_path = "1Nov2024.xlsx"

# Read all sheets from the Excel file
excel_sheets = pd.read_excel(excel_file_path, sheet_name=None)

for sheet_name, data in excel_sheets.items():
    # Sanitize table name
    table_name = sheet_name.lower().replace(" ", "_").replace("-", "_")

    # Write each sheet to the SQL table
    data.to_sql(table_name, con=engine, if_exists="replace", index=False)
    print(f"Sheet '{sheet_name}' has been written to table '{table_name}'.")

print("All sheets have been written to the SQL Server database.")
