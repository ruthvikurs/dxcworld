import pandas as pd

# Load the Excel file
input_file = "Kafka-Cluster-Auth-Details.-test.xlsx"
excel = pd.ExcelFile(input_file)

# Create a new ExcelWriter to save pivot tables
with pd.ExcelWriter("Kafka-Cluster-Auth-Details-PivotTables.xlsx", engine="openpyxl") as writer:
    for sheet_name in excel.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        if not df.empty:
            # Pivot on the first column as index and count all other columns
            first_col = df.columns[0]
            pivot = pd.pivot_table(df, index=[first_col], aggfunc='count')
            pivot.to_excel(writer, sheet_name=f"{sheet_name}_Pivot")