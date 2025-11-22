import pandas as pd

# ------------------------------------------------------
# Helper Functions
# ------------------------------------------------------

def choose_sheet(file_path):
    xls = pd.ExcelFile(file_path)
    print("\nAvailable sheets:")
    for i, s in enumerate(xls.sheet_names, start=1):
        print(f"{i}. {s}")

    choice = int(input("Choose sheet number: "))
    return xls.sheet_names[choice - 1]


def show_columns_numbered(df):
    print("\nColumns:")
    col_map = {}
    for i, col in enumerate(df.columns, start=1):
        uniques = df[col].dropna().unique()[:5]
        print(f"{i}. {col} → {list(uniques)}")
        col_map[i] = col
    return col_map


# ------------------------------------------------------
# MAIN PROGRAM
# ------------------------------------------------------

file_path = input("Enter Excel file path: ")

# Choose which sheet to load
sheet_name = choose_sheet(file_path)
df = pd.read_excel(file_path, sheet_name=sheet_name)

print(f"\nLoaded sheet: {sheet_name}")
print(df.head())

# ------------------------------------------------------
# Main Menu
# ------------------------------------------------------
print("\nWhat will you do?")
print("1. Delete column")
print("2. Append data")
print("3. Filter")

action = input("Choose option (1/2/3): ").strip()

# ------------------------------------------------------
# 1. DELETE COLUMN
# ------------------------------------------------------
if action == "1":
    col_map = show_columns_numbered(df)

    col_num = int(input("\nType the column NUMBER to DELETE: "))
    col_to_delete = col_map.get(col_num)

    if not col_to_delete:
        print("Invalid column number!")
    else:
        df = df.drop(columns=[col_to_delete])
        base_name = file_path.split("\\")[-1].replace(".xlsx", "")
        export_name = f"deletedColumn-{base_name}.csv"
        df.to_csv(export_name, index=False)
        print(f"\nColumn '{col_to_delete}' deleted. Exported as {export_name}")


# ------------------------------------------------------
# 2. APPEND DATA
# ------------------------------------------------------
elif action == "2":
    other_path = input("Enter the OTHER Excel file path to append: ")
    other_sheet = choose_sheet(other_path)
    df2 = pd.read_excel(other_path, sheet_name=other_sheet)

    df_combined = pd.concat([df, df2], ignore_index=True)

    base_name = file_path.split("\\")[-1].replace(".xlsx", "")
    export_name = f"appended-{base_name}.csv"
    df_combined.to_csv(export_name, index=False)

    print(f"\nData appended. Exported as {export_name}")


# ------------------------------------------------------
# 3. FILTER DATA
# ------------------------------------------------------
elif action == "3":
    col_map = show_columns_numbered(df)

    col_num = int(input("\nSelect column NUMBER to filter: "))
    col_filter = col_map.get(col_num)

    if not col_filter:
        print("Invalid column number!")
    else:
        uniques = df[col_filter].dropna().unique()[:5]
        print(f"\nUnique values in '{col_filter}' (max 5 shown): {list(uniques)}")

        value = input(f"Type EXACT value to filter by: ")

        df_filtered = df[df[col_filter] == value]

        base_name = file_path.split("\\")[-1].replace(".xlsx", "")
        export_name = f"Filtered-{base_name}.csv"
        df_filtered.to_csv(export_name, index=False)

        print(f"\nData filtered. Exported as {export_name}")

else:
    print("Invalid choice.")
