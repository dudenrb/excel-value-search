import pandas as pd

# ✅ Step 1: Load the Excel file
file_path = r"C:\Users\Hp\Downloads\bhaiya.xlsx"  # <-- Use raw string or double backslashes
df = pd.read_excel(file_path, sheet_name=0)

# ✅ Step 2: Function to find cells containing a substring
def find_cells_containing(df, substring):
    matches = df.stack().astype(str).str.contains(substring, case=False, na=False)
    match_locations = matches[matches].index.tolist()
    return [{"Row": row + 2, "Column": col} for row, col in match_locations]  # +2 to match Excel (header + 1-based)

# ✅ Step 3: Define the search term and run the function
search_term = "dis"
results = find_cells_containing(df, search_term)

# ✅ Step 4: Output the results
if results:
    for match in results:
        print(f"Row: {match['Row']}, Column: {match['Column']}")
else:
    print("Value not found.")
