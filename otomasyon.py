import pandas as pd

# Load both Excel files
first_file = pd.ExcelFile('test.xlsx')
second_file = pd.read_excel('control.xlsx')

# Load the relevant sheet from the first file, starting from row 11
fy2325_table = pd.read_excel(first_file, sheet_name='Baby Care', skiprows=10)

# Load the relevant columns from the second sheet
second_sheet = second_file

def convert_percentage_to_float(value):
    """Converts a percentage string to a float."""
    if isinstance(value, str):
        value = value.strip('%')
    try:
        return float(value)
    except ValueError as e:
        print(f"Error converting value to float: {e}")
        return None

def compare_values(index, distributor, measure, year, value):
    # Only allow 'KBD1' or 'KBD2', skip others
    if measure not in ['KBD1', 'KBD2']:
        return

    # Match distributor in FY2325 JAS KBD Structure row
    distributor_column = [col for col in fy2325_table.columns if distributor in col]
    if not distributor_column:
        return

    # Match measure in the rows of the FY2325 table
    measure_row = fy2325_table[fy2325_table.iloc[:, 0].str.contains(measure, na=False)]
    if measure_row.empty:
        return

    # Extract the scalar value from the measure_row for the distributor
    fy2325_value = measure_row[distributor_column[0]].values
    if len(fy2325_value) > 0:
        fy2325_value = fy2325_value[0]
    else:
        return

    # Convert values to float
    fy2325_value = convert_percentage_to_float(fy2325_value)
    value = convert_percentage_to_float(value)

    if fy2325_value is None or value is None:
        return

    # Compare the values
    if abs(fy2325_value - value) >= 1e-6:  # Allowing for minor floating-point precision issues
        print(f"Mismatch at row index {index + 2}:")
        print(f"FY2325 Row Data: {measure_row.iloc[0].to_dict()}")
        print(f"File Row Data: {row.to_dict()}")
        print(f"Distributor = {distributor}, Measure = {measure}, FY2325 Value = {fy2325_value}%, File Value = {value}%")

# Iterate over the rows in the second sheet
for index, row in second_sheet.iterrows():
    distributor = row['Distributor']
    measure = row['Measure']
    year = row['Year']
    value = row['Value']

    if year == 2024:  # Assuming we're only interested in 2024 data
        compare_values(index, distributor, measure, year, value)
