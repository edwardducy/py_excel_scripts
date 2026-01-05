import pandas as pd
import xlwings as xw
import re
import time

start_time = time.time()

file_path = input('Enter full file path: ')
sheet_name = input('Enter sheet name: ')
start_row = int(input('Enter starting row number: '))
last_row = int(input('Enter last row number: '))
start_column = input('Enter starting column: ')
last_column = input('Enter last column: ')
target_column = input('Enter target column: ')

def convert_alphabet_to_number(column):
    column = column.upper()
    num = 0
    for char in column:
        num = num * 26 + ord(char) - 64
    return num

start_col_idx = convert_alphabet_to_number(start_column)
last_col_idx = convert_alphabet_to_number(last_column)
num_columns = last_col_idx - start_col_idx + 1

if not target_column.isdigit():
    target_column = convert_alphabet_to_number(target_column)
else:
    target_column = int(target_column)

def initialize_excel_app():
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = 'manual'
    return app

def get_workbook(app, file_path):
    return app.books.open(file_path)

def get_worksheet(wb, sheet_name):
    return wb.sheets[sheet_name]

def get_data_range(ws, start_column, start_row, last_column, last_row):
    return ws.range(f"{start_column}{start_row}:{last_column}{last_row}").value

def convert_to_dataframe(data):
    return pd.DataFrame(data).dropna(how="all")

def get_filtered_dataframe(df, target_column):
    return df[df.iloc[:, target_column-1].astype(str).str.rstrip().str.match(r".*\d$")]

def clear_contents(ws, start_column, start_row, last_column, last_row):
    ws.range(f"{start_column}{start_row}:{last_column}{last_row}").clear_contents()

def set_range_values(ws, start_column, start_row, filtered_df, num_columns):
    ws.range(f"{start_column}{start_row}").value = filtered_df.iloc[:, :num_columns].values

def close_excel_app(wb, app):
    wb.save()
    app.books.close()
    app.quit()

app = initialize_excel_app()
wb = get_workbook(app, file_path)
ws = get_worksheet(wb, sheet_name)

data = get_data_range(ws, start_column, start_row, last_column, last_row)
df = convert_to_dataframe(data)
filtered_df = get_filtered_dataframe(df, target_column)

clear_contents(ws, start_column, start_row, last_column, last_row)
set_range_values(ws, start_column, start_row, filtered_df, num_columns)

close_excel_app(wb, app)

print(f"Total elapsed time: {time.time() - start_time:.2f} seconds")
