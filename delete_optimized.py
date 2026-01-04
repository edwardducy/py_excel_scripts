import xlwings as xw
import pandas as pd
import time

# ---------------- Configuration ----------------

# Change to your file path
file_path = r"D:\Downloads\2025 10Oct List of Transaction detailed per Sources.xlsb"

# Change to your sheet name
sheet_name = "DN100-DVR128_OCT2025"

start_row = 14        # first row of the actual data
column_letter = 'I'   # column to filter
# ------------------------------------------------

start_time = time.time()
app = xw.App(visible=True)
app.display_alerts = False
app.screen_updating = False
app.calculation = 'manual'
wb = None

try:
    wb = xw.Book(file_path)
    ws = wb.sheets[sheet_name]

    if ws.api.AutoFilterMode:
        ws.api.AutoFilterMode = False

    last_row = ws.range(f"{column_letter}{ws.cells.last_cell.row}").end('up').row

    if last_row < start_row:
        print("No data rows found; nothing to delete.")
        wb.save()
        wb.close()
        wb = None
        raise SystemExit

    col_values = ws.range(f"{column_letter}{start_row}:{column_letter}{last_row}").value
    df = pd.DataFrame({column_letter: col_values})

    def ends_with_number(val):
        if val is None:
            return False
        s = ''.join(c for c in str(val) if c.isalnum())
        return s != "" and s[-1].isdigit()

    mask = df[column_letter].apply(ends_with_number)
    rows_to_keep = mask[mask].index.tolist()

    all_rows = list(range(start_row, start_row + len(df)))
    rows_to_delete = sorted(set(all_rows) - set([r + start_row for r in rows_to_keep]))

    print(f"Total rows to delete: {len(rows_to_delete)}")

    if rows_to_delete:
        ranges = []
        start = prev = rows_to_delete[0]
        for r in rows_to_delete[1:]:
            if r == prev + 1:
                prev = r
            else:
                ranges.append((start, prev))
                start = prev = r
        ranges.append((start, prev))

        for start_row_range, end_row_range in reversed(ranges):
            ws.range(f"{start_row_range}:{end_row_range}").api.Delete()

    wb.save()

finally:
    if wb is not None:
        wb.close()
    app.quit()

end_time = time.time()
print(f"Done.")
print(f"Time taken: {end_time - start_time:.2f} seconds")
