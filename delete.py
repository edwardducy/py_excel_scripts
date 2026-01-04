import xlwings as xw
import time

# ---------------- Configuration ----------------
file_path = r"D:\Downloads\2025 10Oct List of Transaction detailed per Sources.xlsb"
sheet_name = "DN100-DVR128_OCT2025"
start_row = 14      # first row of actual data
column_letter = 'I' # column to check
progress_step = 1000 # print progress every N deletions
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

    # Turn off any filter
    if ws.api.AutoFilterMode:
        ws.api.AutoFilterMode = False

    # Find last used row in column I
    last_row = ws.range(f"{column_letter}{ws.cells.last_cell.row}").end('up').row

    # Read all values from column I
    values = ws.range(f"{column_letter}{start_row}:{column_letter}{last_row}").value

    rows_to_delete = []

    # Determine rows to delete
    for i, val in enumerate(values):
        excel_row = start_row + i
        if val is None or (isinstance(val, str) and val.strip() == ""):
            rows_to_delete.append(excel_row)
            continue
        s = ''.join(c for c in str(val) if c.isalnum())
        if s == "" or not s[-1].isdigit():
            rows_to_delete.append(excel_row)

        # Progress print
        if len(rows_to_delete) % progress_step == 0:
            print(f"{len(rows_to_delete)} rows marked for deletion so far...")

    print(f"\nTotal rows to delete: {len(rows_to_delete)}")

    # Delete in bulk ranges, bottom-up
    if rows_to_delete:
        rows_to_delete.sort()
        ranges = []
        start = prev = rows_to_delete[0]
        for r in rows_to_delete[1:]:
            if r == prev + 1:
                prev = r
            else:
                ranges.append((start, prev))
                start = prev = r
        ranges.append((start, prev))

        print(f"Deleting {len(ranges)} consecutive row ranges...")
        for idx, (start_row_range, end_row_range) in enumerate(reversed(ranges), 1):
            ws.range(f"{start_row_range}:{end_row_range}").api.Delete()
            if idx % 10 == 0:  # progress every 10 ranges
                print(f"Deleted {idx}/{len(ranges)} ranges...")

    wb.save()

finally:
    if wb is not None:
        wb.close()
    app.quit()

end_time = time.time()
print(f"Done. All non-number-ending rows deleted with formatting preserved.")
print(f"Time taken: {end_time - start_time:.2f} seconds")
