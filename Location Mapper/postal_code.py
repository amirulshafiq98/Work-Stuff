import sys
import requests
import time
import xlwings as xw

HEADER_ROW = 5
DATA_START_ROW = 6


def get_coordinates(postal_code):
    postal_code = str(postal_code).strip().replace(".0", "").zfill(6)

    url = (
        "https://www.onemap.gov.sg/api/common/elastic/search"
        f"?searchVal={postal_code}&returnGeom=Y&getAddrDetails=Y"
    )

    response = requests.get(url, timeout=10)
    data = response.json()
    time.sleep(0.3)

    if data.get("found", 0) > 0:
        result = data["results"][0]
        return result["LATITUDE"], result["LONGITUDE"]
    else:
        return None, None


def get_open_book_by_fullname(workbook_path):
    normalized_target = workbook_path.lower()

    for app in xw.apps:
        for book in app.books:
            try:
                if book.fullname and book.fullname.lower() == normalized_target:
                    return book, False
            except Exception:
                continue

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    book = app.books.open(workbook_path)
    return book, True


def run_postal_coordinates(workbook_path):
    book, should_close = get_open_book_by_fullname(workbook_path)

    try:
        sheet = book.sheets[0]

        # ========================
        # READ HEADERS (ROW 5)
        # ========================
        last_col = sheet.used_range.last_cell.column
        headers = sheet.range((HEADER_ROW, 1), (HEADER_ROW, last_col)).value

        if not isinstance(headers, list):
            headers = [headers]

        headers = [str(h).strip() if h is not None else "" for h in headers]

        if "Postal Code" not in headers:
            raise ValueError("Could not find a 'Postal Code' column in row 5.")

        postal_col = headers.index("Postal Code") + 1

        # ========================
        # CREATE / FIND LAT/LON COLUMNS
        # ========================
        if "Latitude" in headers:
            latitude_col = headers.index("Latitude") + 1
        else:
            latitude_col = last_col + 1
            sheet.cells(HEADER_ROW, latitude_col).value = "Latitude"
            last_col += 1
            headers.append("Latitude")

        if "Longitude" in headers:
            longitude_col = headers.index("Longitude") + 1
        else:
            longitude_col = last_col + 1
            sheet.cells(HEADER_ROW, longitude_col).value = "Longitude"

        # ========================
        # FIND LAST ROW BASED ON POSTAL COLUMN
        # ========================
        last_row = sheet.cells(sheet.rows.count, postal_col).end("up").row

        if last_row < DATA_START_ROW:
            print("No data rows found.")
            return

        total_rows = last_row - (DATA_START_ROW - 1)

        # ========================
        # MAIN LOOP
        # ========================
        for i, row_num in enumerate(range(DATA_START_ROW, last_row + 1), start=1):
            postal = sheet.cells(row_num, postal_col).value

            if postal in (None, ""):
                sheet.cells(row_num, latitude_col).value = None
                sheet.cells(row_num, longitude_col).value = None
                print(f"Checked {i} out of {total_rows}")
                continue

            lat, lon = get_coordinates(str(postal))

            sheet.cells(row_num, latitude_col).value = lat
            sheet.cells(row_num, longitude_col).value = lon

            print(f"Checked {i} out of {total_rows}")

        book.save()
        print("Done!")

    finally:
        if should_close:
            app = book.app
            book.close()
            app.quit()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit("Usage: postal_coordinates_xlwings.exe <full_workbook_path>")

    workbook_path = sys.argv[1]
    run_postal_coordinates(workbook_path)
