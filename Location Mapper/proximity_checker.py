import sys
import time
import requests
import xlwings as xw
from math import radians, sin, cos, sqrt, atan2

TOP_N = 3
USER_INPUT_SHEET = "User Input"
CENTRE_INFO_SHEET = "Centre Info"
OUTPUT_SHEET = "Output"
DATA_START_ROW = 6


def clean_postal_code(postal_code):
    """Turn the postal code into a clean 6-digit string."""
    if postal_code is None:
        return ""
    return str(postal_code).strip().replace(".0", "").zfill(6)


def get_coordinates(postal_code):
    """Get latitude and longitude from the OneMap API."""
    postal_code = clean_postal_code(postal_code)
    if not postal_code:
        return None, None

    url = (
        "https://www.onemap.gov.sg/api/common/elastic/search"
        f"?searchVal={postal_code}&returnGeom=Y&getAddrDetails=Y"
    )

    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        data = response.json()
        time.sleep(0.3)

        if data.get("found", 0) > 0 and data.get("results"):
            result = data["results"][0]
            return float(result["LATITUDE"]), float(result["LONGITUDE"])

        return None, None
    except Exception as e:
        print(f"Error fetching postal code {postal_code}: {e}")
        return None, None


def haversine_km(lat1, lon1, lat2, lon2):
    """Calculate the distance between two coordinate points in kilometres."""
    radius_km = 6371.0
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
    return radius_km * 2 * atan2(sqrt(a), sqrt(1 - a))


def get_open_book_by_fullname(workbook_path):
    """Try to attach to a workbook that is already open in Excel."""
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


def read_centres(ws_centres):
    """Read centres from the Centre Info sheet."""
    last_row = ws_centres.range((ws_centres.cells.last_cell.row, 1)).end("up").row
    if last_row < DATA_START_ROW:
        return []

    rows = ws_centres.range(f"A{DATA_START_ROW}:D{last_row}").value
    if not isinstance(rows, list):
        return []
    if rows and not isinstance(rows[0], list):
        rows = [rows]

    centres = []
    for row in rows:
        centre_name = row[0]
        centre_lat = row[2]
        centre_lon = row[3]

        if centre_name in (None, ""):
            continue
        if centre_lat in (None, "") or centre_lon in (None, ""):
            continue

        centres.append(
            {
                "name": centre_name,
                "lat": float(centre_lat),
                "lon": float(centre_lon),
            }
        )

    return centres


def read_tutors(ws_input):
    """Read tutors from the User Input sheet."""
    last_row = ws_input.range((ws_input.cells.last_cell.row, 1)).end("up").row
    if last_row < DATA_START_ROW:
        return []

    rows = ws_input.range(f"A{DATA_START_ROW}:B{last_row}").value
    if not isinstance(rows, list):
        return []
    if rows and not isinstance(rows[0], list):
        rows = [rows]

    tutors = []
    for row in rows:
        tutor_name = row[0]
        tutor_postal = row[1]

        if tutor_name in (None, ""):
            continue

        tutors.append(
            {
                "name": tutor_name,
                "postal": clean_postal_code(tutor_postal),
            }
        )

    return tutors


def build_results(tutors, centres):
    """Get tutor coordinates, calculate distances, and prepare output rows."""
    results = []

    for index, tutor in enumerate(tutors, start=1):
        print(f"Processing {index} of {len(tutors)}: {tutor['name']} ({tutor['postal']})")

        lat, lon = get_coordinates(tutor["postal"])

        if lat is None or lon is None:
            results.append(
                [
                    tutor["name"],
                    tutor["postal"],
                    "Postal code not found",
                    None,
                    None,
                    None,
                    None,
                    None,
                ]
            )
            continue

        distances = []
        for centre in centres:
            distance_km = haversine_km(lat, lon, centre["lat"], centre["lon"])
            distances.append((centre["name"], round(distance_km, 2)))

        distances.sort(key=lambda item: item[1])
        top_matches = distances[:TOP_N]

        output_row = [tutor["name"], tutor["postal"]]
        for i in range(TOP_N):
            if i < len(top_matches):
                output_row.extend([top_matches[i][0], top_matches[i][1]])
            else:
                output_row.extend([None, None])

        results.append(output_row)

    return results


def write_results(ws_output, results):
    """Clear old rows and write the new output rows."""
    last_used_row = ws_output.range((ws_output.cells.last_cell.row, 1)).end("up").row
    clear_to_row = max(last_used_row, DATA_START_ROW)
    ws_output.range(f"A{DATA_START_ROW}:H{clear_to_row}").clear_contents()

    if results:
        end_row = DATA_START_ROW + len(results) - 1
        ws_output.range(f"A{DATA_START_ROW}:H{end_row}").value = results


def run_distance_checker(workbook_path):
    """Main function that reads the workbook, calculates results, and writes them back."""
    book, should_close = get_open_book_by_fullname(workbook_path)

    try:
        ws_centres = book.sheets[CENTRE_INFO_SHEET]
        ws_input = book.sheets[USER_INPUT_SHEET]
        ws_output = book.sheets[OUTPUT_SHEET]

        centres = read_centres(ws_centres)
        tutors = read_tutors(ws_input)

        print(f"Loaded {len(centres)} centres.")
        print(f"Loaded {len(tutors)} tutors.")

        results = build_results(tutors, centres)
        write_results(ws_output, results)

        book.save()
        print("Done. Output sheet has been updated.")
    finally:
        if should_close:
            app = book.app
            book.close()
            app.quit()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit("Usage: proximity_checker.exe <full_workbook_path>")

    workbook_path = sys.argv[1]
    run_distance_checker(workbook_path)