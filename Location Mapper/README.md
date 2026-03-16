<img width="1200" height="799" alt="image" src="https://github.com/user-attachments/assets/9fb5f4c5-ef6e-4f1f-9d54-7f137527976d" />

# Proximity Checker in Excel

An Excel automation tool that calculates the closest geographic matches between two sets of locations using Singapore postal codes and [OneMap](https://www.onemap.gov.sg) Singapore.

**Download:**
➡️ [Download] the latest ZIP from the **Releases** section of this repository.
Extract the ZIP and ensure that the `.exe` and all necessary files remain in the same ZIP folder.

---

## What This Tool Does

This tool automates the process of manually calculating distances between addresses by retrieving precise coordinates and performing a proximity analysis.

| Feature | Description |
|---|---|
| ✅ Real-time Excel Integration | Connects to an active workbook via `xlwings` to read and write data live |
| ✅ OneMap API Integration | Fetches accurate latitude/longitude coordinates from the Singapore Land Authority (SLA) |
| ✅ Automatic Data Cleaning | Standardises postal codes into valid 6-digit strings (handles leading zeros and formatting) |
| ✅ Top-N Matching | Identifies and ranks the **Top 3** closest centres for every individual in your list |
| ✅ Automated Output | Clears old data and generates a fresh report in the "Output" sheet with a single click |

---

<img width="1351" height="311" alt="image" src="https://github.com/user-attachments/assets/c8aaa06a-77f6-41c5-a656-934b71394e63" />


## How to Use

1. **Prepare your data** — Enter the data of locations you are interested in **"User Input"** sheet and maintain the data in **"Centre Info"** sheet as that will be used as reference ***(Note: if the centre info needs to be changed, check the `Update` section)***.
2. **Click the button** — Click **"Click This to Populate!"** inside the Excel file. This triggers the `.exe` file in the ZIP file you opened the workbook from.
3. **Wait for processing** — The tool cleans postal codes, calls the OneMap API, calculates distances, and finds the Top 3 matches per participant.
4. **Check the Output sheet** — The **"Output"** sheet is automatically cleared and populated with centre names and distances in kilometres.

---

## Data Validation & Error Handling

The tool includes built-in handling for bad or missing postal codes, so the script won't crash on messy data.

```python
def clean_postal_code(postal_code):
    """Turn the postal code into a clean 6-digit string."""
    if postal_code is None:
        return ""
    return str(postal_code).strip().replace(".0", "").zfill(6)
```

**Internal Cleaning (`clean_postal_code`)**
- Strips extra spaces and removes decimal points (e.g., `"310123.0"` → `"310123"`)
- Uses `.zfill(6)` to pad 5-digit codes with a leading zero (e.g., `"10123"` → `"010123"`)

**Missing Inputs**
- If a postal code cell is blank or empty, the tool skips the API call entirely and moves on — no crash, no delay.

**Invalid Postal Codes**
- If a valid-looking 6-digit code doesn't exist in the OneMap database (e.g., `"999999"`), the API returns no results. The tool catches this gracefully.

**Output for Problem Records**
- If coordinates can't be retrieved, the row will show `"Postal code not found"` in the Output sheet with blank distance fields, so you can manually follow up on those records.

---

## How Distance is Calculated

The tool uses the [**Haversine formula**](https://en.wikipedia.org/wiki/Haversine_formula) to calculate straight-line ("as-the-crow-flies") distance between two geographic points.

**What this means in plain terms:** given two postal codes, the tool converts each into a latitude/longitude coordinate pair, then calculates the shortest possible distance between them on the surface of the Earth.

> **Note:** This is *not* a road distance or travel time. It's the geometric distance between two points, which may differ from actual travel routes.

**Formula (for reference):**

$$a = \sin^2\left(\frac{\Delta\text{lat}}{2}\right) + \cos(\text{lat}_1)\cdot\cos(\text{lat}_2)\cdot\sin^2\left(\frac{\Delta\text{lon}}{2}\right)$$

$$d = R \cdot 2 \cdot \arctan2\left(\sqrt{a}\,\sqrt{1-a}\right)$$

Where $d$ is the distance in kilometres, $R = 6371.0\,\text{km}$ (Earth's radius), and $\frac{\Delta \text{lat}}{\Delta \text{lon}}$ are coordinate differences in radians.

**Python Implementation:**

```python
def haversine_km(lat1, lon1, lat2, lon2):
    """Calculate the distance between two coordinate points in kilometres."""
    radius_km = 6371.0
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
    return radius_km * 2 * atan2(sqrt(a), sqrt(1 - a))
```

---

## Technical Details

| Item | Detail |
|---|---|
| Language | Python |
| Key Libraries | `xlwings` (Excel automation), `requests` (API calls), `math` (distance formula) |
| API | [OneMap API](https://www.onemap.gov.sg/apidocs/) by the Singapore Land Authority (SLA) |
| Internet Required | Yes — needed to reach the OneMap API endpoint |
| Data Privacy | Only 6-digit postal codes are sent to the API. No names, NRICs, or personal identifiers leave your machine or the Excel file. All processing is done locally. |

---

## How to Update Centre Postal Codes and the Coordinates?

If the existing postal codes are outdated, follow the instructions below to update them:

1. **Download the [Postal Code] ZIP file** — This ZIP file contains the necessary files to regenerate the coordinates based on inputted postal codes.
2. **Input Postal Codes and the Names of the buildings in `Postal Code.xlsm`** — Once updated, click the `Generate Output` button.
3. **Wait for Processing** — The tool cleans postal codes, calls the OneMap API and fills in the longitude and latitude of the buildings based on the postal code.
4. **Check the Output** — If any blanks are present, it indicates that the postal code was not written correctly. Re-enter the details, clear the output, and regenerate the coordinates.
5. **After Processing** — Copy all rows **_excluding the headers_** under `Centre Name`, `Postal Code`, `Latitude`, `Longitude` and paste them into the `Centre Info` sheet in `file_with_centres.xlsm`

---

## Limitations

- **Internet required** — The tool cannot run offline; it needs to reach the OneMap API.
- **Straight-line distance only** — Does not account for actual road routes or public transport travel times.
- **Singapore postal codes only** — Only valid 6-digit Singapore postal codes are supported. Other formats will result in missing output data.
- **No Preference Assigning** — The assignment does not account for duplicate postal codes. So if 2 people have the same postal code, they will both have the same 3 closest centres
