# Google Geolocation Script

## Overview

This script processes latitude and longitude coordinates from an Excel file and retrieves location details using the Google Geolocation API. It ensures all relevant address components are available and updates the missing fields in the sheet.

## Features

- Reads an Excel file containing latitude and longitude values.
- Fetches location details from the Google Geolocation API.
- Supports missing column detection and prompts the user to add them.
- Converts latitude/longitude to:
  - DMS (Degrees, Minutes, Seconds)
  - USNG (U.S. National Grid) (only for U.S. locations)
- Handles timeouts and request failures gracefully.
- Saves processed locations incrementally to prevent data loss.
- Uses multithreading to speed up API requests.

## Dependencies

Ensure you have the following dependencies installed before running the script:

```sh
pip install requests pandas openpyxl pyproj
```

### Required Python Modules

- `requests` - For making API requests.
- `pandas` - For reading and modifying Excel files.
- `openpyxl` - For working with Excel files.
- `pyproj` - For coordinate conversions (USNG calculation).

## Setup

1. **Obtain a Google API Key:**

   - Get your API key from the [Google Cloud Console](https://console.cloud.google.com/).
   - Replace `YOUR_GOOGLE_API_KEY` in the script with your actual API key.

2. **Prepare Your Input File:**

   - Ensure your Excel file (`latlong.xlsx`) contains at least `Latitude` and `Longitude` columns.
   - Place the file in the same directory as the script.

3. **Run the Script:**

   ```sh
   python3 bulk-geo-data-mapping-utility.py
   ```

## Output

- The script updates the Excel file (`latlong.xlsx`) with missing details:

  - `City`
  - `State`
  - `Country`
  - `Zip`
  - `County`
  - `Street Address`
  - `DMS` (Degrees, Minutes, Seconds)
  - `USNG` (if in the U.S.)

- If a column is missing in the sheet, the script will prompt the user to add it.

- If a component is not available in the API response, it will be marked as `N/A`.

## Error Handling

- **Timeouts & Failures:** The script catches timeouts and request failures and marks those rows accordingly.
- **Incremental Saving:** The script saves processed locations after each entry to prevent data loss.

## Notes

- The script is optimized to run multiple API requests in parallel.
- It only calculates USNG coordinates for locations within the United States.
- If the Google API does not return certain details, they are marked as `N/A` instead of reprocessing repeatedly.

## License

This script is licensed under the Apache 2.0 License. See the LICENSE file for more details.

---

**Author:** [Your Name]
**Last Updated:** [Date]

