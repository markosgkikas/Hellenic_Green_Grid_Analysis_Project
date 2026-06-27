# This script downloads all raw data files from the ADMIE (Greek TSO) public API.
# For each file category, it makes one API call to get all files in the date range,
# then downloads each file into its own folder.

import requests
import os
import time
import zipfile

DATE_START = "2019-01-01"
DATE_END   = "2026-06-25"

FILE_CATEGORIES = [
    "SystemRealizationSCADA",
    "RealTimeSCADARES",
    "RealTimeSCADASystemLoad",
    "RealTimeSCADAImportsExports",
    "ReservoirFillingRate",
    "ISP1DayAheadRESForecast",
    "ISP2DayAheadRESForecast",
    "ISP3IntraDayRESForecast",
    "UnitProduction",
    "DayAheadRESForecast",
]

OUTPUT_FOLDER = "data/raw/admie"


# loop through each category
for category in FILE_CATEGORIES:
    print("=" * 40)
    print("Category:", category)

    # call the API to get the list of files for this category
    url = "https://www.admie.gr/getOperationMarketFilewRange"
    params = {
        "dateStart":    DATE_START,
        "dateEnd":      DATE_END,
        "FileCategory": category,
    }

    try:
        response = requests.get(url, params=params, timeout=60)
        records = response.json()
        print("  found", len(records), "files")
    except Exception as e:
        print("  ERROR getting file list:", e)
        continue

    # create the folder for this category if it doesn't exist
    category_folder = os.path.join(OUTPUT_FOLDER, category)
    if not os.path.exists(category_folder):
        os.makedirs(category_folder)

    # collect the download URL and filename for each file in the list
    # (the API already returns only the latest revision per date)
    files_to_get = []
    for record in records:
        download_url = record.get("file_path")
        if not download_url:
            continue
        filename = download_url.split("/")[-1]
        files_to_get.append((download_url, filename))

    print("  files to download:", len(files_to_get))

    # download each file
    for download_url, filename in files_to_get:
        destination = os.path.join(category_folder, filename)

        # download it (unless we already have it)
        if os.path.exists(destination):
            print("  already have:", filename, "- skipping")
        else:
            try:
                print("  downloading:", filename)
                file_response = requests.get(download_url, timeout=60)
                with open(destination, "wb") as f:
                    f.write(file_response.content)
                time.sleep(0.5)  # small pause so we don't overload the server
            except Exception as e:
                print("  ERROR downloading", filename, ":", e)
                continue

        # some categories (like UnitProduction) come as .zip archives.
        # extract each one into a folder next to it with the daily report*.xls files inside.
        if filename.lower().endswith(".zip"):
            extract_folder = destination[:-4]  # drop the ".zip" part
            if not os.path.exists(extract_folder):
                try:
                    with zipfile.ZipFile(destination) as z:
                        z.extractall(extract_folder)
                    print("  extracted:", filename)
                except Exception as e:
                    print("  ERROR extracting", filename, ":", e)

print()
print("Done!")
