import requests
import csv
import pandas as pd
import os
import concurrent.futures
import time
from pyproj import Proj, transform

# Replace with your actual Google API Key
API_KEY = "YOUR_GOOGLE_API_KEY"

# Specify the input file path (same directory as script)
input_file = os.path.join(os.path.dirname(__file__), "latlong.xlsx")

# Load the Excel file
df = pd.read_excel(input_file)

# Function to convert latitude and longitude to DMS (Degrees, Minutes, Seconds)
def decimal_to_dms(deg, is_lat=True):
    direction = "N" if is_lat and deg >= 0 else "S" if is_lat else "E" if deg >= 0 else "W"
    deg = abs(deg)
    d = int(deg)
    m = int((deg - d) * 60)
    s = round((deg - d - m/60) * 3600, 2)
    return f"{d}°{m}'{s}\"{direction}"

# Function to convert lat/lon to USNG (U.S. National Grid)
def latlon_to_usng(lat, lon):
    try:
        proj_utm = Proj(proj="utm", zone=int((lon + 180) / 6) + 1, datum="WGS84")
        proj_latlon = Proj(proj="latlong", datum="WGS84")
        easting, northing = transform(proj_latlon, proj_utm, lon, lat)
        return f"{int(easting)},{int(northing)}"
    except Exception as e:
        return "Conversion Error"

# Ensure required columns exist in the sheet
required_columns = ["City", "State", "Country", "Zip", "County", "Street Address", "DMS", "USNG"]
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    user_input = input(f"Missing columns: {missing_columns}. Do you want to add them? (yes/no): ")
    if user_input.lower() == "yes":
        for col in missing_columns:
            df[col] = ""
    else:
        print("Skipping column addition. Some data may not be stored correctly.")

def get_location(lat, lon):
    """Fetch location details using Google Geolocation API."""
    url = f"https://maps.googleapis.com/maps/api/geocode/json?latlng={lat},{lon}&key={API_KEY}"
    try:
        response = requests.get(url, timeout=10)
        data = response.json()
    
        if data["status"] == "OK":
            address_components = data["results"][0]["address_components"]
            city, state, country, zip_code, county, street_address = "", "", "", "", "", ""
            
            for component in address_components:
                if "locality" in component["types"]:
                    city = component["long_name"]
                elif "administrative_area_level_1" in component["types"]:
                    state = component["long_name"]
                elif "country" in component["types"]:
                    country = component["long_name"]
                elif "postal_code" in component["types"]:
                    zip_code = component["long_name"]
                elif "administrative_area_level_2" in component["types"]:
                    county = component["long_name"]
                elif "route" in component["types"] or "street_number" in component["types"]:
                    street_address = component["long_name"]
            
            dms = decimal_to_dms(lat, True) + ", " + decimal_to_dms(lon, False)
            
            # Only calculate USNG if the country is the United States
            usng = latlon_to_usng(lat, lon) if country.lower() in ["united states", "us", "usa"] else ""
            
            return city, state, country, zip_code, county, street_address, dms, usng, ""
        else:
            return "", "", "", "", "", "", "", "", data.get("status", "Unknown Error")
    except requests.exceptions.Timeout:
        return "", "", "", "", "", "", "", "", "Timeout Error"
    except requests.exceptions.RequestException as e:
        return "", "", "", "", "", "", "", "", str(e)

def process_locations():
    total_locations = len(df)
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        future_to_index = {
            executor.submit(get_location, row["Latitude"], row["Longitude"]): index 
            for index, row in df.iterrows() if any(pd.isna(row[col]) or row[col] == "" for col in ["City", "State", "Country", "Zip", "County", "Street Address"])
        }
        
        for future in concurrent.futures.as_completed(future_to_index):
            index = future_to_index[future]
            try:
                city, state, country, zip_code, county, street_address, dms, usng, error_reason = future.result()
                df.at[index, "City"] = city if city else error_reason
                df.at[index, "State"] = state
                df.at[index, "Country"] = country
                df.at[index, "Zip"] = zip_code
                df.at[index, "County"] = county
                df.at[index, "Street Address"] = street_address
                df.at[index, "DMS"] = dms
                df.at[index, "USNG"] = usng
                print(f"Processed {index + 1}/{total_locations}: {city}, {state}, {country}, {zip_code}, {county}, {street_address}, {dms}, {usng} {f'(Error: {error_reason})' if error_reason else ''}")
            except Exception as e:
                print(f"Error processing index {index}: {e}")
                df.at[index, "City"] = f"Error: {e}"
            finally:
                # Save progress to file after every processed location
                df.to_excel(input_file, index=False)
    
# Run processing
process_locations()

# Final save to ensure everything is stored
df.to_excel(input_file, index=False)
print(f"Updated Excel file '{input_file}' with location data!")
