
# Import pandas library and read Excel file
import pandas as pd

excel_file = pd.ExcelFile('VI_New_raw_data_update_final.xlsx')

# Create an empty dictionary to store column values for each sheet
sheet_column_values = {}

# Loop over all sheets in the Excel file
for sheet_name in excel_file.sheet_names:
    # Read the sheet into a DataFrame, skipping the first row
    df = pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=1)
    # Create an empty dictionary to store column values for this sheet
    column_values = {}
    # Loop over the columns in the DataFrame
    for column in df.columns:
        # Store the values in the column as a list
        values = df[column].tolist()
        # Add the list to the dictionary with the column name as the key
        column_values[column] = values
    # Add the column values dictionary to the sheet dictionary with the sheet name as the key
    sheet_column_values[sheet_name] = column_values

CustomerAddress = sheet_column_values["CustomerAddress"]


address = CustomerAddress["address"]
postcode = CustomerAddress["postcode"]
postcode = [str(num) for num in postcode]
state = CustomerAddress["state"]
country = CustomerAddress["country"]

combined_list = [", ".join(entry) for entry in zip(address, state, postcode, country)]

print(combined_list)

shopAddress = "Bennelong Point, NSW, 2000, Australia"

import geocoder

def get_lat_lng(address):
    g = geocoder.osm(address, user_agent="my-application")
    if g.ok:
        return g.latlng
    else:
        return None


import math

def calculate_distance(locA, locB):

    coordinatesA = get_lat_lng(locA)
    coordinatesB = get_lat_lng(locB)

    # Radius of the Earth in kilometers
    radius = 6371.0

    # Convert latitude and longitude to radians
    lat1_rad = math.radians(coordinatesA[0])
    lon1_rad = math.radians(coordinatesA[1])
    lat2_rad = math.radians(coordinatesB[0])
    lon2_rad = math.radians(coordinatesB[1])

    # Difference between latitudes and longitudes
    dlat = lat2_rad - lat1_rad
    dlon = lon2_rad - lon1_rad

    # Haversine formula
    a = math.sin(dlat / 2) ** 2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon / 2) ** 2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

    # Calculate the distance
    distance = radius * c

    return distance

locA = shopAddress
locB = combined_list


distances = []
for i in range(len(locB)):
    print(locB[i])
    distance = calculate_distance(locA, locB[i])
    distances.append(distance)
    
print(distances)


        


#distance = calculate_distance(locA, locB)
#print("The distance between the two cities is", distance, "km.")


