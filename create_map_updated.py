import pandas as pd
from geopy.geocoders import Nominatim
import folium
import webbrowser
import os

# Load the Excel file
file_path = r'C:\\Users\\DELL\\Desktop\\map\\Book copy copy copy copy copy copy copy.xlsx'
data = pd.read_excel(file_path)

# Define the geolocator
geolocator = Nominatim(user_agent="geoapiExercises")

# Clean and organize the data
cleaned_data = data.copy()

# Function to get coordinates
def get_coordinates(location):
    try:
        loc = geolocator.geocode(location + ", Jordan")
        if loc and loc.address.split(',')[-1].strip() == "Jordan":
            return loc.latitude, loc.longitude
        else:
            return None, None
    except Exception as e:
        print(f"Error in geocoding {location}: {e}")
        return None, None

# Check if Latitude and Longitude columns exist and fill missing values
if 'Latitude' not in cleaned_data.columns or 'Longitude' not in cleaned_data.columns:
    cleaned_data['Latitude'], cleaned_data['Longitude'] = zip(*cleaned_data['الاسواق / فروع الموسسة'].apply(get_coordinates))

# Create a base map of Jordan
jordan_map = folium.Map(location=[31.5, 35.9], zoom_start=8)

# Add OpenTopoMap layer for better terrain details
folium.TileLayer(
    tiles='https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png',
    attr='Map data © OpenStreetMap contributors, CC-BY-SA, Imagery © OpenTopoMap',
    name='OpenTopoMap',
    max_zoom=17,
).add_to(jordan_map)

# Set the map boundaries to Jordan more accurately
jordan_map.fit_bounds([[29.185, 34.959], [33.375, 39.301]])

# Function to add markers to the map
def add_markers(row):
    if pd.notna(row['Latitude']) and pd.notna(row['Longitude']):
        location = [row['Latitude'], row['Longitude']]
        popup_content = folium.Popup(f"""
            <b>{row['الاسواق / فروع الموسسة']}</b><br>
            {'مواقع و مساحات استثمار متوفرة: ' + str(row['مواقع و مساحات استثمار متوفرة']) if not pd.isna(row['مواقع و مساحات استثمار متوفرة']) else ''}
            {'<br>مواقع / مساحات مستغله حاليا: ' + str(row['مواقع / مساحات مستغله حاليا']) if not pd.isna(row['مواقع / مساحات مستغله حاليا']) else ''}
            {'<br>لا يوجد مواقع و مساحات للاستغلال: ' + str(row['لا يوجد مواقع و مساحات للاستغلال']) if not pd.isna(row['لا يوجد مواقع و مساحات للاستغلال']) else ''}
            <br><a href="https://www.google.com/maps/dir/?api=1&origin=YOUR_LOCATION&destination={location[0]},{location[1]}" target="_blank">احصل على الاتجاهات إلى هنا</a>
        """, max_width=300)

        icon_color = 'gray'  # Default color
        if not pd.isna(row['مواقع و مساحات استثمار متوفرة']):
            icon_color = 'green'
        elif not pd.isna(row['مواقع / مساحات مستغله حاليا']):
            icon_color = 'red'

        folium.Marker(
            location=location,
            popup=popup_content,
            icon=folium.Icon(color=icon_color)
        ).add_to(jordan_map)

# Apply the function to each row
cleaned_data.apply(add_markers, axis=1)

# Save the map to an HTML file
output_file = 'jordan_map_updated.html'
jordan_map.save(output_file)

# Open the HTML file in the default web browser
webbrowser.open(f'file://{os.path.realpath(output_file)}')

# Print a success message
print("The map has been created and viewed successfully.")
