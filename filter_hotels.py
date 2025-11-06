import json
import pandas as pd
from openpyxl import load_workbook
import os

# Read the compset distance JSON file
with open(r'C:\Users\reservations\Desktop\Compset Tool\Compset by distance.json', 'r', encoding='utf-8') as f:
    content = f.read()
    # Extract JSON from the file (it's embedded in a larger text)
    json_start = content.find('{')
    json_end = content.rfind('}') + 1
    compset_data = json.loads(content[json_start:json_end])

# Filter hotels: within 5 km and not luxury
hotels = compset_data['hotels']
filtered_hotels = [
    hotel for hotel in hotels
    if hotel['distance_km'] <= 5 and hotel['category'] != 'Luxury'
]

print(f"Total hotels in JSON: {len(hotels)}")
print(f"Hotels within 5 km: {len([h for h in hotels if h['distance_km'] <= 5])}")
print(f"Luxury hotels within 5 km: {len([h for h in hotels if h['distance_km'] <= 5 and h['category'] == 'Luxury'])}")
print(f"Non-luxury hotels within 5 km: {len(filtered_hotels)}")
print("\nFiltered hotels (non-luxury, within 5 km):")
print("=" * 80)

for hotel in filtered_hotels:
    print(f"{hotel['id']}. {hotel['name']}")
    print(f"   Star Rating: {hotel['star_rating']} | Category: {hotel['category']}")
    print(f"   Distance: {hotel['distance_km']} km | Area: {hotel['area']}")
    print()

# Save filtered hotels to a JSON file for later use
with open(r'C:\Users\reservations\Desktop\Compset Tool\filtered_hotels.json', 'w', encoding='utf-8') as f:
    json.dump(filtered_hotels, f, indent=2, ensure_ascii=False)

print(f"\nFiltered hotels saved to 'filtered_hotels.json'")
