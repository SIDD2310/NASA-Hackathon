from pyowm import OWM
from datetime import datetime, timedelta
import pytz
from tzwhere import tzwhere
import xlrd
# from pyowm.utils.geo import Point
# from pyowm.commons.tile import Tile
# from pyowm.tiles.enums import MapLayerEnum


class color:
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    END = '\033[0m'


cities = []
radiation = []
brands = []
efficiency = []
area = []
price = []
cost = []

x = y = z = p = c = 0
pr = 0.75

owm = OWM('fe1385007f86f775eec151c0b17495a0')
wmgr = owm.weather_manager()
gmgr = owm.geocoding_manager()
tz = tzwhere.tzwhere()

city = input("Enter the city: ")
country = input("Enter the country: ")
brand = input("Enter the brand of the solar panel: ")
land_area = input("Enter the estimated area of the plot (in m²): ")
years = input("Enter the number of years: ")

lol = gmgr.geocode(city, country=country)
place = lol[0]
timezone = pytz.timezone(tz.tzNameAt(round(place.lat, 2), round(place.lon, 2)))
observation = wmgr.one_call(lat=place.lat, lon=place.lon)

print(color.BOLD + color.UNDERLINE + "Hourly Forcast" + color.END)
for i in range(12):
    print("Time: " + str((datetime.now(timezone) + timedelta(hours=i)).strftime("%H:%M")) + " Zone: " + str(timezone))
    print("Temperature: " + str(observation.forecast_hourly[i].temperature('celsius').get('temp')))
    print("Humidity: " + str(observation.forecast_hourly[i].humidity) + "\n")

print(color.BOLD + color.UNDERLINE + "\n\n~Daily Forcast~" + color.END)
for i in range(7):
    print("Time: " + str((datetime.now(timezone) + timedelta(days=i)).strftime("%d %B %Y")) + " Zone: " + str(timezone))
    print("Temperature: " + "Day " + str(observation.forecast_daily[i].temperature('celsius').get('day')) + "  Night " + str(observation.forecast_daily[i].temperature('celsius').get('night')))
    print("Humidity: " + str(observation.forecast_daily[i].humidity) + "\n")

c_loc = xlrd.open_workbook("Cities.xls")
b_loc = xlrd.open_workbook("Brands.xls")
cities_sheet = c_loc.sheet_by_index(0)
brands_sheet = b_loc.sheet_by_index(0)

for i in range(1, cities_sheet.nrows):
    cities.append(cities_sheet.cell_value(i, 1))
    radiation.append(cities_sheet.cell_value(i, 15))
    cost.append(cities_sheet.cell_value(i, 10))

for i in range(0, brands_sheet.nrows):
    brands.append(brands_sheet.cell_value(i, 0))
    area.append(brands_sheet.cell_value(i, 2))
    efficiency.append(brands_sheet.cell_value(i, 4))
    price.append(brands_sheet.cell_value(i, 1))

for i in range(len(cities)):
    if city == cities[i]:
        z = radiation[i]
        c = cost[i]

for i in range(len(brands)):
    if brand == brands[i]:
        x = area[i]
        y = efficiency[i]
        p = price[i]

number_of_panels = int(float(land_area) / x)
po = (x * y * z * pr)*number_of_panels
print("Power Output: " + str(po))


total_cost = number_of_panels * p
earning = po * c * int(years)

roi = (earning/total_cost)*100
print("Return on investment in " + years + ": " + str(round(roi, 2)) + "%")

roy=total_cost/(earning)   # put cities yearly cost
print("Return on investment in years " +str(roy)+" yrs")

'''Power Output: 25805510.21902873
Return on investment in 10: 2625.46%
Return on investment in years 0.00234805890227577'''