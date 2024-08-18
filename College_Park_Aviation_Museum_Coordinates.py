import pandas as pd
import xlwings as xw
import numpy as np
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

#Reading the Excel file
xls = pd.ExcelFile('Tours_data Log_UMD.xlsx')

#Create dictionary of dataframes
df_dict = {}
for year in range(2018, 2025):
  df_dict[f'tours{year}'] = pd.read_excel(xls, f'FY {year}')

#Function to convert floats to strings that do not have a decimal point
def str_no_decimals(value):
  try:
    return str(int(value))
  except (ValueError, TypeError):
    return str(value)

for year in range(2018, 2025):
  df_dict[f'tours{year}']['Zipcode'] = df_dict[f'tours{year}']['Zipcode'].apply(str_no_decimals)

#Masking for rows that contain an address in DC
for year in range(2018, 2025):
  df_dict[f'tours{year}']['Full Address'] = np.nan

  dc = df_dict[f'tours{year}']['State'] == 'DC'
  df_dict[f'tours{year}']['Full Address'][dc] = (df_dict[f'tours{year}']['Street Address'] + ', ' + df_dict[f'tours{year}']['State'] + ', ' + df_dict[f'tours{year}']['Zipcode'])

  not_dc = ~dc
  df_dict[f'tours{year}']['Full Address'][not_dc] = (df_dict[f'tours{year}']['Street Address'] + ', ' + df_dict[f'tours{year}']['County'] + ' County' + ', ' + df_dict[f'tours{year}']['State'] + ', ' + df_dict[f'tours{year}']['Zipcode'])

#Retrieving coordinates from geocoding API
geolocater = Nominatim(user_agent="CPAM")
geocode = RateLimiter(geolocater.geocode, min_delay_seconds=1)
for year in range(2018, 2025):
  for index, row in df_dict[f'tours{year}'].iterrows():
    if pd.notna(row['Full Address']):
      location = row['Full Address']
      geocoded_location = geocode(location)
      if geocoded_location:
          df_dict[f'tours{year}'].at[index, 'longitude'] = geocoded_location.longitude
          df_dict[f'tours{year}'].at[index, 'latitude'] = geocoded_location.latitude

#Saving changes made to the Excel file
try:
    app = xw.App()
    wb = xw.Book('Tours_data Log_UMD.xlsx')
    
    for year in range(2018, 2025):
      ws = wb.sheets[f'FY {year}']
      ws.range('A1').options(index=False).value = df_dict[f'tours{year}']
    
    wb.save()
    wb.close()
    app.quit()
except: 
    pass