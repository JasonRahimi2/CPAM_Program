import pandas as pd
import numpy as np
import folium
from folium.plugins import HeatMap
from folium.plugins import MarkerCluster
import tkinter as tk
from tkinter import ttk
import webbrowser
from tkcalendar import Calendar
import tkinter.messagebox
import xlwings as xw
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import re

#Loads Sheets from the Excel file and places their titles as years in fy_sheets
def load_sheets():
  xls = pd.ExcelFile('Tours_Data Log_UMD.xlsx')

  sheets = xls.sheet_names
  sheets = [title.replace(' ', '') for title in sheets]
  sheets = [title.upper() for title in sheets]
  fy_sheets = [title for title in sheets if title.startswith('FY') and title[2:].isdigit()]
  fy_sheets.reverse()
  fy_sheets = [title.replace('FY', '') for title in fy_sheets]
  xls.close()
  return fy_sheets

#Loads each Excel sheet into df_dict to create a dictionary of dataframes
def load_dictionary(fy_sheets):
  xls = pd.ExcelFile('Tours_Data Log_UMD.xlsx')
  df_dict = {}
  for year in fy_sheets:
    df_dict[f'tours{year}'] = pd.read_excel(xls, f'FY {year}')
  xls.close()
  return df_dict

#Creates a dataframe containing every entry in every Excel sheet
def combined_tours(df_dict):
  fy_sheets = load_sheets()
  df_dict = load_dictionary(fy_sheets)

  concat_dict = []
  for year in fy_sheets:
    concat_dict.append(df_dict[f'tours{year}'])
  combined_tours_df = pd.concat(concat_dict)

  return combined_tours_df

#Retrieves coordinates from a geocoding API and stores coordinates in the Excel file and develops filter screen for heatmap
def heatmap():
  fy_sheets = load_sheets()

  clear_screen()

#Developing a loading bar to track progress of retrieving coordinates
  load = ttk.Progressbar(root, orient="horizontal", length=150, mode='determinate')
  load.grid(row = 0, column = 0)
  root.geometry('150x22')

#Updates the loading bar
  def progression(value):
    load['value'] = value
    root.update_idletasks()
  progress_value = 0

#Converts to string without a decimal point
  def str_no_decimals(value):
    try:
      return str(int(value))
    except (ValueError, TypeError):
      return str(value)
    
  progress_value = 10
  progression(progress_value)

  df_dict = load_dictionary(fy_sheets)

#Converts values under Zipcode column to strings without a decimal point
  for year in fy_sheets:
    df_dict[f'tours{year}']['Zipcode'] = df_dict[f'tours{year}']['Zipcode'].apply(str_no_decimals)

#Masking for DC addresses so that they have a different format under the Full Address column
  for year in fy_sheets:
    df_dict[f'tours{year}']['Full Address'] = np.nan

    dc = df_dict[f'tours{year}']['State'] == 'DC'
    df_dict[f'tours{year}']['Full Address'][dc] = (df_dict[f'tours{year}']['Street Address'] + ', ' + df_dict[f'tours{year}']['State'] + ', ' + df_dict[f'tours{year}']['Zipcode'])

    not_dc = ~dc
    df_dict[f'tours{year}']['Full Address'][not_dc] = (df_dict[f'tours{year}']['Street Address'] + ', ' + df_dict[f'tours{year}']['County'] + ' County' + ', ' + df_dict[f'tours{year}']['State'] + ', ' + df_dict[f'tours{year}']['Zipcode'])


  progress_value = progress_value + 10
  progression(progress_value)

  #Getting coordinates from the geocoding API
  geolocater = Nominatim(user_agent="CPAM")
  geocode = RateLimiter(geolocater.geocode, min_delay_seconds=1)

#Getting coordinates from rows with dates within 100 days prior to the most recent date in the 2 newest sheets 
  fy_two = fy_sheets[-2:]
  for year in fy_two:
    last_date = df_dict[f'tours{year}']['Date'].max()
    start_date = last_date - pd.DateOffset(days=100)


    for index, row in df_dict[f'tours{year}'][df_dict[f'tours{year}']['Date'].between(start_date, last_date)].iterrows():
      #Adds to loading bar
      progress_value += 0.5
      progression(progress_value)

#Setting conditions for which rows to get coordinates for and what should be sent to the API
      if pd.notna(row['Full Address']) and pd.isna(row['longitude']) and pd.isna(row['latitude']):
        Location = row['Full Address']
        if '#' in Location:
          geocoded_location = geocode(re.sub(r'#[^,]*', '', Location))
        elif 'Suite' in Location:
          geocoded_location = geocode(re.sub(r'Suite[^,]*', '', Location))
        elif 'Unit' in Location:
          geocoded_location = geocode(re.sub(r'Unit[^,]*', '', Location))
        elif 'Apt' in Location:
          geocoded_location = geocode(re.sub(r'Apt[^,]*', '', Location))
        else:
          geocoded_location = geocode(Location)

        if geocoded_location:
          df_dict[f'tours{year}'].at[index, 'longitude'] = geocoded_location.longitude
          df_dict[f'tours{year}'].at[index, 'latitude'] = geocoded_location.latitude

#Saving the changes to the Excel file
  try: 
    with xw.App(visible=False) as app:
      wb = app.books.open('Tours_Data Log_UMD.xlsx')

      for year in fy_two:
        ws = wb.sheets[f'FY {year}']
        ws.range('A1').options(index=False).value = df_dict[f'tours{year}']

      wb.save('Tours_Data Log_UMD.xlsx')
      wb.close()
      app.quit()
  except: 
    pass

  xls = pd.ExcelFile('Tours_Data Log_UMD.xlsx')

  df_dict = load_dictionary(fy_sheets)
  combined_tours_df = combined_tours(df_dict)

  cities = combined_tours_df['City'].dropna()
  cities = cities.drop_duplicates()

#Setting up lists for dropdown menus

  cities = sorted(list(cities))
  cities.insert(0, 'Select City')

  counties = combined_tours_df['County'].dropna()
  counties = counties.drop_duplicates()

  counties = sorted(list(counties))
  counties.insert(0, 'Select County')

  states = combined_tours_df['State'].dropna()
  states = states.drop_duplicates()

  states = sorted(list(states))
  states.insert(0, 'Select State')

  latest_date = combined_tours_df.iloc[len(combined_tours_df) - 1]['Date']
  latest_day = latest_date.day
  latest_month = latest_date.month
  latest_year = latest_date.year

  school_types = combined_tours_df['School Type'].dropna()
  school_types = school_types.drop_duplicates()

  school_types = sorted(list(school_types))
  school_types.insert(0, 'Select School Type')

  progression(99.99)
  clear_screen()
  root.geometry('1200x280')

#Developing the dropdown menus

#City dropdown menu filter
  tkvar_city = tk.StringVar(root)
  tkvar_city.set(cities[0])

  city_menu = ttk.Combobox(root, textvariable=tkvar_city, values = [*cities])
  city_menu.config(state="readonly")
  city_menu.grid(row = 1, column = 0, padx = 10)

#County dropdown menu filter
  tkvar_county = tk.StringVar(root)
  tkvar_county.set(counties[0])
        
  county_menu = ttk.Combobox(root, textvariable=tkvar_county, values = [*counties])
  county_menu.config(state="readonly")
  county_menu.grid(row = 1, column = 1, padx = 10)

#State dropdown menu filter
  tkvar_state = tk.StringVar(root)
  tkvar_state.set(states[0])

  state_menu = ttk.Combobox(root, textvariable=tkvar_state, values = [*states])
  state_menu.config(state="readonly")
  state_menu.grid(row = 1, column = 2, padx = 10)

#School Type dropdown menu filter
  tkvar_school_type = tk.StringVar(root)
  tkvar_school_type.set(school_types[0])

  school_type_menu = ttk.Combobox(root, textvariable=tkvar_school_type, values = [*school_types])
  school_type_menu.config(state="readonly")
  school_type_menu.grid(row = 1, column = 3, padx = 10)

#Start date calendar widget
  cal_start = Calendar(root, selectmode = 'day', year = 2017, month = 7, day = 3)
  cal_start.grid(row = 1, column = 4, padx = 10)

  start_label = tk.Label(root, text='Start Date')
  start_label.grid(row = 0, column = 4)

#End date calendar widget
  cal_end = Calendar(root, selectmode ='day', year = latest_year, month = latest_month, day = latest_day)
  cal_end.grid(row = 1, column = 5, padx = 10)

  end_label = tk.Label(root, text='End Date')
  end_label.grid(row = 0, column = 5)

#Button to open the heatmap
  gen_heatmap_button = ttk.Button(root, text="Generate Heatmap", command=lambda: open_heatmap(tkvar_city, tkvar_county, tkvar_state, cal_start, cal_end, tkvar_school_type))
  gen_heatmap_button.grid(row = 2, column = 2, columnspan = 2)

#Button to go back to start screen
  go_back = back_button()
  go_back.grid(row = 2, column = 0)

#Generates the heatmap when Generate Heatmap button is clicked
def open_heatmap(tkvar_city, tkvar_county, tkvar_state, cal_start, cal_end, tkvar_school_type):
  city_select = tkvar_city.get()
  county_select = tkvar_county.get()
  state_select = tkvar_state.get()
  start_date = cal_start.selection_get()
  end_date = cal_end.selection_get()
  school_type_select = tkvar_school_type.get()

#Error message for calendar widgets
  if (start_date > end_date):
    calendar_message()
    return

  def str_no_decimals(value):
    try:
      return str(int(value))
    except (ValueError, TypeError):
      return str(value)
    
  fy_sheets = load_sheets()

  df_dict = load_dictionary(fy_sheets)

  for year in fy_sheets:
    df_dict[f'tours{year}']['Zipcode'] = df_dict[f'tours{year}']['Zipcode'].apply(str_no_decimals)
    
  concat_dict = []
  for year in fy_sheets:
    concat_dict.append(df_dict[f'tours{year}'])
  combined_tours_df = pd.concat(concat_dict)

#Applying filters to heatmap
  start_date = pd.to_datetime(start_date)
  end_date = pd.to_datetime(end_date)
  combined_tours_df = combined_tours_df.loc[(combined_tours_df['Date'] >= start_date) & (combined_tours_df['Date'] <= end_date)]

#Apply city filter if a value other than Select City has been selected from the city dropdown menu
  if city_select != 'Select City':
    combined_tours_df = combined_tours_df[combined_tours_df['City'] == city_select]

#Apply county filter if a value other than Select County has been selected from the county dropdown menu
  if county_select != 'Select County':
    combined_tours_df = combined_tours_df[combined_tours_df['County'] == county_select]

#Apply state filter if a value other than Select State has been selected from the state dropdown menu
  if state_select != 'Select State':
    combined_tours_df = combined_tours_df[combined_tours_df['State'] == state_select]

#Apply school type filter if a value other than Select School Type has been selected from the school type dropdown menu
  if school_type_select != 'Select School Type':
    combined_tours_df = combined_tours_df[combined_tours_df['School Type'] == school_type_select]

#Dataframe for components of the heatmap
  df_heatmap = combined_tours_df.loc[:, ['latitude', 'longitude', 'Organization', 'Zipcode']]
  df_heatmap = df_heatmap.dropna()
  df_heatmap = df_heatmap.groupby(df_heatmap.columns.tolist(), as_index=False).size()
  df_heatmap = df_heatmap[['latitude', 'longitude', 'size', 'Organization', 'Zipcode']]

#Generating heatmap
  heatm = folium.Map(location=[38.9779099, -76.925903], tiles='OpenStreetMap', zoom_start=10)
  coordinates = [[row['latitude'], row['longitude'], row['size']] for index, row in df_heatmap.iterrows()]
  HeatMap(coordinates, min_opacity=0.4, blur=15).add_to(folium.FeatureGroup(name='Heat Map').add_to(heatm))

#Adding makers to the heatmap
  marker_feature = folium.FeatureGroup(name='Markers')
  marker_cluster = MarkerCluster().add_to(marker_feature)
  for i in range(0, len(df_heatmap)):
    text1 = df_heatmap.iloc[i]['Organization']
    text2 = df_heatmap.iloc[i]['size']
    text3 = df_heatmap.iloc[i]['Zipcode']
    folium.Marker([df_heatmap.iloc[i]['latitude'], df_heatmap.iloc[i]['longitude']], popup=f'Organization: {text1}<br> Visits: {text2}<br> Zip Code: {text3}').add_to(marker_cluster)
  marker_feature.add_to(heatm)

  folium.LayerControl().add_to(heatm)

#Saving then opening heatmap
  heatm.save("heatmap.html")
    
  webbrowser.open("heatmap.html")

#Generates visits by year visualization
def visits_by_year():
  clear_screen()
  root.geometry('1400x625')

  fy_sheets = load_sheets()
  df_dict = load_dictionary(fy_sheets)

  size_dict = {}
  size_list = []

  for year in fy_sheets:
    size_dict[f'tours{year}'] = df_dict[f'tours{year}'].loc[:, ['Organization',  'Date']]
    size_dict[f'tours{year}'] = size_dict[f'tours{year}'].dropna()
    size_dict[f'tours{year}']  = size_dict[f'tours{year}'].groupby(size_dict[f'tours{year}'].columns.tolist(), as_index=False).size()
    size_dict[f'tours{year}'] = len(size_dict[f'tours{year}'])
    size_list.append(size_dict[f'tours{year}'])

  size_index = 0
  fig = Figure(figsize = (14,6))
  plot = fig.add_subplot(111)

  for year in fy_sheets:
    organizations = size_list[size_index]
    plot.bar('FY ' + fy_sheets[size_index], organizations)
    size_index += 1

  plot.set_xlabel('Years')
  plot.set_ylabel('Visits')
  plot.set_title('Visits by Year')

  canvas = FigureCanvasTkAgg(fig, master=root)
  canvas.get_tk_widget().grid(row = 0, column = 0, columnspan = 15)
  canvas.draw()

#Adds button to save the visualization
  save = ttk.Button(root, text="Save Chart", command=lambda: save_chart(fig, 'VisitsbyYear'))
  save.grid(row = 1, column = 6, columnspan = 2)

  go_back = back_button()
  go_back.grid(row = 1, column = 1)

#Generates visualization for the top 10 organizations
def top_10_organizations():
  clear_screen()
  root.geometry('1680x625')

  fy_sheets = load_sheets()
  df_dict = load_dictionary(fy_sheets)
  combined_tours_df = combined_tours(df_dict)

  top10_visualization_df = combined_tours_df.loc[:, ['Organization']]
  top10_visualization_df = top10_visualization_df.dropna()
  top10_visualization_df = top10_visualization_df.groupby(top10_visualization_df.columns.tolist(), as_index=False).size()
  top10_df = top10_visualization_df.nlargest(10, 'size')

  fig = Figure(figsize = (18, 6))
  plot = fig.add_subplot(111)

  plot.bar(top10_df['Organization'], top10_df['size'])
  plot.set_xlabel('Organizations')
  plot.set_ylabel('Visits')
  plot.set_title('Top 10 Organizations by Visits')
  plot.set_xticks(range(0, 10))
  plot.set_xticklabels(plot.get_xticklabels(), fontsize = 6, rotation=5)

  canvas = FigureCanvasTkAgg(fig, master=root)
  canvas.get_tk_widget().grid(row = 0, column = 0, columnspan = 15, sticky='nsew')
  canvas.draw()

  go_back = back_button()
  go_back.grid(row = 2, column = 1)

  save = ttk.Button(root, text="Save Chart", command=lambda: save_chart(fig, 'Top10Organizations'))
  save.grid(row = 2, column = 6, columnspan = 2)

#Develops filter screen for breakdown by month visualization
def breakdown_by_month_filter():
  clear_screen()
  root.geometry('650x100')

  fy_sheets = load_sheets()

  years = fy_sheets
  years.insert(0, 'Select Year')

  tkvar_year = tk.StringVar(root)
  tkvar_year.set(years[0])

#Adds dropdown menu to filter for a specific fiscal year
  year_menu = ttk.Combobox(root, textvariable=tkvar_year, values = [*years])
  year_menu.config(state="readonly")
  year_menu.grid(row = 0, column = 1, padx = 100)

  gen_breakdown_button = ttk.Button(root, text="Generate Breakdown by Month", command=lambda: breakdown_by_month(tkvar_year))
  gen_breakdown_button.grid(row = 1, column = 1, padx = 100, pady = 50)

  go_back = back_button()
  go_back.grid(row = 1, column = 0, padx = 100, pady = 50)

#Generates breakdown by month visualization
def breakdown_by_month(tkvar_year):
  year_select = tkvar_year.get()

  if year_select == 'Select Year':
    year_message()
    return

  clear_screen()
  root.geometry('1000x625')

  fy_sheets = load_sheets()
  df_dict = load_dictionary(fy_sheets)

#Develops range for visualization to start from the beginning of the fiscal year in July to the end of the fiscal year in June
  def month_to_list(start):
    month_dict[f'tours{month}'] = df_dict[f'tours{year_select}'][df_dict[f'tours{year_select}']['Date'].dt.month == start]
    month_dict[f'tours{month}'] = month_dict[f'tours{month}'].loc[:, ['Organization',  'Date']]
    month_dict[f'tours{month}'] = month_dict[f'tours{month}'].dropna()
    month_dict[f'tours{month}']  = month_dict[f'tours{month}'].groupby(month_dict[f'tours{month}'].columns.tolist(), as_index=False).size()
    month_dict[f'tours{month}'] = len(month_dict[f'tours{month}'])
    month_list.append(month_dict[f'tours{month}'])

  month_dict = {}
  month_list = []
  july_start = 7
  january_start = 1
  for month in range(1, 13):
    if july_start <= 12:
        month_to_list(july_start)
        july_start += 1
    else:
        month_to_list(january_start)
        january_start += 1

  months = ['July', 'August', 'September', 'October', 'November', 'December', 'January', 'February', 'March', 'April', 'May', 'June']

  fig = Figure(figsize = (10,6))
  plot = fig.add_subplot(111)

  for month in range(1, 13):
    plot.bar(months[month - 1], month_list[month - 1])

  plot.set_yticks(range(0, max(month_list), 2))
  plot.set_xlabel('Months')
  plot.set_ylabel('Visits')
  plot.set_title(f'Visits by Month for FY {year_select}')
  plot.set_xticks(range(0, 12))
  plot.set_xticklabels(plot.get_xticklabels(), fontsize = 8)


  canvas = FigureCanvasTkAgg(fig, master=root)
  canvas.get_tk_widget().grid(row = 0, column = 0, columnspan = 15)
  canvas.draw()

  go_back = back_button_breakdown_by_month_filter()
  go_back.grid(row = 1, column = 1)

  save = ttk.Button(root, text="Save Chart", command=lambda: save_chart(fig, f'BreakdownbyMonthFY{year_select}'))
  save.grid(row = 1, column = 6, columnspan = 2)

#Generates visualization for county visits
def county_visits():
  clear_screen()
  root.geometry('1000x625')

  fy_sheets = load_sheets()
  df_dict = load_dictionary(fy_sheets)
  combined_tours_df = combined_tours(df_dict)

#Finds number of visits for each county and develops dataframe with County and Visits
  county_df = pd.DataFrame(combined_tours_df['County'].groupby(combined_tours_df['County']).size())
  county_df = county_df.rename(columns = {'County':'Visits'})
  county_df = county_df.sort_values(by='Visits', ascending=False)
  county_df = county_df.reset_index()

#If visits from a county are less than 10 then the county will not be displayed as a separate slice on the chart
  less_than_10 = county_df[county_df['Visits'] < 10]
  less_than_10 = less_than_10['Visits'].sum()
  county_row = pd.DataFrame({'County': 'Other', 'Visits': [less_than_10]})
  county_df = county_df[county_df['Visits'] >= 10]
  county_df = pd.concat([county_df, county_row], ignore_index=True)

  fig = Figure(figsize = (10,6))
  plot = fig.add_subplot(111)

  plot.pie(county_df['Visits'], labels = county_df['County'], autopct ='%1.1f%%', pctdistance = 1.12, labeldistance = 1.28)
  plot.set_title('Percentage of Visits by County')

  canvas = FigureCanvasTkAgg(fig, master=root)
  canvas.get_tk_widget().grid(row = 0, column = 0, columnspan = 15)
  canvas.draw()

  go_back = back_button()
  go_back.grid(row = 1, column = 1)

  save = ttk.Button(root, text="Save Chart", command=lambda: save_chart(fig, f'PercentageofVisitsbyCounty'))
  save.grid(row = 1, column = 6, columnspan = 2)

#Generates visualization for tour types
def tour_types():
  clear_screen()
  root.geometry('1000x625')

  fy_sheets = load_sheets()
  df_dict = load_dictionary(fy_sheets)
  combined_tours_df = combined_tours(df_dict) 

#Develops dataframe for top 5 most popular tour types
  tour_types = combined_tours_df['Type of Tour'].dropna()
  tour_types_df = pd.DataFrame(tour_types.value_counts())
  tour_types_df = tour_types_df.iloc[:5, :] 
  tour_types_df = tour_types_df.reset_index()

  fig = Figure(figsize = (10,6))
  plot = fig.add_subplot(111)  

  plot.bar(tour_types_df.iloc[:, 0], tour_types_df.iloc[:, 1])
  plot.set_xlabel('Tour Types')
  plot.set_ylabel('Count of Tour Types')
  plot.set_title('Top 5 Tour Types')

  canvas = FigureCanvasTkAgg(fig, master=root)
  canvas.get_tk_widget().grid(row = 0, column = 0, columnspan = 15)
  canvas.draw()

  go_back = back_button()
  go_back.grid(row = 1, column = 1)

  save = ttk.Button(root, text="Save Chart", command=lambda: save_chart(fig, f'Top5TourTypes'))
  save.grid(row = 1, column = 6, columnspan = 2)

#Develops filter screen for cumulative revenue visualization
def cumulative_revenue_filter():
  clear_screen()
  root.geometry('650x100')

  fy_sheets = load_sheets()

  years = fy_sheets[5:]
  years.insert(0, 'Select Year')

  tkvar_year = tk.StringVar(root)
  tkvar_year.set(years[0])
  
  year_menu = ttk.Combobox(root, textvariable=tkvar_year, values = [*years])
  year_menu.config(state="readonly")
  year_menu.grid(row = 0, column = 1, padx = 100)

  gen_revenue_button = ttk.Button(root, text="Generate Cumulative Revenue", command=lambda: cumulative_revenue(tkvar_year))
  gen_revenue_button.grid(row = 1, column = 1, padx = 100, pady = 50)

  go_back = back_button()
  go_back.grid(row = 1, column = 0, padx = 100, pady = 50)

#Generates cumulative revenue visualization
def cumulative_revenue(tkvar_year):
  year_select = tkvar_year.get()

  if year_select == 'Select Year':
    year_message()
    return

  clear_screen()
  root.geometry('1000x625')

  fy_sheets = load_sheets()
  df_dict = load_dictionary(fy_sheets)

  revenue_df = pd.DataFrame(df_dict[f'tours{year_select}'], columns = ['Date', 'Total Revenue ($)'])
  revenue_df = revenue_df.dropna()

  fig = Figure(figsize = (10,6))
  plot = fig.add_subplot(111)

  plot.plot(revenue_df['Date'], revenue_df['Total Revenue ($)'].cumsum())

  plot.set_xlabel('Dates')
  plot.set_ylabel('Cumulative Revenue ($)')
  plot.set_title(f'Cumulative Revenue Over Time for FY {year_select}')

  canvas = FigureCanvasTkAgg(fig, master=root)
  canvas.get_tk_widget().grid(row = 0, column = 0, columnspan = 15)
  canvas.draw()

  go_back = back_button_cumulative_filter()
  go_back.grid(row = 1, column = 1)

  save = ttk.Button(root, text="Save Chart", command=lambda: save_chart(fig, f'CumulativeRevenueFY{year_select}'))
  save.grid(row = 1, column = 6, columnspan = 2)

#Generates county groups visualization for FY 2024
def county_groups():
  clear_screen()
  root.geometry('1400x625')

  fy_sheets = load_sheets()
  df_dict = load_dictionary(fy_sheets)
  
  pgcps_count = len(df_dict['tours2024'].loc[:, ['County', 'School Type']][(df_dict['tours2024']['County'] == 'Prince George\'s') & 
                                                                           (df_dict['tours2024']['School Type'] == 'Public')])
  pg_county_private_count = len(df_dict['tours2024'].loc[:, ['County', 'School Type']][(df_dict['tours2024']['County'] == 'Prince George\'s') &
                                                                                        (df_dict['tours2024']['School Type'] == 'Private')])
  outside_pg_county_public_count = len(df_dict['tours2024'].loc[:, ['County', 'School Type']][~(df_dict['tours2024']['County'] == 'Prince George\'s') &
                                                                                               (df_dict['tours2024']['School Type'] == 'Public')])
  outside_pg_county_private_count = len(df_dict['tours2024'].loc[:, ['County', 'School Type']][~(df_dict['tours2024']['County'] == 'Prince George\'s') &
                                                                                                (df_dict['tours2024']['School Type'] == 'Private')])

  pg_county_schools_df = pd.DataFrame({'Categories': ['PGCPS', 'PG County Private Schools', 'Outside PG County Public Schools',
                                                        'Outside PG County Private Schools'],
                                                        'Counts': [pgcps_count, pg_county_private_count, outside_pg_county_public_count, outside_pg_county_private_count]})
  
  fig = Figure(figsize = (14,6))
  plot = fig.add_subplot(111)
  
  plot.bar(pg_county_schools_df['Categories'], pg_county_schools_df['Counts'])

  plot.set_xlabel('Categories')
  plot.set_ylabel('Visits')
  plot.set_title('Visits for Categories of Schools FY 2024')

  canvas = FigureCanvasTkAgg(fig, master=root)
  canvas.get_tk_widget().grid(row = 0, column = 0, columnspan = 15)
  canvas.draw()

  go_back = back_button()
  go_back.grid(row = 1, column = 1)

  save = ttk.Button(root, text="Save Chart", command=lambda: save_chart(fig, f'VisitsforCategoriesofSchools'))
  save.grid(row = 1, column = 6, columnspan = 2)

#Saves a visualization and displays a message when the visualization has been saved
def save_chart(chart, file_name):
  chart.savefig(f'{file_name}.png')
  tkinter.messagebox.showinfo('Save Message', f'Chart has been saved as \"{file_name}\"')

#Clears the screen
def clear_screen():
    for widget in root.winfo_children():
        widget.destroy()

#Develops the start screen
def start_screen():
  clear_screen()
  root.geometry('970x140')
#Adds button for heatmap
  heatmap_button = ttk.Button(root, text="Heatmap", width = 32, command=heatmap)
  heatmap_button.grid(row = 0, column = 0, padx = 20, pady = 20)

#Adds button for visits by year visualization
  visits_by_year_button = ttk.Button(root, text="Visits by Year", width = 32, command=visits_by_year)
  visits_by_year_button.grid(row = 0, column = 1, padx = 20, pady = 20)

#Adds button for top organizations visualization
  top_organizations_button = ttk.Button(root, text="Top Organizations by Visits", width = 32, command=top_10_organizations)
  top_organizations_button.grid(row = 0, column = 2, padx = 20, pady = 20)

#Adds button for breakdown by month visualization
  breakdown_by_month_button = ttk.Button(root, text="Visits Breakdown by Month", width = 32, command=breakdown_by_month_filter)
  breakdown_by_month_button.grid(row = 0, column = 3, padx = 20, pady = 20)

#Adds button for county visits visualization
  county_visits_button = ttk.Button(root, text="Percentage of Visits by County", width = 32, command=county_visits)
  county_visits_button.grid(row = 1, column = 0, padx = 20, pady = 20)

#Adds button for tour types visualization
  tour_types_button = ttk.Button(root, text="Tour Types", width = 32, command=tour_types)
  tour_types_button.grid(row = 1, column = 1, padx = 20, pady = 20)

#Adds button for cumulative revenue visualization
  cumulative_revenue_button = ttk.Button(root, text="Cumulative Revenue", width = 32, command=cumulative_revenue_filter)
  cumulative_revenue_button.grid(row = 1, column = 2, padx = 20, pady = 20)

#Adds button for categories of schools visualization
  categories_of_schools_button = ttk.Button(root, text="Visits for Categories of Schools", width = 32, command=county_groups)
  categories_of_schools_button.grid(row = 1, column = 3, padx = 20, pady = 20)
  
#Develops a back button
def back_button():
  back = ttk.Button(root, text="Go Back", command=start_screen)
  return back

#Develops a back button for the breakdown by month visualization
def back_button_breakdown_by_month_filter():
  back = ttk.Button(root, text="Go Back", command=breakdown_by_month_filter)
  return back

#Develops a back button for the cumulative revenue visualization
def back_button_cumulative_filter():
  back = ttk.Button(root, text="Go Back", command=cumulative_revenue_filter)
  return back

#Displays an error message for the calendar widgets
def calendar_message():
  tkinter.messagebox.showerror('Error Message', 'Start and end dates are in reverse order. Start date must be less than end date')

#Displays an error message message for year dropdown menus
def year_message():
  tkinter.messagebox.showerror('Error Message', 'A year must be selected from the dropdown menu')

root = tk.Tk()
root.title("College Park Aviation Museum Visualizations")

start_screen()

root.minsize(22, 22)

root.maxsize(1800, 1800)

root.mainloop()