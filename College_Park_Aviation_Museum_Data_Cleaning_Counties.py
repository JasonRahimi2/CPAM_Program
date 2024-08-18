import pandas as pd
import xlwings as xw

#Reading the Excel file
xls = pd.ExcelFile('Tours_Data Log_UMD.xlsx')

#Creating a class to instantiate objects where each object represents an Excel sheet
class excel_sheets:
    def __init__(self, FY):
        self.WS = pd.read_excel(xls, f'FY {FY}')

#Creating a dictionary to store dataframes
tours_df_dict = {}
for i in range(2018,2025):
    tours_df_dict[f'tours{i}'] = excel_sheets(i)

#Concatenating dataframes
concat_df = []
for i in range(2018,2025):
    concat_df.append(tours_df_dict[f'tours{i}'].WS)
combined_tours_df = pd.concat(concat_df)

#Making cities easier to read
combined_tours_df.City = combined_tours_df.City.str.strip()
combined_tours_df.City = combined_tours_df.City.str.lower()

#Creating a dictionary to connect each city with a county 
city_dict = {
    'silver spring': 'Montgomery',
    'rockville': 'Montgomery',
    'washington': 'DC',
    'college park': 'Prince George\'s',
    'beltsville': 'Prince George\'s',
    'bowie': 'Prince George\'s',
    'lanham': 'Prince George\'s',
    'gaithersburg': 'Montgomery',
    'wheaton': 'Montgomery',
    'bethesda': 'Montgomery',
    'takoma park': 'Montgomery',
    'clinton': 'Prince George\'s',
    'odenton': 'Anne Arundel',
    'greenbelt': 'Prince George\'s',
    'hyattsville': 'Prince George\'s',
    'mitchellville': 'Prince George\'s',
    'severna park': 'Anne Arundel',
    'columbia': 'Howard',
    'annapolis': 'Anne Arundel',
    'edgewater': 'Anne Arundel',
    'kensington': 'Montgomery',
    'sandy spring': 'Montgomery',
    'olney': 'Montgomery',
    'potomac': 'Montgomery',
    'riverdale': 'Prince George\'s',
    'severn': 'Anne Arundel',
    'baltimore': 'Baltimore',
    'crofton': 'Anne Arundel',
    'glenn dale': 'Prince George\'s',
    'ellicott city': 'Howard',
    'clarksville': 'Howard',
    'university park': 'Prince George\'s',
    'adelphi': 'Prince George\'s',
    'pasadena': 'Anne Arundel',
    'washington dc': 'DC',
    'camp springs': 'Prince George\'s'
}

#Changing abbreviations, fix casing
county_fix_dict = {
    'Prince George\'S': 'Prince George\'s',
    'Pg': 'Prince George\'s',
    'Dc': 'DC',
    'Mc': 'Montgomery',
    'Aa': 'Anne Arundel',
    'Hc': 'Howard',
    'Bc': 'Baltimore',
    'Cc': 'Carroll'
}

#Changing abbreviations, fix casing
city_fix_dict = {
    'Nw': 'NW',
    'Ne': 'NE',
    'Sw': 'SW',
    'Se': 'SE'
}
#Opening the Excel file
try: 
    app = xw.App()
    wb = xw.Book('Tours_Data Log_UMD.xlsx')
    
    #Loop to clean dataframes, to map cities to counties in each dataframe, then loading each dataframe to their respective sheet
    for i in range(2018,2025):
        tours_df_dict[f'tours{i}'].WS.City = tours_df_dict[f'tours{i}'].WS.City.str.strip()
        tours_df_dict[f'tours{i}'].WS.City = tours_df_dict[f'tours{i}'].WS.City.str.lower()
        tours_df_dict[f'tours{i}'].WS.County = tours_df_dict[f'tours{i}'].WS.County.fillna(tours_df_dict[f'tours{i}'].WS.City.map(city_dict))
        tours_df_dict[f'tours{i}'].WS.County = tours_df_dict[f'tours{i}'].WS.County.str.title()
        tours_df_dict[f'tours{i}'].WS.County = tours_df_dict[f'tours{i}'].WS.County.replace(county_fix_dict)
        tours_df_dict[f'tours{i}'].WS.City = tours_df_dict[f'tours{i}'].WS.City.str.title()
        tours_df_dict[f'tours{i}'].WS.City = tours_df_dict[f'tours{i}'].WS.City.replace(city_fix_dict)
        ws = wb.sheets[f'FY {i}']
        ws.range('A1').options(index=False).value = tours_df_dict[f'tours{i}'].WS
    
    #Saving the Excel master sheet
    wb.save()
    wb.close()
    app.quit()
except:
    pass
