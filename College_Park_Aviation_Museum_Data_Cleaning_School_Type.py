import pandas as pd
import xlwings as xw

#Reading the Excel file
xls = pd.ExcelFile('Tours_data Log_UMD.xlsx')

#Function that creates a dataframe with a school type column
def get_tours_df(FY):
    start_year = FY - 1
    #Read in a sheet
    tours_df = pd.read_excel(xls, f'FY {FY}')
    #Ensure Date column type is datetime
    tours_df['Date'] = pd.to_datetime(tours_df['Date'], errors='coerce')
    #Start on August 25th for each FY
    new_tours_df = tours_df.loc[(tours_df['Date'] > f'{start_year}-08-25')]
    #Get Email column as list and then in a dataframe
    email_list = new_tours_df.Email
    school_type_df = pd.DataFrame({"email": email_list})
    #Split each email address at '@' and get the domain
    school_type_df['Domain'] = school_type_df['email'].str.split('@').str[1]
    school_type_df = school_type_df.drop(columns=['email'])
    #Get Group Type column and remove whitespaces
    school_type_df['Group Type'] = new_tours_df['Group Type']
    school_type_df['Group Type'] = school_type_df['Group Type'].str.strip()
    #Create dictionary for group types and map it to dataframe
    school_type_group_dict = {
        'school': 'Private',
        'School': 'Private',
        'homeschool': 'Private',
        'Homeschool': 'Private'
    }

    school_type_df['School Type Group'] = school_type_df['Group Type'].map(school_type_group_dict)
    #Create dictionary for domains and map it to dataframe
    school_type_domain_dict = {
        'mcpsmd.net': 'Public',
        'pgcps.org': 'Public',
        'k12.dc.gov': 'Public',
        'mcpsmd.org': 'Public',
        'dc.gov': 'Public',
        'aacps.org': 'Public',
        'hcpss.org': 'Public'
    }

    school_type_df['School Type Domain'] = school_type_df['Domain'].map(school_type_domain_dict)
    #Combines the two school type columns
    school_type_df['School Type'] = school_type_df['School Type Domain'].combine_first(school_type_df['School Type Group'])
    school_type_df = school_type_df.drop(['School Type Group', 'School Type Domain'], axis=1)
    #Adds school type column to dataframe representing the Excel sheet
    tours_df['School Type'] = school_type_df['School Type']
    School_Type_col = tours_df.pop('School Type')
    tours_df.insert(tours_df.columns.get_loc('# of tours per month') + 1, 'School Type', School_Type_col)
    return tours_df

#Open the Excel file
app = xw.App()
wb = xw.Book('Tours_data Log_UMD.xlsx')

#Create function to get updated dataframe for each sheet and adds updated dataframes to their respective sheets
def add_to_excel(FY):
    tours_df = get_tours_df(FY)
    ws = wb.sheets[f'FY {FY}']
    ws.range('A1').options(index=False).value = tours_df

FY = 2018
#Loop that calls add_to_excel function to add dataframes to Excel master sheet
while(FY <= 2024):
    add_to_excel(FY)
    FY+=1
#Saves the changes the Excel file
try:
    wb.save()
    wb.close()
    app.quit()
except:
    pass