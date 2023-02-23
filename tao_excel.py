import datetime
from typing import Dict, List
import pandas as pd
import os

# change these variables
excel_file_location = 'C:\\TMP\\import_leerlingen\\codes 23022023.xlsx'
form_date = '01-01-2023'

from_data_toegevoegd = datetime.datetime.strptime(form_date, '%m-%d-%Y').date()


excel = pd.read_excel(excel_file_location)
# create a dict of dataframes
excel_list: dict[str, pd.DataFrame] = {}
# float to string with .0
def float_to_string(value):
    return str(value).split('.')[0]

def float_to_string_two_digits(value):
    return str(value).split('.')[0].zfill(2)

def get_boekje(value):
    # if value <= 12 the return BAO_(value) else return SBO_(value -12)
    if value <= 12:
        return f'BAO_{float_to_string_two_digits(value)}'
    else:
        return f'SBO_{float_to_string_two_digits(value - 12)}'

# filter out rows where column 6 is nan
excel = excel.loc[excel[excel.columns[6]].notna()]
excel = excel.reset_index(drop=True)
# excel = excel[excel['boekje'].notna()]
# list all rows where column 6 is not unique
dup = excel.duplicated(excel.columns[1])
# loop dup with index
duplicate_list = []
for index, value in dup.iteritems():
    if value == True:
        # print row where column 6 is not unique
        duplicate_student = float_to_string(excel.loc[index][excel.columns[1]])
        if not duplicate_student in duplicate_list:
            print(duplicate_student)
            duplicate_list.append(duplicate_student)
duplicate_count = duplicate_list.__len__()
if duplicate_count > 0:
    print(f'Found {duplicate_count} duplicate students, stopping script')
else:
    for index, row in excel.iterrows():
        if index > 0:
            ww = row[0]
            ll_code = float_to_string(row[1])
            school = float_to_string(row[2])
            klas_id = float_to_string(row[3])
            school_naam = row[4]
            klas_naam = row[5]
            boekje = row[6]
            datum_toegevoegd = row[7]
            
            if not pd.isna(boekje) and boekje != pd.NaT and datetime.datetime(datum_toegevoegd.year, datum_toegevoegd.month, datum_toegevoegd.day).date() >= from_data_toegevoegd:
                boekje = get_boekje(boekje)
                if boekje not in excel_list:
                    excel_list[boekje] = pd.DataFrame()
                
                new_row = {'Label': ll_code, 'Group': boekje, 'Lastname': klas_naam,  'Password': ww,'Login': ll_code, 'SchoolId': school, 'ClassId': klas_id, 'Firstname': school_naam, 'Booklet': boekje}
                # add row to dataframe
                new_df = pd.DataFrame(new_row, index=[0])
                excel_list[boekje] = pd.concat([excel_list[boekje], new_df], ignore_index=True)

    for boekje in excel_list:
        # string with current yearmonthday with leading zeros
        yearmonthday = datetime.datetime.now().strftime('%Y%m%d')
        # get directory form excel file location
        directory = os.path.dirname(excel_file_location)
        excel_list[boekje].to_csv(os.path.join(directory, f'boekje_{boekje}_{yearmonthday}.csv'), index=False, sep=';')
