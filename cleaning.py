import pandas as pd
import pyxlsb

file_path = '230421 Trans Tasman Data .xlsb'

# input years like 2022 and 2023
def filter_date(start: str, end: str, tocsv: bool = False):

    with pyxlsb.open_workbook(file_path) as workbook:
        with workbook.get_sheet('Data') as sheet:
            data = [[item.v for item in row] for row in sheet.rows()]

    df = pd.DataFrame(data[1:], columns=data[0])

    df['Time series'] = pd.to_datetime(df['Time series'], format='%Y-%m')

    start_date = pd.to_datetime(start)
    end_date = pd.to_datetime(end)
    filtered_df = df[(df['Time series'] >= start_date) & (df['Time series'] < end_date)]
    if tocsv:
        filtered_df.to_csv('filtered_date.csv', index=False)
    filtered_df = filtered_df.reset_index(drop=True)
    return filtered_df

def filter_operation_days(df: pd.DataFrame, day_in_week: str, tocsv: bool = False):
    filtered_df = df[df['Op Days'].str.contains(day_in_week)]
    if tocsv:
        filtered_df.to_csv('filtered_operation_days.csv', index=False)
    filtered_df = filtered_df.reset_index(drop=True)
    return filtered_df

'''
Define times:
Morning: 6am - 12pm
Afternoon: 12pm - 6pm
Evening: 6pm - 12am
Night: 12am - 6am
'''
# param time as defined above, using departure time
def filter_time_of_day(df: pd.DataFrame, time: str, tocsv: bool = False):
    if time == 'morning':
        filtered_df = df[(df['Departure Time'] >= '0600') & (df['Departure Time'] < '1200')]
    elif time == 'afternoon':
        filtered_df = df[(df['Departure Time'] >= '1200') & (df['Departure Time'] < '1800')]
    elif time == 'evening':
        filtered_df = df[(df['Departure Time'] >= '1800') & (df['Departure Time'] < '2400')]
    elif time == 'night':
        filtered_df = df[(df['Departure Time'] >= '0000') & (df['Departure Time'] < '0600')]
    else:
        raise ValueError('Time not defined')
    if tocsv:
        filtered_df.to_csv('filtered_time_of_day.csv', index=False)
    filtered_df = filtered_df.reset_index(drop=True)
    return filtered_df

'''
airlines:
NZ
QF
VA
JQ
CI
LA
EK
'''
def filter_airline(df: pd.DataFrame, airline: str, tocsv: bool = False):
    filtered_df = df[df['Published Carrier'] == airline]
    if tocsv:
        filtered_df.to_csv('filtered_airline.csv', index=False)
    filtered_df = filtered_df.reset_index(drop=True)
    return filtered_df
    

filter_df = filter_date('2022', '2023')
filter_df = filter_operation_days(filter_df, '6')
filter_df = filter_time_of_day(filter_df, 'morning')
filter_df = filter_airline(filter_df, 'NZ')
print(filter_df)