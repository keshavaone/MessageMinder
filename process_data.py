import whatsapp_excel_import as whatsapp_from_excel
import datetime
import pandas as pd
import openpyxl

sheetName = 'Messages'
HOUR = 'Hour'
MINUTE = 'Minute'
AM_PM = 'AM/PM'
DATE = 'Date'
STATUS = 'Status'
reject_message = 'Sent Rejected'
TIME = 'Time'
SENT = 'Sent'
NAME = 'Name'


def load_workbook(file_location):
    return openpyxl.load_workbook(file_location)


def fit(file_location: object) -> object:
    df = whatsapp_from_excel.import_to_df(file_location, sheetName)
    df = whatsapp_from_excel.pre_process_data(df)
    df[DATE] = pd.to_datetime(df[DATE], format='%d-%m-%Y')
    df[DATE] = df[DATE].astype('datetime64')
    df.drop(['Time Stamp'], axis=1, inplace=True)
    today_df = df[df[DATE].dt.date == pd.Timestamp.today().date()]
    today_df = today_df.sort_values(by=[AM_PM, HOUR, MINUTE])
    print(today_df)
    return df, today_df


def check_outstanding_pings(sheet, workbook, today_df, time_out_condition, file_location):
    time_out_df = today_df[time_out_condition]
    print(time_out_df)
    ask_verification = str(input('above data messages are not sent. so, can i send to all the above now?:(y/n) '))
    print('ok')
    if ask_verification.lower() == 'y':
        today_df.loc[time_out_condition, HOUR] = datetime.datetime.now().hour
        today_df.loc[time_out_condition, MINUTE] = datetime.datetime.now().minute
    else:
        for i in time_out_df.index:
            sheet['H' + str(i + 2)].value = reject_message
            today_df.loc[i, STATUS] = reject_message
            sheet['I' + str(int(i) + 2)].value = datetime.datetime.now().strftime('%d-%m-%Y %H:%M')
        workbook.save(file_location)
        print('Updated the Sheet Data')
    return today_df


def second_fit(today_df):
    today_df = today_df[today_df[STATUS].isna()]
    today_df = today_df.sort_values(by=[AM_PM, HOUR, MINUTE])
    today_df[TIME] = today_df[NAME].copy()
    today_df.reset_index(inplace=True)
    i: int
    for i in range(len(today_df)):
        today_df.loc[i, TIME] = '{0}-{1}-{2} {3}:{4}'.format(str(datetime.datetime.now().day),
                                                             str(datetime.datetime.now().month),
                                                             str(datetime.datetime.now().year),
                                                             str(int(today_df[HOUR][i])),
                                                             str(int(today_df[MINUTE][i])))
    assert isinstance(today_df, object)
    today_df[TIME] = pd.to_datetime(today_df[TIME], format='%d-%m-%Y %H:%M')
    return today_df


def get_main_data(today_df, i):
    numbers = [str(i) for i in today_df[today_df[TIME].dt.time == i].Phno]
    indexes = [str(i) for i in today_df[today_df[TIME].dt.time == i]['index']]
    names = [str(i) for i in today_df[today_df[TIME].dt.time == i].Name]
    messages = list(today_df[today_df[TIME].dt.time == i].Message)
    return numbers, indexes, names, messages
