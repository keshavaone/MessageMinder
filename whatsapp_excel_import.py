import pandas as pd
import datetime

HOUR = 'Hour'
MINUTE = 'Minute'
AM_PM = 'AM/PM'


def import_to_df(file_location, sheet_name):
    return pd.read_excel(file_location, sheet_name)


def pre_process_data(current_df):

    for i in range(len(current_df)):
        if current_df.isnull().loc[i, HOUR]:
            current_df.loc[i, HOUR] = datetime.datetime.now().hour
            current_df.loc[i, MINUTE] = datetime.datetime.now().minute

        if current_df.isnull().loc[i, MINUTE]:
            current_df.loc[i, MINUTE] = 0

        if current_df.isnull().loc[i, AM_PM]:
            if current_df.loc[i, HOUR] <= 23:
                if current_df.loc[i, HOUR] >= 12:
                    current_df.loc[i, AM_PM] = 'PM'
                else:
                    current_df.loc[i, AM_PM] = 'AM'
        if current_df.loc[i,AM_PM]:
            if current_df.loc[i, AM_PM] == 'PM' and current_df.loc[i, HOUR] < 12:
                current_df.loc[i, HOUR] = current_df.loc[i, HOUR] + 12
    return current_df
