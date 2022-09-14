import whatsapp_base as whatsapp
import datetime
import process_data

total_messages_sent = 0
file_location: str = r"C:\Users\iamke\OneDrive\Desktop\AutoMessages.xlsx"


def status_update() -> object:
    if status_check[-1]:
        if len(status_check[1]) > 0:
            print('Communications Not Sent to ', status_check[1])
            for y in status_check[1]:
                indexes.remove(str(today_df.index[y[0]]))
        else:
            pass
    else:
        raise Exception('Please Login to WhatsApp!!!')
    for j in indexes:
        sheet['H' + str(int(j) + 2)].value = process_data.SENT
        sheet['I' + str(int(j) + 2)].value = str(today_df.loc[today_df[today_df['index'] == int(j)].
                                                 index[0], process_data.TIME])
        today_df.loc[today_df[today_df['index'] == int(j)].index, process_data.STATUS] = process_data.SENT
    print(today_df)
    workbook.save(file_location)
    return today_df[today_df[process_data.STATUS].isna()]


if __name__ == '__main__':
    workbook, sheet = process_data.load_workbook(file_location)
    df, today_df = process_data.fit(file_location)
    assert isinstance(today_df, object)
    time_out_condition = today_df[process_data.STATUS].isna() & (
            today_df[process_data.HOUR].le(datetime.datetime.now().hour) &
            today_df[process_data.MINUTE].le(datetime.datetime.now().minute) | (today_df[process_data.HOUR].isna() &
                                                                                today_df[process_data.MINUTE].isna()))
    if len(today_df[time_out_condition]) > 0:
        today_df = process_data.check_outstanding_pings(sheet, workbook, today_df, time_out_condition,
                                                        file_location)  # re verifying the timely messages are missed
        # out or not
    today_df = process_data.re_ordering_data(today_df)  # sorting the data by AM/PM, Hour, Minute
    while 0 < len(today_df[today_df[process_data.STATUS].isna()]):
        for i in today_df[today_df[process_data.STATUS].isna()][process_data.TIME].dt.time:
            try:
                if max(today_df.Time.dt.time).hour >= datetime.datetime.now().hour:
                    print(
                        "Waiting for... {0}:00 to equal {1}".format(str(datetime.datetime.now().strftime('%H:%M')),
                                                                    str(i)), end='\r')

                    if str(i) == str(datetime.datetime.now().strftime('%H:%M')) + ':00':
                        numbers, indexes, names, messages = process_data.get_main_data(today_df, i)
                        print('Sending Message/s to', ', '.join(names))
                        status_check = whatsapp.whatsapp_communicate(numbers, messages)
                        today_df = status_update()
                        total_messages_sent += len(indexes)
                else:
                    break
            except ValueError:
                print('WooHooH!!!!, Done with Messaging Today.\nTotal Messages Sent: ', total_messages_sent)
                break
