import pandas as pd
import datetime

# import from excel
data = pd.read_excel('nqtt_actions_ncc.xlsx', sheet_name='nqtt_011021')
data = data.sort_values(by=['starttime'])
data = data.reset_index()
# add additional columns
data['duration'] = pd.Series(dtype='datetime64[ns]')
data['endtime'] = pd.Series(dtype='datetime64[ns]')
data['duration_new'] = pd.Series(dtype='datetime64[ns]')

# change minut timedifference to 'h:m:s' format
for i in data.index:
    data['duration'][i] = datetime.timedelta(seconds=int(data['timedifference'][i])*60)
# calculate endtime
for i in data.index:
    data['endtime'][i] = data['starttime'][i] + data['duration'][i]

startdate = datetime.datetime.date(data.iloc[0,4]) - datetime. timedelta(days=1)
enddate = datetime.datetime.date(data.iloc[-1,4]) + datetime. timedelta(days=1)

#create not working time periods dataframe
start = []
fin = []
current_date = startdate
while current_date <= enddate:
    if current_date.weekday() in [0,1,2,3,4]:
        start.append(datetime.datetime.combine(current_date, datetime.time(hour=0, minute=0, second=0)))
        fin.append(datetime.datetime.combine(current_date, datetime.time(hour=8, minute=59, second=59)))
        start.append(datetime.datetime.combine(current_date, datetime.time(hour=18, minute=0, second=1)))
        fin.append(datetime.datetime.combine(current_date, datetime.time(hour=23, minute=59, second=59)))
    else:
        start.append(datetime.datetime.combine(current_date, datetime.time(hour=0, minute=0, second=0)))
        fin.append(datetime.datetime.combine(current_date, datetime.time(hour=23, minute=59, second=59)))
    current_date = current_date + datetime.timedelta(days=1)

list_of_tuples = list(zip(start, fin))
df_notworking = pd.DataFrame(list_of_tuples, columns = ['start', 'finish'])

#remove notworking time periods from durations
for i in data.index:
    data['duration_new'][i] = data['duration'][i]
    for j in df_notworking.index:
        if df_notworking['start'][j] > data['endtime'][i]:
            break
        elif df_notworking['finish'][j] < data['starttime'][i]:
            continue
        if data['starttime'][i] <= df_notworking['start'][j] and data['endtime'][i] >= df_notworking['finish'][j]:
            data['duration_new'][i] = data['duration_new'][i] - (df_notworking['finish'][j] - df_notworking['start'][j])
        elif data['starttime'][i] >= df_notworking['start'][j] and data['endtime'][i] <= df_notworking['finish'][j]:
            data['duration_new'][i] = data['duration_new'][i] - data['duration_new'][i]
        elif data['starttime'][i] >= df_notworking['start'][j] and data['endtime'][i] >= df_notworking['finish'][j]:
            data['duration_new'][i] = data['duration_new'][i] - (df_notworking['finish'][j] - data['starttime'][i])
        elif data['starttime'][i] <= df_notworking['start'][j] and data['endtime'][i] <= df_notworking['finish'][j]:
            data['duration_new'][i] = data['duration_new'][i] - (data['endtime'][i] - df_notworking['start'][j])

data.to_excel('ncc_nqtt.xlsx', index=False)