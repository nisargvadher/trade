import alpha_vantage
from alpha_vantage.techindicators import TechIndicators
import pandas as pd
import datetime
import pytz


def get_processeddata(nse_symbol):

    ti = TechIndicators(key='BI5WXUMB5FD3XHU1', output_format='pandas')
    data, metadata = ti.get_rsi(symbol=nse_symbol, interval='5min',
                                time_period=14,  series_type='close')

    # Data Processing Begins. Object: data is of DataFrame type & metadata is of dictionary type.

    # Timezone conversion for Object: data
    data.index = data.index.tz_localize(
        'US/Eastern').tz_convert('Asia/Kolkata')
    data.index = data.index.tz_localize(None)

    total_rows = int(len(data.index) - 1)

    def reduce_rows(data):
        for num in range(total_rows, 0, -75):
            yield data.iloc[num]

    data = pd.DataFrame(reduce_rows(data))

    data = data.loc[lambda x: (x['RSI'] > 80) | (x['RSI'] < 20)]

    #  Timezone conversion for Object: Metadata
    str_time_temp = metadata['3: Last Refreshed']
    dt_obj_temp = datetime.datetime.strptime(
        str_time_temp, '%Y-%m-%d %H:%M:%S')
    dt_obj_temp = pytz.timezone('US/Eastern').localize(dt_obj_temp)
    dt_obj_temp = dt_obj_temp.astimezone(pytz.timezone('Asia/Kolkata'))
    dt_obj_temp = dt_obj_temp.replace(tzinfo=None)
    metadata['3: Last Refreshed'] = str(dt_obj_temp)
    metadata['7: Time Zone'] = 'Asia/Kolkata'

    return data, metadata


def export_to_excel(data, metadata):
    filename = str(metadata['1: Symbol'] + '.xlsx')
    filename = filename[4:]
    writer = pd.ExcelWriter(filename, sheet_name='RSI')
    data.to_excel(writer)
    writer.save()
    print('Last Refreshed')
    print(metadata['3: Last Refreshed'])
    print(str('Data written successfully to the Excel File:' + filename))


print("Enter NSE Symbol for which data needs to be generated:")
nse_symbol = str('NSE:' + input())
data, metadata = get_processeddata(nse_symbol)
if(data.empty):
    print("No Data")
else:
    print(data)
    export_to_excel(data, metadata)
