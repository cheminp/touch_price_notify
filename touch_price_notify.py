import twstock
import time
import datetime
import requests
import pandas as pd
from itertools import cycle
from pandas import DataFrame
from colorama import init, Fore, Back


# LINE Notify token  (Plz ref https://officeguide.cc/python-line-notify-send-messages-images-tutorial-examples/ to get the token)
token = 'your_line_token'

#------------------ time setting -------------------
now = datetime.datetime.now()

program_stop_time = "13:30:00"
my_datetime = datetime.datetime.strptime(program_stop_time, "%H:%M:%S")

#suppose that the date must be the same as now
my_datetime = now.replace(hour=my_datetime.time().hour, minute=my_datetime.time().minute, second=my_datetime.time().second, microsecond=0)

#---------------------------------------------------

data = pd.read_excel("alarm.xlsx", sheet_name="工作表1",dtype = str)

is_first = int(input('Are you first time executing this program(0 is False / 1 is True): '))
if is_first == 1:
    data['ticker_symbol'] = None
    data['name'] = None
    data['target_price'] = None
    data['current_price'] = None

print('current track Inf is:')
print(data)
print('-----------------------------')
print('-----OPERATION---GUIDE-------')
print('-----------------------------')
print('---No need to modify: 0------')
print('---Add a new target: 1-------')
print('---Change a target price: 2--')
print('---Delete the target: 3------')
print('-----------------------------')
print('-----------------------------')
is_add_or_change_a_target = int(input('Plz input the operation: '))

while is_add_or_change_a_target > 0:
    match is_add_or_change_a_target:
        case 1: #add a new target
            input_ticker_symbol = input("Plz input the ticker_symbol you want to search: ")
            stock = twstock.realtime.get(input_ticker_symbol)
            input_target_price = input(f"Plz input the {input_ticker_symbol} target_price: ")
            row_to_add = pd.DataFrame({'ticker_symbol':[input_ticker_symbol],'name':stock['info']['name'],'target_price':[input_target_price]})
            data = pd.concat([data,row_to_add],ignore_index = True)
        case 2: #change a old target price
            input_change_price_symbol = input("Plz input the symbol you want to change it's price: ")
            input_new_price = input(f"Plz input {input_change_price_symbol} the new price: ")
            print(f"{type(data.index[data['ticker_symbol'] == input_change_price_symbol].tolist()[0])}")
            data.at[data.index[data['ticker_symbol'] == input_change_price_symbol].tolist()[0],'target_price'] = input_new_price
        case 3: #delete the target
            input_to_delete_ticker_symbol = input('Plz input the ticker symbol that you want to delete: ')
            data.drop(data.index[data['ticker_symbol'] == input_to_delete_ticker_symbol].tolist(),inplace = True)
            data = data.reset_index(drop = True)  #use this function to reset the index

    print(f'change data {data}')
    print('-----------------------------')
    print('-----OPERATION---GUIDE-------')
    print('-----------------------------')
    print('---No need to modify: 0------')
    print('---Add a new target: 1-------')
    print('---Change a target price: 2--')
    print('---Delete the target: 3------')
    print('-----------------------------')
    print('-----------------------------')
    is_add_or_change_a_target = int(input('Do you want to add or change a new target(0 is False 1 is add a new target 2 is change a target price): '))

print('Below is the save data')
print(data)
DataFrame(data).to_excel('alarm.xlsx', sheet_name='工作表1', index=False, header=True)

col_one_list = data['ticker_symbol'].tolist()   #get a list from the column
print(f'col_one_list is {col_one_list}')


sleep_time = 3 #default sleep time 3s  
minutes_tmp = 0

while 1:
    now = datetime.datetime.now()
    
    if(now >= my_datetime):
        print(f"----------{program_stop_time}----------")
        break
    else:
        stocks = twstock.realtime.get(col_one_list)
        print(stocks)

        count = 0
        for item in col_one_list:
            print(f"item is {item}({stocks[item]['info']['name']})")
            print(f"{item}'s higher best_bid_price {stocks[item]['realtime']['best_bid_price'][0]}")
            print(f"{item}'s latest_trade_price {stocks[item]['realtime']['latest_trade_price']}")
            print(f"{item}'s target_price is {data.at[count,'target_price']}/Is_arrive_the_target:{float(stocks[item]['realtime']['best_bid_price'][0]) >= float(data.at[count,'target_price'])}")
            data.at[count,'current_price'] = stocks[item]['realtime']['best_bid_price'][0]  #set the current price as the higher best_bid_price

           
            if float(stocks[item]['realtime']['best_bid_price'][0]) >= float(data.at[count,'target_price']):
                
                if (minutes_tmp != now.minute):  #send a message per minute
                    # to send message
                    message = f"\nTicker symbol {item}({stocks[item]['info']['name']}): Touch the target price {data.at[count,'target_price']}"

                    # HTTP 標頭參數與資料
                    headers = { "Authorization": "Bearer " + token }
                    send_message = { 'message': message }

                    # 以 requests 發送 POST 請求
                    requests.post("https://notify-api.line.me/api/notify",headers = headers, data = send_message)
               

            count += 1
        
        time.sleep(sleep_time)
        minutes_tmp = now.minute
 
print('Below is the save data')
print(data)
DataFrame(data).to_excel('alarm.xlsx', sheet_name='工作表1', index=False, header=True)  #final write to alarm.xlsx