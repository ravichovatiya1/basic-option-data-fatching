import requests
import pandas as pd
import json
import math
import xlsxwriter
import pytz
from datetime import datetime

# data
headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36',
            'accept-language': 'en,gu;q=0.9,hi;q=0.8',
            'accept-encoding': 'gzip, deflate, br'}

index_url = "https://www.nseindia.com/api/option-chain-indices?symbol=BANKNIFTY"
url_oc      = "https://www.nseindia.com/option-chain"
url_indices = "https://www.nseindia.com/api/allIndices"

sess = requests.Session()
cookies = dict()

def round_nearest(x,num=50): return int(math.ceil(float(x)/num)*num)
def nearest_strike_bnf(x): return round_nearest(x,100)
def nearest_strike_nf(x): return round_nearest(x,50)


# Local methods
def set_cookie():
    try:
        request = sess.get(url_oc, headers=headers, timeout=5)
        cookies = dict(request.cookies)
    except requests.exceptions.Timeout:
        print("Timeout occurred")
    
def get_data(url):
    set_cookie()
    response = sess.get(url, headers=headers, timeout=5, cookies=cookies)
    if(response.status_code==401):
        set_cookie()
        response = sess.get(url_nf, headers=headers, timeout=5, cookies=cookies)
    if(response.status_code==200):
        return response.text
    return ""


response_text = get_data(url_indices)
data = json.loads(response_text)

# get last trade price of option chain(underlying price or current price)
for index in data["data"]:
    
    if index["index"]=="NIFTY 50":
        #underlying price of nifty
        nf_ul = index["last"]
        print("nifty",nf_ul,',',round(nf_ul))
        
    if index["index"]=="NIFTY BANK":
        #underlying price of banknifty
        bnf_ul = index["last"]
        print("banknifty",bnf_ul,',',round(bnf_ul))
        
bnf_nearest=nearest_strike_bnf(bnf_ul)
nf_nearest=nearest_strike_nf(nf_ul)
        
print('-------------------------------------------------------------\n')

# option chain data
data = None
response_text = get_data(index_url)
data = json.loads(response_text)
excel_data=[]
currExpiryDate = data["records"]["expiryDates"][0]
call_put_bnf_nearest_counter = 0
for item in data['records']['data']:
    if item["expiryDate"] == currExpiryDate:
        try:
            if item['strikePrice'] == bnf_nearest:
                nearest_seperate_price = item['strikePrice']
                nearest_count_for_banknifty_background_color = call_put_bnf_nearest_counter
            row_data= [
                    item['CE']['openInterest'],
                    item['CE']['changeinOpenInterest'],
                    item['CE']['totalTradedVolume'],
                    item['CE']['impliedVolatility'],
                    item['CE']['lastPrice'],
                    item['CE']['change'],
                    item['CE']['bidprice'],
                    item['CE']['askPrice'],
                    item['strikePrice'],
                    item['PE']['bidprice'],
                    item['PE']['askPrice'],
                    item['PE']['change'],
                    item['PE']['lastPrice'],
                    item['PE']['impliedVolatility'],
                    item['PE']['totalTradedVolume'],
                    item['PE']['changeinOpenInterest'],
                    item['PE']['openInterest']]
            excel_data.append(row_data)
            call_put_bnf_nearest_counter += 1
        except:
            pass

# making excel format
columns =pd.MultiIndex.from_tuples(zip(['CALL','CALL','CALL','CALL','CALL','CALL','CALL','CALL',
          'EXP: '+str(currExpiryDate),
          'PUT','PUT','PUT','PUT','PUT','PUT','PUT','PUT'],
          ['CE_OI','CE_CHANGE_IN_OI','CE_VOLUME','CE_IV','CE_LTP','CE_CHNG','CE_BID_PRICE','CE_ASK_PRICE',
          'STRIKE_PRICE',
          'PE_BID_PRICE','PE_ASK_PRICE','PE_CHNG','PE_LTP','PE_IV','PE_VOLUME','PE_CHANGE_IN_OI','PE_OI']))
    
df = pd.DataFrame(excel_data,columns= columns)


# time zone settings
tz_NY = pytz.timezone('Asia/Kolkata')   
datetime_NY = datetime.now(tz_NY)  
dt_string = str(datetime_NY.strftime("%Y%m%d%H%M%S"))


sheet_name = 'Sheet1'
writer = pd.ExcelWriter(dt_string+'.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name=sheet_name,index=True,header=[0, 1],freeze_panes=(3,1),startrow=1,float_format = "%0.2f",encoding='utf-8')
writer.sheets[sheet_name].set_row(3, None, None, {'hidden': True})

workbook = writer.book
worksheet = writer.sheets[sheet_name]



#header settings
header_format = workbook.add_format({
    'align': 'center',
    'font_name': 'Roboto', 
    'font_size': 9, 
    'color' : 'white', 
    'fg_color': '#3A2D7D',
    'bold': False, 
    'border' : 1})


for col_num, value in enumerate(df):
    worksheet.write(1, col_num + 1, value[0], header_format)
    worksheet.write(2, col_num + 1, value[1], header_format)
    

# marge first line
merge_format = workbook.add_format({'align': 'center','font_name': 'Roboto', 'font_size': 10,'fg_color': '#D7E4BC','bold': False, 'border' : 1})

# datetime and price sheet formating
context_dt_string = str(datetime_NY.strftime("%d-%b-%Y- %H:%M:%S"))
context_to_display = "Underlying Index: BANKNIFTY: ( {} ) As on {} IST".format(bnf_ul,context_dt_string)
worksheet.merge_range('B1:R1', context_to_display, merge_format)


# making the text center of all excel
new_format = workbook.add_format()
new_format.set_align('center')
worksheet.set_column('A:R', 9, new_format)

bule_color_ltp_strike_format = workbook.add_format({
    'font_name': 'Roboto',
    'border' : 0,
    'align':'center',
    'font_size': 9, 
    'color' : '#295C89', 
    'fg_color': 'white',
    'bold': False, })


worksheet.set_column('C:C',12)
worksheet.set_column('J:J',15,bule_color_ltp_strike_format)
worksheet.set_column('Q:Q',12)
worksheet.set_column('F:F',9,bule_color_ltp_strike_format)
worksheet.set_column('N:N',9,bule_color_ltp_strike_format)


# making condition if chg is < 0 or > 0
df.columns = df.columns.droplevel(0)
shape = df.shape
print(nearest_seperate_price)

red_color_chg_strike_format = workbook.add_format({
    'font_name': 'Roboto',
    'border' : 0,
    'align':'center',
    'font_size': 9, 
    'color' : '#cb0505', 
    'fg_color': 'white',
    'bold': False, })

green_color_chg_strike_format = workbook.add_format({
    'font_name': 'Roboto',
    'border' : 0,
    'align':'center',
    'font_size': 9, 
    'color' : '#007a00', 
    'fg_color': 'white',
    'bold': False, })

call_put_in_money_background_color_format = workbook.add_format({
    'font_name': 'Roboto',
    'border' : 0,
    'align':'center',
    'font_size': 9,  
    'bg_color': '#D7E4BC',
    'bold': False, })


worksheet.conditional_format('G4:G'+str(shape[0]+4), {'type': 'cell',
                                       'criteria': '<',
                                       'value': 0,
                                       'format': red_color_chg_strike_format,})

worksheet.conditional_format('G4:G'+str(shape[0]+4), {'type': 'cell',
                                       'criteria': '>',
                                       'value': 0,
                                       'format': green_color_chg_strike_format,})

worksheet.conditional_format('M4:M'+str(shape[0]+4), {'type': 'cell',
                                       'criteria': '<',
                                       'value': 0,
                                       'format': red_color_chg_strike_format,})

worksheet.conditional_format('M4:M'+str(shape[0]+4), {'type': 'cell',
                                       'criteria': '>',
                                       'value': 0,
                                       'format': green_color_chg_strike_format,})

worksheet.conditional_format('B4:I'+str(nearest_count_for_banknifty_background_color+4), {'type': 'no_blanks',
                                       'format': call_put_in_money_background_color_format,})

worksheet.conditional_format('K'+str(nearest_count_for_banknifty_background_color+5)+':'+'R'+str(shape[0]+4), {'type': 'no_blanks',
                                       'format': call_put_in_money_background_color_format,})


# border format code 
border_fmt = workbook.add_format({'bottom':1, 'top':0, 'left':1, 'right':0})
worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(df)+3, len(df.columns)), {'type': 'no_errors', 'format': border_fmt})


df =df.replace(0,'-')


# file saving
writer.save()