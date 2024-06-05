from utils.bonds_lib import *
from bs4 import BeautifulSoup
from enum import Enum
import requests

for bd_idx in range (bonds_num):
    web = requests.get(f'{list_urls[bd_idx]}')
    if web.status_code == 403: ERROR(WebException, f"請檢察'{list_name[bd_idx]}'網頁連結正確 {list_urls[bd_idx]}")
    soup = BeautifulSoup(web.text, "html.parser")
    table_30days=soup.find_all('table')
    print(f'Capture {list_name[bd_idx]} data Start ...')

    # 
    # Found table contain '近30日淨值'
    # Next 3 table will be data
    # 
    # output:
    #   table_30ays will found title table index
    #   and next 3 tables would be bond's info 
    # 
    try:
        for i in range (len(table_30days)):
            foundBondTitleName = TITLE_NAME in table_30days[i].find('td').text
            if foundBondTitleName:
                TALBE_30DAYS_TITLE_IDX=i
                TALBE_30DAYS_1ST_IDX = TALBE_30DAYS_TITLE_IDX+1
                TALBE_30DAYS_2ND_IDX = TALBE_30DAYS_TITLE_IDX+2
                TALBE_30DAYS_3RD_IDX = TALBE_30DAYS_TITLE_IDX+3
                list_table = [TALBE_30DAYS_1ST_IDX, TALBE_30DAYS_2ND_IDX, TALBE_30DAYS_3RD_IDX]
                break
        table_30days=soup.find_all('table')[TALBE_30DAYS_TITLE_IDX]
        assert(TITLE_NAME in table_30days.find('td').text and list_name[bd_idx] in table_30days.find('td').text)
        DEBUG(f'Choose title idx: {TALBE_30DAYS_TITLE_IDX}, table list idx: {list_table}')
        DONE(f'Found title table')
    except AssertionError:
        ERROR(AssertionError, f'DOUBLE CHECK TITLE NOT CORRECT: {TALBE_30DAYS_TITLE_IDX}')
    
    # 
    # Catch data from web
    # 1st td should be '日期'
    # 2nd td should be '淨值'
    # class contain t3n0 should be date
    # class contain t3n1 should be value
    # 
    # output:
    #  dates: bonds date's from website
    #  val:   bonds value's from website
    # 
    try:
        for list_idx in list_table:
            table_30days=soup.find_all('table')[list_idx]
            td_30days=table_30days.find_all('td')
            for idx in range (len(td_30days)):
                td=td_30days[idx]
                if   idx==0: assert(dict_idx[idx] in td.text)
                elif idx==1: assert(dict_idx[idx] in td.text)
                elif "t3n0" in str(td): dates.append(td.text)
                elif "t3n1" in str(td): val.append(f'{float(td.text):.3f}')
                else: raise ValueError
        DONE(f'Collect date and value')
    except AssertionError:
        ERROR(AssertionError, f'LIST IDX: {idx} IS NOT {dict_idx[idx]}')
    except ValueError:
        ERROR(ValueError, f'TABLE DATA: CANNOT FOUND -> {td}')
    
    # 
    # Assert length is same
    # Append date & value into list
    # 
    try:
        assert(len(dates) == len(val))
        list_date.append(dates)
        list_val.append(val)
    except AssertionError:
        ERROR(AssertionError, f'Length \'date\' and \'val\' is not same')
    finally:
        DEBUG(f'date: {dates}\n  val:  {val}')
        DONE(f'Capture {list_name[bd_idx]} data Start', end='\n')
        dates, val = [], []

# return dict
retDict = {
    'bondsName': list_name,
    'bondsURL': list_urls,
    'bondsDates': list_date,
    'bondsVals': list_val,
    'bondsRowNum': list_rowNum,
    'bondsRowSMA': list_rowSMA,
}

class bondsEnum(Enum):
    name = retDict['bondsName']
    URL = retDict['bondsURL']
    dates = retDict['bondsDates']
    vals = retDict['bondsVals']
    rowNum = retDict['bondsRowNum']
    rowSMA = retDict['bondsRowSMA']