from bs4 import BeautifulSoup
import requests
import os
from utils import *
from datetime import datetime


USD_URL=f'https://hk.investing.com/currencies/usd-twd-historical-data'
# sys.stdout=open(f'{os.getcwd()}\\log.txt', 'w')

def writeTDCCUSDSellPrice(price:float):
    wb, ws = open_workbook(sheet_name=SHEET_SUMMARY_NAME)
    USDSellColLetter='L'
    row1_data = '美元匯率'
    L1, L2 = USDSellColLetter+'1', USDSellColLetter+'2'
    L1_val = returnVal(ws[L1])
    L2_val = returnVal(ws[L2]) # sell price
    try:
        assert(L1_val == row1_data)
        if float(L2_val) == price:
            DONE(f'{L2} == {price}, 不更新')
        else:
            writeVal(ws[L2], price)
            wb.save(file_name)
            DONE(f'{L2_val} 取代成 {price}')
        wb.close()
    except AssertionError:
        ERROR(AssertionError, f'{L1} 不是 {row1_data}')

def formula_SMA(letterCol_bond:str, row_num:int):
    return f"=AVERAGE({letterCol_bond}{row_num}:{(letterCol_bond)}{row_num+4})"

def getCurDateIdx(ws, max_rows, list_bond_date:list): 
    for excel_idx in range (2, max_rows+1):
        col_A = str_A
        getDate = returnVal(ws[col_A+str(excel_idx)])
        getDateNotNone = getDate is not None
        bond_lastDate_MMDD = list_bond_date[0]
        if getDateNotNone: 
            if type(getDate) is not str: sameYear = getDate.strftime("%Y") == date.today().strftime("%Y")
            else: break
        # bond_lastDate_MMDD = date.today().strftime("%m/%d")
        if sameYear and getDateNotNone:
            getDateMMDD = getDate.strftime("%m/%d") 
            if getDateMMDD == bond_lastDate_MMDD:
                return excel_idx
    return 2

def checkWeekendLast5day(ws, last_idx):
    for i in range(last_idx, last_idx+5):
        dt=returnVal(ws[str_A+str(i)])
        if not isinstance(dt, datetime): dt=datetime.strptime(dt, "%Y/%m/%d")
        if checkIsSaturdyOrSunday(dt): ERROR(DateException, f'{dt.strftime("%Y/%m/%d")} is Saturday or Sunday')

def checkIsSaturdyOrSunday(dt):
    isSaturday=dt.isoweekday() == 6
    isSunday=dt.isoweekday() == 7
    return isSaturday or isSunday

if __name__ == '__main__':
    # 
    # Open excel data sheet
    # 
    wb, ws = open_workbook(sheet_name=SHEET_DATA_NAME)
    max_row=ws.max_row
    max_col=ws.max_column

    # 
    # Update title name
    # A1 B1 write excel_fix_name
    # C1 E1 G1 ... write bonds name
    # D1 F1 H1 ... write "SMA"
    # 
    print(f'Update Title name Start ...')
    last_title_idx=bonds_list_rowSMA[-1]
    list_title_name = [returnVal(ws[f'{returnExcelColumnLetter(i)}1']) for i in range(1, returnExcelColumnIndex(last_title_idx)+1)]
    try:
        for title_idx in range(1, returnExcelColumnIndex(last_title_idx)+1):
            if title_idx == 1:
                assert(returnVal(ws[f'{str_A}1']) == excel_fix_name[title_idx-1])
                continue
            elif title_idx == 2:
                assert(returnVal(ws[f'{str_B}1']) == excel_fix_name[title_idx-1])
                continue
            elif list_title_name[title_idx-1] == None:
                col_letter=returnExcelColumnLetter(title_idx)
                isBondsCol=(title_idx-START_ROW)%CONTENT_INTERVAL==0
                BondsCol=(title_idx-START_ROW)/CONTENT_INTERVAL
                isSMACol=(title_idx-START_ROW)%CONTENT_INTERVAL==1
                SMACol=(title_idx-START_ROW)/CONTENT_INTERVAL
                if isBondsCol: 
                    writeVal(ws[f'{col_letter}1'], bonds_list_name[int(BondsCol)])
                    print(f"Write title name {bonds_list_name[int(BondsCol)]} in excel Col {col_letter}")
                if isSMACol: 
                    writeVal(ws[f'{col_letter}1'], 'SMA')
                    print(f"Write title name SMA in excel Col {col_letter}")
    except AssertionError:
        ERROR(AssertionError, f'Title idx {title_idx} assert {returnVal(ws[f'{str_A}1'])} == {excel_fix_name[title_idx-1]} Failed')
    finally:
        DONE(f'Update Title name Start ...')

    # 
    # Get USD table
    # 
    print(f'\nCapture USD data Start ...')
    web = requests.get(f'{USD_URL}')
    if web.status_code == 403: ERROR(WebException, f"請檢察'美元'網頁連結正確 {USD_URL}")
    soup = BeautifulSoup(web.text, "html.parser")
    USD_table=soup.find_all('table')[1]
    list_table_date=USD_table.find_all('time')
    list_table_val=[USD_table.find_all('td', {'dir':'ltr'})[i] for i in range(0, len(USD_table.find_all('td', {'dir':'ltr'})), 2)]
    list_USD_val=[val.text for val in list_table_val]
    list_USD_date=[val.text[5:].replace("月", "/").replace("日", "") for val in list_table_date]
    for idx, USD_date in enumerate(list_USD_date):
        m=int(USD_date.split("/")[0])
        d=int(USD_date.split("/")[1])
        str_m=f'0{m}' if m < 10 else str(m)
        str_d=f'0{d}' if d < 10 else str(d)
        list_USD_date[idx]=f'{str_m}/{str_d}'
    DEBUG(f'USD date: {list_USD_date}\n  USD val:  {list_USD_val}')
    DONE(f'Capture USD data Start ...')

    # 
    # Check last 5 days is Saturday or Sunday in excel
    # 
    last_idx=getCurDateIdx(ws, max_row, list_USD_date)
    DEBUG(f'last_idx = {last_idx}')
    checkWeekendLast5day(ws, last_idx)

    # 
    # Update USD value
    # 
    loop=[i for i in range(last_idx, max_row+1)]
    loop.reverse()
    list_excel_date=[returnVal(ws[str_A+str(idx)]).strftime("%m/%d") if type(returnVal(ws[str_A+str(idx)])) is not None else returnVal(ws[str_A+str(idx)]) for idx in loop]
    DEBUG(f"USD date already in exl{list_excel_date}")

    # for idx in range(len(list_USD_val)):
    tmp_date, tmp_val=[], []
    USDINSERTLIST, BONDINSERTLIST = [], []
    for USD_date, USD_val in zip(list_USD_date, list_USD_val):
        if USD_date in list_excel_date:
            excel_date_idx=loop[list_excel_date.index(USD_date)]
            val_address=ws[str_B+str(excel_date_idx)]
            status=writeVal(val_address, float(USD_val), align="center")
            DEBUG(f"write USD/TWD to excel {str_B+str(excel_date_idx)}: {USD_val}")
        else:
            str_dt=list_table_date[list_USD_date.index(USD_date)].text
            DEBUG(f"str_dt: {str_dt}")
            dt=datetime.strptime(str_dt.replace("年", "-").replace("月", "-").replace("日", ""), "%Y-%m-%d")
            if not checkIsSaturdyOrSunday(dt):
                DEBUG(f"Need to Insert USD column: {USD_date} -> {str_dt} -> {dt} --> Is Saturday/Sunday: {checkIsSaturdyOrSunday(dt)}")
                if USD_date not in tmp_date:
                        tmp_date.insert(0, dt)
                        tmp_val.insert(0, USD_val)
    USDINSERTLIST.append(tmp_date)
    USDINSERTLIST.append(tmp_val)
    tmp_date, tmp_val=[], []

    # 
    # Update bonds value
    # 
    for idx, bond_name in enumerate(bonds_list_name):
        print(f'\nWrite data {bond_name} into excel')
        col_val=bonds_list_rowNum[idx]
        col_SMA=bonds_list_rowSMA[idx]
        for date_from_web, val_from_web in zip(bonds_list_date[idx], bonds_list_val[idx]):
            if date_from_web in list_excel_date:
                excel_date_idx=loop[list_excel_date.index(date_from_web)]
                val_address=ws[col_val+str(excel_date_idx)]
                SMA_address=ws[col_SMA+str(excel_date_idx)]
                val_from_excel=returnVal(val_address)
                if val_from_excel is not None and float(val_from_excel) == float(val_from_web):
                    # print(f'No write: {val_from_excel}{excel_date_idx}, {val_from_web}')
                    pass
                else:
                    val_SMA=formula_SMA(col_val, excel_date_idx)
                    status=writeVal(val_address, float(val_from_web))
                    status=writeVal(SMA_address, val_SMA)
                    DEBUG(f"write data to excel {col_val}{excel_date_idx}: {val_from_excel}, {val_from_web}, SMA: {col_SMA}{excel_date_idx} {val_SMA}")
            else:
                # str_dt=list_table_date[list_USD_date.index(USD_date)].text
                dt=datetime.strptime(date_from_web, "%m/%d")
                if not checkIsSaturdyOrSunday(dt):
                    if date_from_web not in tmp_date: 
                        # if checkIsSaturdyOrSunday():
                            tmp_date.insert(0, date_from_web)
                            tmp_val.insert(0, val_from_web)
        BONDINSERTLIST.append(tmp_date)
        BONDINSERTLIST.append(tmp_val)
        tmp_date, tmp_val=[], []
        DONE(f'Write data {bond_name} into excel')

    wb.save(file_name)
    wb.close()
    # 
    # Insert value if not in excel
    # 
    del_idx, del_list=[], [datetime(2024, 3, 29, 0, 0),datetime(2024, 4, 1, 0, 0)]
    for dt in USDINSERTLIST[0]:
        if dt in del_list:
            del_idx.append(USDINSERTLIST[0].index(dt))
    del_idx.reverse()
    for i in range(0, len(USDINSERTLIST)):
        for idx in del_idx:
            USDINSERTLIST[i].pop(idx)
    if len(USDINSERTLIST[0]) != 0:
        #  
        # get USD buy in/selling price
        # Parm: USD date not in excel
        # 
        web = WEBDRIVER(USDINSERTLIST[0])
        list_USDPrice =   web.list_USDPrice
        list_buyInPrice = list_USDPrice[0]
        list_sellPrice =  list_USDPrice[1]
        # 
        # Insert rows and update all columns data
        # 
        print(f'\nInsert data to excel row 2 ~ {2+len(USDINSERTLIST[0])}')
        DEBUG(f'USD date not in excel {USDINSERTLIST}')
        DEBUG(f'Bond date not in excel {BONDINSERTLIST}')
        DEBUG(f"buyIn price {list_buyInPrice}")
        DEBUG(f"sell price {list_sellPrice}")
        ws.insert_rows(2, len(USDINSERTLIST[0]))
        last_row_num=2+len(USDINSERTLIST[0])
        offset=1
        for dt, val, buyInVal, sellVal in zip(USDINSERTLIST[0], USDINSERTLIST[1], list_buyInPrice, list_sellPrice):
            row_num=last_row_num-offset
            writeVal(ws[str_A+str(row_num)], dt)
            writeVal(ws[str_B+str(row_num)], float(val), "center")
            writeVal(ws[str_C+str(row_num)], float(buyInVal), "center") if type(buyInVal) is not str else writeVal(ws[str_C+str(row_num)], buyInVal)
            writeVal(ws[str_D+str(row_num)], float(sellVal), "center")  if type(sellVal)  is not str else writeVal(ws[str_D+str(row_num)], sellVal)
            print(f'Insert data {str_A}{row_num}: {dt.strftime("%Y/%m/%d")}, {str_B}{row_num}: {val}', end='')
            for idx in range (0, bonds_num):
                if dt.strftime("%m/%d") in BONDINSERTLIST[idx*CONTENT_INTERVAL]:
                    val_idx = BONDINSERTLIST[idx*CONTENT_INTERVAL].index(dt.strftime("%m/%d"))
                    value =   BONDINSERTLIST[1+idx*CONTENT_INTERVAL][val_idx]
                    letterCol_bond = returnExcelColumnLetter(START_ROW+idx*CONTENT_INTERVAL)
                    letterCol_SMA =  returnExcelColumnLetter(START_ROW+idx*CONTENT_INTERVAL+1)
                    writeVal(ws[letterCol_bond+str(row_num)], float(value))
                    writeVal(ws[letterCol_SMA+str(row_num)], formula_SMA(letterCol_bond, row_num))
                    print(f', {letterCol_bond}{row_num}: {value}, {letterCol_SMA}{row_num}: {formula_SMA(letterCol_bond, row_num)}', end='')
            offset=offset+1
            print('')
        DONE(f'Insert data to excel row 2 ~ {len(USDINSERTLIST[0])}')
        # 
        # Update all SMA
        # 
        print(f"\nUpdate SMA formula from '{', '.join(bonds_list_rowSMA)}' {last_row_num} ~ {last_row_num+max_row-1}")
        for rowNum, rowSMA in zip(bonds_list_rowNum, bonds_list_rowSMA):
            debug_msg=''
            for i in range(last_row_num, last_row_num+max_row-1):
                cur_val=returnVal(ws[rowSMA+str(i)])
                if cur_val is not None:
                    if "=AVERAGE" in cur_val.upper():
                        writeVal(ws[rowSMA+str(i)], formula_SMA(rowNum, i))
                        debug_msg = debug_msg + f'{rowSMA}{i}: {formula_SMA(rowNum, i)}, '
            DEBUG(f'Update {debug_msg}')
        DONE(f"Update SMA formula from '{', '.join(bonds_list_rowSMA)}' {last_row_num} ~ {last_row_num+max_row-1}")
    else:
        DEBUG(f'NO INSERT: Due to USDINSERTLIST[0] is {len(USDINSERTLIST[0])}')
    wb.save(file_name)
    wb.close()
    # sys.stdout.close()
    os.startfile(file_name)

    os.system("pause")