from utils.Exception_lib import *
from utils.openpyxl_lib import *

dict_idx = {0: '日期', 1: '淨值'}
dates, val = [], []

TITLE_NAME = '近30日淨值'


START_ROW=len(excel_fix_name)+1
CONTENT_INTERVAL=2
# Open excel bonds sheet
wb, ws=open_workbook(sheet_name=SHEET_REFERWEBSITE_NAME)
bonds_num=ws.max_row

list_name=[ws[str_A+str(i)].value for i in range(1, bonds_num+1)] # bonds name
list_urls=[ws[str_B+str(i)].value for i in range(1, bonds_num+1)] # bonds url
list_date=[]
list_val=[]
list_rowNum=[returnExcelColumnLetter(START_ROW+i*CONTENT_INTERVAL)   for i in range(0, bonds_num)]
list_rowSMA=[returnExcelColumnLetter(START_ROW+i*CONTENT_INTERVAL+1) for i in range(0, bonds_num)]

DEBUG(f'債券名稱: {list_name}, 債券網址: {list_urls}, 債券對應欄位: {list_rowNum}, 債券SMA對應欄位: {list_rowSMA}')
