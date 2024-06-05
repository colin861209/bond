from datetime import date, datetime
from utils.Exception_lib import *
from utils.openpyxl_lib import *
from utils.USDSelling_lib import *

# 
# bonds price and other information
# 
from utils.bonds_lib.main import bondsEnum
from utils.bonds_lib import *

DEBUG([bondKey.name for bondKey in bondsEnum])
DEBUG(bondsEnum.name.value)
DEBUG(bondsEnum.URL.value)
DEBUG(bondsEnum.rowNum.value)
DEBUG(bondsEnum.rowSMA.value)

bonds_list_name    = bondsEnum.name.value
bonds_list_date    = bondsEnum.dates.value
bonds_list_val     = bondsEnum.vals.value
bonds_list_rowNum  = bondsEnum.rowNum.value
bonds_list_rowSMA  = bondsEnum.rowSMA.value
# print(bondsEnum.name.value)