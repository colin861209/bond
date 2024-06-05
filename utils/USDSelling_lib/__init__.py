from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException, SessionNotCreatedException
from selenium.webdriver.chrome.service import Service
from utils.Exception_lib import *
import os


class WEBDRIVER:
    def __init__(self, USDDateNotInExcel) -> None:
        self.dateNotInExcel = USDDateNotInExcel
        self.__timeout=30
        url = f'https://announce.fundclear.com.tw/MOPSFundWeb/R03.jsp'
        service = Service(executable_path=r'./chromedriver.exe')
        try:
            self.chrome = webdriver.Chrome(service=service)
        except SessionNotCreatedException:
            ERROR(WebException, f"""ChromeDriver is not support\n\t請去連結下載新版本chromedriver win32在解壓縮至對應資料夾\n\t連結:https://googlechromelabs.github.io/chrome-for-testing/#stable""")
        self.chrome.get(url)
        self.wait = WebDriverWait(self.chrome, timeout=self.__timeout)
        __sell_price = []
        __buyIn_price = []
        self.basicCheck()
        for dt in USDDateNotInExcel:
            self.clickBtn(dt.strftime('%Y/%m/%d'))
            # retVal0 sell price
            # retVal1 buy in price
            USDPrice_tmp = self.getUSDPrice()
            __sell_price.append(USDPrice_tmp[0])
            __buyIn_price.append(USDPrice_tmp[1])
        self.list_USDPrice = [__sell_price, __buyIn_price]
        self.chrome.close()

    def basicCheck(self):
        try:
            xpath1 = f'/html/body/table[1]/tbody/tr/td/form/table[1]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[2]'
            web_compare_str = self.chrome.find_element("xpath", xpath1).text
            expect_str = '集保基金交易平台匯率查詢'
            assert(web_compare_str == expect_str)
        except AssertionError:
            print(f'ASSERT ERROR:: {web_compare_str} == {expect_str}?')
            os.system("pause")


    def clickBtn(self, str_splashDt):
        self.chrome.execute_script(f"document.getElementById('myDate').value='{str_splashDt}'")
        target_btn_xpath = f'/html/body/table[1]/tbody/tr/td/form/table[3]/tbody/tr/td[2]/input[3]'
        self.wait.until(expected_conditions.element_to_be_clickable((By.XPATH, target_btn_xpath)))
        self.chrome.find_element("xpath", target_btn_xpath).click()

    def getUSDPrice(self) -> tuple:
        expect_str = '美元(USD)'
        target_USD_row_xpath =  f"//td[text()='{expect_str}']"
        USD_buyIn_price_xpath = f'/html/body/table[1]/tbody/tr/td/form/table[4]/tbody/tr[15]/td[2]'
        USD_sell_price_xpath =  f'/html/body/table[1]/tbody/tr/td/form/table[4]/tbody/tr[15]/td[3]'
        try:
            self.chrome.find_element("xpath", target_USD_row_xpath).text # if element not found will occur UnexpectedAlertPresentException
            USD_buyIn_price = self.chrome.find_element("xpath", USD_buyIn_price_xpath).text
            USD_sell_price = self.chrome.find_element("xpath", USD_sell_price_xpath).text
            return USD_buyIn_price, USD_sell_price
        except UnexpectedAlertPresentException:
            print(f'ELEMENT WARMING:: {expect_str} NOT FOUND')
            return ['', '']
        # os.system("pause")
