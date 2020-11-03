from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import xlwt
import xlrd
from xlutils.copy import copy
from progressbar import ProgressBar
import time
import re


def getPdbDate(n, pdb):
    url_pdb = Url + pdb
    driver.get(url_pdb)
    try:
        dep = driver.find_element_by_xpath('//*[@id="header_deposited-released-dates"]').text
        dep = re.findall(" (.*?) ", dep)[0]
        excel_table.write(n, 11, dep)
    except:
        excel_table.write(n, 11, 'Error')
    excel.save('PDB_Excel.xls')


if __name__ == '__main__':
    driver = webdriver.Chrome('./chromedriver.exe')
    driver.implicitly_wait(5)

    Url = 'http://www.rcsb.org/structure/'

    data = xlrd.open_workbook('PDB_Excel.xls', formatting_info=True)
    excel = copy(wb=data)  # 完成xlrd对象向xlwt对象转换
    excel_table = excel.get_sheet(0)  # 获得要操作的页

    table = data.sheet_by_index(0)
    rows = table.nrows

    excel_table.write(0, 11, 'Deposited')

    pbar = ProgressBar()
    for i in pbar(range(1, rows)):
        Pdb = table.cell(i, 0).value
        getPdbDate(i, Pdb)
    excel.save('PDB_Excel.xls')
