#encoding=utf-8

#author-夏晓旭

from selenium import webdriver

import time

from selenium.webdriver.support.ui import WebDriverWait

from selenium.webdriver.common.by import By

from selenium.common.exceptions import TimeoutException, NoSuchElementException

import traceback

from chapt_15.DataDrivenFrameWork.pageObjects.LoginPage import *

from chapt_15.DataDrivenFrameWork.pageObjects.AddressBookPage  import *


from chapt_15.DataDrivenFrameWork.appModules.LoginAction import LoginAction
from chapt_15.DataDrivenFrameWork.util.ParseExcel import   *
from chapt_15.DataDrivenFrameWork.util.temp import   parseExcel
from chapt_15.DataDrivenFrameWork.appModules.AddContactPersonAction import AddContactPerson
from chapt_15.DataDrivenFrameWork.config.VarConfig import *


driver =webdriver.Firefox(executable_path="/Users/wangyanhua/Documents/学习/Python/webdrive/geckodriver")

driver.get('http://mail.126.com')



pe =parseExcel(u'/Users/wangyanhua/Documents/学习/Python/sele_frame/chapt_15/testData/126邮箱联系人3.xlsx')

pe.set_sheet_by_name(u"126账号")

print pe.get_default_sheet()

rows=pe.get_all_rows()[1:]

for id,row in enumerate(rows):

    if row[4].value =='y':

        username = row[1].value

        password = row[2].value

        print username, password

        try:

            LoginAction.login(driver,username,password)

            pe.set_sheet_by_name(u"联系人")

            print pe.get_default_sheet()

            rows1=pe.get_all_rows()[1:]

            print "rows1:",rows1

            for id1,row in enumerate(rows1):

                if row[7].value == 'y':

                    try:

                        #print row[1].value,row[2].value,row[3].value,row[4].value,row[5].value

                        #print "execute1"

                        add_contact(driver,row[1].value,row[2].value,row[3].value,row[4].value,row[5].value)

                        print "assert word:",row[6].value in driver.page_source

                        print row[6].value

                        pe.write_cell_content(id1+2,10,"pass")

                    except Exception,e:

                        print u"异常信息：",e.message

                        pe.write_cell_content(id1+2,10,"fail")

                else:

                    pe.write_cell_content(id1+2,10,u"忽略")

                    continue

        except Exception,e:

            pe.set_sheet_by_name(u"126账号")

            print u"异常信息：",e

            pe.write_cell_content(id+2,5,"fail")

    else:

        pe.set_sheet_by_name(u"126账号")

        pe.write_cell_content(id+2,6,u"忽略")

        continue