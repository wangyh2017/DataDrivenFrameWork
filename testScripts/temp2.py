#encoding=utf-8
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from chapt_15.DataDrivenFrameWork.pageObjects.LoginPage import LoginPage
from chapt_15.DataDrivenFrameWork.appModules.LoginAction import LoginAction
from chapt_15.DataDrivenFrameWork.util.ParseExcel import  ParseExcel
from chapt_15.DataDrivenFrameWork.appModules.AddContactPersonAction import AddContactPerson
from chapt_15.DataDrivenFrameWork.config.VarConfig import *
import  time
from time import sleep
import  unittest
import  traceback

#设置此次测试环境的编码环境为utf-8
import  sys
reload(sys)
sys.setdefaultencoding("utf-8")
import sys


# 创建解析Excel对象
excelObj = ParseExcel()
#将Excel数据文件加载到内存
excelObj.loadWorkBook(dataFilePath)  #获取不到


def LauchBrowser():
    #创建Chrome浏览器的一个option实例的一个对象
    chrome_options = Options()
    #向option实例中添加禁用扩展插件的设置参数项
    chrome_options.add_argument("--disable-extensions")
    #添加屏蔽ignore-certificate-errors提示信息的设置参数项
    chrome_options.add_experimental_option("excludeSwitches",["ignore-certificate-errors"])
    #添加浏览器最大化设置参数，已启动就是最大化
    chrome_options.add_argument('--start-maximized')
    #启动带有自定义设置的Chrome浏览器
    driver = webdriver.Chrome(executable_path="/Users/wangyanhua/Documents/学习/Python/webdrive/chromedriver" )
    #访问126邮箱
    driver.get("http://mail.126.com")
    sleep(3)
    return driver

def test126MailAddContacts():
    try:
        #根据Excel文件中的sheet名称获取sheet对象
        userSheet = excelObj.getSheetByName(u"126账号")
        #获取126账号sheet是否执行列
        isExecuteUser = excelObj.getColumn(userSheet,account_isExecute)
        #获取126账号sheet中的数据表列
        dataBookColunm = excelObj.getColumn(userSheet,account_dataBook)
        print u"测试为126邮箱添加联系人执行开始..."
        for idx,i in enumerate(isExecuteUser[1:]):
            #循环遍历126账号表中的账号，为需要执行的账号添加联系人
            if i.value == 'y': #表示要执行
                #获取第i行的数据
                userRow = excelObj.getRow(userSheet,idx + 2)
                #获取第i行中的用户名
                username = userRow[account_username - 1].value
                password = str(userRow[account_password -1].value)
                print username,password
                #创建浏览器实例对象
                driver = LauchBrowser()
                #登录126邮箱
                LoginAction.login(driver,username,password)
                #等待三秒，让浏览器启动完成，以便正常进行后续操作
                sleep(3)
                #获取第i行中用户添加的联系人数据表sheet名
                dataBookName = dataBookColunm[idx +1].value
                #获取对应的数据表对象
                dataSheet = excelObj.getSheetByName(dataBookName)
                #获取联系人数据表中是否执行列对象
                isExecuteData = excelObj.getColumn(dataSheet,contacts_isExecute)
                contactsNum =1 #记录添加成功联系人的个数
                isExecuteNum = 0 #记录需要执行联系人个数
                for id,data in enumerate(isExecuteData[1:]):
                    #循环遍历是否执行添加联系人列
                    #如果被设置添加，则进行联系人添加邮件
                    if data.value == 'y':
                        #如果第id行的的联系人被设置为执行，则isExecuteNum自增1
                        isExecuteNum +=1
                        #获取联系人表第id +2行对象
                        rowContent = excelObj.getRow(dataSheet, id+2)
                        #获取联系人姓名
                        contactPersonName = rowContent[contacts_contactPersonName -1].value
                        #获取联系人邮箱
                        contactPersonEmail = rowContent[contacts_contactPersonEmail -1].value
                        #获取是够设置为星标联系人
                        isStr = rowContent[contacts_isStr -1].value
                        #获取联系人手机号
                        contactPersonPhone = rowContent[contacts_contactPersonMobile -1].value
                        #获取联系人备注信息
                        contactPersonComment = rowContent[contacts_contactPersonComment -1].value
                        #添加联系人成功后，断言关键字
                        assertKeyWord = rowContent[contacts_assertKeyWords -1].value
                        print contactPersonName,contactPersonEmail,assertKeyWord
                        print contactPersonPhone,contacts_contactPersonComment,isStr
                        #执行新建联系人操作
                        AddContactPerson.add(driver,
                            contactPersonName,
                            contactPersonEmail,
                            isStr,
                            contactPersonPhone,
                            contactPersonComment)

                        sleep(1)
                        #在联系人工作表中写入联系人执行时间
                        excelObj.writeCellCurrentTime(dataSheet,
                            rowNo= id +2,colsNo=contacts_runTime)
                        try:
                            #断言给定的关键字是否出现在页面中
                            assert  assertKeyWord in driver.page_source
                        except AssertionError,e:
                            #断言失败，在联系人工作表中写入联系人测试失败信息
                            excelObj.writeCell(dataSheet,"failed",rowNo=id+2,colsNo=contacts_testResult)
                        else:
                            #断言成功，写入添加联系人成功信息
                            excelObj.writeCell(dataSheet, "pass", rowNo=id + 2, colsNo=contacts_testResult)
                            contactsNum +=1
                print "contactNum= %s,isExecuteNum= %s" %(contactsNum,isExecuteNum)
                if contactsNum == isExecuteNum:
                    #如果成功添加的联系人人数和需要添加的联系人数相等
                    #说明第i个用户添加联系人测试用例执行成功
                    #在126邮箱账号工作表中写入成功消息，否则写入失败消息
                    excelObj.writeCell(userSheet,"pass",rowNo=idx +2,colsNo=account_testResult)
                    print u"为用户 %s 添加 %d 个联系人，测试通过" % (username,contactsNum)
                else:
                    excelObj.writeCell(userSheet,"failed",rowNo= idx +2,colsNo=account_testResult,)
            else:
                print u"用户 %s 被设置为忽略执行！" %excelObj.getCellOfValue(userSheet,rowNo=idx +2,colsNo = account_username)
    except Exception,e:
        print u"数据驱动框架主程序发生异常，异常信息为："
        #打印异常信息
        print  traceback.print_exc()


if __name__ =='__main__':
    test126MailAddContacts()
    print u"登录126邮箱成功"









""" 

def testMailLogin():
    try:
        #启用Firefox浏览器
        driver = webdriver.Firefox(executable_path="/Users/wangyanhua/Documents/学习/Python/webdrive/geckodriver")
        #访问126邮箱首页
        driver.get("http://mail.126.com")
        #driver.implicitly_wait(30)
        driver.maximize_window()

        LoginAction.login(driver, username="wangyh_0516ok", passsword="1502070803")
        time.sleep(8)

        #assert u"未读邮件" in driver.page_source

    except Exception,e:
        raise e
    finally:
        #退出浏览器
        driver.quit()


if __name__ =='__main__':
    testMailLogin()
    print u"登录126邮箱成功！"

"""
