#encoding=utf-8
import sys

reload(sys)
sys.setdefaultencoding('utf-8')
import os
#获取当前文件的绝对路径
parentDirPath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
print parentDirPath


#绝对路径是不是/config/PageElementLocator.ini???
#获取存放页面元素的定位表达式文件的绝对路径
pageElementLocatorPath = parentDirPath + u"/config/PageElementLocator.ini"

#获取存放数据文件的绝对路径
#dataFilePath =  parentDirPath  + u"/testData/126邮箱联系人3.xlsx"
dataFilePath =  u"/Users/wangyanhua/Documents/学习/Python/sele_frame/chapt_15/testData/126邮箱联系人3.xlsx"
#126账户工作表中，每列对应的数字序号

account_username = 2
account_password = 3
account_dataBook = 4
account_isExecute = 5
account_testResult = 6

#联系人工作列表中，每列对应的数字符号
contacts_contactPersonName = 2
contacts_contactPersonEmail = 3
contacts_isStr = 4
contacts_contactPersonMobile = 5
contacts_contactPersonComment = 6
contacts_assertKeyWords = 7
contacts_isExecute = 8
contacts_runTime = 9
contacts_testResult = 10






