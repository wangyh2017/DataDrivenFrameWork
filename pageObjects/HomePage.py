#encoding=utf-8
from chapt_15.DataDrivenFrameWork.util.ObjectMap import *
from chapt_15.DataDrivenFrameWork.util.ParseConfigurationFile import ParseConfigFile

class HomePage(object):
    def __init__(self,driver):
        self.driver = driver
        self.parseCF =ParseConfigFile()


    def addressLink(self):
        try:
            #从定位表达式配置文件中读取定位通讯录按钮的定位方式和表达式
            locateType,locatorExpression = self.parseCF.getOptionValue\
                ("126mail_homePage","homePage.addressbook").split(">")
            #获取登录成功页面的通讯录页面元素，并返回给调用者
            elementObject = getElement(self.driver,locateType,locatorExpression)
            return elementObject

        except Exception,e:
            raise  e