#encoding=utf-8
from chapt_15.DataDrivenFrameWork.util.ObjectMap import *
from chapt_15.DataDrivenFrameWork.util.ParseConfigurationFile import ParseConfigFile

class AddressBookPage(object):
    def __init__(self,driver):
        self.driver = driver
        self.parseCF =ParseConfigFile()
        self.addContactsOptions = self.parseCF.getItemsSection("126mail_addContactsPage")
        print self.addContactsOptions

    def createContactPersonButton(self):
        #获取新建联系人按钮
        try:
            #从定位表达式配置文件中读取定位新建联系人按钮的定位方式和表达式
            locateType,locatorExpression = self.addContactsOptions["addContactsPage.createContactsBtn".lower()].split(">")
            #获取新建联系人按钮页面元素，并返回给调用者
            elementObj = getElement(self.driver,locateType,locatorExpression)
            return  elementObj

        except Exception,e:
            raise e

    def contactPersonName(self):
        #获取新建联系人界面中的姓名输入框
        try:
            #从定位表达式配置文件中读取定位新建联系人姓名输入框定位方式和表达式
            locateType,locatorExpression = self.addContactsOptions["addContactsPage.contactPersonName".lower()].split(">")
            #获取新建联系人界面输入框页面元素，并返回给调用者
            elementObj = getElement(self.driver,locateType,locatorExpression)
            return  elementObj

        except Exception,e:
            raise e

    def contactPersonEmail(self):
        #获取新建联系人界面中的电子邮件输入框
        try:
            #从定位表达式配置文件中读取定位新建联系人邮箱输入框的定位方式和表达式
            locateType,locatorExpression = self.addContactsOptions["addContactsPage.contactPersonEmail".lower()].split(">")
            #获取新建联系人界面的邮箱输入框页面元素，并返回给调用者
            elementObj = getElement(self.driver,locateType,locatorExpression)
            return  elementObj

        except Exception,e:
            raise e

    def startContacts(self):
        #获取新建联系人界面中星标联系人选择框
        try:
            #从定位表达式配置文件中读取定位星标联系人的复选框定位方式和表达式
            locateType,locatorExpression = self.addContactsOptions["addContactsPage.startContacts".lower()].split(">")
            #获取新建联系人界面的星标联系人复选框页面元素，并返回给调用者
            elementObj = getElement(self.driver,locateType,locatorExpression)
            return  elementObj

        except Exception,e:
            raise e

    def contactPersonMobile(self):
        # 获取新建联系人界面中联系人手机号输入框
        try:
            # 从定位表达式配置文件中读取定位联系人手机号输入框定位方式和表达式
            locateType, locatorExpression = self.addContactsOptions["addContactsPage.contactPersonMobile".lower()].split(">")
            # 获取新建联系人界面的联系人手机号输入框页面元素，并返回给调用者
            elementObj = getElement(self.driver, locateType, locatorExpression)
            return elementObj

        except Exception, e:
            raise e

    def contactPersonComment(self):
        # 获取新建联系人界面中联系人备注信息输入框
        try:
            # 从定位表达式配置文件中读取定位联系人备注信息输入框定位方式和表达式
            locateType, locatorExpression = self.addContactsOptions["addContactsPage.contactPersonComment".lower()].split(">")
            # 获取新建联系人界面的联系人备注信息输入框页面元素，并返回给调用者
            elementObj = getElement(self.driver, locateType, locatorExpression)
            return elementObj

        except Exception, e:
            raise e

    def savecontacePerson(self):
        # 获取新建联系人界面中保存联系人按钮
        try:
            # 从定位表达式配置文件中读取定位保存联系人按钮定位方式和表达式
            locateType, locatorExpression = self.addContactsOptions["addContactsPage.saveContacePerson".lower()].split(">")
            # 获取新建联系人界面的保存联系人按钮页面元素，并返回给调用者
            elementObj = getElement(self.driver, locateType, locatorExpression)
            return elementObj

        except Exception, e:
            raise e

