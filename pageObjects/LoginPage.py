#encoding=utf-8
import sys
import os
from chapt_15.DataDrivenFrameWork.util.ObjectMap import *
from chapt_15.DataDrivenFrameWork.util.ParseConfigurationFile import ParseConfigFile

class LoginPage(object):

    def __init__(self,driver):
        self.driver = driver
        self.parseCF= ParseConfigFile()
        self.loginOptions = self.parseCF.getItemsSection("126mail_login")
        print self.loginOptions


    def switchToFrame(self):

        try:
            #从定位表达式配置文件中读取frame的定位表达式
            #locatorExpression = self.loginOptions["loginPage.frame".lower()].split(">")[1]
            self.driver.switch_to.frame(self.driver.find_element_by_tag_name("iframe"))
            #self.driver.switch_to.frame(locatorExpression)
        except Exception,e:
            raise e

    def switchToDefaultFrame(self):
        try:

            self.driver.switch_to.default_content()
        except Exception,e:
            raise e

    def userNameObj(self):
        try:
            #从定位表达式配置文件中读取定位用户名输入框的定位方式和表达式和表达式
            locateType,locatorExpression = self.loginOptions["loginPage.username".lower()].split(">")
            #获取登录页面的用户输入框页面对象，并返回给调用者
            elmentObj = getElement(self.driver,locateType,locatorExpression)
            return  elmentObj
        except Exception,e:
            raise  e

    def passwordObj(self):
        try:
            #从定位表达式配置文件中读取定位密码输入框的定位方式和表达式和表达式
            locateType,locatorExpression = self.loginOptions["loginPage.password".lower()].split(">")
            #获取登录页面的用户输入框页面对象，并返回给调用者
            elmentObj = getElement(self.driver,locateType,locatorExpression)
            return  elmentObj
        except Exception, e:
            raise e

    def loginButton(self):
        try:
            # 从定位表达式配置文件中读取定位登录按钮表达式和表达式
            locateType, locatorExpression = self.loginOptions["loginPage.loginbutton".lower()].split(">")
            # 获取登录页面的用户输入框页面对象，并返回给调用者
            elmentObj = getElement(self.driver, locateType, locatorExpression)
            return elmentObj
        except Exception, e:
            raise e

if __name__ =='__main__':
    from selenium import webdriver
    driver = webdriver.Firefox(executable_path="/Users/wangyanhua/Documents/学习/Python/webdrive/geckodriver")
    #driver = webdriver.Chrome(executable_path="/Users/wangyanhua/Documents/学习/Python/webdrive/chromedriver")
    driver.get("http://mail.126.com")
    import  time

    time.sleep(5)
    login = LoginPage(driver)
    time.sleep(10)
    login.switchToFrame()

    #输入登录用户用
    login.userNameObj().send_keys("wangyh_0516ok")
    #输入登录密码
    login.passwordObj().send_keys("1502070803")

    login.loginButton().click()
    time.sleep(10)
    login.switchToDefaultFrame()
    assert u"未读邮件" in  driver.page_source
    driver.quit()








