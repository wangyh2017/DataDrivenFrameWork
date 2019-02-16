#encoding=utf-8
from chapt_15.DataDrivenFrameWork.pageObjects.HomePage import HomePage
from chapt_15.DataDrivenFrameWork.pageObjects.AddressBookPage import AddressBookPage
import  traceback
import time

class AddContactPerson(object):

    def __init__(self):
        print "add content person"


    @staticmethod
    def add(driver,contactName,contactEmail,isStr,contactPhone,contactComment):
        try:
            #创建主页的实例对象
            hp = HomePage(driver)
            #单击通讯录链接
            hp.addressLink().click()
            time.sleep(3)
            #创建添加联系人页实例对象
            apb = AddressBookPage(driver)
            apb.createContactPersonButton().click()
            if contactName:
                #非必填项
                apb.contactPersonName().send_keys(contactName)
                #必填项
                apb.contactPersonEmail().send_keys(contactEmail)
                if isStr == u"是":
                    #非必填项
                    apb.startContacts().click()
                if contactPhone:
                    #非必填项
                    apb.contactPersonMobile().send_keys(contactPhone)
                if contactComment:
                    apb.contactPersonComment().send_keys(contactComment)

                apb.savecontacePerson().click()

        except Exception,e:
            #打印异常信息
            print traceback.print_exc()
            raise e

if __name__ == '__main__':
    from chapt_15.DataDrivenFrameWork.appModules.LoginAction import LoginAction
    from selenium import webdriver
    import  time
    #启动火狐浏览器
    driver = webdriver.Firefox(executable_path="/Users/wangyanhua/Documents/学习/Python/webdrive/geckodriver")
    #driver = webdriver.Chrome(executable_path="/Users/wangyanhua/Documents/学习/Python/webdrive/chromedriver")
    #访问126邮箱首页
    driver.get("http://mail.126.com")
    #driver.maximize_window()
    time.sleep(5)
    LoginAction.login(driver,username="wangyh_0516ok",passsword="1502070803")
    time.sleep(5)
    AddContactPerson.add(driver,u"张三","zs@qq.com",u"是","","")
    time.sleep(3)
    #assert u"张三" in driver.page_source
    driver.quit()
