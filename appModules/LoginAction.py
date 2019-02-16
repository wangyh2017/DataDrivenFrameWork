#encoding=utf-8
from chapt_15.DataDrivenFrameWork.pageObjects.LoginPage import  LoginPage

class LoginAction(object):
    def __init__(self):
        print "login..."

    @staticmethod
    def login(driver,username,passsword):
        try:
            login = LoginPage(driver)
            login.switchToFrame()

            login.userNameObj().send_keys(username)
            login.passwordObj().send_keys(passsword)
            login.loginButton().click()

            login.switchToDefaultFrame()
        except Exception,e:
            raise e


if  __name__ == '__main__':
    from selenium import webdriver
    import  time

    driver = webdriver.Firefox(executable_path="/Users/wangyanhua/Documents/学习/Python/webdrive/geckodriver")
    driver.get("http://mail.126.com")

    time.sleep(5)
    LoginAction.login(driver,username="wangyh_0516ok",passsword="1502070803")
    time.sleep(5)
    driver.quit()

