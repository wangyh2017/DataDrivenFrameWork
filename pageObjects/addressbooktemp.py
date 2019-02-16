# encoding=utf-8

# author-夏晓旭

from selenium import webdriver

import time

from selenium.webdriver.support.ui import WebDriverWait

from selenium.webdriver.common.by import By

from selenium.common.exceptions import TimeoutException, NoSuchElementException

import traceback

from chapt_15.DataDrivenFrameWork.pageObjects.LoginPage  import *


class AddressBook(object):

    def __init__(self, driver):

        self.driver = driver

    def add_contact(self):

        try:

            wait = WebDriverWait(self.driver, 10, 0.2)  # 显示等待

            address_book_link = wait.until(lambda x: x.find_element_by_xpath("//div[text()='通讯录']"))

            address_book_link.click()

            # assert u"新建联系人" in driver.page_source

            add_contact_button = wait.until(lambda x: x.find_element_by_xpath("//span[text()='新建联系人']"))

            add_contact_button.click()

            contact_name = wait.until(

                lambda x: x.find_element_by_xpath("//a[@title='编辑详细姓名']/preceding-sibling::div/input"))

            contact_name.send_keys(u"徐凤钗")

            contact_email = wait.until(lambda x: x.find_element_by_xpath("//*[@id='iaddress_MAIL_wrap']//input"))

            contact_email.send_keys("593152023@qq.com")

            contact_is_star = wait.until(

                lambda x: x.find_element_by_xpath("//span[text()='设为星标联系人']/preceding-sibling::span/b"))

            contact_is_star.click()

            contact_mobile = wait.until(lambda x: x.find_element_by_xpath("//*[@id='iaddress_TEL_wrap']//dd//input"))

            contact_mobile.send_keys('18141134488')

            contact_other_info = wait.until(lambda x: x.find_element_by_xpath("//textarea"))

            contact_other_info.send_keys('my wife')

            contact_save_button = wait.until(lambda x: x.find_element_by_xpath("//span[.='确 定']"))

            contact_save_button.click()



        except TimeoutException, e:

            # 捕获TimeoutException异常

            print traceback.print_exc()



        except NoSuchElementException, e:

            # 捕获NoSuchElementException异常

            print traceback.print_exc()



        except Exception, e:

            # 捕获其他异常

            print traceback.print_exc()


if __name__ == "__main__":
    driver = webdriver.Firefox(executable_path="/Users/wangyanhua/Documents/学习/Python/webdrive/geckodriver")

    driver.get('http://mail.126.com')

    lp = LoginPage(driver)

    lp.Login()

    ab = AddressBook(driver)

    ab.add_contact()