#encoding=utf-8
from ConfigParser import ConfigParser
from chapt_15.DataDrivenFrameWork.config.VarConfig import pageElementLocatorPath

class ParseConfigFile(object):

    def __init__(self):
        self.cf = ConfigParser()
        self.cf.read(pageElementLocatorPath)

    def getItemsSection(self,sectionName):
        #获取配置文件中指定的section下面所有的option键值对
        #并以字典类型返回给调用者
        """注意
        使用self.cf.items(sectionName)此种方法获取的
        配置文件中的options均被转换成小写
        比如loginPage.frame被转换成了loginpage.frame
        """

        optionsDiction = dict(self.cf.items(sectionName))
        return  optionsDiction

    def getOptionValue(self,sectionName,optionName):
        #被指定的section下的指定option的值
        value = self.cf.get(sectionName,optionName)
        return  value

if __name__ == '__main__':
    pc = ParseConfigFile()
    print pc.getItemsSection("126mail_login")
    print pc.getOptionValue("126mail_login","loginPage.frame")
