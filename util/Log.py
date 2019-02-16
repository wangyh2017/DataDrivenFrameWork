#encoding=utf-8
import  logging
import logging.config
from chapt_15.DataDrivenFrameWork.config.VarConfig import parentDirPath

#读取日志配置文件
logging.config.fileConfig(parentDirPath + u"/config/Logger.conf")

logger = logging.getLogger("example02") #或者example01

def debug(message):
    #定义debuug级别日志打印
    logger.debug(message)


def info(message):
    #定义info级别日志打印方法
    logger.info(message)

def warning(message):
    #定义warning级别日志的方法
    logging.warning(message)


