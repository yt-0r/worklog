import pandas as pd
import configparser
import logging
from requests.auth import HTTPBasicAuth
import json
import requests
import time
import locale
from datetime import timedelta, datetime

#------------------------------------------------------------------------------
def outputPandas_set(): # настройки вывода pandas в консоль 
    pd.set_option('display.max_rows', 2000)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)
    pd.set_option('expand_frame_repr', False)
    pd.options.mode.chained_assignment = None
#------------------------------------------------------------------------------    
def parseArg_set(argv):  #парсим аргументы 
    arg = {}
    for i in range(1, len(argv)):
        arg_str = argv[i].split(':',1)
        arg[arg_str[0]] = arg_str[1]
        arg[arg_str[0]] = arg[arg_str[0]].replace('=', ':')
    return arg
#------------------------------------------------------------------------------  
def arglog_set(arg):       # arglog - копия словаря для вставки в log. Удаляем из него пароли и пути
    arglog = arg.copy()
    del arglog['Login'], arglog['Password'], arglog['config_path']
    return arglog
#------------------------------------------------------------------------------
def config_set(arg): # подключаем конфиг
    config = configparser.ConfigParser()  # создаём объекта парсера
    config.read(arg['config_path'],encoding='UTF-8')  # читаем конфиг arg['config_path']
    return config
#------------------------------------------------------------------------------
# def loggi(arg):
#     logging.debug("year_now " +str(arg['year'])+ " month_now " + str(arg['month']))
#------------------------------------------------------------------------------
def logger_set(arglog, config): # настраиваем логи
    logging.basicConfig(filename=config["log"]["log_path"], 
                        filemode='w',
                        format='[%(asctime)s] [%(levelname)s] => %(message)s', 
                        datefmt='%Y-%m-%d %H:%M:%S',
                        level=logging.DEBUG)
    logging.info("Start scrypt")
    logging.debug("Started arguments "+str(arglog))
#------------------------------------------------------------------------------    
def time_set():     #узнаем дату
    # month_now_num = time.strftime('%m')
    # print(month_now_num)
    now = datetime.now()
    day_delta = timedelta(1)    #если делта = 1, то берем за день до
    time_delta = now - day_delta
    time_string = time_delta.strftime("%d.%m.%Y")
    return time_string

#------------------------------------------------------------------------------
def rest_set(arg, config, time_string):
    if arg['worker'] != 'all':
        rest = config["jira"]["rest"]+'?query=staff&staff='+arg['worker']
    else:
        #rest = config["jira"]["rest"]
        # rest = 'http://jiradev.its-sib.ru/rest/scriptrunner/latest/custom/report_backup'
        
        # rest = 'http://jiradev.its-sib.ru/rest/scriptrunner/latest/custom/test?query=getNew&date=12.04.2023'
        # rest = 'http://jiradev.its-sib.ru/rest/scriptrunner/latest/custom/test?query=getNew&date=25.03.2023'
        
        rest = config["jira"]["rest"]+'?query=getNew&date='+time_string
        # rest = config["jira"]["rest"]+'?query=getNew&date='+'17.04.2023'
    return rest
#------------------------------------------------------------------------------
def response_set(arg, rest):
    return requests.get(rest, auth=HTTPBasicAuth(arg['Login'], arg['Password']))
#------------------------------------------------------------------------------   
def data_set(response):
    return json.loads(response.text)
#------------------------------------------------------------------------------
def year_set(arg):
    if arg['year'] == 'now':
        arg['year'] = time.strftime('%Y')
        return arg['year']
    else:
        return arg['year']
#------------------------------------------------------------------------------
def month_set(arg):
    if arg['month'] == 'now':
        month = ['','Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь']
        i = int(time.strftime('%m')) #текущий месяц
        arg['month'] = month[i]
        return arg['month']
    elif arg['month'] == 'all':
        return arg['month']
    else:
        return arg['month']
#------------------------------------------------------------------------------
def locale_set():
    locale.setlocale(category=locale.LC_ALL, locale="ru_RU.utf8")
#------------------------------------------------------------------------------
