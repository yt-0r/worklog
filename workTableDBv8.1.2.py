# Скрипт формирует Exel таблицу на основе Json данных переденных Jira
# нужно установить библиотеку pip install openpyxl
# pip install mysql-connector-python
# pip install pandas==1.5.3
# и, возможно pip install numpy

#своё
from process_JS import processJson
from calc import calcJson, calcTotal
from exell import create_exel
from jira_attach import JiraAttach
import settings
import db_connect
from sys import argv
import pandas as pd

import warnings

#отключаем предупреждения об pandas.append()

warnings.filterwarnings('ignore')

def main():
#==============================================================================
    # производим первоначальные действия
    
    settings.outputPandas_set()
    
    arg = settings.parseArg_set(argv) #парсим аргументы
    
    arglog = settings.arglog_set(arg) #делаем копию для логов
    
    config = settings.config_set(arg) #подгружаем конфиг

    settings.logger_set(arglog, config) #настраиваем логи
    
    time_string = settings.time_set() #здесь берем текущую дату, если надо другую, надо зайти в settings и там поиграться с delta
    
    rest = settings.rest_set(arg, config, time_string) # ссылка
    
    response = settings.response_set(arg, rest) 
    
    data = settings.data_set(response) #считываем JSON

    data = [{
        "period": {
            "month": "Май",
            "year": 2024
        },
        "job": {
            "name": "Водяников Сергей Васильевич",
            "department": "Группа содержания",
            "position": "Ведущий инженер"
        },
        "days": [
            {
                "work_calendar": {
                    "day": 26,
                    "daytype": 1
                },
                "kontrakt": [
                    {
                        "name": "Общехозяйственный"
                    },
                    {
                        "name": "Контракт № 644 - АО Автодор"
                    },
                    {
                        "name": "Контракт № 633 - ПАО Ростелеком Красноярск"
                    },
                    {
                        "name": "Контракт № 632 - ГКУ НСО ТУАД"
                    },
                    {
                        "name": "Контракт № 631 - ГКУ НСО ТУАД"
                    },
                    {
                        "name": "Контракт № 630 - ГКУ НСО ТУАД"
                    },
                    {
                        "name": "Контракт № 620 - МариинскАвтодор"
                    },
                    {
                        "name": "Контракт № 621 - Томскавтодор"
                    },
                    {
                        "name": "Контракт № 592 - ООО Восток-М"
                    },
                    {
                        "name": "Контракт № 599 - Восток-М"
                    },
                    {
                        "name": "Контракт № 591 - ООО Восток-М"
                    },
                    {
                        "name": "Контракт № 579 - ООО Восток-М"
                    },
                    {
                        "name": "Контракт № 585 - ФКУ Сибуправтодор"
                    },
                    {
                        "name": "Контракт № 580 - ОГКУ «Томскавтодор»"
                    },
                    {
                        "name": "Контракт № 576 - Восток-М"
                    },
                    {
                        "name": "Контракт № 625 - Нижний Новгород",
                        "issuecount": 2,
                        "timetracking": [
                            {
                                "TECHWIM-4373": [
                                    1716688800,
                                    1716721200
                                ]
                            },
                            {
                                "TECHWIM-4372": [
                                    1716688800,
                                    1716721200
                                ]
                            }
                        ],
                        "filter": "TECHWIM-4373,TECHWIM-4372"
                    }
                ],
                "event": "shift",
                "day_night": "day",
                "work_time": 0.0
            }
        ]
    }]


#------------------------------------------------------------------------------   
    # делаем нормальную дату, вносим в лог 
    
    settings.locale_set()
    
    arg['year'] = settings.year_set(arg)  # для параметра "year" можно указать год конкретно (цифрой) или "now" (текущий год)
   
    arg['month'] = settings.month_set(arg) # для "month" можно указать любой один месяц (строкой) или "now" (текущий) или "all" (все месяцы в году)
    
#==============================================================================
    # коннектимся к БД

    workers = db_connect.workers_db()
    
    kontrakts = db_connect.kontrakts_db()
    
    clocks = db_connect.clocks_db()
    
    total = db_connect.total_db()
    
    raw =  db_connect.raw_db()
    
    engine = db_connect.engine_db(config)

    db_connect.create_all_db(engine)
  
    conn = db_connect.conn_db(engine)
    
    # engine.connect()
    
    session = db_connect.session_db(engine)
    
#==============================================================================    
    # формируем DataFrame и вставляем в базу данных
    
    df_json = processJson(data, engine, conn, workers, kontrakts, raw, clocks) # в этом DataFrame весь JSON  
    
    calcJson(df_json, engine, conn, clocks)
        
    df_base = db_connect.df_base_set(conn)
                              
    df_total = calcTotal(df_json, total, engine, df_base, conn)
    
#==============================================================================
    # создаем exel
    
    flag = True  #если True - дописываем табель, если False то создаем заноvо 
    
    create_exel(df_base, df_total, df_json, arg, config, flag) 
    
#==============================================================================
    # атачим
    
    flag = True #если тру, то не удаляем
    file_path = config["excel"]["excel_path"]+config["excel"]["excel_name"]+' '+arg['worker']+config["excel"]["typefile"]
    

    JiraAttach(config, arg['issuekey'], file_path, arg['Login'], arg['Password'], flag)

main()
