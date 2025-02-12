import pandas as pd

#это модуль преобразования JSON в pandas DataFrame
#ВНИМАНИЕ!!!
#Для того, чтобы нормально пользоваться данной штукой, нужно заранее ознакомиться с обрабатываемым JSON
#============================================================================================================================================================
#ПОРЯДОК ВАЖЕН
#на вход функции поступает - 
#------------------------------------------------------------------------------------------------------------------------------------------------------------
#1)data - содержит JSON 
#------------------------------------------------------------------------------------------------------------------------------------------------------------
#2)list_path - содержит пути к спискам записей в JSON. ТИП - список строк / список списков строк.
#------------------------------------------------------------------------------------------------------------------------------------------------------------
#3)list_meta - содержит пути к определенным полям в JSON, которые будут использоваться в DataFrame. ТИП - список строк / список списков строк
#если такого нет, то передаем пустой список list_meta=[]
#------------------------------------------------------------------------------------------------------------------------------------------------------------
#4)cols_to_norm - содержит пути к спискам записей в JSON, которые были в списках записей (вложены). ТИП - строка / список строк
#если вложенностей нет, то пишем cols_to_norm=[]
#------------------------------------------------------------------------------------------------------------------------------------------------------------
#5)dict_fillna - содержит словарь, ключи которого, содержат столбцы DataFrame, в которых надо поменять отсутсвующие данные(Nan) на те, которые указаны в словаре.
#если замена не требуется, то передаем пустой словарь dict_fillna={}
#------------------------------------------------------------------------------------------------------------------------------------------------------------
#6)merge_method - содержит один из методов слияния: "left","right","outer","inner","cross"
#left - использовать только ключи из левого фрейма, аналогично левому внешнему соединению SQL; сохранить порядок ключей.
#right - использовать только ключи из правого фрейма, аналогично правому внешнему соединению SQL; сохранить порядок ключей
#outer - использовать объединение ключей из обоих фреймов, аналогично полному внешнему соединению SQL; сортировать ключи лексикографически.
#inner - использовать пересечение ключей из обоих фреймов, аналогично внутреннему соединению SQL; сохранить порядок левых клавиш.
#cross - создает декартово произведение из обоих фреймов, сохраняет порядок левых ключей.

#merge_method используется только при обработке JSON, в которых есть множество записей.
#если множества записей нет, то передаем пустую строку merge_method=''

#При выборе метода мержа, рекомендуется потыкать разные методы и посмотреть, какой подойдёт

#ВНИМАНИЕ!!! 
#Информация ниже требует проверки. 
#Если в списках записей нет одинаковых ключей, то cross. Если в списках записей есть одинковые ключи, то outer

#для большей информации смотри инструкцию!!!!!
#=============================================================================================================================================================

def js_norm(data, list_path, list_meta, cols_to_norm, dict_fillna, merge_method):
    if len(list_path) == 0:  
        df_json = pd.json_normalize(data, meta = list_meta, sep ='_', errors="ignore")  
        df_json = df_json.apply(lambda x: x.explode()).reset_index(drop=True)  
    else:    
        df_json = pd.json_normalize(data, record_path = list_path[0], meta = list_meta, sep = '_', errors="ignore") 
        df_json = df_json.apply(lambda x: x.explode()).reset_index(drop=True)
        for i in range(1,len(list_path)):   
            # print('ffdfsdfdsfdsf')
            df_temp = pd.json_normalize(data, record_path = list_path[i], sep =' ', errors="ignore") 
            df_temp = df_temp.apply(lambda x: x.explode()).reset_index(drop=True)
            df_json = df_json.merge(df_temp, how = merge_method) 
    if len(cols_to_norm) != 0:
        normalized = list()
        for col in cols_to_norm:
            d = pd.json_normalize(df_json[col], sep='_')
            d.columns = [f'{col}_{v}' for v in d.columns]
            normalized.append(d.copy())
        df_json = pd.concat([df_json] + normalized, axis=1).drop(columns=cols_to_norm)
    for i in dict_fillna: 
        df_json = df_json.fillna({i:dict_fillna[i]})
    return df_json 