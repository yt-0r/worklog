# Стандартные библиотеки
import time
# Сторонние библиотеки
import pandas as pd
from sqlalchemy import insert
import numpy as np

def calcJson(df_json, engine, conn, clocks):
    def time_tracking_2(df_json, dff, id_worker, work_calendar_day, period_month, period_year):
        # dff = df_json['kontrakt_timetracking'].loc[(df_json['id_worker'] == id_worker) & (df_json['work_calendar_day'] == work_calendar_day) & (df_json['period_month'] == period_month)]
        
        # print(df_json['kontrakt_timetracking'].loc[(df_json['id_worker'] == id_worker) & (df_json['work_calendar_day'] == work_calendar_day) & (df_json['period_month'] == period_month)],'\n\n')
        
        return_dict = {}

#----------------------------------------------------------------------------
#заменяем 'Nan' на нижеприведенную конструкцию, для того, чтобы ничего не сломалось при подсчете
        rep = {'Nan':[{'0':[0,0]}]}
        
        df_json['kontrakt_timetracking'].loc[(df_json['id_worker'] == id_worker) 
            & (df_json['work_calendar_day'] == work_calendar_day) 
            & (df_json['period_month'] == period_month) 
            & (df_json['kontrakt_timetracking'] == 'Nan')] = df_json['kontrakt_timetracking'].loc[(df_json['id_worker'] == id_worker) 
                          & (df_json['work_calendar_day'] == work_calendar_day) 
                          & (df_json['period_month'] == period_month) & (df_json['kontrakt_timetracking'] == 'Nan')].map(rep)
                                                       
        # print(df_json['kontrakt_timetracking'].loc[(df_json['id_worker'] == id_worker) & (df_json['work_calendar_day'] == work_calendar_day) & (df_json['period_month'] == period_month)],'\n\n\n')
#----------------------------------------------------------------------------- 
#превращаем kontrakt_timetracking в массив с единицами
        
        # print(df_json['kontrakt_timetracking'].loc[(df_json['id_worker'] == id_worker) & (df_json['work_calendar_day'] == work_calendar_day) & (df_json['period_month'] == period_month)],'\n\n\n')

        kontraktdict = df_json['kontrakt_timetracking'].loc[(df_json['id_worker'] == id_worker) 
            & (df_json['work_calendar_day'] == work_calendar_day) 
            & (df_json['period_month'] == period_month)].to_dict()
        
        for i in kontraktdict:
            # print('i >>>>> ', i)
            for j in kontraktdict[i]:
                # print('j >>>>> ', j)
                min_list = [0 for i in range(0, 1440)]
                # заполняем список единицами, там, где нашлось время (тут мы переводим unix time в нормальные минуты и на каждую минуту ставим "1")
                start_min = time.gmtime(list(j.values())[0][0]).tm_hour*60+time.gmtime(list(j.values())[0][0]).tm_min
                stop_min = time.gmtime(list(j.values())[0][1]).tm_hour*60+time.gmtime(list(j.values())[0][1]).tm_min
                for t in range(start_min, stop_min):
                    min_list[t] = 1
                # print('min list >>>>>>',min_list)
                # складываем все назад
                j[list(j.keys())[0]] = min_list
                # print('after j >>>>>>', j)
          
        # print(kontraktdict)
        # print(df_json['kontrakt_timetracking'].loc[(df_json['id_worker'] == id_worker) & (df_json['work_calendar_day'] == work_calendar_day) & (df_json['period_month'] == period_month)],'\n\n\n')
        # здесь мы сложили все единички в каждом контракте в один массив, 
        # таким образом мы получили четкое распределение в этом контракте когда человек работал
        
        for i in kontraktdict:
            summin_list = [0 for i in range(0, 1440)]
            for j in range(len(kontraktdict[i])):
                for si in range(len(summin_list)):
                    summin_list[si] = summin_list[si] + list(kontraktdict[i][j].values())[0][si]
            return_dict[i] = summin_list
            
        # print('return_dict 1 >>>>', return_dict,'\n\n')
            
#------------------------------------------------------------------------------
        # print(df_json['id_worker'].loc[(df_json['id_worker'] == id_worker) 
        #         & (df_json['work_calendar_day'] == work_calendar_day) 
        #         & (df_json['period_month'] == period_month)])
        multiplierlist = [0 for i in range(0, 1440)] 
        for key, val in df_json['work_calendar_day'].loc[(df_json['id_worker'] == id_worker) 
                & (df_json['work_calendar_day'] == work_calendar_day) 
                & (df_json['period_month'] == period_month)].to_dict().items(): 
                # print(key, val)
                if df_json['id_worker'].loc[(df_json['id_worker'] == id_worker) 
                        & (df_json['work_calendar_day'] == work_calendar_day) 
                        & (df_json['period_month'] == period_month)][key] == id_worker: 
                    if val == work_calendar_day:
                        multiplierlist = list(map(sum, zip(multiplierlist, return_dict[key])))

                # тут делим единицы на их сумму
        # print('multi >>>>>',multiplierlist,'\n\n')
        for key, val in df_json['work_calendar_day'].loc[(df_json['id_worker'] == id_worker) 
                & (df_json['work_calendar_day'] == work_calendar_day) 
                & (df_json['period_month'] == period_month)].to_dict().items(): 
                if df_json['id_worker'].loc[(df_json['id_worker'] == id_worker) 
                        & (df_json['work_calendar_day'] == work_calendar_day) 
                        & (df_json['period_month'] == period_month)][key] == id_worker: 
                        for j in range(len(return_dict[key])):
                            if multiplierlist[j] != 0:
                                return_dict[key][j] = return_dict[key][j]/multiplierlist[j]
                                        
                        # складываем все назад
                        return_dict[key] = (time.gmtime(sum(return_dict[key])*60).tm_hour*60 + time.gmtime(sum(return_dict[key])*60).tm_min)/60
        # print('return_dict 2 >>>>', return_dict,'\n\n')
#------------------------------------------------------------------------------
        # здесь мы определяем корневой/родной/основной контракт, на который будем списывать все недоработанные часы
        # если его нет(командировка, переработка, то ничего списывать не будем)
        
        df3_merged = df_json.loc[(df_json['id_worker'] == id_worker) 
            & (df_json['work_calendar_day'] == work_calendar_day) 
            & (df_json['period_month'] == period_month) 
            ].loc[(((df_json['kontrakt_name'] == 'Общехозяйственный') & (df_json['event'] == 'shift')) | 
            ((df_json['kontrakt_name'] == 'Общехозяйственный') & (df_json['event'] == 'correction'))|
            ((df_json['kontrakt_name'] == 'Офис') & (df_json['event'] == 'shift')) | 
            ((df_json['kontrakt_name'] == 'Офис') & (df_json['event'] == 'correction')))]       
        
        index = df3_merged.index.tolist()
       #--------------------------------------
       #-cписываем
       
        if len(index) != 0:
            for key, val in df_json['work_calendar_day'].loc[(df_json['id_worker'] == id_worker) 
                    & (df_json['work_calendar_day'] == work_calendar_day) 
                    & (df_json['period_month'] == period_month)].to_dict().items():
                # print('key-->', key)
                if key == index[0]:
                    # print('value-->>', val)
                    target_day = val
                    target_index = key
            # sdfsdf
            sumtime = 0
            for key, val in df_json['work_calendar_day'].loc[(df_json['id_worker'] == id_worker) 
                    & (df_json['work_calendar_day'] == work_calendar_day) 
                    & (df_json['period_month'] == period_month)].to_dict().items(): 
                if val == target_day:
                    sumtime = sumtime + return_dict[key]
            
            return_dict[target_index] = 9 - sumtime
        
        df_json['kontrakt_timetracking'].loc[(df_json['id_worker'] == id_worker) 
            & (df_json['work_calendar_day'] == work_calendar_day) 
            & (df_json['period_month'] == period_month)] = list(return_dict.values())
        
        
        # print(df_json.loc[(df_json['id_worker'] == id_worker) 
        #     & (df_json['work_calendar_day'] == work_calendar_day) 
        #     & (df_json['period_month'] == period_month)])
        
        # print()
        # print()
        
        return df_json.loc[(df_json['id_worker'] == id_worker) 
            & (df_json['work_calendar_day'] == work_calendar_day) 
            & (df_json['period_month'] == period_month)]

        
    def time_timetracking(kontrakt_timetracking, kontrakt_name, event, work_calendar_day, id_worker):
        
        # тут есть один неприятный момент - в Json прилетают данные не только за один день
        # а за рандомное кол-во дней назад. надо что-то с этим делать
        # еще ВАЖНЫЙ МОМЕНТ - тот несчастный, кто задумает что-то здесь модифицировать, 
        # знай - повторять структуру DataFrame для результирующего массива, который мы ретёрним НЕНУЖНО! 
        # мы передаём, по большому счету, Series. если простыми словами - индекс: значение
        # первым делом мы проходимся по нашим контрактам и словарь с временными метками заменяем на массив с единичками
        
        # print('df_json-=--==-==-=-=-=->>>>>>>>>>', df_json)
        # print('id_worker.drop_duplicates()--->>', id_worker.drop_duplicates())
        # print('id_worker--->>', id_worker)
        
        kontraktdict = kontrakt_timetracking.to_dict()
        # print('контракты нового дня kontraktdict-->>>', kontraktdict)
        return_dict = {}
        # в качестве аргумента прилетает список словарей списков словарей.
        # в этих циклах мы докапываемся до сути - списка с временными отметками
        # print('-->>', kontraktdict.items())
        for i, v in kontraktdict.items(): #for i, v in range(len(kontraktdict)):
            if kontraktdict[i] == 'Nan':
                kontraktdict[i] = [{'0': [0, 0]}]
        # print('контракты нового дня kontraktdict-->>>', kontraktdict)
        
        
        # здесь мы наши time_tracking превращаем в массивы с единичками
        for i in kontraktdict:
            # print('in I--->', i)
            for j in kontraktdict[i]:
                # на каждом новом проходе формируем новый массив с нулями
                min_list = [0 for i in range(0, 1440)]
                
                # заполняем список единицами, там, где нашлось время (тут мы переводим unix time в нормальные минуты и на каждую минуту ставим "1")
                start_min = time.gmtime(list(j.values())[0][0]).tm_hour*60+time.gmtime(list(j.values())[0][0]).tm_min
                stop_min =time.gmtime(list(j.values())[0][1]).tm_hour*60+time.gmtime(list(j.values())[0][1]).tm_min
                for t in range(start_min, stop_min):
                    min_list[t] = 1
                
                # складываем все назад
                j[list(j.keys())[0]] = min_list
                
        # print('kontraktdict=====>', kontraktdict)
        
        # print('df_json-->>', df_json.loc[(df_json['id_worker'] == 100) ])
        # print('id_worker.drop_duplicates()--->>', id_worker.drop_duplicates())
        # print('id_worker--->>', id_worker)
        
        
        
        
        
        # здесь мы сложили все единички в каждом контракте в один массив, 
        # таким образом мы получили четкое распределение в этом контракте когда человек работал
        id_worker_dict = id_worker.to_dict()
        # print('id_worker_dict--->>', id_worker_dict)
        # for work in id_worker.drop_duplicates():
        #     print('work-->>', work)
        for i in kontraktdict:
            # print('i-->>', i)
            # if id_worker_dict[i] == work:
            # print('id_worker_dict[i]--->', id_worker_dict[i])
            # print('in I--->', i)
            summin_list = [0 for i in range(0, 1440)]
            for j in range(len(kontraktdict[i])):
                # print('j-->>', j)
                # print('len(kontraktdict[i])-->>', len(kontraktdict[i]))
                # print('kontraktdict[i]-->>', kontraktdict[i])
                for si in range(len(summin_list)):
                    # print('si-->>', si)
                    # print('kontraktdict[i][j]-->>>', list(kontraktdict[i][j].values())[0][si])
                    summin_list[si] = summin_list[si] + list(kontraktdict[i][j].values())[0][si]
                # print('summin_list-->>', summin_list)
            return_dict[i] = summin_list
        # print('return_dict-->', return_dict)
        
        # тут хуйня!!!! все работает слишьком сложно и непонятно как
        
        # теперь мы должны каким-то образом определить контракты в каждом дне и все скалькулировать
        # а именно мы должны сложить все единички из каждого TimeTracking в каждом дне по всем контрактам и каждый TimeTracking разделить на эту сумму
        # после этого все сложить, превратив в доли часов
        
        # print('df_json[work_calendar_day]-->>', df_json['work_calendar_day'])
        
        # print('list-->>', list(df_json['work_calendar_day'].drop_duplicates().to_dict().values()))
        
        # print('id_worker.drop_duplicates()--->', id_worker.drop_duplicates().sort_values())
        
        # !!! скорее всего тут то же не хватает контекста месяца и года, но как его сюда впихнуть - ХЗ
        for df_day in list(df_json['work_calendar_day'].drop_duplicates().to_dict().values()):
            # print('df_day-->>', df_day)
            for work in id_worker.drop_duplicates().sort_values():
                multiplierlist = [0 for i in range(0, 1440)]
                for key, val in df_json['work_calendar_day'].to_dict().items():
                    
                    if id_worker_dict[key] == work:   
                        if val == df_day:
                            multiplierlist = list(map(sum, zip(multiplierlist, return_dict[key])))
                    # тут делим единицы на их сумму
                for key, val in df_json['work_calendar_day'].to_dict().items():
                    # print('key-->>', key)
                    # print('val-->>', val)
                    if id_worker_dict[key] == work:
                        if val == df_day:
                            for j in range(len(return_dict[key])):
                                if multiplierlist[j] != 0:
                                    return_dict[key][j] = return_dict[key][j]/multiplierlist[j]
                                            
                                            # складываем все назад
                            return_dict[key] = (time.gmtime(sum(return_dict[key])*60).tm_hour*60 + time.gmtime(sum(return_dict[key])*60).tm_min)/60
        # print('return_dict--->>>', return_dict)

        
        # здесь мы определяем корневой/родной/основной контракт, на который будем списывать все недоработанные часы
        # df3_merged = pd.merge(kontrakt_name, event, left_index=True, right_index=True)
        
        # print('df3_merged-->>', df3_merged)
        # print('df_json-->', df_json)
        df3_merged = df_json.loc[(
                                ((df_json['kontrakt_name'] == 'Общехозяйственный') & (df_json['event'] == 'shift') & (df_json['kontrakt_issuecount'] == 0)) | 
                                ((df_json['kontrakt_name'] == 'Общехозяйственный') & (df_json['event'] == 'correction') & (df_json['kontrakt_issuecount'] == 0))|
                                ((df_json['kontrakt_name'] == 'Офис') & (df_json['event'] == 'shift') & (df_json['kontrakt_issuecount'] == 0)) | 
                                ((df_json['kontrakt_name'] == 'Офис') & (df_json['event'] == 'correction') & (df_json['kontrakt_issuecount'] == 0))
                              )]
        
        # !!!!!!!! здесь выяснилась такая неприятная история - если у работника появилась заявка на один из основных контрактов, то скрипт путается. 
        # к примеру Егоров с "Общехозом" выполняет заявку на "офис" и все идет по пизде. 
        # ниже написанная функция определения основного контракта написано неудачно, в ней нет контекста таймтрекинга, что бы действительно определять 
        # основной контракт от неосновного. 
        # нужно менять подход. как и на какой его менять не совсем понятно
        
        index = df3_merged.index.tolist()
        # print('index++++++++===>', index)
        # print('=================')
        
        # dfs = [total,truck,passenger_car,noweight,zerroweight,badcalc,badmetric, notvalidweight,notvalidsize, notvalidfull, notvalidtruck,  minlength, maxlength, height_zero, lenght_zero, maxemptycount, norecognize_front, norecognize_rear, norecognizefull,recognize[['RecognizeFront', 'RecognizeRear', 'Dev']], tempasp]

        # df_final = ft.reduce(lambda left, right: pd.merge(left, right, on='Dev'), dfs)

                # ---------------------------------------------------------------------------------------------------------
        # теперь мы складываем все часы в других контрактах, отнимаем от 9 и результат пишем в словарь return_dict в соответствующий индекс
        # тут очень важный косяк!!!! я отнимаю сумму времени от 9, что неверно. я должен отнимать не от 9, а от worktime. причем ворктайм надо брать полный 12 часов или 9
        # а откуда брать его непонятно, если у меня есть только корректировка!
        # print('--->>>>>', df_json['work_calendar_day'])
        for key, val in df_json['work_calendar_day'].to_dict().items():
            # print('key-->', key)
            if key == index[0]:
                # print('value-->>', val)
                target_day = val
                target_index = key
                # print('target_index----->>>>', target_index)
        sumtime= 0
        for key, val in df_json['work_calendar_day'].to_dict().items():
            if val == target_day:
                sumtime = sumtime + return_dict[key]
        # print('parent_kontrakt-->>', sumtime)
        return_dict[target_index] = 9 - sumtime
        # print('return_dict--->>>', return_dict)
        return return_dict.values()
    # ---------------------------------------------------------------------------------------------------------------------------------
    
    #нижеследующуюий блок вызывает функцию time_tracking_2()
    
    #функция по сути дублирует time_timetracking()
    #была написана в отчаянье, от непонимания проблемы 9, имеет место быть

    # # ========================================
    dff = pd.DataFrame(columns=['event','day_night','work_time','work_calendar_day','work_calendar_daytype','job_name','job_department','job_position','period_month','period_year','kontrakt_name','kontrakt_issuecount','kontrakt_filter','id_kontrakts','id_worker','kontrakt_timetracking'])
    # dfff = pd.DataFrame(columns=['event','day_night','work_time','work_calendar_day','work_calendar_daytype','job_name','job_department','job_position','period_month','period_year','kontrakt_name','kontrakt_issuecount','kontrakt_filter','id_kontrakts','id_worker','kontrakt_timetracking'])
    for i in df_json['period_year'].drop_duplicates():
        # print(i)
        for j in df_json['period_month'].loc[(df_json['period_year'] == i)].drop_duplicates():
            # print(j)
            for l in df_json['work_calendar_day'].loc[(df_json['period_year'] == i) & (df_json['period_month'] == j)].drop_duplicates():
                # print(l)
                for k in df_json['id_worker'].loc[(df_json['period_year'] == i) & (df_json['period_month'] == j) & (df_json['work_calendar_day'] == l)].drop_duplicates():
                    dff = dff.append(time_tracking_2(df_json, dff, k, l, j, i))
    # print(dff.loc[['event','work_time']])#'work_calendar_day','job_name','job_position','period_month','kontrakt_name','kontrakt_filter','id_kontrakts','id_worker','kontrakt_timetracking']])               
    # print(dff[['event','kontrakt_timetracking','job_name','kontrakt_name','work_calendar_day']])
    df_json = dff
    
    # =======================================
    
    # print('df_json in calc >>>',df_json[['event','kontrakt_timetracking','job_name','kontrakt_name','work_calendar_day']])
    
    
    # df_json = df_json.assign(timetrackng = lambda x: time_timetracking(x['kontrakt_timetracking'], x['kontrakt_name'], x['event'], x['work_calendar_day'], x['id_worker']))
    # df_json = df_json.drop(['kontrakt_timetracking'], axis=1)
    # df_json = df_json.rename(columns={'timetrackng': 'kontrakt_timetracking'})
    
    
    # print()
    # print()
    # print('df_json after sum>>>',df_json[['event','kontrakt_timetracking','job_name','kontrakt_name','work_calendar_day','kontrakt_issuecount']])
    # # dfs
    # # print(df_json[['event','kontrakt_timetracking','job_name','kontrakt_name','work_calendar_day']])
    # print(df_json[['event','kontrakt_timetracking','job_name','kontrakt_name','work_calendar_day']])
    
    # есть мнение что именно здесь нужно дополнительно повозиться с DataFrame
    # а именно, избавиться от переработок, которые идут вместе с шифтами или корректировками
    # + перенести время переработок из work_time в kontrakt_timetracking
    # ++ надо добавить перенос времени из work_time в kontrakt_timetracking для командировок. но тут вопрос - должно ли стоять в качестве времени 8 часов? может 9???
    
    # -------------------- перенос времени командировок -------------------
    # выясняем дни, работников и все правильно делим 
    df_json_trip = df_json.loc[(df_json['event'] == 'trip')].to_dict()
    # здесь надобавлял везде контекс месяца, что бы командировки в разных месяцах не путались
    for period_month in list(set(df_json_trip['period_month'].values())):
        for id_worker in list(set(df_json_trip['id_worker'].values())):
            for day in list(set(df_json_trip['work_calendar_day'].values())):
                search_trip = df_json.loc[(df_json['id_worker'] == id_worker) & (df_json['event'] == 'trip') & (df_json['work_calendar_day'] == day) & (df_json['period_month'] == period_month)].to_dict()
                for trip in search_trip['event'].keys():
                    df_json.loc[[trip], ['kontrakt_timetracking']] = search_trip['work_time'][trip]/len(df_json.loc[(df_json['id_worker'] == id_worker) & (df_json['event'] == 'trip') & (df_json['work_calendar_day'] == day) & (df_json['period_month'] == period_month)])
                    
    # ---------------------------------------------------------------------
 
    # --- переносим часы в переработках----------------------------------------
    # здесь дополнительно делаем проверку и делим время на кол-во конрактов. 
    # тема такая - переработку можно оформить на несколько контрактов сразу, а в Json мне приходит оно время на все контракты
    # тупо 3 контракта и один work_time. это значит, что чуваки работали, к примеру 2 часа но по всем контрактам, а не по 2 часа на каждом из них.
    # и значит мы должны взять work_time и поделить его ровненько между всеми контрактами
    
    # готовим список с человеками
    # print('search worker-->>', list(df_json['id_worker'].loc[(df_json['event'] == 'permit')].drop_duplicates().to_dict().values()))
    # готовим список с kontrakt_filter
    # print('search worker-->>', list(df_json['kontrakt_filter'].loc[(df_json['event'] == 'permit')].drop_duplicates().to_dict().values()))
    
    # выясняем кол-во контрактов на одном документе
    # print('search-->', df_json.loc[(df_json['id_worker'] == 3) & (df_json['kontrakt_filter'] == 'DOCCORP-16511')& (df_json['event'] == 'permit')].drop_duplicates(subset = ['kontrakt_name']).count())
    
    for worker in list(df_json['id_worker'].loc[(df_json['event'] == 'permit')].drop_duplicates().to_dict().values()):
        
        for kontrakt_filter in list(df_json['kontrakt_filter'].loc[(df_json['event'] == 'permit')].drop_duplicates().to_dict().values()):
            
            count = df_json.loc[(df_json['id_worker'] == worker) & (df_json['kontrakt_filter'] == kontrakt_filter)& (df_json['event'] == 'permit')].drop_duplicates(subset = ['kontrakt_name']).count().to_dict()
            # print('c-->', count['id_worker'])
            # print('df_json-->', df_json)
            df_json['work_time'].loc[(df_json['id_worker'] == worker) & (df_json['kontrakt_filter'] == kontrakt_filter)& (df_json['event'] == 'permit')] /= count['id_worker']
            # print('df_json-->', df_json)
   
    
    df_json_permit = df_json.loc[(df_json['event'] == 'permit')].to_dict()
    
    # кроме всего прочего, нам нужно делить время между переработками на разных контрактах
    # и складывать, если контракт один и удалять лишьние строки с пермитами
    
    # здесь переносим время в переработках из work_time в kontrakt_timetracking
    
    for permit in df_json_permit['event'].keys():
        df_json.loc[[permit], ['kontrakt_timetracking']] = df_json_permit['work_time'][permit]
        
        
    # print(df_json[['event','kontrakt_timetracking','job_name','kontrakt_name','work_calendar_day']])
    
        
    # -------------------------------------------------------------------------
    
    # перед тем, как фиксить лишьние переработки, нам нужно немного поправить время у шифтов/переработок.
    # меня не поняли и коллегиально надовили... короче, в Json приходит максимальное время по TimeTracking 9 часов (с 9:00 до 18:00)
    # нужно время привезти к 8-и часам (типа час на обед, который неоплачивается)
    # общая логика такая - если есть время на общехозяйственном/офис контракте, то отнимаем максимально возможное от него,
    # остаток пропорционально отнимаем от остальных заявок. 
    # хочу здесь добавить, что очень хочется отнимать не от всех заявок тупо по, к примеру 10 минут, 
    # а как-то по-умному, больше времени отнимать от больших заявок и меньше от коротких, причем именно здесь надо вычислить пропорцию, что не просто....
    # после не долгих рассуждений выродилась формула, которая работает, только если на общехозяйственном после вычитания нужного кол-ва времени получилось < 0: 
        # x*(y/z)=w, где: 
                        # x (delta_time) - остаток от вычитания времени от основного контракта (x = √(sqr(общехоз. - (z - work_time))))
                        # y - кол-во часов в дне на контракте
                        # z - сумма всего затраченного на работу времени по заявкам на всех контрактах в дне (общехозяйственный не учитываем)
                        # w - кол-во часов, которое нужно отнять от суммы времени на контракте в дне (общехозяйственный не учитываем, от него отнимать ничего не надо)
    
    # !!!!!!! с ? по ? строку происходит хуйня - если работник получил таймтрекинг по одному из основных контрактов, то скрипт ломается (в таймстрекинг пишется Nan)
    # готовим циклы по людям и дням
    
    for i in df_json['kontrakt_issuecount']:
        df_json['kontrakt_issuecoun'] = 0

    
    df_json = df_json.fillna({'kontrakt_issuecount':0})
    
    #ВНИМАНИЕ!!!!
    #в некоторых моментах был полностью убран kontrakt issuecount, таким образом удалось фиксануть проблему связанную с 9
    #возможно это вызовет непредсказуемые последствия, но сейчас всё ок :)
    
    
    df_json_shift = df_json.loc[(df_json['event'] == 'shift') | (df_json['event'] == 'correction')]#.to_dict()
        
    for year in list(set(df_json_shift['period_year'].drop_duplicates().to_dict().values())):
        for month in list(df_json_shift['period_month'].loc[(df_json_shift['period_year'] == year)].drop_duplicates().to_dict().values()):
            for day in list(df_json_shift['work_calendar_day'].loc[(df_json_shift['period_year'] == year) & (df_json_shift['period_month'] == month)].drop_duplicates().to_dict().values()):
                for id_worker in list(df_json_shift['id_worker'].loc[(df_json_shift['period_year'] == year) & (df_json_shift['period_month'] == month) & (df_json_shift['work_calendar_day'] == day)].drop_duplicates().to_dict().values()):
                    
                    alltime = df_json['kontrakt_timetracking'].loc[
                                                                   (df_json['id_worker'] == id_worker) & 
                                                                   ((df_json['event'] == 'shift') | (df_json['event'] == 'correction')) & 
                                                                   (df_json['period_year'] == year) & (df_json['period_month'] == month) & 
                                                                   (df_json['work_calendar_day'] == day)
                                                                   ].sum()
                    
                    # при дебаге выяснилось, что тут есть косячина косячная. 
                    # если в этом дне хоть у какого-нибудь сотрудника есть шифт, этот день обязательно будет в списке дней на обработку
                    # и значит, что если у другого в этот день отпуск или больничный (а он на основном контракте), скрипт ломается, потому что all_time и work_time пустые
                    # значит нужно сделать проверку. добавим коротенький if ))) 
                    # обещаю, потом все переписать )))
                    if len(df_json.loc[
                                       (df_json['id_worker'] == id_worker) & 
                                       ((df_json['event'] == 'shift') | (df_json['event'] == 'correction')) & 
                                       (df_json['period_year'] == year) & 
                                       (df_json['period_month'] == month) & 
                                       (df_json['work_calendar_day'] == day)
                                       ]) > 0:
                        
                        work_time = df_json['work_time'].loc[
                                                             (df_json['id_worker'] == id_worker) & 
                                                             ((df_json['event'] == 'shift') | (df_json['event'] == 'correction')) & 
                                                             (df_json['period_year'] == year) & 
                                                             (df_json['period_month'] == month) & 
                                                             (df_json['work_calendar_day'] == day) & 
                                                             ((df_json['kontrakt_name'] == 'Общехозяйственный') | (df_json['kontrakt_name'] == 'Офис'))
                                                             # & (df_json['kontrakt_issuecount'] == 0)
                                                             ].values#.to_dict()
                        
                        # extra_time - результат вычисления. то, сколько часов нужно срезать с контрактов
                        extra_time = alltime - work_time[0]
                        # base_kontrakt_time - время, которое записано на основной контракт (общехозяйственный или офис)
                        

                        base_kontrakt_time = df_json['kontrakt_timetracking'].loc[
                                                                                  (df_json['id_worker'] == id_worker) & 
                                                                                  ((df_json['event'] == 'shift') | (df_json['event'] == 'correction')) & 
                                                                                  (df_json['period_year'] == year) & 
                                                                                  (df_json['period_month'] == month) & 
                                                                                  (df_json['work_calendar_day'] == day) & 
                                                                                  (((df_json['kontrakt_name'] == 'Общехозяйственный') | (df_json['kontrakt_name'] == 'Офис')))
                                                                                  # & (df_json['kontrakt_issuecount'] == 0)
                                                                                  ].values
                      
                        if (base_kontrakt_time[0] - extra_time) < 0:
                            # присваиваем основному контракту ноль
                            df_json['kontrakt_timetracking'].loc[
                                                                 (df_json['id_worker'] == id_worker) & 
                                                                 ((df_json['event'] == 'shift') | (df_json['event'] == 'correction')) & 
                                                                 (df_json['period_year'] == year) & 
                                                                 (df_json['period_month'] == month) & 
                                                                 (df_json['work_calendar_day'] == day) & 
                                                                 (((df_json['kontrakt_name'] == 'Общехозяйственный') | (df_json['kontrakt_name'] == 'Офис')))
                                                                 # & (df_json['kontrakt_issuecount'] == 0)
                                                                 ] = 0
                           
                            # возводим ранее вычисленное кол-во часов в квадрат и вычитаем корень квадратный. нужно что бы мы не попали на комплексные числа
                            delta_time = np.sqrt(np.power(base_kontrakt_time[0] - extra_time, 2))
                            # подсчитываем сумму времени на контрактах. это не всегда равно 9, ведь иногда на общехозяйственном будут какие-то часы
                            all_kontrakt_time = df_json['kontrakt_timetracking'].loc[
                                                                                     (df_json['id_worker'] == id_worker) & 
                                                                                     ((df_json['event'] == 'shift') | (df_json['event'] == 'correction')) & 
                                                                                     (df_json['period_year'] == year) & 
                                                                                     (df_json['period_month'] == month) & 
                                                                                     (df_json['work_calendar_day'] == day) & 
                                                                                     (((df_json['kontrakt_name'] != 'Общехозяйственный') | (df_json['kontrakt_name'] != 'Офис')))
                                                                                      # & (df_json['kontrakt_issuecount'] == 0)
                                                                                     ].sum()
                            
                           
                            # бежим по каждому контракту и получаем затраченное на него время
                            # для этого получаем список контрактов
                            all_kontrakts= df_json['kontrakt_name'].loc[
                                                                        (df_json['id_worker'] == id_worker) & 
                                                                        ((df_json['event'] == 'shift') | (df_json['event'] == 'correction')) & 
                                                                        (df_json['period_year'] == year) & 
                                                                        (df_json['period_month'] == month) & 
                                                                        (df_json['work_calendar_day'] == day) & 
                                                                        # ((df_json['kontrakt_name'] != 'Общехозяйственный') | (df_json['kontrakt_name'] != 'Офис'))
                                                                        (((df_json['kontrakt_name'] != 'Общехозяйственный') | (df_json['kontrakt_name'] != 'Офис')))
                                                                        ].to_dict()

                            for kont in all_kontrakts:
                               
                                kont_time = df_json['kontrakt_timetracking'].loc[kont]
                               
                                overtime = delta_time * (df_json['kontrakt_timetracking'].loc[kont] / all_kontrakt_time)
                                
                                df_json['kontrakt_timetracking'].loc[kont] = df_json['kontrakt_timetracking'].loc[kont] - overtime
                               
                                
                        else:
                            df_json['kontrakt_timetracking'].loc[
                                                                 (df_json['id_worker'] == id_worker) & 
                                                                 ((df_json['event'] == 'shift') | (df_json['event'] == 'correction')) & 
                                                                 (df_json['period_year'] == year) & 
                                                                 (df_json['period_month'] == month) & 
                                                                 (df_json['work_calendar_day'] == day) & 
                                                                 ((df_json['kontrakt_name'] == 'Общехозяйственный') | (df_json['kontrakt_name'] == 'Офис'))
                                                                 ] = base_kontrakt_time[0] - extra_time
                  
    # print()
    # print()                        
    # print('df after fi1x >>',df_json[['event','kontrakt_timetracking','job_name','kontrakt_name','work_calendar_day','kontrakt_issuecount']])
    # ---- здесь избавляемся от переработки, которая вместе с шифтом или корректировкой--
    # нам нужно получить DataFrame в котором бы была и переработка и шифт/корректировка
    for permit in df_json_permit['event'].keys():

        df_json_calc = df_json.loc[(df_json_permit['day_night'][permit] == df_json['day_night']) & 
                        (df_json_permit['work_calendar_day'][permit] == df_json['work_calendar_day']) & 
                        (df_json_permit['job_name'][permit] == df_json['job_name']) & 
                        (df_json_permit['job_department'][permit] == df_json['job_department']) & 
                        (df_json_permit['job_position'][permit] == df_json['job_position']) & 
                        (df_json_permit['period_month'][permit] == df_json['period_month']) & 
                        (df_json_permit['period_year'][permit] == df_json['period_year']) & 
                        ((df_json['event'] == 'correction') | (df_json['event'] == 'shift')) & 
                        (df_json_permit['kontrakt_name'][permit] == df_json['kontrakt_name'])].append(
                            df_json.loc[(df_json_permit['day_night'][permit] == df_json['day_night']) & 
                                            (df_json_permit['work_calendar_day'][permit] == df_json['work_calendar_day']) & 
                                            (df_json_permit['job_name'][permit] == df_json['job_name']) & 
                                            (df_json_permit['job_department'][permit] == df_json['job_department']) & 
                                            (df_json_permit['job_position'][permit] == df_json['job_position']) & 
                                            (df_json_permit['period_month'][permit] == df_json['period_month']) & 
                                            (df_json_permit['period_year'][permit] == df_json['period_year']) & 
                                            (df_json['event'] == 'permit') & 
                                            (df_json_permit['kontrakt_name'][permit] == df_json['kontrakt_name'])], ignore_index=False)
                            
       
        if len(df_json_calc) > 1:
            # df_json_shift_inner находится индекс нашего шифта или корректировки
            df_json_shift_inner = list(df_json_calc.loc[(df_json_calc['event'] == 'shift') | (df_json_calc['event'] == 'correction')].index)
            
            # df_json_permit_inner находится вся стата по переработке
            df_json_permit_inner = df_json_calc.loc[(df_json_calc['event'] == 'permit')].to_dict()
            
            for permit_inner in df_json_permit_inner['event'].keys():
               
                df_json.loc[[df_json_shift_inner[0]], ['kontrakt_timetracking']] = df_json.loc[[df_json_shift_inner[0]], ['kontrakt_timetracking']] + df_json_permit_inner['kontrakt_timetracking'][permit_inner]
                
                df_json.loc[[df_json_shift_inner[0]], ['kontrakt_filter']] = df_json.loc[[df_json_shift_inner[0]], ['kontrakt_filter']] + ',' + df_json_permit_inner['kontrakt_filter'][permit_inner]
                
                # удаляем наш уже ненужный permit
                df_json = df_json.drop(permit_inner)
    # -------------------------------------------------------------------------------------
    
    # --- удилим лишнее из DataFrames, оставим только данные для добавления в БД ---
    # ----------------- получаем в DataFrame данные из БД clocks--------
    df_db_clocks = pd.read_sql("""SELECT *
                      FROM clocks 
                      """, con=engine, index_col='index')
    dbclocks = df_db_clocks.to_dict('records') # split, records
    # здесь удаляем "левые" столбцы из DataFrame потому что они мешают алхимии делать инсерты
    
    jclocks = df_json.to_dict('records')
    
    inclocks=[]
    for j in range(len(jclocks)):
        check = 0
        for i in range(len(dbclocks)):
            if ((dbclocks[i]['id_worker'] == jclocks[j]['id_worker']) &
            (dbclocks[i]['id_kontrakts'] == jclocks[j]['id_kontrakts']) &
            (dbclocks[i]['period_year'] == jclocks[j]['period_year']) &
            (dbclocks[i]['period_month'] == jclocks[j]['period_month']) &
            (dbclocks[i]['work_calendar_day'] == jclocks[j]['work_calendar_day']) &
            (dbclocks[i]['work_calendar_daytype'] == jclocks[j]['work_calendar_daytype']) &
            (dbclocks[i]['work_time'] == jclocks[j]['work_time']) &
            # (dbclocks[i]['kontrakt_issuecount'] == jclocks[j]['kontrakt_issuecount']) &
            (dbclocks[i]['kontrakt_filter'] == jclocks[j]['kontrakt_filter']) &
            (dbclocks[i]['event'] == jclocks[j]['event']) &
            # (dbclocks[i]['kontrakt_timetracking'] == jclocks[j]['kontrakt_timetracking']) &
            (dbclocks[i]['day_night'] == jclocks[j]['day_night'])
            ):
                check = 1
                if check == 1:
                    break
        if check == 0:
            inclocks.append(jclocks[j])
    jclocks = None
    dbclocks = None
    
    # ------------- инсертим в БД наш список, если он не пустой ---------------
    if len(inclocks) > 0:
        # print(inclocks)
        conn.execute(insert(clocks),inclocks)
    inclocks = None
    return df_json
#------------------------------------------------------------------------------
def calcTotal(df_json, total, engine, df_base, conn):
    # ===== здесь займемся DataFrame Total и в сем, что с ним связано =====
    # !!!!
    # здесь косяк!!!!! нужно добавлять контекст месяца, года
    # более того. эти тоталики плохо вставляются в excel и в февраль (т.к. они там затерлись), пишутся мартовские ))))) это пиздец залет....
    # work = df_json[['id_worker','work_calendar_day','period_month']].drop_duplicates().to_dict()
    
    # из таблицы Total в БД удаляем все тоталики за пришедшие в Json даты на пришедших сотрудников
    # dele = total.delete().where(total.c.work_calendar_day.in_(list(set(work['work_calendar_day'].values()))) & total.c.id_worker.in_(list(set(work['id_worker'].values()))) & total.c.period_month.in_(list(set(work['period_month'].values()))))
    
    
    work = df_json[['id_worker','work_calendar_day', 'period_month']].drop_duplicates().to_dict('split') 
    work = work["data"]  
    sql = ''
    for i in work:
        strochka = '(total.id_worker = ' + str(i[0]) + ' AND total.work_calendar_day = ' + str(i[1]) + ' AND total.period_month = "' + str(i[2]) + '")'
        sql = sql + strochka + ' or '
    sql = sql[:-4]
    dele = 'DELETE FROM total WHERE '+ sql
    # print(dele)
    # dfsdf
    engine.execute(dele)
    
    
    # ищем в пришедшем Json шифты, корректировки и прочее и записываем это в DataFrame
    df_total = df_json.loc[(df_json['event'] == 'shift') | (df_json['event'] == 'correction') | (df_json['event'] == 'trip') | (df_json['event'] == 'hospital') | (df_json['event'] == 'otpusk') | (df_json['event'] == 'compensatory')].copy()
    # print('df_total--->>>',df_total.loc[df_total['id_worker'] == 49])
    
    # здесь нам пришлось сделать два почти одинаковых DataFrame. один для инсерта в базу, второй для добавления в Excel. дело в том, что поля, которые мы выводим в excel и которые храним в БД разные
    
    df_total.drop_duplicates(subset=['event', 'day_night', 'work_time', 'work_calendar_day', 'work_calendar_daytype', 'period_month', 'period_year', 'id_worker'], inplace=True)
    
    # здесь должно быть что-то вроде: находим все пермиты сотрудника в дне и к шифтам/корректировкам/трипам (к их ворктайму) прибавляем ворктайм пермитов
    dict_total = df_total.to_dict('records')
    
    for d_t in dict_total:
        df_json_1 = df_base.loc[# df_json заменил на df_base
                        (d_t['work_calendar_day'] == df_base['work_calendar_day']) & 
                        (d_t['job_name'] == df_base['job_name']) & 
                        (d_t['job_department'] == df_base['job_department']) & 
                        (d_t['job_position'] == df_base['job_position']) & 
                        (d_t['period_month'] == df_base['period_month']) & 
                        (d_t['period_year'] == df_base['period_year']) & 
                        (df_base['event'] == 'permit')]
        
        d_t['work_time'] = d_t['work_time'] + df_json_1['work_time'].sum()
        
        # эта конструкция вызывает у Pandas вопросы. Выдается предупреждение:
        # C:\Users\Pafnuty\anaconda3\lib\site-packages\pandas\core\indexing.py:1732: SettingWithCopyWarning: 
        # A value is trying to be set on a copy of a slice from a DataFrame
        # See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
        # self._setitem_single_block(indexer, value, name)
        # которое говорит нам, что неочень хорошо искать по фрэйму данных и менять его налету в одно движение ))))
        
        df_total['work_time'].loc[(df_total['work_calendar_day'] == d_t['work_calendar_day']) &
                     (df_total['work_calendar_daytype'] == d_t['work_calendar_daytype']) &
                     (df_total['job_name'] == d_t['job_name']) &
                     (df_total['job_department'] == d_t['job_department']) &
                     (df_total['job_position'] == d_t['job_position']) &
                     (df_total['period_month'] == d_t['period_month']) &
                     (df_total['period_year'] == d_t['period_year']) &
                     (df_total['id_worker'] == d_t['id_worker'])] = d_t['work_time']
        
    # -------------------------------------------------------------------------
    df_total_ins = df_total.copy()
    
    df_total.drop(['event', 'day_night', 'kontrakt_name', 'kontrakt_filter', 'kontrakt_timetracking', 'id_kontrakts', 'kontrakt_issuecount'], axis='columns', inplace=True)
   
    df_total_ins.drop(['event', 'day_night', 'job_name', 'job_department', 'job_position', 'kontrakt_name', 'kontrakt_filter', 'kontrakt_timetracking', 'id_kontrakts'], axis='columns', inplace=True)

    dict_total_ins = df_total_ins.to_dict('records')
    
    if len(dict_total_ins) > 0:
        conn.execute(insert(total),dict_total_ins)
    dict_total_ins = None
    
    # когда я получаю json за определенную дату, total имеет в себе только эту дату. 
    # нужно поправить и дополнительно к этой дате подмахнуть все тоталики, что есть в БД
    totaldb = pd.read_sql("""SELECT tl.index, tl.work_time, tl.work_calendar_day, tl.work_calendar_daytype, wk.job_name, wk.job_department, wk.job_position, tl.period_month, tl.period_year, tl.id_worker
                          FROM total tl
                          LEFT JOIN workers wk ON (tl.id_worker=wk.index) 
                          """, con=engine, index_col='index')
    df_total = df_total.append(totaldb, ignore_index=True) 
    
    # -------------------------------------------------------------------------
    
    return df_total
