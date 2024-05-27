# Стандартные библиотеки
import ast
# Сторонние библиотеки
import pandas as pd
from sqlalchemy import insert
from normalize import js_norm
import time

def processJson(data, engine, conn, workers, kontrakts, raw, clocks):
        #----------------------------------------------------------------------
        #вводим параметры под формирование датафрейма
        
        # print(data)
        
        list_path = ["days"]
        
        list_meta = [["job","name"],["job","department"],["job","position"],['period', 'month'], ["period", "year"]]
        
        cols_to_norm = ["kontrakt"]
        
        dict_fillna = {'kontrakt_issuecount': 0,'kontrakt_filter': '0', 'kontrakt_timetracking': 'Nan'}
        
        merge_method = 'cross'
        #----------------------------------------------------------------------
        #формируем датафрейм
        
        df_json = js_norm(data, list_path, list_meta, cols_to_norm, dict_fillna, merge_method)

        df_json.drop_duplicates(subset=['event','day_night','work_time','work_calendar_day','work_calendar_daytype','job_name','job_department','job_position','period_month','period_year','kontrakt_name','kontrakt_issuecount','kontrakt_filter'], inplace=True)
        
        # print('df_json >>>>>', df_json[['event','kontrakt_timetracking','job_name','kontrakt_name','work_calendar_day']],'\n\n')
        #----------------------------------------------------------------------
        
        if 'kontrakt_timetracking' in df_json.columns:
            pass
        else:
            df_json['kontrakt_timetracking'] = 'Nan'#np.nan
        
        # ====== пишем в БД сотрудников из JSON ===============================
            
        # ------- в этом DataFrame содержится список работников из JSON -----------
        df_json_workers=pd.concat([df_json['job_name'], df_json['job_department'], df_json['job_position']], axis=1).drop_duplicates(keep='first')
        jworkers = df_json_workers.to_dict('records')
        # -------------------------------------------------------------------------
            
        # ----------------- получаем в DataFrame данные о работниках из БД --------
        df_db_workers = pd.read_sql("""SELECT *
                              FROM workers 
                              """, con=engine, index_col='index')
        dbworkers = df_db_workers.to_dict('records') # split, records
        # -------------------------------------------------------------------------
            
        # --- удилим лишнее из DataFrames, оставим только список работников для добавления в БД ---
        inworkers=[]
        for j in range(len(jworkers)):
            check = 0
            for i in range(len(dbworkers)):
                if (dbworkers[i]['job_name'] == jworkers[j]['job_name']) & (dbworkers[i]['job_department'] == jworkers[j]['job_department']) & (dbworkers[i]['job_position'] == jworkers[j]['job_position']):
                    check = 1
            if check == 0:
                inworkers.append(jworkers[j])
        jworkers = None
        dbworkers = None
        # ----------------------------------------------------------------------------------------
            
        # ------------- инсертим в БД наш список, если он не пустой ---------------
        if len(inworkers) > 0:
            # print(inworkers)
            conn.execute(insert(workers),inworkers)
            pass
        inworkers = None
        # -------------------------------------------------------------------------
        # =========================================================================
        # ====== пишем в БД контракты из JSON =====================================
        
        # ------- в этом DataFrame содержится список контрактов из JSON -----------
        df_json_kontrakts=pd.concat([df_json['kontrakt_name']], axis=1).drop_duplicates(keep='first')
        jkontrakts = df_json_kontrakts.to_dict('records')
        # -------------------------------------------------------------------------
        
        # ----------------- получаем в DataFrame данные о контрактах из БД --------
        df_db_kontrakts = pd.read_sql("""SELECT *
                              FROM kontrakts 
                              """, con=engine, index_col='index')
        dbkontrakts = df_db_kontrakts.to_dict('records') # split, records

        # -------------------------------------------------------------------------
            
        # --- удилим лишнее из DataFrames, оставим только список контрактов для добавления в БД ---
        inkontrakts=[]
        for j in range(len(jkontrakts)):
            check = 0
            for i in range(len(dbkontrakts)):
                if dbkontrakts[i]['kontrakt_name'] == jkontrakts[j]['kontrakt_name']:
                    check = 1
            if check == 0:
                inkontrakts.append(jkontrakts[j])
        jkontrakts = None
        dbkontrakts = None
        # -----------------------------------------------------------------------------------------
        
        # ------------- инсертим в БД наш список, если он не пустой ---------------
        if len(inkontrakts) > 0:
            conn.execute(insert(kontrakts),inkontrakts)
            pass
        inkontrakts = None
        # -------------------------------------------------------------------------
        # =========================================================================
        
        # ==== добиваем наш DataFrame id-шниками контрактов и работников ======
        
        # спрашиваем у базы все контракты и всех работников и подставим из в DataFrame
        
        df_db_kontrakts = pd.read_sql("""SELECT *
                              FROM kontrakts 
                              """, con=engine, index_col='index')
        dbkontrakts = df_db_kontrakts.to_dict('dict')
        # print('dbkontrakt  ',dbkontrakts)
        df_db_workers = pd.read_sql("""SELECT *
                              FROM workers 
                              """, con=engine, index_col='index')
        dbworkers = df_db_workers.to_dict('split')
        
        def id_kontrakts(kontrakt_name):
            # print('in-> ', kontrakt_name)
            inverse_dic={}
            for key,val in dbkontrakts['kontrakt_name'].items():
                inverse_dic[val]=key
            # print('inverse_dic-->', inverse_dic)
            dfkontrakrs = kontrakt_name.to_dict()
            # print('dfkontrakrs-->', dfkontrakrs)
            for k, v in dfkontrakrs.items():
                if inverse_dic[v]:
                    # print(inverse_dic[v])
                    dfkontrakrs[k] = inverse_dic[v]
            # print(dfkontrakrs.values())
            return dfkontrakrs.values()
        df_json = df_json.assign(id_kontrakts = lambda x: id_kontrakts(x['kontrakt_name']))
        # print(df_json)
        
        # ---- то же самое с работниками---
        def id_worker(job_name, job_department, job_position):
            df_workers = pd.concat([job_name, job_department, job_position], axis=1)
            dfworkers = df_workers.to_dict('split')
            
            id_workers = {}
            key = 0
            for i in dfworkers['data']:
                
                index = 0
                for j in dbworkers['data']:
                    index = index + 1
                    if (i[0] == j[0]) & (i[1] == j[1]) & (i[2] == j[2]):
                        id_workers[key] = index
                key = key+1
                        
            return id_workers.values()
        
        
        df_json = df_json.assign(id_worker = lambda x: id_worker(x['job_name'], x['job_department'], x['job_position']))
        jclocks = df_json.to_dict('records')
        
        # print('df_json >>>>>', df_json[['event','kontrakt_timetracking','job_name','kontrakt_name','work_calendar_day']],'\n\n')
        
        # =====================================================================
        
        # work = df_json[['id_worker','work_calendar_day', 'period_month']].drop_duplicates().to_dict()
        
        #======================================================================
        #нижеследующей конструкцией получаем ид работника, день и месяц, в формате списка списков
        #формируем запрос
       
        work1 = df_json[['id_worker','work_calendar_day', 'period_month']].drop_duplicates().to_dict('split')
       
        work1 = work1['data']

        sql = ''
        
        for i in work1:
            strochka = '(id_worker = ' + str(i[0]) + ' and work_calendar_day = ' + str(i[1]) + ' and period_month = "' + str(i[2]) + '")'
            sql = sql + strochka + ' or '
        sql = sql[:-4]
        sql = '('+ sql + ') AND raw.id_worker = workers.index and raw.id_kontrakts = kontrakts.index'
        # print(sql)
        
        query = 'SELECT raw.index, event, day_night, work_time, work_calendar_day, work_calendar_daytype, job_name, job_department, job_position, period_month, period_year, kontrakt_name, kontrakt_filter, kontrakt_timetracking, id_kontrakts, id_worker FROM raw, workers, kontrakts WHERE'+ sql
        
        # print('query >>> ',query,'\n\n')
        
        # == дополним наш DataFrame df_fson данными из таблицы Raw на соответствующие в Json даты ===
        # --- получаем данные из таблицы Raw преобразуем строку в kontrakt_timetracking к списку словарей ---
        
        # здесь в последний отлов багни добавил  AND period_month IN (' + str(list(set(work['period_month'].values()))).strip('[]') + ') это должно помочь получать из raw не всякую хуйню, а то что нужно в нужную дату 
        
        df_db_raw_date = pd.read_sql(query, con=engine, index_col='index')
        dbraw = df_db_raw_date.to_dict('index') #index
        for k, v in dbraw.items():
            if v['kontrakt_timetracking'] != 'Nan':
                v['kontrakt_timetracking']=ast.literal_eval(v['kontrakt_timetracking'])
                df_db_raw_date.at[k, 'kontrakt_timetracking'] = v['kontrakt_timetracking']
                
         # --- переместим все строки из df_db_raw_date в df_json ---------------
         # и, на всякий случай, удалим все дубликаты
         
        df_json_append = df_json.append(df_db_raw_date, ignore_index=True)
         
        df_json_append.drop_duplicates(subset=['event', 'day_night', 'work_time', 'work_calendar_day', 'work_calendar_daytype', 'job_name', 'job_department', 'job_position', 'period_month', 'period_year', 'kontrakt_name', 'kontrakt_filter', 'id_kontrakts', 'id_worker'], inplace=True)
                
        # print('df >>', df_json[['event','kontrakt_timetracking','job_name','kontrakt_name','work_calendar_day']],'\n\n')
        
        # print('df_db >>', df_db_raw_date[['event','kontrakt_timetracking','job_name','kontrakt_name','work_calendar_day']],'\n\n')
        
        # print('df_appened >>', df_json_append[['event','kontrakt_timetracking','job_name','kontrakt_name','work_calendar_day']],'\n\n')
        
        # sdf
        
        #----------------------------------------------------------------------
        
        # здесь нам нужно из raw удалить шифты и корректировки за даты пришедшие в json. 
        # нужно для того что бы работали корректировки задним числом
        
        # здесь в последний отлов багни добавил  & raw.c.id_worker.in_(list(set(work['period_month'].values()))) это должно помочь не удалять старые даты из raw
        # !!!!
        # Новые трудности (( выяснилось, что конструкция написанная ниже (для удаления всякого из таблицы Raw) ни куда не годится
        # она удаляет все у всех. дело с том, что ей на вход передаются тюплы с данными и эти самые данные сами с собой ни как не вязаны
        # таким образом, если хоть у одного сотрудника будет в прошлом дне хоть одна запись - у всех этот дени почистится
        
        year = df_json_append[['period_year']].drop_duplicates().to_dict()
        
        for y in list(set(year['period_year'].values())):
            for m in list(df_json_append['period_month'].loc[(df_json_append['period_year'] == y)].drop_duplicates().to_dict().values()):
                for d in list(df_json_append['work_calendar_day'].loc[(df_json_append['period_year'] == y) & (df_json_append['period_month'] == m)].drop_duplicates().to_dict().values()):
                    idw = []
                    for w in list(df_json_append['id_worker'].loc[(df_json_append['period_year'] == y) & (df_json_append['period_month'] == m) & (df_json_append['work_calendar_day'] == d)].drop_duplicates().to_dict().values()):
                        idw.append(w)
                   
                    dele = raw.delete().where((raw.c.work_calendar_day == d) & 
                                              (raw.c.period_month == m) & 
                                              (raw.c.period_year == y) & 
                                              (raw.c.id_worker.in_(idw)) & 
                                              (raw.c.event == 'shift') | (raw.c.event == 'correction'))
                    
                    # print("work_calendar_day >>>> ", d)
                    # print("period month >>>>",m)
                    # print("year >>>>",y)
                    # print("workers >>>>",idw)
                    # print("dele >>>>>", dele)
                    # print('\n\n')
                    engine.execute(dele)           
        # ---------------------------------------------------------------------------------------------------
        
        # print('df_json_appened ----->>>>', df_json_append[['event','kontrakt_timetracking','job_name','kontrakt_name','work_calendar_day']],'\n\n')
        
        # ---------------------------------------------------------------------
        
        # -- удалим из таблицы всё clocks за пришедшие даты на нужных сотрудников -------
        # здесь после последнего отлова багни так же был добавлен контекст месяца
        work1 = df_json[['id_worker','work_calendar_day', 'period_month']].drop_duplicates().to_dict('split')
        sql = ''
        work1 = work1['data']
        for i in work1:
            strochka = '(clocks.id_worker = ' + str(i[0]) + ' and clocks.work_calendar_day = ' + str(i[1]) + ' AND clocks.period_month = "' + str(i[2]) + '")'
            sql = sql + strochka + ' or '
        sql = sql[:-4]
        dele1 = 'DELETE FROM clocks WHERE '+ sql
        engine.execute(dele1)
        # -------------------------------------------------------------------------------
   
        # ================ пишем таблицу Raw =================================
        # здесь надо понимать важный момент - в kontrakt_timetracking лежит СТРОКА!!!
        # --- удилим лишнее из DataFrames, оставим только данные для добавления в БД ---
        # ----------------- получаем в DataFrame данные из БД clocks--------
        
        df_db_raw = pd.read_sql("""SELECT *
                              FROM raw 
                              """, con=engine, index_col='index')
        dbraw = df_db_raw.to_dict('records') # split, records
    
            
            # --- удилим лишнее из DataFrames, оставим только список для добавления в БД ---
        inclocks=[]
        jclocks = df_json_append.to_dict('records')
        for j in range(len(jclocks)):
            check = 0
            for i in range(len(dbraw)):
                if ((dbraw[i]['id_worker'] == jclocks[j]['id_worker']) &
                    (dbraw[i]['id_kontrakts'] == jclocks[j]['id_kontrakts']) &
                    (dbraw[i]['period_year'] == jclocks[j]['period_year']) &
                    (dbraw[i]['period_month'] == jclocks[j]['period_month']) &
                    (dbraw[i]['work_calendar_day'] == jclocks[j]['work_calendar_day']) &
                    (dbraw[i]['work_calendar_daytype'] == jclocks[j]['work_calendar_daytype']) &
                    (dbraw[i]['work_time'] == jclocks[j]['work_time']) &
                    # (dbraw[i]['kontrakt_issuecount'] == jclocks[j]['kontrakt_issuecount']) &
                    (dbraw[i]['kontrakt_filter'] == jclocks[j]['kontrakt_filter']) &
                    (dbraw[i]['event'] == jclocks[j]['event']) &
                    # (dbraw[i]['kontrakt_timetracking'] == jclocks[j]['kontrakt_timetracking']) &
                    (dbraw[i]['day_night'] == jclocks[j]['day_night'])
                    ):
                    check = 1
                    if check == 1:
                        break
            if check == 0:
                # print('-->>', jclocks[j]['kontrakt_timetracking'])
                jclocks[j]['kontrakt_timetracking'] = str(jclocks[j]['kontrakt_timetracking'])
                inclocks.append(jclocks[j])

        
        # ------------- инсертим в БД наш список, если он не пустой ---------------
        # print(inclocks)
        if len(inclocks) > 0:
            # print(jclocks)
            conn.execute(insert(raw),inclocks)
            # -------------------------------------------------------------------------
        
        return df_json_append

# def month_s():
#         months = ['','Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь']
#         i = int(time.strftime('%m'))  #текущий месяц
#         return months[i-1]  #минус месяц


def month_to_insert(df_json):
    # month_now = []
    # if "permit_month" in df_json:
    #     month_now.append(month_s())
    #     print(month_s())
    # # month_now.append(df_json["period_month"].drop_duplicates().drop_index())
    # month_now.append(list(df_json["period_month"].drop_duplicates().to_dict())[1])
    # print(month_now)
    
    return df_json["period_month"].drop_duplicates().to_dict()
