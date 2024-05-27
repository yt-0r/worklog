from sqlalchemy.orm import sessionmaker
from sqlalchemy import create_engine, Table, Column, BIGINT, INT, VARCHAR, DECIMAL, MetaData, ForeignKeyConstraint, PrimaryKeyConstraint
import pandas as pd

meta = MetaData()  
# коннектимся к БД. основной коннект для скрипта
#------------------------------------------------------------------------------
def workers_db():
    workers = Table('workers', meta,
                Column('index', BIGINT),
                Column('job_name', VARCHAR(50), nullable=False),
                Column('job_department', VARCHAR(255), nullable=False),
                Column('job_position', VARCHAR(255), nullable=False),
                PrimaryKeyConstraint('index', name='uniq_1')
                )
    return workers
#------------------------------------------------------------------------------
def kontrakts_db():
    kontrakts = Table('kontrakts', meta,
                Column('index', BIGINT),
                Column('kontrakt_name', VARCHAR(255), nullable=False),
                PrimaryKeyConstraint('index', name='uniq_2')
                )
    return kontrakts
#------------------------------------------------------------------------------
def clocks_db():
    clocks = Table('clocks', meta,
                Column('index', BIGINT, primary_key=True),
                Column('id_worker', BIGINT, nullable=True, default=None),
                Column('id_kontrakts', BIGINT, nullable=True, default=None),
                Column('period_year', INT, nullable=False),
                Column('period_month', VARCHAR(10), nullable=False),
                Column('work_calendar_day', INT, nullable=False),
                Column('work_calendar_daytype', INT, nullable=False),
                Column('work_time', DECIMAL(10,5), nullable=False),
                # Column('kontrakt_issuecount', INT, nullable=False),
                Column('kontrakt_filter', VARCHAR(1000), nullable=False),
                Column('event', VARCHAR(20), nullable=False),
                Column('kontrakt_timetracking', DECIMAL(10,5), nullable=False),
                Column('day_night', VARCHAR(10), nullable=False),
                # Column('in_shift_time', DECIMAL(10,5), nullable=False, default=0),
                Column('status', VARCHAR(3), nullable=False, default='new'),
                ForeignKeyConstraint(['id_worker'],['workers.index']),
                ForeignKeyConstraint(['id_kontrakts'],['kontrakts.index'])
                )
    return clocks
#------------------------------------------------------------------------------
def total_db():
    total = Table('total', meta,
                Column('index', BIGINT, primary_key=True),
                Column('id_worker', BIGINT, nullable=True, default=None),
                Column('period_year', INT, nullable=False),
                Column('period_month', VARCHAR(10), nullable=False),
                Column('work_calendar_day', INT, nullable=False),
                Column('work_calendar_daytype', INT, nullable=False),
                Column('work_time', DECIMAL(10,5), nullable=False),
                ForeignKeyConstraint(['id_worker'],['workers.index'])
                )
    return total
#------------------------------------------------------------------------------
def raw_db():
    raw = Table('raw', meta,
            Column('index', BIGINT, primary_key=True),
            Column('id_worker', BIGINT, nullable=True, default=None),
            Column('id_kontrakts', BIGINT, nullable=True, default=None),
            Column('period_year', INT, nullable=False),
            Column('period_month', VARCHAR(10), nullable=False),
            Column('work_calendar_day', INT, nullable=False),
            Column('work_calendar_daytype', INT, nullable=False),
            Column('work_time', DECIMAL(10,5), nullable=False),
            # Column('kontrakt_issuecount', INT, nullable=False),
            Column('kontrakt_filter', VARCHAR(1000), nullable=False),
            Column('event', VARCHAR(20), nullable=False),
            Column('kontrakt_timetracking', VARCHAR(1000), nullable=False),
            Column('day_night', VARCHAR(10), nullable=False),
            # Column('in_shift_time', DECIMAL(10,5), nullable=False, default=0),
            Column('status', VARCHAR(3), nullable=False, default='new'),
            ForeignKeyConstraint(['id_worker'],['workers.index']),
            ForeignKeyConstraint(['id_kontrakts'],['kontrakts.index'])
            )
    return raw
#------------------------------------------------------------------------------
def engine_db(config):
    #старье
    # return create_engine("mysql+mysqlconnector://workloguser:Gsv248754@jiradev.its-sib.ru:3306/worklog?charset=utf8", echo=False, pool_size=6, max_overflow=10, encoding='latin1')
    return create_engine("mysql+mysqlconnector://"+config["worklog_db"]["user"]+":"+config["worklog_db"]["password"]+"@"+config["worklog_db"]["host"]+":3306/worklog?charset=utf8", echo=False, pool_size=6, max_overflow=10, encoding='latin1')
#------------------------------------------------------------------------------
def create_all_db(engine):
    meta.create_all(engine)
#------------------------------------------------------------------------------
def conn_db(engine):
    return engine.connect()
#------------------------------------------------------------------------------
def session_db(engine):
    return sessionmaker(bind=engine, autocommit=True)
#------------------------------------------------------------------------------
def df_base_set(conn):
    return pd.read_sql("""SELECT 
                          id_worker, id_kontrakts, period_year, period_month, work_calendar_day, work_calendar_daytype, 
                          work_time, kontrakt_filter, event, kontrakt_timetracking, day_night, 
                          STATUS, job_name, job_department, job_position, kontrakt_name FROM clocks, workers, kontrakts 
                          WHERE clocks.id_worker = workers.index AND clocks.id_kontrakts = kontrakts.index""", conn)