
from datetime import datetime, date
from numpy import busday_offset, datetime_as_string  # busday_count,


def now_str():
    now = datetime.now()
    fmt = '%d.%m.%Y:%H:%M:%S'
    str_now = now.strftime(fmt)
    return str_now


def make_date(usr_date):
    if not usr_date:
        return date.today()
    if usr_date.find(".") > -1:
        lst_date = usr_date[0:10].split(".")
    elif usr_date.find("/") > -1:
        lst_date = usr_date[0:10].split("/")
    else:
        return date.today()
    return date(int(lst_date[2]), int(lst_date[1]), int(lst_date[0]))


def read_date_for_project(ss, end_stage, len_stage, holydays):
    ''' считываем с таблицы дату завершения этапа и по длительности этапа
        определяем дату задачи в этапе.
    '''
    date_stage = make_date(end_stage)  # из поля забираем только дату
    date_for_task = busday_offset(date_stage, -1 * (len_stage - 1),
                                  weekmask="1111100",
                                  holidays=holydays,
                                  roll="forward")
    date_for_task = datetime_as_string(date_for_task)
    # print("Дата старта проекта", date_for_task)
    return date_for_task
