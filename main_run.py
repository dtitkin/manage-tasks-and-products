
from pprint import pprint
from datetime import datetime, timedelta, date
import time
import os

from numpy import busday_offset, busday_count, datetime_as_string

import Wrike
import Spreadsheet


CREDENTIALS_FILE = "creds.json"
TABLE_ID = os.getenv("gogletableid")
TOKEN = os.getenv("wriketoken")
ONE_DAY = timedelta(days=1)
HOLYDAY = []
COLOR_FINISH = {}


def now_str():
    now = datetime.now()
    fmt = '%d.%m.%Y:%H:%M:%S'
    str_now = now.strftime(fmt)
    return str_now


def progress(percent=0, width=30):
    left = int(width * percent) // 1
    right = width - left
    print('\r[', '#' * left, ' ' * right, ']',
          f' {percent * 100:.0f}%',
          sep='', end='', flush=True)


def log(msg, prn_time=False, one_str=False):
    str_now = now_str()
    if prn_time:
        print(str_now, end=" ")
        if one_str:
            print('\r', msg, sep='', end='', flush=True)
        else:
            pprint(msg)
    else:
        if one_str:
            print('\r', msg, sep='', end='', flush=True)
        else:
            pprint(msg)


def log_ss(ss, msg, cells_range):
    value = [[msg]]
    my_range = f"{cells_range}:{cells_range}"
    ss.prepare_setvalues(my_range, value)
    ss.run_prepared()


def make_date(usr_date):
    if not usr_date:
        return date.today()
    lst_date = usr_date[0:10].split(".")
    return date(int(lst_date[2]), int(lst_date[1]), int(lst_date[0]))


def get_user(ss, wr):
    ''' возвращает два словаря пользователей
    ключ по имени из гугл таблицы
    ключ по id из wrike
    '''
    def find_name_from_email(email, users_from_name):
        for user_name, user_val, in users_from_name.items():
            if user_val["email"] == email:
                return user_name, user_val["group"]

    ss.sheetTitle = "Ресурсы"
    table = ss.values_get("B50:D89")
    users_from_id = {}
    users_from_name = {}
    lst_mail = []
    for row in table:
        users_from_name[row[0]] = {}  # имя пользователя
        users_from_name[row[0]]["id"] = ""
        users_from_name[row[0]]["group"] = row[1].strip(" ")
        users_from_name[row[0]]["email"] = row[2].strip(" ")
        lst_mail.append(row[2])
    id_dict = wr.id_contacts_on_email(lst_mail)
    for email, id_user in id_dict.items():
        if email in lst_mail and id_user is not None:
            users_from_id[id_user] = {}
            users_from_id[id_user]["email"] = email
            name_user, gr_user = find_name_from_email(email, users_from_name)
            users_from_id[id_user]["name"] = name_user
            users_from_id[id_user]["group"] = gr_user
            users_from_name[name_user]["id"] = id_user
    return users_from_name, users_from_id


def chek_old_session(ss, wr):
    '''Проверяем не заавершенные результаты предыдущей сессии
    '''
    pass


def new_product(ss, wr, row_id, num_row, template_id, folder_id,
                users_from_name, date_start):
    ''' По признаку G в строке продукта создаем новый проект во Wrike
    '''
    # обозначим в гугл таблице начало работы
    log_ss(ss, "new product:", f"F{num_row}")
    # копируем шаблон в новый проект
    name_stage = row_id[10] + " " + row_id[11]

    cfd = {"Номер этапа": "",
           "Номер задачи": "",
           "Норматив часы": 0,
           "Стратегическая группа": row_id[2],
           "Проект": row_id[3],
           "Руководитель проекта": row_id[4],
           "Технолог": row_id[5],
           "Код-1С": row_id[10],
           "Название рабочее": row_id[11],
           "Группа": row_id[21],
           "Линейка": row_id[12],
           "Клиент": row_id[27],
           "Бренд": row_id[28]}

    resp = wr.copy_folder(template_id, folder_id, name_stage.upper(),
                          f"{name_stage} ", rescheduleDate=date_start,
                          rescheduleMode="Start")
    id_project = resp[0]["id"]
    id_manager = users_from_name[row_id[4]]["id"]

    cf = wr.custom_field_arr(cfd)
    pr = {"ownersAdd": [id_manager]}
    resp = wr.update_folder(id_project, customFields=cf, project=pr)
    # сохраним в таблице ID
    log_ss(ss, id_project, f"G{num_row}")

    # обзначим в гугл таблице завершение работы
    log_ss(ss, "Finish new product:" + now_str(), f"F{num_row}")
    ok = True
    return (id_project, cfd), ok  # id созданного проекта , заполненные поля


def find_cf(wr, resp_cf, name_cf):
    ''' Ищем в списке полей поле с нужным id по имени
        возвращеем значение
    '''
    return_value = ""
    id_field = wr.customfields[name_cf]
    for cf in resp_cf:
        if cf["id"] == id_field:
            return_value = cf["value"]
            break
    return return_value


def find_r_bles(task_user, users_from_id, users_from_name, own_teh):
    '''Возворащает список пользователей которых нужно удалить из шаблона
    '''
    return_list = []
    manager_id = users_from_name[own_teh[0]]["id"]
    teh_id = users_from_name[own_teh[1]]["id"]

    group_user = {}
    group_user["RP"] = []
    group_user["RP_helper"] = []
    group_user["Tech"] = []
    group_user["Other"] = []

    for num, user in enumerate(task_user, 1):
        group = users_from_id[user]["group"]
        slice_pos = group.find("[")
        if slice_pos > -1:
            group = group[0:slice_pos]
        if not group:
            group_user["Other"].append(user)
        elif group == "РП" and slice_pos > -1:
            group_user["RP_helper"].append(user)
        elif group == "РП" and slice_pos == -1:
            group_user["RP"].append(user)
        elif group == "Технолог":
            group_user["Tech"].append(user)

    if len(group_user["Tech"]) > 0:
        tmp_lst = [vl for vl in group_user["Tech"] if vl != teh_id]
        return_list.extend(tmp_lst)
    if len(group_user["RP"]) > 0:
        tmp_lst = [vl for vl in group_user["RP"] if vl != manager_id]
        return_list.extend(tmp_lst)

    return return_list


def delete_products_recurs(wr, id_task, num=0):
    ''' удаляем все задачи и вехи по продукту

    '''
    pass


def update_sub_task(ss, wr, parent_id, cfd, users_from_id, users_from_name,
                    own_teh, num_row, finish_status):
    ''' Обновление задач в проекте. Установка исполнителей, пользовательских
        полей, статуса выполненно
    '''
    # обозначим в таблице что начали этап
    log_ss(ss, "update sub task:", f"F{num_row}")
    #  читаем из `Wrike задачи
    fields = ["responsibleIds", "customFields"]
    resp = wr.get_tasks(f"folders/{parent_id}/tasks", subTasks="true",
                        fields=fields)
    # перебираем задачи и обновляем
    n = 0
    len_sub = len(resp)
    for task in resp:
        n += 1
        percent = n / len_sub
        progress(percent)
        # пользовательские поля
        resp_cf = task["customFields"]
        cfd["Номер этапа"] = find_cf(wr, resp_cf, "Номер этапа")
        cfd["Номер задачи"] = find_cf(wr, resp_cf, "Номер задачи")
        cfd["Норматив часы"] = find_cf(wr, resp_cf, "Норматив часы")
        cf = wr.custom_field_arr(cfd)
        # определям список пользователей которых нужно исключить из задачи
        r_bles = find_r_bles(task["responsibleIds"], users_from_id,
                             users_from_name, own_teh)
        #  проверяем на статус выполненно
        num_stage = find_cf(wr, resp_cf, "num_stage")
        if int(num_stage) in finish_status:
            status = "Completed"
        else:
            status = None
        #  обновляем задачу
        resp_upd = wr.update_task(task["id"], removeResponsibles=r_bles,
                                  customFields=cf, status=status)
        if len(resp_upd) == 0:
            log(task["id"] + " ошибка обновления")
            break
    else:
        #  если обработали все задачи обозначим в таблице выполнение
        log_ss(ss, "Finish update sub task:" + now_str(), f"F{num_row}")
        return True
    return False


def get_len_stage(num_stage, num_template="#1"):
    template = {}
    template["#1"] = {}
    template["#1"]["1"] = 4
    template["#1"]["2"] = 6
    template["#1"]["3"] = 5
    template["#1"]["4"] = 1
    template["#1"]["5"] = 8
    template["#1"]["6"] = 92
    template["#1"]["7"] = 7
    template["#1"]["8"] = 2

    my_tmpl = template[num_template]
    return my_tmpl.get(num_stage)


def read_date_for_project(ss, end_stage, num_stage="1"):
    ''' считываем с таблицы дату завершения этапа и по длительности этапа
        определяем дату задачи в этапе.
    '''
    date_stage = make_date(end_stage)  # из поля забираем только дату
    len_stage = get_len_stage(num_stage)
    date_for_task = busday_offset(date_stage, -1 * (len_stage - 1),
                                  weekmask="1111100", holidays=HOLYDAY)
    date_for_task = datetime_as_string(date_for_task)
    # print("Дата старта проекта", date_for_task)
    return date_for_task


def set_date_on_task(ss, wr, num_row, num_task, end_stage, num_stage="1"):
    ''' устанавливаем дату у задачи с определенным номером
        дату у задачи вычисляем в зависимости от длительности этапа
        длительность этапа берем из функции get_len_stage
    '''
    '''log_ss(ss, "set date:", f"F{num_row}")
    date_stage = make_date(end_stage)
    len_stage = get_len_stage(num_stage)
    date_for_task = busday_offset(date_stage, -1 * len_stage - 1,
                                  weekmask="1111100", holidays=HOLYDAY)
    date_for_task = datetime_as_string(date_for_task)
    # найдем задачу у которой нужно поменять день

    id_task = num_id_task[num_task]["id"]
    duration = num_id_task[num_task]["duration"]
    if id_task:
        log(f"     меняем дату у задачи {num_task}:{id_task}")
        dt = wr.dates_arr(type_="Planned", start=date_for_task,
                          duration=duration)
        wr.update_task(id_task, dates=dt)
    log_ss(ss, "Finish set date:" + now_str(), f"F{num_row}")
    return date_for_task'''


def read_finish(ss, wr, num_row):
    ''' устанвливаем статус выполнено
        в зависимости от статуса в Гугл таблице
    '''
    return_list = []
    lst_range = ["BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE"]
    for k, cl in enumerate(lst_range, 1):
        cells_range = f"Рабочая таблица №1!{cl}{num_row}:{cl}{num_row}"
        lst_status = read_color_cells(ss, cells_range)
        if len(lst_status) > 0:
            color_stage = lst_status[0][0]
            if str(color_stage) == str(COLOR_FINISH):
                # запоминаем у каких этапов нужно установить Finish
                return_list.append(k)
    return return_list


def test_all_parametr(row_project, row_id):
    '''Проверяем все параметры в строке
        - менеджер
        - технолог
        - название продукта
        - даты этапов
        - наличие типовой длительности для шаблона задач
        - ????
    '''
    return True

def load_from_google_to_wrike(ss, wr, users_from_name, users_from_id,
                              template_id, folder_id):

    ss.sheetTitle = "Рабочая таблица №1"
    table_id = ss.values_get("F:AH")
    table_project = ss.values_get("BV:CE")
    num_row = 19
    for row_project in table_project[19:]:
        num_row += 1
        if len(row_project) == 0:
            continue
        if len(table_id) < num_row:
            row_id = ["" for x in range(0, 29)]
        else:
            row_id = table_id[num_row - 1]
        if row_project[0] == "G":
            ok = test_all_parametr(row_project, row_id)
            if not ok:
                continue
            m = f"Создаем продукт #{num_row} {row_id[10]} {row_id[11]}"
            log(m, True, False)
            #  определяем дату для старта проекта
            date_start = read_date_for_project(ss, row_project[2])
            #  создаем новый проект
            id_and_cfd, ok = new_product(ss, wr, row_id, num_row, template_id,
                                         folder_id, users_from_name,
                                         date_start)
            if not ok:
                log(" Выполнение прервано")
                return False
            # установим исполнителей, пользовательские поля, признак выполненно
            log("   Обновление задач в проекте")

            # считываем статус выполенно из  таблицы
            finish_status = read_finish(ss, wr, num_row)
            # обновляем задачи с учетом всех статусов
            ok = update_sub_task(ss, wr, id_and_cfd[0],
                                 id_and_cfd[1], users_from_id,
                                 users_from_name, row_id[4:6],
                                 num_row, finish_status)
            if not ok:
                log("     Выполнение прервано")
                return False
            # проверяем дату первой задачи в этапе после выполненных этапов

            # переносим вехи в соответсвии с датами во Wrike

            # меняем статуc проекта - в работе

        elif row_project[0] == "P":
            # удаляем проект из Wrike если он там есть
            id_product = row_id[1]
            if id_product:
                m = f"Статус на Отменен #{num_row} {row_id[10]} {row_id[11]}"
                log(m, True, False)
                log_ss(ss, "set canсel :", f"F{num_row}")
                # delete_products_recurs(wr, id_product)
                log_ss(ss, "Finish set cancel:" + + now_str(), f"F{num_row}")


def read_holiday(ss):
    holidays_str = ss.values_get("Рабочий календарь!A:A")
    global HOLYDAY
    do_it = False
    for hd in holidays_str:
        if (len(hd) == 0):
            continue
        if hd[0] == "Шаблон{":
            do_it = True
            continue
        elif hd[0] == "Шаблон}":
            break
        if do_it:
            d_str = hd[0].split(".")
            HOLYDAY.append(date(int(d_str[2]), int(d_str[1]), int(d_str[0])))


def read_color_cells(ss, cells_range):
    '''возвращает список, элементы - строка
        строка - список, элементы - кортеж из двух словарей
        1 - цвет фона
        2 - цвет шрифта
    '''
    return_list = []
    resp = ss.sh_get(cells_range)[0]  # было без [0]
    # sh = resp["sheets"][0]
    data = resp["data"][0]

    row_data = data.get("rowData", [])
    n = 0
    for row in row_data:
        return_list.append([])
        row_lst = return_list[n]
        n += 1
        vls = row["values"]
        for column in vls:
            ef_format = column["effectiveFormat"]
            bg_color = ef_format["backgroundColor"]
            fnt_color = ef_format["textFormat"]["foregroundColor"]
            row_lst.append((bg_color, fnt_color))

    return return_list


def main():
    '''Создает шаблон во Wrike на основе Гугл таблицы
    '''
    t_start = time.time()
    log("Приосоединяемся к Гугл", True, False)
    ss = Spreadsheet.Spreadsheet(CREDENTIALS_FILE)
    ss.set_spreadsheet_byid(TABLE_ID)
    read_holiday(ss)
    global COLOR_FINISH
    COLOR_FINISH = read_color_cells(ss, 'Рабочая таблица №1!BV2:BV2')[0][0]
    log("Приосоединяемся к Wrike")
    wr = Wrike.Wrike(TOKEN)

    log("Получить id шаблона")
    permalink = "https://www.wrike.com/open.htm?id=637661621"
    #  "#1 1CКОД РАБОЧЕЕ НАЗВАНИЕ"
    template_id = wr.get_folders("folders", permalink=permalink)[0]["id"]
    log(f"ID шаблона {template_id}")

    permalink = "https://www.wrike.com/open.htm?id=632246796"
    # 000 НОВЫЕ ПРОДУКТЫ
    parent_id = wr.get_folders("folders", permalink=permalink)[0]["id"]
    log(f"ID папки для размещения проектов {parent_id}")

    log("Получить ID, email  пользователей")
    users_from_name, users_from_id = get_user(ss, wr)

    log("Проверка и отчистка результатов предыдущих сессий", True, False)
    chek_old_session(ss, wr)

    log("Выгрузка проектов из Гугл во Wrike", True, False)
    load_from_google_to_wrike(ss, wr, users_from_name, users_from_id,
                              template_id, parent_id)

    t_finish = time.time()
    log(f"Выполненно за: {int(t_finish - t_start)} секунд")


if __name__ == '__main__':
    main()
