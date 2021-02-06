
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
                users_from_name):
    ''' По признаку G в строке продукта грузим проект во Wrike
        Если проект уже есть то стираем его и создаем новый с новыми датами

    '''
    # обозначим в гугл таблице начало работы
    log_ss(ss, "N:", f"F{num_row}")
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
                          f"{name_stage} ")
    id_project = resp[0]["id"]
    id_manager = users_from_name[row_id[4]]["id"]

    cf = wr.custom_field_arr(cfd)
    pr = {"ownersAdd": [id_manager]}
    resp = wr.update_folder(id_project, customFields=cf, project=pr)
    # сохраним в таблице ID
    log_ss(ss, id_project, f"G{num_row}")

    # обзначим в гугл таблице завершение работы
    log_ss(ss, "NF:" + now_str(), f"F{num_row}")
    ok = True
    return (id_project, cfd), ok  # id созданного проекта , заполненные поля


def find_cf(wr, resp_cf, name_cf):
    ''' Ищем в списке полей поле с нужным id
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
                    own_teh, num_row):
    ''' Обновление задач в продукте. Установка исполнителей и пользовательских
        полей
    '''
    log_ss(ss, "ST:", f"F{num_row}")
    fields = ["responsibleIds", "customFields"]
    resp = wr.get_tasks(f"folders/{parent_id}/tasks", subTasks="true",
                        fields=fields)

    n = 0
    len_sub = len(resp)
    num_id_task = {}
    for task in resp:
        n += 1
        percent = n / len_sub
        progress(percent)

        resp_cf = task["customFields"]
        cfd["Номер этапа"] = find_cf(wr, resp_cf, "Номер этапа")
        cfd["Номер задачи"] = find_cf(wr, resp_cf, "Номер задачи")
        cfd["Норматив часы"] = find_cf(wr, resp_cf, "Норматив часы")
        nt = cfd["Номер задачи"]
        num_id_task[nt] = {"id": task["id"],
                           "duration": task["dates"].get("duration", 0),
                           "num_stage": cfd["Номер этапа"]}

        r_bles = find_r_bles(task["responsibleIds"], users_from_id,
                             users_from_name, own_teh)
        cf = wr.custom_field_arr(cfd)
        resp_upd = wr.update_task(task["id"], removeResponsibles=r_bles,
                                  customFields=cf)
        if len(resp_upd) == 0:
            break
    else:
        log_ss(ss, "STF:", f"F{num_row}")
        return True, num_id_task
    return False, ""


def get_len_stage(num_stage, num_template="#1"):
    template = {}
    template["#1"] = {}
    template["#1"]["001"] = 4
    template["#1"]["002"] = 6
    template["#1"]["003"] = 5
    template["#1"]["004"] = 1
    template["#1"]["005"] = 8
    template["#1"]["006"] = 92
    template["#1"]["007"] = 7
    template["#1"]["008"] = 2

    my_tmpl = template[num_template]
    return my_tmpl.get(num_stage)


def set_date_on_task(ss, num_row, wr, num_task, end_stage, num_id_task,
                     num_stage="001"):
    ''' устанавливаем дату у задачи с определенным номером
        дату у задачи вычисляем в зависимости от длительности этапа
        длительность этапа берем из функции get_len_stage
    '''
    log_ss(ss, "D:", f"F{num_row}")
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
    log_ss(ss, "DF:", f"F{num_row}")
    return date_for_task


def test_all_parametr(row_project, row_id):
    '''Проверяем все параметры в строке
        - менеджер
        - технолог
        - название продукта
        - даты этапов
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
            id_and_cfd, ok = new_product(ss, wr, row_id, num_row, template_id,
                                         folder_id, users_from_name)
            if not ok:
                log(" Выполнение прервано")
                return False
            # установим исполнителей и пользовательские поля
            log("   Обновление задач в проекте")
            ok, num_id_task = update_sub_task(ss, wr, id_and_cfd[0],
                                              id_and_cfd[1], users_from_id,
                                              users_from_name, row_id[4:6],
                                              num_row)
            if not ok:
                log("     Выполнение прервано")
                return False
            log("")
            # устанавливаем дату первой задачи у первого этапа
            set_date_on_task(ss, num_row, wr, "1", row_project[2], num_id_task)
            # отмечаем выполненно


            # проверяем дату первой задачи в этапе после выполненных этапов
            # переносим вехи в соответсвии с датами во Wrike

            # меняем стату проекта - в работе


        elif row_project[0] == "P":
            # удаляем проект из Wrike если он там есть
            id_product = row_id[1]
            if id_product:
                m = f"Статус на Отменен #{num_row} {row_id[10]} {row_id[11]}"
                log(m, True, False)
                log_ss(ss, "С:", f"F{num_row}")
                # delete_products_recurs(wr, id_product)
                log_ss(ss, "СF:", f"F{num_row}")


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


def main():
    '''Создает шаблон во Wrike на основе Гугл таблицы
    '''
    t_start = time.time()
    log("Приосоединяемся к Гугл", True, False)
    ss = Spreadsheet.Spreadsheet(CREDENTIALS_FILE)
    ss.set_spreadsheet_byid(TABLE_ID)
    read_holiday(ss)
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
