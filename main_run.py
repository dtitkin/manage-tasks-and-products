
from pprint import pprint
from datetime import datetime, timedelta, date
import time
import os
import sys

from numpy import busday_offset, datetime_as_string  # busday_count,

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
        if len(row) < 3:
            continue
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

    # найдем помошника менеджера и запишем его к менеджеру
    for key, value in users_from_name.items():
        if value["group"].find("[") > 0:
            group, name = value["group"].split("[")
            name = name.strip("]")
            user = users_from_name.get(name)
            if user:
                user["idhelper"] = value["id"]

    return users_from_name, users_from_id


def chek_old_session(ss, wr):
    '''Проверяем не заавершенные результаты предыдущей сессии
    '''
    pass


def new_product(ss, wr, row_id, num_row, template_id, folder_id,
                users_from_name, date_start, personal_template):
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

    if personal_template:
        id_template = personal_template
    else:
        id_template = template_id
    resp = wr.copy_folder(id_template, folder_id, name_stage.upper(),
                          "", rescheduleDate=date_start,
                          rescheduleMode="Start", copyStatuses="false",
                          copyResponsibles="true")
    id_project = resp[0]["id"]
    id_manager = users_from_name[row_id[4]]["id"]

    cf = wr.custom_field_arr(cfd)
    resp_project = resp[0]["project"]
    ownerIds = resp_project["ownerIds"]
    if ownerIds == id_manager:
        pr = None
    else:
        pr = {"ownersAdd": [id_manager]}
        if ownerIds:
            pr["ownersRemove"] = ownerIds
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


def find_r_bles(task_user, users_from_id, users_from_name, own_teh, num_task,
                remove=True):
    '''Возвращает список пользователей которых нужно удалить из шаблона
        или установить в шаблон
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

    if remove:
        if len(group_user["Tech"]) > 0:
            tmp_lst = [vl for vl in group_user["Tech"] if vl != teh_id]
            return_list.extend(tmp_lst)
        if len(group_user["RP"]) > 0:
            tmp_lst = [vl for vl in group_user["RP"] if vl != manager_id]
            return_list.extend(tmp_lst)
            if len(group_user["RP_helper"]) > 0:
                id_helper = users_from_name[own_teh[0]].get("idhelper")
                if not id_helper:
                    return_list.extend(group_user["RP_helper"])

    else:
        if len(group_user["RP"]) > 0:
            return_list.append(manager_id)
            id_helper = users_from_name[own_teh[0]].get("idhelper")
            if id_helper and num_task.find("*") > 0:
                return_list.append(id_helper)
        if len(group_user["Tech"]) > 0:
            return_list.append(teh_id)
        return_list.extend(group_user["Other"])

    return return_list


def delete_product(ss, wr, id_project, num_row):
    ''' удаляем весь проект и стираем его ID

    '''
    log_ss(ss, "Del  project:", f"F{num_row}")
    resp = wr.rs_del(f"folders/{id_project}")
    if resp:
        log_ss(ss, "Finish Del:" + now_str(), f"F{num_row}")
        log_ss(ss, "", f"G{num_row}")
        log_ss(ss, "", f"BV{num_row}")
        return True
    else:
        return False


def update_sub_task(ss, wr, parent_id, cfd, users_from_id, users_from_name,
                    own_teh, num_row, finish_status, dates_stage,
                    personal_template):
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
        ownersid = task["responsibleIds"]
        num_stage = find_cf(wr, resp_cf, "num_stage")
        num_task = find_cf(wr, resp_cf, "num_task")

        if personal_template:
            remove_r_bles = find_r_bles(ownersid, users_from_id,
                                        users_from_name, own_teh, num_task)
            add_r_bles = find_r_bles(ownersid, users_from_id,
                                     users_from_name, own_teh, num_task,
                                     False)
        else:
            add_r_bles = None
            remove_r_bles = find_r_bles(ownersid, users_from_id,
                                        users_from_name, own_teh, num_task)
        #  проверяем на статус выполненно
        if num_stage in finish_status:
            status = "Completed"
        else:
            status = "Active"
        # у 'этапов' устанавливаем дату из таблицы
        type_task = task["dates"]["type"]
        dt = None
        if type_task == "Milestone" and num_task[0:2] == "00":
            if dates_stage.get(num_stage):
                dt = wr.dates_arr(type_="Milestone",
                                  due=dates_stage[num_stage].isoformat())
        #  обновляем задачу
        resp_upd = wr.update_task(task["id"], removeResponsibles=remove_r_bles,
                                  addResponsibles=add_r_bles,
                                  customFields=cf, status=status, dates=dt)
        if len(resp_upd) == 0:
            log(task["id"] + " ошибка обновления")
            break
    else:
        print()
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
    template["#1"]["5"] = 10
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


def read_stage_info(ss, wr, num_row):
    ''' запоминаем из таблицы значение, цвет шрифта и цвет ячейки
    '''
    finish_list = []
    dates_stage = {}
    lst_range = ["BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE"]
    cells_range = f"Рабочая таблица №1!BX{num_row}:CE{num_row}"
    lst_status = read_color_cells(ss, cells_range)[0]
    for k, cl in enumerate(lst_range, 1):
        color_stage = (lst_status[k - 1][0], lst_status[k - 1][1])
        stage_value = lst_status[k - 1][2]
        if str(color_stage) == str(COLOR_FINISH):
            # запоминаем у каких этапов нужно установить Finish
            finish_list.append(str(k))
        dates_stage[str(k)] = make_date(stage_value)
    return finish_list, dates_stage


def set_color_W(ss, num_row, finish_status):
    ''' устанавливает цвет - признак того что даты обновляются из Wrike
    '''
    ss.sheetTitle = "Рабочая таблица №1"
    ss.sheetId = 1375512515

    lst_range = ["BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE"]
    max_finish_stage = 0
    if len(finish_status) > 0:
        max_finish_stage = int(max(finish_status))
    if max_finish_stage < 8:
        s_cl = lst_range[max_finish_stage]
        e_cl = "CE"
        cl = f"{s_cl}{num_row}:{e_cl}{num_row}"
        bg_colr = {"backgroundColor": Spreadsheet.htmlColorToJSON("#cfe2f3")}
        ss.prepare_setcells_format(cl, bg_colr,
                                   fields="userEnteredFormat.backgroundColor")

        ss.run_prepared()


def test_all_parametr(row_project, row_id, num_row, users_from_name,
                      users_from_id):
    '''Проверяем все параметры в строке
        - менеджер - ок
        - технолог - ок
        - название продукта -ок
        - наличие типовой длительности для шаблона задач
        - наличие дат во всех этапах -ок
        - ????
    '''
    ok = True
    if len(row_id[11]) == 0:
        ok = False
        log(f"{num_row} нет названия продукта")
    # руководитель проекта
    id_user = users_from_name.get(row_id[4])
    if not id_user or not id_user.get("id"):
        ok = False
        rp = row_id[4]
        log(f"{num_row} руководителя проекта {rp} нет во Wrike")
    # технолог
    id_user = users_from_name.get(row_id[5])
    if not id_user or not id_user.get("id"):
        ok = False
        rp = row_id[5]
        log(f"{num_row} технолог {rp} нет во Wrike")
    in_all = True
    if len(row_project) < 10:
        in_all = False
    else:
        for x in row_project[2:10]:
            if not x:
                in_all = False
    if not in_all:
        ok = False
        log(f"{num_row} не установленны даты этапов")

    return ok


def write_date_to_google(ss, wr, num_row, project_id):
    ''' устанавливает даты из wrike Гугл таблицу
        Отмечаает этапы выполеннными
    '''
    fields = ["customFields"]
    resp = wr.get_tasks(f"folders/{project_id}/tasks", type="Milestone",
                        fields=fields)
    lst = ["" for x in range(0, 8)]
    lst_finish = ["" for x in range(0, 8)]
    lst_values = []
    lst_values.append(lst)

    for task in resp:
        resp_cf = task["customFields"]
        num_stage = find_cf(wr, resp_cf, "num_stage")
        num_task = find_cf(wr, resp_cf, "num_task")
        if num_task[0:2] == "00" and len(num_stage) == 1:
            due_date = task["dates"]["due"]
            y, m, d = due_date[0:10].split("-")
            value = f"{d}.{m}.{y}"
            lst[int(num_stage) - 1] = value
            if task["status"] == "Completed":
                lst_finish[int(num_stage) - 1] = "Finish"
    ss.sheetTitle = "Рабочая таблица №1"
    ss.sheetId = 1375512515
    cells = f"BX{num_row}:CE{num_row}"
    ss.prepare_setvalues(cells, lst_values)
    lst_range = ["BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE"]
    for n, f in enumerate(lst_finish, 0):
        column = lst_range[n]
        txt_field = "userEnteredFormat.textFormat"
        bcg_field = "userEnteredFormat.backgroundColor"
        cells = f"{column}{num_row}:{column}{num_row}"
        if f != "Finish":
            color_bc = Spreadsheet.htmlColorToJSON("#cfe2f3")
            color_font = Spreadsheet.htmlColorToJSON("#000000")
        else:
            color_bc = Spreadsheet.htmlColorToJSON("#6aa84f")
            color_font = Spreadsheet.htmlColorToJSON("#ffffff")
        bg_color = {"backgroundColor": color_bc}
        txt_color = {"textFormat": {"foregroundColor": color_font}}
        ss.prepare_setcells_format(cells, bg_color, fields=bcg_field)
        ss.prepare_setcells_format(cells, txt_color, fields=txt_field)
    ss.run_prepared()
    return True


def load_from_google_to_wrike(ss, wr, users_from_name, users_from_id,
                              template_id, folder_id):

    lst_argv = sys.argv
    rp_filter = ""
    if lst_argv > 1:
        rp_filter = lst_argv[1]
        log(f"Обрабатываем строки РП {rp_filter}")
    ss.sheetTitle = "Рабочая таблица №1"
    table_id = ss.values_get("F:AH")
    table_project = ss.values_get("BV:CF")
    num_row = 19
    for row_project in table_project[19:]:
        num_row += 1
        if len(row_project) == 0:
            continue
        if len(table_id) < num_row:
            row_id = ["" for x in range(0, 29)]
        else:
            row_id = table_id[num_row - 1]
        if rp_filter:
            if row_id[4] != rp_filter:
                continue
        if row_project[0] == "G":
            ok = test_all_parametr(row_project,
                                   row_id, num_row, users_from_name,
                                   users_from_id)
            if not ok:
                continue
            m = f"Создаем продукт #{num_row} {row_id[10]} {row_id[11]}"
            log(m, True, False)
            #  определяем дату для старта проекта
            date_start = read_date_for_project(ss, row_project[2])
            # найдем персональный шаблон
            personal_template = None
            if len(row_project) == 11:
                nrt = row_project[10]
                if len(nrt) > 0 and nrt.isdigit():
                    ss.sheetTitle = "Рабочая таблица №1"
                    personal_template = ss.values_get(f"G{nrt}:G{nrt}")[0][0]
                    if not personal_template:
                        log(f"Для {num_row} в строке {nrt} нет шаблона.")
                        log(f"Строка  {num_row} не обрабатывается.")
                        continue
            #  создаем новый проект
            id_and_cfd, ok = new_product(ss, wr, row_id, num_row, template_id,
                                         folder_id, users_from_name,
                                         date_start, personal_template)

            if not ok:
                log("!Выполнение прервано!")
                return False
            # установим исполнителей, пользовательские поля,
            # признак выполненно, перенесем вехи на нужные даты
            log("   Обновление задач в проекте")

            # считываем статусы и даты этапов из таблицы
            finish_status, dates_stage = read_stage_info(ss, wr, num_row)
            # обновляем задачи с учетом всех статусов
            ok = update_sub_task(ss, wr, id_and_cfd[0],
                                 id_and_cfd[1], users_from_id,
                                 users_from_name, row_id[4:6],
                                 num_row, finish_status, dates_stage,
                                 personal_template)
            if not ok:
                log("!Выполнение прервано!")
                return False
            # Устанавливаем в таблицу W  вместо G
            log_ss(ss, "W", f"BV{num_row}")
            set_color_W(ss, num_row, finish_status)
        elif row_project[0] == "W":
            # обновление дат в гугл и признака выполенно
            m = f"Обновляем даты из Wrike #{num_row} {row_id[10]} {row_id[11]}"
            log(m, True, False)
            ok = write_date_to_google(ss, wr, num_row, row_id[1])

        elif row_project[0] == "A":
            # удаляем проект из Wrike если он там есть
            id_project = row_id[1]
            if id_project:
                m = f"Удаляем проект #{num_row} {row_id[10]} {row_id[11]}"
                log(m, True, False)
                ok = delete_product(ss, wr, id_project, num_row)
                if not ok:
                    m = f"Ошибка удаления #{num_row} {id_project}"
                    log(m, True, False)


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
        строка - список, элементы - кортеж из трех словарей
        1 - цвет фона
        2 - цвет шрифта
        3 - значение в ячейке
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
            value = column["formattedValue"]
            row_lst.append((bg_color, fnt_color, value))
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
    fl = read_color_cells(ss, 'Рабочая таблица №1!BV2:BV2')[0][0]
    COLOR_FINISH = (fl[0], fl[1])
    log("Приосоединяемся к Wrike")
    wr = Wrike.Wrike(TOKEN)

    log("Получить id шаблонoв")
    permalink = "https://www.wrike.com/open.htm?id=637661621"
    #  "#1 1CКОД РАБОЧЕЕ НАЗВАНИЕ"
    template = wr.get_folders("folders", permalink=permalink)[0]
    template_id = template["id"]
    template_title = template["title"]
    log(f"ID шаблона {template_title} {template_id}")

    permalink = "https://www.wrike.com/open.htm?id=632246796"
    # 000 НОВЫЕ ПРОДУКТЫ
    parent = wr.get_folders("folders", permalink=permalink)[0]
    parent_id = parent["id"]
    parent_title = parent["title"]
    log(f"ID папки  {parent_title} {parent_id}")

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
