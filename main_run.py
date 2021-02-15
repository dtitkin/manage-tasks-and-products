
import time
import configparser
import sys

from numpy import busday_offset, datetime_as_string  # busday_count,

import Wrike
import Spreadsheet
import Baselog
from sub_func import now_str, progress, log_ss, make_date, read_holiday
from sub_func import read_color_cells, read_stage_info, get_user
from sub_func import find_cf, get_len_stage

VERSION = '0.61'
HOLYDAY = []
COLOR_FINISH = []
db = None  # объект для лога в базу или в терминал


def chek_old_session(ss, wr, row_id, num_row):
    '''Проверяем не завершенные результаты предыдущей сессии
    new product:
    Del project:
    update sub task: '''

    if row_id[0] == "new product:":
        pass
    elif row_id[0] == "Del project:":
        pass
    elif row_id[0] == "update sub task:":
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
    if not personal_template:
        resp = wr.copy_folder(id_template, folder_id, name_stage.upper(),
                              "", rescheduleDate=date_start,
                              rescheduleMode="Start", copyStatuses="false",
                              copyResponsibles="true")
    else:
        resp = wr.copy_folder(id_template, folder_id, name_stage.upper(),
                              "", rescheduleDate=None,
                              rescheduleMode=None, copyStatuses="true",
                              copyResponsibles="true")
    id_project = resp[0]["id"]
    id_manager = users_from_name[row_id[4]]["id"]

    cf = wr.custom_field_arr(cfd)
    resp_project = resp[0]["project"]
    ownerIds = resp_project["ownerIds"]
    pr = {"ownersAdd": [id_manager]}
    if ownerIds and ownerIds != id_manager:
        pr["ownersRemove"] = ownerIds
    resp = wr.update_folder(id_project, customFields=cf, project=pr)
    # сохраним в таблице ID
    log_ss(ss, id_project, f"G{num_row}")

    # обзначим в гугл таблице завершение работы
    log_ss(ss, "Finish new product:" + now_str(), f"F{num_row}")
    ok = True
    return (id_project, cfd), ok  # id созданного проекта , заполненные поля


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
    log_ss(ss, "Del project:", f"F{num_row}")
    resp = wr.rs_del(f"folders/{id_project}")
    if resp:
        log_ss(ss, "Finish Del project:" + now_str(), f"F{num_row}")
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
    pred_proc = 0
    len_sub = len(resp)
    for task in resp:
        n += 1
        percent = n / len_sub
        if db.log == "all" or db.log == "terminal":
            progress(percent)
        if pred_proc + 1 == int(percent * 10):
            pred_proc += 1
            log_ss(ss, f"{pred_proc*10}%", f"BW{num_row}")
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

        type_task = task["dates"]["type"]
        dt = None
        status = None
        if not personal_template:
            #  проверяем на статус выполненно
            if num_stage in finish_status:
                status = "Completed"
            else:
                status = "Active"
            # у 'этапов' устанавливаем дату из таблицы
            if type_task == "Milestone" and num_task[0:2] == "00":
                if dates_stage.get(num_stage):
                    dt = wr.dates_arr(type_="Milestone",
                                      due=dates_stage[num_stage].isoformat())
        #  обновляем задачу
        resp_upd = wr.update_task(task["id"], removeResponsibles=remove_r_bles,
                                  addResponsibles=add_r_bles,
                                  customFields=cf, status=status, dates=dt)
        if len(resp_upd) == 0:
            db.out(task["id"] + " ошибка обновления", num_row=num_row,
                   runtime_error="y", error_type="обновление задачи")
            break
    else:
        print()
        #  если обработали все задачи обозначим в таблице выполнение
        log_ss(ss, "Finish update sub task:" + now_str(), f"F{num_row}")
        return True
    return False


def read_date_for_project(ss, end_stage, num_stage="1"):
    ''' считываем с таблицы дату завершения этапа и по длительности этапа
        определяем дату задачи в этапе.
    '''
    date_stage = make_date(end_stage)  # из поля забираем только дату
    len_stage = get_len_stage(num_stage)
    date_for_task = busday_offset(date_stage, -1 * (len_stage - 1),
                                  weekmask="1111100", holidays=HOLYDAY,
                                  roll="forward")
    date_for_task = datetime_as_string(date_for_task)
    # print("Дата старта проекта", date_for_task)
    return date_for_task


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

        time.sleep(1)
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
    if len(row_id[1]) != 0:
        ok = False
        db.out("Проект уже выгружен во Wrike", num_row=num_row,
               runtime_error="y", error_type="исходные данные")
    if len(row_id[11]) == 0:
        ok = False
        # log(f"{num_row} нет названия продукта")
        db.out("нет названия продукта", num_row=num_row,
               runtime_error="y", error_type="исходные данные")
    # руководитель проекта
    id_user = users_from_name.get(row_id[4])
    if not id_user or not id_user.get("id"):
        ok = False
        rp = row_id[4]
        # log(f"{num_row} руководителя проекта {rp} нет во Wrike")
        db.out(f"руководителя проекта {rp} нет во Wrike", num_row=num_row,
               runtime_error="y", error_type="исходные данные")
    # технолог
    id_user = users_from_name.get(row_id[5])
    if not id_user or not id_user.get("id"):
        ok = False
        rp = row_id[5]
        # log(f"{num_row} технолог {rp} нет во Wrike")
        db.out(f"технолог {rp} нет во Wrike", num_row=num_row,
               runtime_error="y", error_type="исходные данные")
    in_all = True
    if len(row_project) < 10:
        in_all = False
    else:
        for x in row_project[2:10]:
            if not x:
                in_all = False
    if not in_all:
        ok = False
        # log(f"{num_row} не установленны даты этапов")
        db.out("не установленны даты этапов", num_row=num_row,
               runtime_error="y", error_type="исходные данные")

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
    time.sleep(1)
    ss.run_prepared()
    return True


def load_from_google_to_wrike(ss, wr, users_from_name, users_from_id,
                              template_id, folder_id, rp_filter, row_filter):

    ss.sheetTitle = "Рабочая таблица №1"
    table_id = ss.values_get("F:AH")
    table_project = ss.values_get("BV:CF")
    num_row = 19
    for row_project in table_project[19:]:
        num_row += 1
        if row_filter:
            if row_filter != num_row:
                continue
        if len(row_project) == 0:
            continue
        if len(table_id) < num_row:
            row_id = ["" for x in range(0, 29)]
        else:
            row_id = table_id[num_row - 1]
        if rp_filter:
            if row_id[4] != rp_filter:
                continue

        db.out("Проверка и отчистка результатов предыдущих сессий",
               prn_time=True)
        chek_old_session(ss, wr, row_id, num_row)
        if row_project[0] == "G":
            ok = test_all_parametr(row_project,
                                   row_id, num_row, users_from_name,
                                   users_from_id)
            if not ok:
                continue
            m = f"Создаем продукт #{row_id[10]} {row_id[11]}"
            # log(m, True, False)
            db.out(m, num_row=num_row, prn_time=True)
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
                        # log(f"Для {num_row} в строке {nrt} нет шаблона.")
                        # log(f"Строка  {num_row} не обрабатывается.")
                        db.out(f"в строке {nrt} нет шаблона",
                               num_row=num_row,
                               runtime_error="y", error_type="ошибка шаблона")

                        continue
                    # log(f"   используем шаблон из строки {nrt}")
                    m = f"используем шаблон из строки {nrt}"
                    db.out(m, num_row=num_row)
            #  создаем новый проект
            id_and_cfd, ok = new_product(ss, wr, row_id, num_row, template_id,
                                         folder_id, users_from_name,
                                         date_start, personal_template)

            if not ok:
                # log("!Выполнение прервано!")
                db.out("!Выполнение прервано!",
                       num_row=num_row,
                       runtime_error="y", error_type="ошибка создания проекта")

                return False
            # установим исполнителей, пользовательские поля,
            # признак выполненно, перенесем вехи на нужные даты
            # log("   Обновление задач в проекте")
            db.out("Обновление задач в проекте", num_row=num_row)

            # считываем статусы и даты этапов из таблицы
            finish_status, dates_stage = read_stage_info(ss, wr, num_row,
                                                         COLOR_FINISH)
            # обновляем задачи с учетом всех статусов
            ok = update_sub_task(ss, wr, id_and_cfd[0],
                                 id_and_cfd[1], users_from_id,
                                 users_from_name, row_id[4:6],
                                 num_row, finish_status, dates_stage,
                                 personal_template)
            if not ok:
                # log("!Выполнение прервано!")
                db.out("!Выполнение прервано!",
                       num_row=num_row,
                       runtime_error="y", error_type="ошибка обновления задач")
                return False
            # Устанавливаем в таблицу W  вместо G
            log_ss(ss, "W", f"BV{num_row}")
            log_ss(ss, "", f"BW{num_row}")
            ok = write_date_to_google(ss, wr, num_row, id_and_cfd[0])
            # set_color_W(ss, num_row, finish_status)
        elif row_project[0] == "W":
            # обновление дат в гугл и признака выполенно
            m = f"Обновляем даты #{row_id[10]} {row_id[11]}"
            # log(m, True, False)
            db.out(m, num_row=num_row, prn_time=True)
            ok = write_date_to_google(ss, wr, num_row, row_id[1])

        elif row_project[0] == "A":
            # удаляем проект из Wrike если он там есть
            id_project = row_id[1]
            if id_project:
                m = f"Удаляем проект #{num_row} {row_id[10]} {row_id[11]}"
                # log(m, True, False)
                db.out(m, num_row=num_row, prn_time=True)
                ok = delete_product(ss, wr, id_project, num_row)
                if not ok:
                    m = f"Не удалось удалить проект {id_project}"
                    # log(m, True, False)
                    db.out(m,
                           num_row=num_row,
                           runtime_error="y", error_type="ошибка удаления")


def main():
    '''Создает шаблон во Wrike на основе Гугл таблицы
    '''
    t_start = time.time()
    CREDENTIALS_FILE = "creds.json"
    cfg = configparser.ConfigParser()
    cfg.read('settings.cfg')
    TABLE_ID = cfg["spreadsheet"]["gogletableid"]
    TOKEN = cfg["wrike"]["wriketoken"]
    OUT = cfg["log"]["out"]
    NAME = cfg["log"]["name"]
    PASS = cfg["log"]["pass"]
    HOST = cfg["log"]["host"]
    BASE = cfg["log"]["base"]

    # обрабатываем список аргументов
    # 0 - сам скрипт
    # 1-й токен пользователя или admin для админа - логи только в терминал
    # 1   help - помощь по параметрам запускаа
    # 2-й -rp
    # 3-й читаем если есть -rp -  менеджер проекта по которому фильтр
    # 4-й -n
    # 5-й читаем еслить есть -n номер строки котору обрабатывать
    rp_filter = None
    row_filter = None
    global db
    lst_argv = sys.argv
    if len(lst_argv) == 1:
        print("Укажите токен пользователя или help")
        return
    else:
        user_token = lst_argv[1]
        if user_token == "help" or user_token == "-h":
            print(" user_token или help или admin. Обязательный параметр")
            print(" -rp <'имя менеджера проекта для фильтра'>."
                  " Не обязательный.")
            print(" -n <номер_строки для фильтра>. Не обязательный.")
            return
        if user_token == "admin":
            OUT = "terminal"
            db = Baselog.Baselog(NAME, PASS, HOST, BASE, OUT, admin_mode=True)
        else:
            db = Baselog.Baselog(NAME, PASS, HOST, BASE, OUT)
            if not db.is_connected():
                print("Нет коннекта к базе с логом")
                return
            ok = db.get_user(user_token)
            if not ok:
                print("Пользователь по токену", user_token, " не найден")
                return
    if len(lst_argv) >= 4:
        key = lst_argv[2]
        if key == "-rp":
            rp_filter = lst_argv[3]
        elif key == "-n":
            row_filter = int(lst_argv[3])
    if len(lst_argv) == 6:
        key = lst_argv[4]
        if key == "-rp":
            rp_filter = lst_argv[5]
        elif key == "-n":
            row_filter = int(lst_argv[5])
    if db.rp_filter and rp_filter is None:
        rp_filter = db.rp_filter
    if rp_filter:
        db.out(f"Обрабатываем строки у РП {rp_filter}")
    if row_filter:
        db.out(f"Обрабатываем строку #{row_filter}")
    if db.log == "terminal":
        db.out(f"Запуск скрипта верси {VERSION}", prn_time=True)
    db.out("Приосоединяемся к Гугл")
    ss = Spreadsheet.Spreadsheet(CREDENTIALS_FILE)
    ss.set_spreadsheet_byid(TABLE_ID)
    global HOLYDAY
    HOLYDAY = read_holiday(ss)
    global COLOR_FINISH
    fl = read_color_cells(ss, 'Рабочая таблица №1!BV2:BV2')[0][0]
    COLOR_FINISH = (fl[0], fl[1])
    db.out("Приосоединяемся к Wrike")
    wr = Wrike.Wrike(TOKEN)

    db.out("Получить id шаблонoв")
    permalink = "https://www.wrike.com/open.htm?id=637661621"
    #  "#1 1CКОД РАБОЧЕЕ НАЗВАНИЕ"
    template = wr.get_folders("folders", permalink=permalink)[0]
    template_id = template["id"]
    template_title = template["title"]
    db.out(f"ID шаблона {template_title} {template_id}")

    permalink = "https://www.wrike.com/open.htm?id=632246796"
    # 000 НОВЫЕ ПРОДУКТЫ
    parent = wr.get_folders("folders", permalink=permalink)[0]
    parent_id = parent["id"]
    parent_title = parent["title"]
    db.out(f"ID папки  {parent_title} {parent_id}")

    db.out("Получить ID, email  пользователей")
    users_from_name, users_from_id = get_user(ss, wr)

    db.out("Выгрузка проектов из Гугл во Wrike", prn_time=True)
    load_from_google_to_wrike(ss, wr, users_from_name, users_from_id,
                              template_id, parent_id, rp_filter, row_filter)

    t_finish = time.time()
    db.out(f"Выполненно за: {int(t_finish - t_start)} секунд")


if __name__ == '__main__':
    main()
