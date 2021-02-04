
from pprint import pprint
from datetime import datetime, timedelta, date
import time
import os


import Wrike
import Spreadsheet


CREDENTIALS_FILE = "creds.json"
TABLE_ID = os.getenv("gogletableid")
TOKEN = os.getenv("wriketoken")
ONE_DAY = timedelta(days=1)


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


def log(msg, prn_time=False, one_str=True):
    str_now = now_str()
    if prn_time:
        print(str_now, end=" ")
        if not one_str:
            pprint(msg)
        else:
            print('\r', msg, sep='', end='', flush=True)
    else:
        if not one_str:
            pprint(msg)
        else:
            print('\r', msg, sep='', end='', flush=True)


def log_ss(ss, msg, cells_range):
    value = [[msg]]
    my_range = f"{cells_range}:{cells_range}"
    ss.prepare_setvalues(my_range, value)
    ss.run_prepared()


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


def create_product(wr, row_id, folder_id, users_from_name):
    # создаем этап или веху с продуктом или веху внутри этапа

    task_date = date.today()
    name_stage = row_id[10] + " " + row_id[11]
    name_stage = name_stage.upper()

    dt = wr.dates_arr(type_="Milestone", due=task_date.isoformat())
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

    manager = users_from_name.get(row_id[4], "")
    r_bles = []
    if manager:
        id_manager = manager["id"]
        r_bles = [id_manager]
    cf = wr.custom_field_arr(cfd)
    resp = wr.create_task(folder_id, name_stage, dates=dt,
                          responsibles=r_bles, customFields=cf)
    return resp[0]["id"], cfd  # id созданной задачи, заполненные поля


def new_product(ss, wr, row_id, num_row, folder_id, users_from_name):
    ''' По признаку G в строке продукта грузим проект во Wrike
        Если проект уже есть то стираем его и создаем новый с новыми датами

    '''
    # обозначим в гугл таблице начало работы
    log_ss(ss, "N:", f"F{num_row}")
    # создадим задачу с продуктом
    id_and_cfd = create_product(wr, row_id, folder_id, users_from_name)
    # сохраним в таблице ID
    log_ss(ss, id_and_cfd[0], f"G{num_row}")

    # обзначим в гугл таблице завершение работы
    log_ss(ss, "NF:" + now_str(), f"F{num_row}")

    return id_and_cfd


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
    '''Подбирает из списка пользоваатателей одного или нескольких
       по соответсвию рабочей группы и заменителей []
    '''
    return_list = []
    manager_name = own_teh[0]
    manager_id = users_from_name.get(manager_name, "")
    if not manager_id:
        log(f"Руководитель проекта {manager_name} не найден в Ресурсах")
    else:
        manager_id = manager_id["id"]
    teh_name = own_teh[1]
    teh_id = users_from_name.get(teh_name, "")
    if not teh_id:
        log(f"Технолог {teh_name} не найден в Ресурсах")
    else:
        teh_id = teh_id["id"]

    all_in_group = True
    group = ""
    group_in_list = ""
    for num, user in enumerate(task_user, 1):
        param_user = users_from_id.get(user)
        if param_user is None:
            log(f"Пользователь {user} не найден в Ресурсах")
            continue
        group = param_user["group"]
        slice_pos = group.find("[")
        if slice_pos > -1:
            group = group[0:slice_pos]
        if not group:
            all_in_group = False
            break
        if num == 1:
            group_in_list = group
        if group_in_list != group:
            all_in_group = False
            break

    if len(task_user) == 1:
        all_in_group = False

    if all_in_group:
        if group == "РП":  # оставим одного рп и его заменителя
            for user in task_user:
                if user == manager_id:
                    return_list.append(user)
                else:
                    param_user = users_from_id.get(user)
                    if param_user:
                        gr = param_user["group"]
                        start_s = gr.find("[")
                        finish_s = gr.find("]")
                        if start_s > -1:
                            helper_name = gr[start_s + 1:finish_s]
                            helper_id = users_from_name[helper_name]["id"]
                            if helper_id == manager_id:
                                return_list.append(user)
        elif group == "Технолог":
            if teh_id:
                return_list.append(teh_id)
    else:
        return_list = task_user.copy()

    return return_list


def delete_products_recurs(wr, id_task, num=0):
    ''' удаляем все задачи и вехи по продукту

    '''
    resp = wr.get_tasks(f"tasks/{id_task}")[0]
    if len(resp) == 0:
        return None
    sub_task = resp["subTaskIds"]
    resp_cf = resp["customFields"]
    num_t = find_cf(wr, resp_cf, "Номер задачи")
    log(str(num_t) + " удаляем", False, True)
    for task in sub_task:
        delete_products_recurs(wr, task, num + 1)
    if num == 0:
        log("")
    wr.rs_del(f"tasks/{id_task}")


def new_sub_task_rekurs(ss, wr, parent_id, cfd, templ_sub_tasks,
                        folder_id, users_from_id, users_from_name, own_teh,
                        level=0):
    '''Создаем новые подзадачи и новые вложенные вехи
       рекурсивно по всему списку задач из шаблона

    '''
    templ_dict = {}
    if len(templ_sub_tasks) == 0:
        return templ_dict
    task_date = date.today()
    n = 0
    len_sub = len(templ_sub_tasks)
    for templ_task in templ_sub_tasks:
        if level == 0:
            n += 1
            percent = n / len_sub
            progress(percent)
        # из родительской задачи нужны: Код-1С, Название рабочее.
        # из шаблона : Название задачи, описание,подзадачи, связи,
        #              Номер этапа, Номер задачи, норматив часы
        #              тип задачи (задача/веха), длительнось, пользоваатели.
        resp = wr.get_tasks(f"tasks/{templ_task}")[0]
        kod = cfd["Код-1С"]
        name = cfd["Название рабочее"]
        tmp_n = resp["title"]
        name_task = tmp_n.replace("[код рабочее название]", f"[{kod} {name}]")
        descr = resp["description"]
        sub_tasks = resp["subTaskIds"]
        dependecy_ids = resp["dependencyIds"]
        resp_cf = resp["customFields"]
        cfd["Номер этапа"] = find_cf(wr, resp_cf, "Номер этапа")
        cfd["Номер задачи"] = find_cf(wr, resp_cf, "Номер задачи")
        cfd["Норматив часы"] = find_cf(wr, resp_cf, "Норматив часы")
        type_task = resp["dates"]["type"]
        duration = 0
        st = [parent_id]
        cf = wr.custom_field_arr(cfd)
        if type_task == "Planned":
            duration = resp["dates"]["duration"]
            dt = wr.dates_arr(type_=type_task, start=task_date.isoformat(),
                              duration=duration)
        elif type_task == "Milestone":
            dt = wr.dates_arr(type_=type_task, due=task_date.isoformat())

        r_bles = find_r_bles(resp["responsibleIds"], users_from_id,
                             users_from_name, own_teh)

        # num_task = cfd["Номер задачи"]
        # log(f"       {num_task} {name_task}")
        resp = wr.create_task(folder_id, name_task, description=descr,
                              dates=dt, responsibles=r_bles, superTasks=st,
                              customFields=cf)

        id_task = resp[0]["id"]
        templ_dict[templ_task] = {}
        templ_dict[templ_task]["new_id"] = id_task
        templ_dict[templ_task]["old_dependecy"] = dependecy_ids

        resp_dict = new_sub_task_rekurs(ss, wr, id_task, cfd,
                                        sub_tasks, folder_id, users_from_id,
                                        users_from_name, own_teh, 1)
        templ_dict.update(resp_dict)
    return templ_dict


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
            m = f"Создаем продукт # {num_row} {row_id[10]} {row_id[11]}"
            log(m, True)
            id_and_cfd = new_product(ss, wr, row_id, num_row, folder_id,
                                     users_from_name)
            # создадим вложенные вехи
            templ_id = {}  # словарь соответсвия id из шаблона с id созданных
            templ_id[template_id] = id_and_cfd[0]
            log_ss(ss, "ST:", f"F{num_row}")
            resp = wr.get_tasks(f"tasks/{template_id}")[0]
            templ_subtask = resp["subTaskIds"]
            log("   Создаем задачи")
            resp_dict = new_sub_task_rekurs(ss, wr, id_and_cfd[0],
                                            id_and_cfd[1], templ_subtask,
                                            folder_id, users_from_id,
                                            users_from_name, row_id[4:6])
            templ_id.update(resp_dict)
            log("")
            log_ss(ss, "STF:", f"F{num_row}")
        elif row_project[0] == "P":
            # удяляем проект из Wrike если он там есть
            id_product = row_id[1]
            if id_product:
                m = f"Удаляем продукт # {num_row} {row_id[10]} {row_id[11]}"
                log(m, True)
                log_ss(ss, "D:", f"F{num_row}")
                delete_products_recurs(wr, id_product)
                # сохраним в таблице ID
                log_ss(ss, "", f"G{num_row}")
                log_ss(ss, "DF:", f"F{num_row}")


def main():
    '''Создает шаблон во Wrike на основе Гугл таблицы
    '''
    t_start = time.time()
    log("Приосоединяемся к Гугл", True)
    ss = Spreadsheet.Spreadsheet(CREDENTIALS_FILE)
    ss.set_spreadsheet_byid(TABLE_ID)

    log("Приосоединяемся к Wrike")
    wr = Wrike.Wrike(TOKEN)

    log("Получить папку и id шаблона")
    name_sheet = "000 НОВЫЕ ПРОДУКТЫ"
    folder_id = wr.id_folders_on_name([name_sheet])[name_sheet]
    log(f"ID папки с проектами {folder_id}")

    name_sheet = "001 ШАБЛОНЫ (новые продукты Рубис)"
    template_folder_id = wr.id_folders_on_name([name_sheet])[name_sheet]
    log(f"ID папки с шаблонами {template_folder_id}")
    api_str = f"folders/{template_folder_id}/tasks"
    resp = wr.get_tasks(api_str, title="V2 1CКОД РАБОЧЕЕ НАЗВАНИЕ")
    template_id = resp[0]["id"]
    log(f"ID шаблона {template_id}")

    log("Получить ID, email  пользователей")
    users_from_name, users_from_id = get_user(ss, wr)

    log("Проверка и отчитска результатов предыдущих сессий", True)
    chek_old_session(ss, wr)

    log("Выгрузка проектов из Гугл во Wrike", True)
    load_from_google_to_wrike(ss, wr, users_from_name, users_from_id,
                              template_id, folder_id)

    t_finish = time.time()
    log("Выполненно за:", int(t_finish - t_start), " секунд")


if __name__ == '__main__':
    main()
