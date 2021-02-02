
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


def log(msg, prn_time=False):
    str_now = now_str()
    if prn_time:
        print(str_now, end=" ")
        pprint(msg)
    else:
        pprint(msg)


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


def create_milestone(wr, row_id, row_project, folder_id, root_task_id="",
                     after_task_id="", root=False):
    # создаем этап или веху с продуктом или веху внутри этапа
    number_stage = ""
    descr = ""
    task_date = date.today()
    st = ""  # [root_task_id]
    if root:
        name_stage = row_id[10] + " " + row_id[11]
        name_stage = name_stage.upper()
    else:
        name_stage = " номер этапа наазвание этапа название товара"
        name_stage = number_stage + ": " + row_id[55]
    dt = wr.dates_arr(type_="Milestone", due=task_date.isoformat())
    cfd = {"Номер этапа": number_stage,
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
    cf = wr.custom_field_arr(cfd)
    resp = wr.create_task(folder_id, name_stage, description=descr,
                          dates=dt, priorityAfter=after_task_id,
                          superTasks=st, customFields=cf)
    return resp[0]["id"], cfd  # id созданной задачи, заполненные поля


def new_product(ss, wr, row_id, row_project, num_row, folder_id):
    ''' По признаку G в строке продукта грузим проект во Wrike
        Если проект уже есть то стираем его и создаем новый с новыми датами

    '''
    # обозначим в гугл таблице начало работы
    log_ss(ss, "N:", f"F{num_row}")
    # создадим задачу с продуктом
    id_and_cfd = create_milestone(wr, row_id, row_project,
                                  folder_id, "", "", True)
    # сохраним в таблице ID
    value = [[id_and_cfd[0]]]
    my_range = f"G{num_row}:G{num_row}"
    ss.prepare_setvalues(my_range, value)
    ss.run_prepared()

    # обзначим в гугл таблице завершение работы
    log_ss(ss, "NF:" + now_str(), f"F{num_row}")

    return id_and_cfd


def new_sub_task_rekurs(ss, wr, id_task, cfd, template_id):
    '''Создаем новые подзадачи и новые вложенные вехи

       рекурсивно по всему списку задач из шаблона
    '''
    resp = wr.get_tasks(f"tasks/{template_id}")
    sub_task_ids = resp[0]["subTaskIds"]
    if len(sub_task_ids) == 0:
        return None


def load_from_google_to_wrike(ss, wr, users_from_name, users_from_id,
                              template_id, folder_id):
    chek_old_session(ss, wr)
    ss.sheetTitle = "Рабочая таблица №1"
    table_id = ss.values_get("F:AH")
    table_project = ss.values_get("BV:CE")
    num_row = 19
    for row_project in table_project[19:]:
        num_row += 1
        if len(row_project) == 0:
            continue
        if len(table_id) < num_row:
            row_id = ["", ""]
        else:
            row_id = table_id[num_row - 1]
        if row_project[0] == "G":
            m = f"Создаем продукт # {num_row} {row_id[10]} {row_id[11]}"
            log(m, True)
            id_and_cfd = new_product(ss, wr, row_id, row_project,
                                     num_row, folder_id)
            # создадим вложенные вехи
            templ_id = {}  # словаррь соответсвия id из шаблона с id созданных
            templ_id[template_id] = id_and_cfd[0]
            new_sub_task_rekurs(ss, wr, id_and_cfd[0],
                                id_and_cfd[1], template_id)


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

    log("Выгрузка проектов из Гугл во Wrike", True)
    load_from_google_to_wrike(ss, wr, users_from_name, users_from_id,
                              template_id, folder_id)

    t_finish = time.time()
    print("Выполненно за:", int(t_finish - t_start), " секунд")


if __name__ == '__main__':
    main()
