import os
from pprint import pprint
from datetime import date, timedelta, datetime
from collections import OrderedDict
import time
import builtins

import Wrike
import Spreadsheet

CREDENTIALS_FILE = "creds.json"
TABLE_ID = os.getenv("gogletableid")
TOKEN = os.getenv("wriketoken")
ONE_DAY = timedelta(days=1)


def log(msg, prn_time=False):
    now = datetime.now()
    fmt = '%d.%m.%Y время %H:%M:%S'
    str_now = now.strftime(fmt)
    if prn_time:
        print(str_now, end=" ")
        pprint(msg)
    else:
        pprint(msg)


def get_user(ss, wr):
    ''' возвращает два словааря пользвоателей
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


def main():
    '''Создает шаблон во Wrike на основе Гугл таблицы
    '''

    t_start = time.time()
    log("Приосоединяемся к Гугл", True)
    ss = Spreadsheet.Spreadsheet(CREDENTIALS_FILE)
    ss.set_spreadsheet_byid(TABLE_ID)

    log("Приосоединяемся к Wrike")
    wr = Wrike.Wrike(TOKEN)

    log("Получить папку")
    name_sheet = "000 НОВЫЕ ПРОДУКТЫ"
    folder_id = wr.id_folders_on_name([name_sheet])[name_sheet]
    log(folder_id)

    log("Получить ID, email  пользователей")
    users_from_name, users_from_id = get_user(ss, wr)
    log(users_from_name)
    log(users_from_id)

    log("Выгрузка проектов из Гугл во Wrike", True)



    t_finish = time.time()
    print("Выполненно за:", int(t_finish - t_start), " секунд")


if __name__ == '__main__':
    main()
