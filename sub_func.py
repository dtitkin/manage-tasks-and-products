
from datetime import datetime, date


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


# def log(msg, prn_time=False, one_str=False):
#    str_now = now_str()
#    if prn_time:
#        print(str_now, end=" ")
#        if one_str:
#            print('\r', msg, sep='', end='', flush=True)
#       else:
#            pprint(msg)
#    else:
#        if one_str:
#            print('\r', msg, sep='', end='', flush=True)
#        else:
#            pprint(msg)


def log_ss(ss, msg, cells_range):
    value = [[msg]]
    my_range = f"{cells_range}:{cells_range}"
    ss.prepare_setvalues(my_range, value)
    ss.run_prepared()


def make_date(usr_date):
    if not usr_date:
        return date.today()
    if usr_date.find(".") > -1:
        lst_date = usr_date[0:10].split(".")
    else:
        lst_date = usr_date[0:10].split("/")
    return date(int(lst_date[2]), int(lst_date[1]), int(lst_date[0]))


def read_holiday(ss):
    ''' считываем из гугл таблицы рабочий календарь
    '''
    holyday = []
    holidays_str = ss.values_get("Рабочий календарь!A:A")
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
            holyday.append(date(int(d_str[2]), int(d_str[1]), int(d_str[0])))
    return holyday


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


def read_stage_info(ss, wr, num_row, color_finish):
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
        if str(color_stage) == str(color_finish):
            # запоминаем у каких этапов нужно установить Finish
            finish_list.append(str(k))
        dates_stage[str(k)] = make_date(stage_value)
    return finish_list, dates_stage


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
