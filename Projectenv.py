# модуль для параметров проекта и общих для всех процедур классов
# все параметры собираем из settings.cfg

import configparser
import time

import Wrike
import Spreadsheet
import Baselog


class Projectenv():
    ''' Содержит пераметры с константами и общими объектами для
        всех процедур проекта
        wr -  API Wrike
        ss - API Google Sheet
        db - вывод лога в базу mySQL и/или на экран
    '''

    def __init__(self, line_argument=None):
        self.CREDENTIALS_FILE = "creds.json"
        cfg = configparser.ConfigParser()
        cfg.read('settings.cfg')
        # spreadsheet секция
        self.TABLE_ID = cfg["spreadsheet"]["gogletableid"]
        self.REPORT_ID = cfg["spreadsheet"]["gogletableid_report"]
        self.cell_color = cfg["spreadsheet"]["cell_color"]
        self.cell_calendar = cfg["spreadsheet"]["cell_calendar"]
        self.cell_user = cfg["spreadsheet"]["cell_user"]
        self.cell_log = cfg["spreadsheet"]["cell_log"]
        self.sheets = {"work_sheet": (cfg["spreadsheet"]["work_sheet"],
                                      cfg["spreadsheet"]["work_sheet_id"])}

        self.sheets["project_sheet"] = (cfg["spreadsheet"]["project_sheet"],
                                        cfg["spreadsheet"]["project_sheet_id"])

        self.sheets["report_stages"] = (cfg["spreadsheet"]["report_stages"],
                                        cfg["spreadsheet"]["report_stages_id"])

        self.column_project = cfg["spreadsheet"]["column_project"]
        self.table_project = cfg["spreadsheet"]["table_project"]
        self.table_params = cfg["spreadsheet"]["table_params"]
        self.table_plan = cfg["spreadsheet"]["table_plan"]
        self.start_work_table = cfg["spreadsheet"]["start_work_table"]
        self.column_id = cfg["spreadsheet"]["column_id"]
        self.column_log = cfg["spreadsheet"]["column_log"]
        self.column_link = cfg["spreadsheet"]["column_link"]
        columns_stage = cfg["spreadsheet"]["columns_stage"]
        self.columns_stage = columns_stage.split(",")
        self.cell_progress = cfg["spreadsheet"]["cell_progress"]
        self.column_comand = cfg["spreadsheet"]["column_comand"]
        self.color_bc_finish = Spreadsheet.htmlColorToJSON("#cfe2f3")
        self.color_font_finish = Spreadsheet.htmlColorToJSON("#000000")
        self.color_bc = Spreadsheet.htmlColorToJSON("#6aa84f")
        self.color_font = Spreadsheet.htmlColorToJSON("#ffffff")
        custom_value_name = cfg["spreadsheet"]["custom_value_name"]
        self.custom_value_name = custom_value_name.split(",")
        custom_value_key = cfg["spreadsheet"]["custom_value_key"]
        self.custom_value_key = custom_value_key.split(",")
        custom_value_adr = cfg["spreadsheet"]["custom_value_adr"]
        self.custom_value_adr = custom_value_adr.split(",")
        # wrike секция
        self.TOKEN = cfg["wrike"]["wriketoken"]
        self.prototip_link = cfg["wrike"]["prototip_link"]
        self.folder_link = cfg["wrike"]["folder_link"]
        # log секция
        self.OUT = cfg["log"]["out"]
        self.NAME = cfg["log"]["name"]
        self.PASS = cfg["log"]["pass"]
        self.HOST = cfg["log"]["host"]
        self.BASE = cfg["log"]["base"]
        # параметры для аргументов запуска руководитель и номер строки
        self.rp_filter = None
        self.row_filter = None
        self.HOLYDAY = None
        self.COLOR_FINISH = None
        self.wr = None
        self.ss = None
        self.db = None
        self.template_id = None
        self.parent_id = None
        self.users_name = None
        self.users_id = None

        self.user_token = None
        self.rp_filter = None
        self.row_filter = None
        self.update_all_cv = None
        self.what_do = "sync"

        # обрабатываем список аргументов
        # 0 - сам скрипт
        # 1-й токен пользователя или admin для админа - логи только в терминал
        # 1   help - помощь по параметрам запускаа
        # 2-й -rp
        # 3-й читаем если есть -rp -  менеджер проекта по которому фильтр
        # 4-й -n
        # 5-й читаем еслить есть -n номер строки котору обрабатывать
        if not line_argument:
            line_argument = ["", "admin"]
        if len(line_argument) == 1:
            self.user_token = None
        else:
            self.user_token = line_argument[1]

        while len(line_argument) > 2:
            key = line_argument.pop(2)
            if key == "-rp":
                self.rp_filter = line_argument.pop(2)
            elif key == "-n":
                self.row_filter = int(line_argument.pop(2))
            if key == "-cv":
                self.update_all_cv = True
            if key == "-do":
                self.what_do = line_argument.pop(2)

    def connect_db(self):
        if self.user_token == "admin" or self.user_token == "help":
            self.OUT = "terminal"
            self.db = Baselog.Baselog(self.OUT, admin_mode=True)
        else:
            self.db = Baselog.Baselog(self.OUT)
            self.db.connect(self.NAME, self.PASS, self.HOST, self.BASE)
        if not self.db.is_connected():
            print("Нет коннекта к базе с логом")
            return False
        else:
            return True

    def user_in_base(self):
        if self.user_token == "admin":
            return True
        if not self.user_token:
            print("Укажите токен пользователя или help")
            return False
        if self.user_token == "help":
            print(" user_token или help или admin. Обязательный параметр")
            print(" -rp <'имя менеджера проекта для фильтра'>."
                  " Не обязательный.")
            print(" -n <номер_строки для фильтра>. Не обязательный.")
            print(" -cv обновить все пользовательские поля  Не обязательный.")
            print((" -do <sync, refl, W, R> Что выполнять."
                   "Не обязательный."))
            print("   sync - тольк синхронизация wrike из Гугл")
            print("   refl - только отражение вех в стратегии")
            print("   W - обновление дат в Гугл и заполнение отчетов")
            print("   R - только заполнение отчетов")
            return False
        ok = self.db.get_user(self.user_token)
        if not ok:
            print("Пользователь по токену", self.user_token, "не найден")
            return False
        else:
            if self.db.rp_filter and self.rp_filter is None:
                self.rp_filter = self.db.rp_filter
            return True

    def connect_spreadsheet(self, tableid=None):
        self.db.out("Приосоединяемся к Гугл")
        self.ss = Spreadsheet.Spreadsheet(self.CREDENTIALS_FILE)
        if not tableid:
            tableid = self.TABLE_ID
        self.ss.set_spreadsheet_byid(tableid)
        # доделать с исключениеями
        return True

    def connect_wrike(self):
        self.db.out("Приосоединяемся к Wrike")
        self.wr = Wrike.Wrike(self.TOKEN)
        del self.TOKEN
        # доделать с исключениеями
        return True

    def read_global_env(self):
        self.HOLYDAY = self.read_holiday(self.cell_calendar)

        fl = self.ss.read_color_cells(self.cell_color)[0][0]
        self.COLOR_FINISH = (fl[0], fl[1])

        self.db.out("Получить id шаблонoв, id email пользователей")
        # Основной прототип в папке прототипов
        template = self.wr.get_folders("folders",
                                       permalink=self.prototip_link)[0]
        self.template_id = template["id"]
        template_title = template["title"]
        self.db.out(f"ID шаблона {template_title} {self.template_id}")
        # 000 НОВЫЕ ПРОДУКТЫ
        parent = self.wr.get_folders("folders", permalink=self.folder_link)[0]
        self.parent_id = parent["id"]
        parent_title = parent["title"]
        self.db.out(f"ID папки  {parent_title} {self.parent_id}")
        self.get_user_list(self.cell_user)

        if self.rp_filter:
            self.db.out(f"Обрабатываем строки у РП {self.rp_filter}")
        if self.row_filter:
            self.db.out(f"Обрабатываем строку #{self.row_filter}")
        # доделать с исключениеями
        return True

    def sheet_now(self, sheet):
        ''' устанавливает рабочую таблицу в ss
            work_sheet
            project_sheet
        '''
        name_and_id = self.sheets.get(sheet)
        if name_and_id:
            self.ss.sheetTitle = name_and_id[0]
            self.ss.sheetId = int(name_and_id[1])
            return True
        else:
            return False

    def get_user_list(self, cell_user):
        ''' возвращает два словаря пользователей
        ключ по имени из гугл таблицы
        ключ по id из wrike
        '''
        def find_name_from_email(email, users_from_name):
            for user_name, user_val, in users_from_name.items():
                if user_val["email"] == email:
                    return user_name, user_val["group"]

        table = self.ss.values_get(cell_user)
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
        id_dict = self.wr.id_contacts_on_email(lst_mail)
        for email, id_user in id_dict.items():
            if email in lst_mail and id_user is not None:
                users_from_id[id_user] = {}
                users_from_id[id_user]["email"] = email
                name_user, gr_user = find_name_from_email(email,
                                                          users_from_name)
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

        self.users_name = users_from_name.copy()
        self.users_id = users_from_id.copy()

    def read_holiday(self, cell_calendar):
        from datetime import date
        ''' считываем из гугл таблицы рабочий календарь
        '''
        holyday = []
        holidays_str = self.ss.values_get(cell_calendar)
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
                holyday.append(date(int(d_str[2]), int(d_str[1]),
                                    int(d_str[0])))
        return holyday

    def print_ss(self, msg, cells_range, save=True, timeout=1):
        value = [[msg]]
        my_range = f"{cells_range}:{cells_range}"
        self.ss.prepare_setvalues(my_range, value)
        if save:
            time.sleep(timeout)
            self.ss.run_prepared()

    def get_project_on_row(self, num_row=None):
        ''' читает из таблицы проект
            если в параметрах запуска передан номер строки
            то читает из него
        '''
        column = self.column_project
        return_value = None
        if not num_row:
            num_row = self.row_filter
        if num_row:
            table = self.ss.values_get(f"{column}{num_row}:{column}{num_row}")
            if len(table) > 0:
                return_value = table[0][0]
        return return_value

    def get_project_list(self):
        ''' забираем из Гугл таблицу проектов
        '''
        self.sheet_now("project_sheet")
        return self.ss.values_get(self.table_project)

    def update_project_name(self, wrike_link, name_project, num_row):
        ''' обновляет название папки-проекта из таблицы
            возвращает id папки
        '''
        return_value = None
        resp = self.wr.get_folders("folders", permalink=wrike_link)
        if len(resp) == 0:
            self.db.out(f"нет папки во wrike {name_project} {wrike_link}",
                        num_row=num_row,
                        runtime_error="y", error_type="создание папки")
        else:
            if resp[0]["title"] != name_project.upper():
                # обновим имя папки во Wrike
                resp = self.wr.update_folder(resp[0]["id"],
                                             title=name_project.upper())
                if len(resp) == 0:
                    self.db.out(f"не обновли папку {name_project}",
                                num_row=num_row,
                                runtime_error="y", error_type="создание папки")
            return_value = resp[0]["id"]
        return return_value

    def make_folder(self, name_project, num_row):
        ''' создаем папку для продуктов из проект таблицы
        '''
        return_value = None
        self.db.out(f"Создаем папку для проекта #{num_row} {name_project}")
        resp = self.wr.create_folder(self.parent_id, name_project.upper())
        if len(resp) == 0:
            self.db.out(f"не создали папку {name_project}", num_row=num_row,
                        runtime_error="y", error_type="создание папки")
        else:
            return_value = resp[0]["id"]
            self.sheet_now("project_sheet")
            self.print_ss(resp[0]["permalink"], f"E{num_row}")
        return return_value

    def get_work_table(self):
        '''  возвращает словарь
            ключ - номер строки
            значения - словарь полей
            словарь возвращается сразу отфильтрованным по руководителю
            и номеру строки
        '''
        def make_dict_from_column(row_id, row_project, num_row):
            return_dict = {}
            return_dict["check"] = row_id[0]
            return_dict["log"] = row_id[1]
            return_dict["id_project"] = row_id[2]
            return_dict["comand"] = row_project[0]
            return_dict["permalink"] = row_project[1]

            for key, index in zip(self.custom_value_key,
                                  self.custom_value_adr):
                return_dict[key] = row_id[int(index)]
            tmp_lst = []
            for k in range(2, len(row_project)):
                tmp_lst.append(row_project[k])
            return_dict["dates"] = tmp_lst.copy()
            return_dict["template_id"] = None
            return_dict["template_str"] = None
            if len(row_project) == 11:
                nrt = row_project[10]
                if len(nrt) > 0 and nrt.isdigit():
                    adr = f"{self.column_id}{nrt}:{self.column_id}{nrt}"
                    p_t = self.ss.values_get(adr)[0][0]
                    if not p_t and return_dict["comand"] == "G":
                        self.db.out(f"в строке {nrt} нет шаблона",
                                    num_row=num_row,
                                    runtime_error="y",
                                    error_type="ошибка шаблона")
                    else:
                        return_dict["template_id"] = p_t
                        return_dict["template_str"] = nrt
            return return_dict
        return_dict = {}
        self .sheet_now("work_sheet")
        table_id = self.ss.values_get(self.table_params)
        table_project = self.ss.values_get(self.table_plan)
        num_row = int(self.start_work_table)
        if self.row_filter:
            if self.row_filter <= len(table_project):
                table_project = [table_project[self.row_filter - 1]]
                num_row = self.row_filter - 1
            else:
                return return_dict
        else:
            table_project = table_project[19:]
        for row_project in table_project:
            num_row += 1
            if len(row_project) == 0:
                continue
            if len(table_id) < num_row:
                row_id = ["" for x in range(0, 30)]
            else:
                row_id = table_id[num_row - 1]
            if len(row_id) < 30:
                row_id.extend(["" for x in range(0, 30 - len(row_id))])

            if row_project[0] and row_project[0] in "GAPWN":
                if self.rp_filter:
                    if row_id[5] != self.rp_filter:
                        continue
                c_v = make_dict_from_column(row_id, row_project, num_row)
                return_dict[num_row] = c_v.copy()
        return return_dict

    def get_len_stage(self, num_stage, num_template="#1"):
        template = {}
        template["#1"] = {}
        template["#1"]["1"] = 4
        template["#1"]["2"] = 5
        template["#1"]["3"] = 5
        template["#1"]["4"] = 1
        template["#1"]["5"] = 10
        template["#1"]["6"] = 92
        template["#1"]["7"] = 7
        template["#1"]["8"] = 2

        my_tmpl = template[num_template]
        return my_tmpl.get(num_stage)

    def progress(self, percent=0, width=30):
        if self.db.log == "all" or self.db.log == "terminal":
            left = int(width * percent) // 1
            right = width - left
            print('\r[', '#' * left, ' ' * right, ']',
                  f' {percent * 100:.0f}%',
                  sep='', end='', flush=True)

    def find_task_user(self, task_user, rp, tech, num_task, remove=True):
        '''Возвращает список пользователей которых нужно удалить из шаблона
            или установить в шаблон
        '''
        return_list = []
        manager_id = self.users_name[rp]["id"]
        teh_id = self.users_name[tech]["id"]

        group_user = {"RP": [], "RP_helper": [], "Tech": [], "Other": []}
        for num, user in enumerate(task_user, 1):
            group = self.users_id[user]["group"]
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
                    id_helper = self.users_name[rp].get("idhelper")
                    if not id_helper:
                        return_list.extend(group_user["RP_helper"])

        else:
            if len(group_user["RP"]) > 0:
                return_list.append(manager_id)
                id_helper = self.users_name[rp].get("idhelper")
                if id_helper and num_task.find("*") > 0:
                    return_list.append(id_helper)
            if len(group_user["Tech"]) > 0:
                return_list.append(teh_id)
            return_list.extend(group_user["Other"])
        return return_list

    def find_cf(self, resp_cf, name_cf):
        ''' Ищем в списке полей поле с нужным id по имени
            возвращеем значение
        '''
        return_value = ""
        id_field = self.wr.customfields[name_cf]
        for cf in resp_cf:
            if cf["id"] == id_field:
                return_value = cf["value"]
                break
        return return_value


if __name__ == '__main__':
    pass
