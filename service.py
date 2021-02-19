import os
from pprint import pprint
from datetime import date, timedelta

import time
import builtins
from numpy import busday_count

import Wrike
import Spreadsheet

CREDENTIALS_FILE = "creds.json"
TABLE_ID = os.getenv("gogletableid")
TOKEN = os.getenv("wriketoken")
ONE_DAY = timedelta(days=1)


def float(x):
    x = x.replace(",", ".")
    return builtins.float(x)


def clear_experiment(root_task_id, folder_id, wr):
    api_str = f"folders/{folder_id}/tasks"
    list_task = wr.get_tasks(api_str)
    len_list = len(list_task)
    print("В списке", len_list, "задач")
    for n, task in enumerate(list_task, 1):
        id_task = task["id"]
        progress(n / len_list)
        if id_task != root_task_id:
            wr.rs_del(f"tasks/{id_task}")
    print()


def find_id_task(stage_and_task, pred_task):
    return_id = None
    for stage in stage_and_task.values():
        if return_id:
            break
        for number_task, task in stage["subtask"].items():
            if number_task == pred_task:
                return_id = task["id"]
                break
    return return_id


def create_template(work, sheet_and_range):
    ''' Создает шаблон во Wrike на основе Гугл таблицы
    '''

    t_start = time.time()
    print("Приосоединяемся к Гугл")
    ss = Spreadsheet.Spreadsheet(CREDENTIALS_FILE)
    ss.set_spreadsheet_byid(TABLE_ID)

    print("Приосоединяемся к Wrike")
    wr = Wrike.Wrike(TOKEN)

    print("Получить шаблон задач из Гугл")
    table = ss.values_get(sheet_and_range)

    print("Получить папку, корневую задачу")
    name_sheet = "001 ШАБЛОНЫ (новые продукты Рубис)"
    folder_id = wr.id_folders_on_name([name_sheet])[name_sheet]
    print(folder_id)
    api_str = f"folders/{folder_id}/tasks"
    resp = wr.get_tasks(api_str, title="1CКОД РАБОЧЕЕ НАЗВАНИЕ")
    if work == "all":
        root_task_id = resp[0]["id"]
    else:
        root_task_id = ""
    print(root_task_id)

    if work == 'all' or work == 'del':
        print("Отчистим результаты предудущих экспериментов")
        clear_experiment(root_task_id, folder_id, wr)

    if work == 'all' or work == 'create':
        print("Создать этапы и задачи")
        # task_in_stage = {}
        stage_and_task = {}
        my_taskid = ""
        predecessor_stage = ""
        max_date = date.today()
        for row in table:
            if len(row) == 0:
                continue
            itstage = row[0]
            if itstage:
                # создаем этап
                number_stage = row[0]
                name_stage = number_stage + ": " + row[1]
                name_stage += " [код рабочее название]"
                if len(row) == 9:  # чек лист не во всех строках
                    descr = wr.deskr_from_str(row[8])
                else:
                    descr = ""
                print("Веха:", name_stage, "# ", end="")
                dt = wr.dates_arr(type_="Milestone", due=max_date.isoformat())
                st = [root_task_id]
                cfd = {"Номер этапа": number_stage,
                       "Название рабочее": "рабочее название",
                       "Код-1С": "код",
                       "Руководитель проекта": "Меуш Николай",
                       "Клиент": "Альпинтех",
                       "Бренд": "TitBit",
                       "Группа": "Группа продукта",
                       "Линейка": "Линейка продукта"}
                cf = wr.custom_field_arr(cfd)
                resp = wr.create_task(folder_id, name_stage, description=descr,
                                      dates=dt, priorityAfter=my_taskid,
                                      superTasks=st, customFields=cf)

                my_taskid = resp[0]["id"]
                print(my_taskid)
                stage_and_task[number_stage + "s"] = {}
                my_stage = stage_and_task[number_stage + "s"]
                my_stage["subtask"] = {}
                my_stage["predecessor_stage"] = predecessor_stage
                my_stage["id"] = my_taskid
                subtask = my_stage["subtask"]
                predecessor_stage = my_taskid
            else:
                # создаем задачу
                number_task = row[2].strip(" ")
                if not number_task:
                    continue
                number_task = int(number_task)
                name_task = row[3] + " [код рабочее название]"
                len_task = int(float(row[6]) * 8 * 60)
                standart_task = float(row[7])
                if len(row) == 9:
                    descr = wr.deskr_from_str(row[8])
                else:
                    descr = ""
                dependency_task = row[4]
                print("   Задача", number_task, name_task, "# ", end="")
                dt = wr.dates_arr(start=max_date.isoformat(),
                                  duration=len_task)
                st = [predecessor_stage]
                cfd = {"Номер этапа": number_stage,
                       "Название рабочее": "рабочее название",
                       "Код-1С": "код",
                       "Номер задачи": number_task,
                       "Норматив часы": standart_task,
                       "Руководитель проекта": "Меуш Николай",
                       "Клиент": "Альпинтех",
                       "Бренд": "TitBit",
                       "Группа": "Группа продукта",
                       "Линейка": "Линейка продукта"}
                cf = wr.custom_field_arr(cfd)
                resp = wr.create_task(folder_id, name_task, description=descr,
                                      dates=dt, priorityAfter=my_taskid,
                                      superTasks=st, customFields=cf)
                my_taskid = resp[0]["id"]
                print(my_taskid)
                subtask[number_task] = {}
                subtask[number_task]["id"] = my_taskid
                subtask[number_task]["dependency_task"] = dependency_task
                subtask[number_task]["stage"] = number_stage + "s"

        #  создаем связи между этапами и задачами
        print("Создадим связи между задачами и вехами")
        for stage in stage_and_task.values():
            print(stage["id"])
            id_predecessor = stage["predecessor_stage"]
            if id_predecessor:
                resp = wr.create_dependency(stage["id"],
                                            predecessorId=id_predecessor,
                                            relationType="FinishToFinish")
            n = 0
            len_sb = len(stage["subtask"])
            for number_task, task in stage["subtask"].items():
                n += 1
                progress(n / len_sb)
                dependency_task = task["dependency_task"].split(";")
                id_task = task["id"]
                for pred_task in dependency_task:
                    pred_task = pred_task.strip(" ")
                    if not pred_task:
                        continue
                    pred_task = int(pred_task)
                    if pred_task == number_task:
                        continue
                    pred_id = find_id_task(stage_and_task, pred_task)
                    resp = wr.create_dependency(id_task,
                                                predecessorId=pred_id,
                                                relationType="FinishToStart")
    # передвинем даты у всех вех на последнюю задачу у вехи
    if work == 'all' or work == 'date':
        print("Перенесем вехи на последнюю дату выполнения задач")
        wr.update_milestone_date(folder_id, root_task_id)
    t_finish = time.time()
    print("Выполненно за:", int(t_finish - t_start), " секунд")


def set_conacts_on_table():
    ''' Загружает в Гугл таблицу пользователей на страницу ресурсы
    '''
    t_start = time.time()
    print("Приосоединяемся к Гугл")
    ss = Spreadsheet.Spreadsheet(CREDENTIALS_FILE)
    ss.set_spreadsheet_byid(TABLE_ID)

    print("Приосоединяемся к Wrike")
    wr = Wrike.Wrike(TOKEN)

    # print("Получить ")
    # table = ss.values_get(sheet_and_range)
    print("Получить список контактов из Wrike")
    resp = wr.rs_get("contacts")
    # pprint(resp)
    value = []
    for row in resp:
        fn = row["firstName"]
        ln = row["lastName"]
        user_name = ln + " " + fn
        # user_id = row["id"]
        user_mail = str(row["profiles"][0].get("email"))
        if user_mail.find("wrike") == -1 and user_mail.find("None") == -1:
            row_list = [user_name, "", user_mail]
            value.append(row_list)

    print("Записать данные в таблицу")
    num_row = 49 + len(value)
    my_range = "B50:D" + str(num_row)
    # тестируем установку одного значения на страницу
    ss.sheetTitle = "Ресурсы"
    ss.prepare_setvalues(my_range, value)
    ss.run_prepared()
    t_finish = time.time()
    print("Выполненно за:", int(t_finish - t_start), " секунд")


def del_in_main(name_sheet="000 НОВЫЕ ПРОДУКТЫ"):
    ''' удалить все задачи из основной папки
    '''
    t_start = time.time()

    print("Приосоединяемся к Wrike")
    wr = Wrike.Wrike(TOKEN)

    print("Получить папку", name_sheet)
    folder_id = wr.id_folders_on_name([name_sheet])[name_sheet]
    print(folder_id)
    root_task_id = ""

    print("Отчистим результаты предудущих экспериментов")
    clear_experiment(root_task_id, folder_id, wr)
    t_finish = time.time()
    print("Выполненно за:", int(t_finish - t_start), " секунд")


def del_in_templ():
    del_in_main(name_sheet="001 ШАБЛОНЫ (новые продукты Рубис)")



def len_mile():

    def find_stage(resp, find_num=""):
        sub = resp[0]["subTaskIds"]
        txt_req = "tasks/"
        for id_t in sub:
            txt_req = txt_req + id_t + ","
        txt_req = txt_req[0:-1]
        resp = wr.get_tasks(txt_req)
        returt_d = {}
        id_001 = ""
        for task in resp:
            resp_cf = task["customFields"]
            num_t = find_cf(wr, resp_cf, "Номер задачи")
            if find_num == "":
                if num_t == "001":
                    id_001 = task["id"]
                due = task["dates"]["due"]
                returt_d[num_t] = due
            elif find_num == num_t:
                start = task["dates"]["start"]
                returt_d[num_t] = start
        return returt_d, id_001

    print("Приосоединяемся к Гугл")
    ss = Spreadsheet.Spreadsheet(CREDENTIALS_FILE)
    ss.set_spreadsheet_byid(TABLE_ID)
    holidays_str = ss.values_get("Рабочий календарь!A:A")
    holidays = []
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
            holidays.append(date(int(d_str[2]), int(d_str[1]), int(d_str[0])))
    print("Приосоединяемся к Wrike")
    wr = Wrike.Wrike(TOKEN)
    print("Получить папку, корневую задачу")
    name_sheet = "001 ШАБЛОНЫ (новые продукты Рубис)"
    folder_id = wr.id_folders_on_name([name_sheet])[name_sheet]
    print(folder_id)
    api_str = f"folders/{folder_id}/tasks"
    resp = wr.get_tasks(api_str, title="1CКОД РАБОЧЕЕ НАЗВАНИЕ")
    root_task_id = resp[0]["id"]
    print(root_task_id)
    resp = wr.rs_get(f"tasks/{root_task_id}")
    d_task, id_001 = find_stage(resp)
    resp = wr.get_tasks(f"tasks/{id_001}")
    d_t1, id_001 = find_stage(resp, "1")
    d_task.update(d_t1)
    # переведем в даты
    for key, value in d_task.items():
        lst = value[0:10].split("-")
        my_dt = date(int(lst[0]), int(lst[1]), int(lst[2]))
        d_task[key] = my_dt
    sort_list = list(d_task.items())
    sort_list.sort(key=lambda i: i[1])
    start_date = sort_list[0][1] - ONE_DAY
    for nxt_stage in sort_list:
        if nxt_stage[0] == "0052" or nxt_stage[0] == "0051":
            continue
        end_date = nxt_stage[1]
        len_stage = busday_count(start_date, end_date, weekmask="1111100",
                                 holidays=holidays)
        print("Этап", nxt_stage[0], "старт:", start_date, "финиш:",
              end_date, "длительность:", len_stage)
        start_date = end_date


def update_name():
    print("Приосоединяемся к Wrike")
    wr = Wrike.Wrike(TOKEN)

    print("Получить id шаблона")
    permalink = "https://www.wrike.com/open.htm?id=637661621"
    #  "#1 1CКОД РАБОЧЕЕ НАЗВАНИЕ"
    template_id = wr.get_folders("folders", permalink=permalink)[0]["id"]
    print(f"ID шаблона {template_id}")

    resp = wr.get_tasks(f"folders/{template_id}/tasks", subTasks="true")
    n = 0
    len_t = len(resp)
    for task in resp:
        n += 1
        progress(n / len_t)
        name_task = task["title"]
        id_task = task["id"]
        end_s = name_task.find("[")
        if end_s > -1:
            name_task = name_task[0:end_s]
            wr.update_task(id_task, title=name_task)


def update_cf():
    print("Приосоединяемся к Wrike")
    wr = Wrike.Wrike(TOKEN)

    print("Получить id шаблона")
    permalink = "https://www.wrike.com/open.htm?id=637661621"
    #  "#1 1CКОД РАБОЧЕЕ НАЗВАНИЕ"
    template_id = wr.get_folders("folders", permalink=permalink)[0]["id"]
    print(f"ID шаблона {template_id}")
    fields = ["customFields"]
    resp = wr.get_tasks(f"folders/{template_id}/tasks", subTasks="true",
                        fields=fields)
    n = 0
    len_t = len(resp)
    for task in resp:
        n += 1
        progress(n / len_t)
        id_task = task["id"]
        resp_cf = task["customFields"]
        cfd = {}
        cfd["num_stage"] = find_cf(wr, resp_cf, "Номер этапа")
        cfd["num_task"] = find_cf(wr, resp_cf, "Номер задачи")
        cfd["Норматив часы"] = find_cf(wr, resp_cf, "Норматив часы")
        cf = wr.custom_field_arr(cfd)
        wr.update_task(id_task, customFields=cf)


def rout_on_sys_argv():
    import sys
    d_globals = globals()
    lst_func = sys.argv
    my_f = lst_func[1]
    # lst_f = ["len_mile", "del_in_main", "del_in_templ"]
    # lst_f = ["update_name"]
    lst_f = ["update_cf"]
    if my_f in lst_f:
        d_globals[my_f]()


if __name__ == '__main__':
    # work = "all"
    # sheet_and_range = "Задачи этапов!A20:I89"
    # create_template(work, sheet_and_range)
    # установка контактов в таблицу
    # set_conacts_on_table()
    # del_in_main()
    rout_on_sys_argv()
    # print("Убери комментарий с нужной функции")
