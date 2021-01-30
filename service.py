import os
from pprint import pprint
from datetime import date, timedelta
#from collections import OrderedDict
import time
import builtins

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


def progress(percent=0, width=30):
    left = int(width * percent) // 1
    right = width - left
    print('\r[', '#' * left, ' ' * right, ']',
          f' {percent * 100:.0f}%',
          sep='', end='', flush=True)


def end_task_date(wr, my_taskid):
    'читает дату заверешения у задачи и переводит ее в формат date'
    resp = wr.get_tasks(f"tasks/{my_taskid}")
    d_str = resp[0]["dates"]["due"]
    y, m, d = map(int, d_str[0:10].split('-'))
    my_date = date(y, m, d)
    return my_date


def create_template():
    ''' Создает шаблон во Wrike на основе Гугл таблицы
    '''
    t_start = time.time()
    print("Приосоединяемся к Гугл")
    ss = Spreadsheet.Spreadsheet(CREDENTIALS_FILE)
    ss.set_spreadsheet_byid(TABLE_ID)

    print("Приосоединяемся к Wrike")
    wr = Wrike.Wrike(TOKEN)

    print("Получить шаблон задач из Гугл")
    table = ss.values_get("Задачи этапов!A20:I91")

    print("Получить папку, корневую задачу, словарь полей из Wrike")
    name_sheet = "001 ШАБЛОНЫ (новые продукты Рубис)"
    folder_id = wr.id_folders_on_name([name_sheet])[name_sheet]
    print(folder_id)
    api_str = f"folders/{folder_id}/tasks"
    resp = wr.get_tasks(api_str, title="код-1с рабочее название")
    root_task_id = []
    root_task_id.append(resp[0]["id"])
    print(root_task_id[0])

    print("Отчистим результаты предудущих экспериментов")
    wr.debugMode = False
    clear_experiment(root_task_id[0], folder_id, wr)

    print("Создать этапы и задачи")
    root_task_id.append("")
    my_taskid = ""
    task_in_stage = {}
    predecessor_stage = ""
    max_date = date.today()
    for row in table:
        if len(row) == 0:
            continue
        itstage = row[0]
        if itstage:
            # меняем дату у предыдущего этапа
            if predecessor_stage:
                dt = wr.dates_arr(type_="Milestone", due=max_date.isoformat())
                resp = wr.update_task(predecessor_stage, dates=dt)
            # создаем этап
            number_stage = row[0]
            name_stage = number_stage + ": " + row[1]
            name_stage += " [код рабочее название]"
            print("Веха:", name_stage, "# ", end="")
            dt = wr.dates_arr(type_="Milestone", due=max_date.isoformat())
            st = [root_task_id[0]]
            cfd = {"Номер этапа": number_stage,
                   "Название рабочее": "рабочее название",
                   "Код-1С": "код",
                   "Руководитель проекта": "Меуш Николай",
                   "Клиент": "Альпинтех",
                   "Бренд": "TitBit",
                   "Группа": "Группа продукта",
                   "Линейка": "Линейка продукта"}
            cf = wr.custom_field_arr(cfd)
            resp = wr.create_task(folder_id, name_stage,
                                  dates=dt, priorityAfter=my_taskid,
                                  superTasks=st, customFields=cf)

            my_taskid = resp[0]["id"]
            print(my_taskid)
            root_task_id[1] = my_taskid
            # создадим связи между этапами
            if predecessor_stage:
                resp = wr.create_dependency(my_taskid,
                                            predecessorId=predecessor_stage,
                                            relationType="FinishToFinish")
            predecessor_stage = my_taskid
            task_in_stage = {}
        else:
            # создаем задачу
            number_task = row[2].strip(" ")
            if not number_task:
                continue
            number_task = int(number_task)
            if number_task == 1:
                max_date += ONE_DAY
                start_date = max_date
            name_task = row[3] + " [код рабочее название]"
            len_task = int(float(row[6]) * 8 * 60)
            standart_task = float(row[7])
            dependency_task = row[4]
            dependency_task = dependency_task.split(";")
            set_date = max_date
            if len(dependency_task[0]) == 0:
                set_date = start_date
            print("   Задача", number_task, name_task, "# ", end="")
            dt = wr.dates_arr(start=set_date.isoformat(),
                              duration=len_task)

            st = [root_task_id[1]]
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
            resp = wr.create_task(folder_id, name_task,
                                  dates=dt, priorityAfter=my_taskid,
                                  superTasks=st, customFields=cf)
            my_taskid = resp[0]["id"]
            print(my_taskid)
            #  создаем связи между задачами
            task_in_stage[number_task] = my_taskid
            for pred_task in dependency_task:
                pred_task = pred_task.strip(" ")
                if not pred_task:
                    continue
                pred_task = int(pred_task)
                if pred_task == number_task or pred_task > number_task:
                    continue
                predid = task_in_stage[pred_task]
                resp = wr.create_dependency(my_taskid, predecessorId=predid,
                                            relationType="FinishToStart")
            my_date = end_task_date(wr, my_taskid)
            if my_date > max_date:
                max_date = my_date
    dt = wr.dates_arr(type_="Milestone", due=max_date.isoformat())
    resp = wr.update_task(predecessor_stage, dates=dt)
    resp = wr.update_task(root_task_id[0], dates=dt)
    t_finish = time.time()
    print("Выполненно за:", int(t_finish - t_start), " секунд")


if __name__ == '__main__':
    create_template()
