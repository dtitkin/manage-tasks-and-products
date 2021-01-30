import os
from pprint import pprint
from datetime import date
import time

import Wrike
import Spreadsheet

CREDENTIALS_FILE = "creds.json"
TABLE_ID = os.getenv("gogletableid")
TOKEN = os.getenv("wriketoken")


def clear_experiment(root_task_id, folder_id, wr):
    api_str = f"folders/{folder_id}/tasks"
    list_task = wr.get_tasks(api_str)
    len_list = len(list_task)
    print("В списке", len_list, "задач")
    for n, task in enumerate(list_task, 1):
        id_task = task["id"]
        if id_task != root_task_id:
            progress(n / len_list)
            wr.rs_del(f"tasks/{id_task}")
    print()


def progress(percent=0, width=30):
    left = int(width * percent) // 1
    right = width - left
    print('\r[', '#' * left, ' ' * right, ']',
          f' {percent * 100:.0f}%',
          sep='', end='', flush=True)



def main():
    print("Приосоединяемся к Гугл")
    ss = Spreadsheet.Spreadsheet(CREDENTIALS_FILE)
    ss.set_spreadsheet_byid(TABLE_ID)
    print("Приосоединяемся к Wrike")
    wr = Wrike.Wrike(TOKEN)
    print("Получить шаблон задач из Гугл")
    table = ss.values_get("Задачи этапов!A20:I32") # I91
    print("Получить папку и корневую задачу из Wrike")
    name_sheet = "001 ШАБЛОНЫ (новые продукты Рубис)"
    folder_id = wr.id_folders_on_name([name_sheet])[name_sheet]
    print(folder_id)
    api_str = f"folders/{folder_id}/tasks"
    resp = wr.get_tasks(api_str, title="[КОД] [РАБОЧЕЕ НАЗВАНИЕ]")
    root_task_id = []
    root_task_id.append(resp[0]["id"])
    print(root_task_id[0])
    print("Отчистим результаты предудущих экспериментов")
    clear_experiment(root_task_id[0], folder_id, wr)
    print("Создать этапы и задачи")
    root_task_id.append("")
    pa_task = ""
    task_in_stage = {}
    dependency_stage = ""
    for row in table:
        if len(row) == 0:
            continue
        number_stage = row[0]
        name_stage = row[1]
        if number_stage:
            # создаем этап
            print("Веха:", number_stage, name_stage, "# ", end="")
            dt = {"type": "Milestone", "due": date.today().isoformat()}
            st = [root_task_id[0]]
            resp = wr.create_task(folder_id, name_stage,
                                  dates=dt, priorityAfter=pa_task,
                                  superTasks=st)
            root_task_id[1] = resp[0]["id"]
            pa_task = root_task_id[1]
            print(pa_task)
            # создадим связи между этапами
            if dependency_stage:
                resp = wr.create_dependency(pa_task,
                                            predecessorId=dependency_stage,
                                            relationType="FinishToFinish")
            dependency_stage = pa_task
            task_in_stage = {}
        else:
            # создаем задачу
            number_task = row[2].strip(" ")
            name_task = row[3]
            len_task = int(float(row[6]) * 24 * 60)
            normal_task = row[6]
            dependency_task = row[4]
            dependency_task = dependency_task.split(";")
            print("   Задача", number_task, name_task, "# ", end="")
            dt = {"type": "Planned", "start": date.today().isoformat(),
                  "duration": len_task}
            st = [root_task_id[1]]
            resp = wr.create_task(folder_id, name_task,
                                  dates=dt, priorityAfter=pa_task,
                                  superTasks=st,)
            pa_task = resp[0]["id"]
            print(pa_task)
            #  создаем связи между задачами
            task_in_stage[number_task] = pa_task
            for pred_task in dependency_task:
                pred_task = pred_task.strip(" ")
                if not pred_task:
                    continue
                predid = task_in_stage[pred_task]
                resp = wr.create_dependency(pa_task, predecessorId=predid,
                                            relationType="FinishToStart")





if __name__ == '__main__':
    main()
