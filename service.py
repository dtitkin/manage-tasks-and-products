import os
from pprint import pprint
from datetime import date
import time

import Wrike
import Spreadsheet

CREDENTIALS_FILE = "creds.json"
TABLE_ID = os.getenv("gogletableid")
TOKEN = os.getenv("wriketoken")


def main():
    print("Приосоединяемся к Гугл")
    ss = Spreadsheet.Spreadsheet(CREDENTIALS_FILE)
    ss.set_spreadsheet_byid(TABLE_ID)
    print("Приосоединяемся к Wrike")
    wr = Wrike.Wrike(TOKEN)
    print("Получить шаблон задач из Гугл")
    table = ss.values_get("Задачи этапов!A20:I91")
    print("Получить папку и корневую задачу из Wrike")
    name_sheet = "001 ШАБЛОНЫ (новые продукты Рубис)"
    folder_id = wr.id_folders_on_name([name_sheet])[name_sheet]
    print(folder_id)
    api_str = f"folders/{folder_id}/tasks"
    root_task_id = wr.get_tasks(api_str, title="Новый продукт Шаблон")[0]["id"]
    print(root_task_id)
    print("Создать этапы и задачи")
    for row in table:
        if len(row) == 0:
            continue
        number_stage = row[0]
        name_stage = row[1]
        if number_stage:
            # создаем этап
            print("------------------------", number_stage, name_stage)
            dt = str({"type": "Milestone", "due": date.today().isoformat()})
            st = str([root_task_id])
            wr.create_task(folder_id, name_stage, dates=dt, superTasks=st)


if __name__ == '__main__':
    main()
