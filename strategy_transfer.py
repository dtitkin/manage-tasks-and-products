''' Модуль для отражения вех в папке Wrike со стретегическими проектами
    Переносив вехи проектов из пакпки "Новые продукты"
    Для переноса используется наастроечная таблица из Google Sheet
    "Проекты"
'''
from datetime import date
from time_func import make_date

env = None  # объект для параметров проекта


def create_milestone(name_project, permalink_project,
                     stages_from_row, name_folder,
                     permalink_folder, name_user):
    ''' Создаем вехи если их нет под каждый проект в терминах таблицы
        возвращаем ID вехи, id подзадач, id папки
    '''
    resp = env.wr.get_folders("folders", permalink=permalink_folder)
    if len(resp) == 0:
        env.db.out(f"Проект {name_folder} не найден", runtime_error="y",
                   error_type="перенос проектов в стратегию")
        return False, False, False
    id_folder = resp[0]['id']
    id_cf = env.wr.customfields["Ссылка на проект НП"]
    cf = {"id": id_cf, "comparator": "EqualTo", "value": permalink_project}

    fields = ["subTaskIds", "customFields"]
    resp = env.wr.get_tasks(f"folders/{id_folder}/tasks", type="Milestone",
                            customField=cf, fields=fields)
    sub_tasks = []
    if len(stages_from_row) > 1:
        start_stage = stages_from_row[0]
        end_stage = stages_from_row[-1]
        name_task = name_project + f" этапы с {start_stage} по {end_stage}"
    else:
        end_stage = stages_from_row[0]
        name_task = name_project + f" этап {end_stage}"
    if len(resp) == 0:
        # нужной вехи нет. создаем ее
        add_user = env.users_name.get(name_user)
        if not add_user:
            env.db.out(f"Проект {name_folder} нет ответственного {name_user}",
                       runtime_error="y",
                       error_type="перенос проектов в стратегию")
            return False, False, False
        add_user = [add_user["id"]]
        now_date = date.today()
        dt = env.wr.dates_arr(type_="Milestone", due=now_date.isoformat())
        cfd = {"Ссылка на проект НП": permalink_project, "М1": 1}
        cfd["Последний этап"] = end_stage
        cf = env.wr.custom_field_arr(cfd)
        env.print_ss(f"Стратегия: создаем {name_task} для {name_user}",
                     env.cell_log)
        env.db.out(f"Стратегия: создаем {name_task} для {name_user}")
        resp = env.wr.create_task(id_folder, name_task, dates=dt,
                                  responsibles=add_user, customFields=cf,
                                  description=permalink_project)
        if len(resp) == 0:
            env.db.out(f"Проект {name_folder} не создал веху {name_task}",
                       runtime_error="y",
                       error_type="перенос проектов в стратегию")
            return False, False, False
    else:
        sub_tasks = resp[0]["subTaskIds"]
        # обновим имя если в настройках изменились этапы
        if name_task != resp[0]["title"]:
            cfd = {"Ссылка на проект НП": permalink_project, "М1": 1}
            cfd["Последний этап"] = end_stage
            cf = env.wr.custom_field_arr(cfd)
            resp = env.wr.update_task(resp[0]["id"], title=name_task,
                                      description=permalink_project,
                                      customFields=cf)
            if len(resp) == 0:
                env.db.out(f"{resp[0]['title']} ошибка записи",
                           runtime_error="y",
                           error_type="перенос проектов в стратегию")

    return resp[0], sub_tasks, id_folder


def delete_stage(task_param, stages_from_row, strat_task):
    num_stage = task_param["num_stage"]

    if num_stage not in stages_from_row:
        m = f"Стратегия: удаляем {strat_task['title']}"
        env.print_ss(m, env.cell_log)
        env.db.out(m)
        resp = env.wr.rs_del(f"tasks/{strat_task['id']}")
        if len(resp) == 0:
            env.db.out(f"{strat_task['title']} ошибка удаления",
                       runtime_error="y",
                       error_type="перенос проектов в стратегию")
            task_param["in_folder_err"] = True
        return True
    return False


def duplicate_stage(milestone, task_from_project, sub_tasks,
                    min_date, max_date, name_user, id_folder, stages_from_row):

    def create_name(task_param):
        name_task = task_param["Название рабочее"]
        name_task += f" -{task_param['title']}"
        return name_task

    # считываем и обновляем задачи
    last_date = date(1990, 1, 1)
    while len(sub_tasks) > 0:
        x = 0
        req = ""
        while x <= 100 and len(sub_tasks):
            x += 1
            req += sub_tasks.pop(0) + ","
        req = "tasks/" + req[:-1]
        resp = env.wr.get_tasks(req)
        # обновляем уже имеющиеся вехи
        for strat_task in resp:
            resp_cf = strat_task["customFields"]
            link_on_task = env.find_cf(resp_cf, "Ссылка на проект НП")
            if not link_on_task:
                continue
            task_param = task_from_project.get(link_on_task)
            if not task_param:
                continue
            if task_param["due"] > last_date:
                last_date = task_param["due"]

            # удаляем этап если в настройка поменяли
            if delete_stage(task_param, stages_from_row, strat_task):
                continue
            # обновляем этап если есть изменения
            status = strat_task["status"]
            strat_date = make_date(strat_task["dates"]["due"])
            strat_name = strat_task["title"]
            name_task = create_name(task_param)

            ok_due = (strat_date != task_param["due"])
            ok_status = (status != task_param["status"])
            ok_name = (strat_name != name_task)
            task_param["in_folder"] = True
            if ok_due or ok_status or ok_name:
                m = f"Стратегия: обновляем {strat_task['title']}"
                env.print_ss(m, env.cell_log)
                env.db.out(m)
                cfd = {"М2": 1}
                status = task_param["status"]
                if status == "Completed":
                    cfd["М2 Ф"] = 1
                else:
                    cfd["М2 Ф"] = 0
                cf = env.wr.custom_field_arr(cfd)
                dt = env.wr.dates_arr(type_="Milestone",
                                      due=task_param["due"].isoformat())
                resp = env.wr.update_task(strat_task["id"], title=name_task,
                                          description=link_on_task,
                                          dates=dt,
                                          status=status,
                                          customFields=cf)
                if len(resp) == 0:
                    env.db.out(f"{strat_task['title']} ошибка записи",
                               runtime_error="y",
                               error_type="перенос проектов в стратегию")
                    task_param["in_folder_err"] = True

    # создаем вехи которых нет
    for link_on_task, task_param in task_from_project.items():
        if task_param["in_folder"] or task_param["in_folder_err"]:
            continue
        # создаем веху
        # проверяем  по номерам этапов, не проверяем на даты
        num_stage = task_param["num_stage"]
        if num_stage not in stages_from_row:
            continue
        now_date = task_param["due"]
        if now_date > last_date:
            last_date = now_date
        st = [milestone["id"]]
        name_task = create_name(task_param)
        dt = env.wr.dates_arr(type_="Milestone", due=now_date.isoformat())
        cfd = {"Ссылка на проект НП": link_on_task, "М2": 1}
        status = task_param["status"]
        if status == "Completed":
            cfd["М2 Ф"] = 1
        else:
            cfd["М2 Ф"] = 0
        cf = env.wr.custom_field_arr(cfd)
        m = f"Стратегия: создаем {name_task} для {name_user}"
        env.print_ss(m, env.cell_log)
        env.db.out(m)
        resp = env.wr.create_task(id_folder, name_task, dates=dt,
                                  customFields=cf, status=status,
                                  superTasks=st, description=link_on_task)
        if len(resp) == 0:
            env.db.out(f"{name_task} ошибка записи", runtime_error="y",
                       error_type="перенос проектов в стратегию")
    # обновим главную веху по статусу всех последних этапов
    # даты у вехи в стратегии меняем только если создали новые вехи
    if last_date > max_date:
        last_date = max_date
    status_m = milestone["status"]
    date_m = make_date(milestone["dates"]["due"])
    one_status = None
    resp_cf = milestone["customFields"]
    end_stage = env.find_cf(resp_cf, "Последний этап")
    for link_on_task, task_param in task_from_project.items():
        now_date = task_param["due"]
        if task_param["num_stage"] == end_stage:
            now_status = task_param["status"]
            if one_status is None:
                one_status = now_status
                continue
            if now_status != one_status:
                one_status = False
                break

    if (one_status and one_status != status_m) or (last_date != date_m):
        if one_status == "Completed":
            cfd = {"М1 Ф": 1}
        else:
            cfd = {"М1 Ф": 0}
        cf = env.wr.custom_field_arr(cfd)
        dt = env.wr.dates_arr(type_="Milestone", due=last_date.isoformat())
        resp = env.wr.update_task(milestone["id"], status=one_status,
                                  customFields=cf, dates=dt)
        if len(resp) == 0:
            env.db.out(f"{milestone['title']} ошибка записи",
                       runtime_error="y",
                       error_type="перенос проектов в стратегию")


def read_all_tasks(name_project, permalink_project):
    resp = env.wr.get_folders("folders", permalink=permalink_project)
    if len(resp) == 0:
        env.db.out(f"Проект {name_project} не найден", runtime_error="y",
                   error_type="перенос проектов в стратегию")
        return False
    id_folder = resp[0]['id']
    fields = ["customFields"]
    resp = env.wr.get_tasks(f"folders/{id_folder}/tasks", type="Milestone",
                            subTasks="True", descendants="True",
                            fields=fields)
    # соберем в словарь все вехи проекта не зависимо от дат и статуса
    return_dict = {}
    for task in resp:
        resp_cf = task["customFields"]
        num_stage = env.find_cf(resp_cf, "num_stage")
        num_task = env.find_cf(resp_cf, "num_task")
        product = env.find_cf(resp_cf, "Название рабочее")
        if num_task[0:2] == "00" and len(num_stage) == 1:
            status = task["status"]
            permalink_task = task["permalink"]
            stage_date = make_date(task["dates"]["due"])
            return_dict[permalink_task] = {}
            return_dict[permalink_task]["due"] = stage_date
            return_dict[permalink_task]["status"] = status
            return_dict[permalink_task]["title"] = task["title"]
            return_dict[permalink_task]["num_stage"] = num_stage
            return_dict[permalink_task]["in_folder"] = False
            return_dict[permalink_task]["in_folder_err"] = False
            return_dict[permalink_task]["Название рабочее"] = product

    return return_dict


def reflect_milestone():
    ''' отображение вех проектов в папках со стратегическими задачами
    '''
    env.sheet_now("project_sheet")
    max_column = env.ss.max_column_get()
    table = env.ss.values_get(f"B24:{max_column}200")
    env.sheet_now("work_sheet")

    for row in table[6:]:
        if len(row) == 0:
            continue
        name_project = row[0]
        permalink_project = row[3]

        if not name_project or not permalink_project:
            continue
        if len(row[5:]) > 0:
            tasks_from_project = read_all_tasks(name_project,
                                                permalink_project)
            if not tasks_from_project:
                tasks_from_project = {}
        else:
            continue
        for num_column, column in enumerate(row[5:], 5):
            if not column:
                continue

            stages_from_row = column.split(";")

            name_user = table[2][num_column]
            name_folder = table[3][num_column]
            permalink_folder = table[4][num_column]
            min_date = make_date(table[0][num_column])
            max_date = make_date(table[1][num_column])

            # проверяем наличие вехи в стартегической папке и создаем если нет

            tpl = create_milestone(name_project, permalink_project,
                                   stages_from_row, name_folder,
                                   permalink_folder, name_user)
            milestone, sub_tasks, id_folder = tpl
            if not milestone:
                continue
            # находим у проекта вехи и дублируем их в milestone
            m = f"Стратегия: проверяем вехи в  {name_project}"
            env.print_ss(m, env.cell_log)
            env.db.out(m)
            duplicate_stage(milestone, tasks_from_project, sub_tasks,
                            min_date, max_date, name_user, id_folder,
                            stages_from_row)


def start_reflect(ext_env):
    global env
    env = ext_env
    reflect_milestone()
