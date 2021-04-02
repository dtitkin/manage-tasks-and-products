''' Модуль для обновления таблиц для очтетов и
    формирования снимков базы
'''

env = None  # объект для параметров проекта


def report_from_stages_dates():
    ''' Обновляет в Гугл таблицу _ОтчетПоЭтапам
    '''
    fields = ["customFields", "parentIds"]
    resp = env.wr.get_tasks(f"folders/{env.parent_id}/tasks",
                            type="Milestone",
                            subTasks="True", descendants="True",
                            fields=fields)
    lst_out = [["Проект", "Продукт", "РП", "Технолог", "Страт. Группа",
                "Группа", "Линейка", "Клиент", "Бренд", "Этап", "Номер",
                "Статус", "Дата завершения"]]
    for task in resp:
        resp_cf = task["customFields"]
        num_stage = env.find_cf(resp_cf, "num_stage")
        num_task = env.find_cf(resp_cf, "num_task")
        if num_task[0:2] == "00" and len(num_stage) == 1:
            project = env.find_cf(resp_cf, "Проект")
            product = env.find_cf(resp_cf, "Название рабочее")
            rp = env.find_cf(resp_cf, "Руководитель проекта")
            tech = env.find_cf(resp_cf, "Технолог")
            strat = env.find_cf(resp_cf, "Стратегическая группа")
            group = env.find_cf(resp_cf, "Группа")
            line = env.find_cf(resp_cf, "Линейка")
            supl = env.find_cf(resp_cf, "Клиент")
            brand = env.find_cf(resp_cf, "Бренд")
            project = env.find_cf(resp_cf, "Проект")
            stage = task["title"]
            status = task["status"]

            if not project or not product or not tech or not rp:
                parentIds = task["parentIds"]
                req = ""
                x = 0
                while x <= 100 and len(parentIds):
                    x += 1
                    req += parentIds.pop(0) + ","
                    req = "folders/" + req[:-1]
                    resp_f = env.wr.get_folders(req)
                    for folder in resp_f:
                        if folder.get("project"):
                            resp_cf = folder["customFields"]
                            project = env.find_cf(resp_cf, "Проект")
                            product = env.find_cf(resp_cf, "Название рабочее")
                            rp = env.find_cf(resp_cf, "Руководитель проекта")
                            tech = env.find_cf(resp_cf, "Технолог")
                            strat = env.find_cf(resp_cf,
                                                "Стратегическая группа")
                            group = env.find_cf(resp_cf, "Группа")
                            line = env.find_cf(resp_cf, "Линейка")
                            supl = env.find_cf(resp_cf, "Клиент")
                            brand = env.find_cf(resp_cf, "Бренд")

            due = task["dates"]["due"]
            if due:
                due = due[0:10]
                lst_line = [project, product, rp, tech, strat, group, line,
                            supl, brand, stage, num_stage, status, due]
                lst_line = [x.strip("\n ") for x in lst_line]
                lst_out.append(lst_line)
    # разместим в гугл таблице
    # отчистим все что было ранее
    env.ss.sheet_clear_all(env.ss.sheetTitle)
    env.ss.prepare_setvalues(f"A1:M{len(lst_out)}", lst_out)
    env.ss.run_prepared()


def create_report_table(Projectenv, argv):
    global env

    env = Projectenv.Projectenv(argv)
    if not env:
        return

    if not env.connect_db():
        return

    if not env.user_in_base():
        return

    if not env.connect_spreadsheet(env.REPORT_ID):
        return

    if not env.connect_wrike():
        return

    parent = env.wr.get_folders("folders", permalink=env.folder_link)[0]
    env.parent_id = parent["id"]

    env.sheet_now("report_stages")
    env.db.out("Обновление отчетов")
    report_from_stages_dates()
    env.db.out("Обновление отчетов завершено")
