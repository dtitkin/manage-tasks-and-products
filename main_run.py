
import time
import sys
from collections import OrderedDict

import Projectenv
from time_func import now_str, make_date, read_date_for_project
import strategy_transfer
import report


VERSION = '1.0'
env = None  # объект для параметров проекта


def chek_old_session(row_project, num_row):
    '''Проверяем не завершенные результаты предыдущей сессии
    new product:
    Del project:
    update sub task: '''

    if row_project["log"] == "new product:":
        pass
    elif row_project["log"] == "Del project:":
        pass
    elif row_project["log"] == "update sub task:":
        pass


def new_product(row_project, num_row, date_start, folders):
    ''' По признаку G в строке продукта создаем новый проект во Wrike
    '''
    # обозначим в гугл таблице начало работы
    env.print_ss("new product:", f"{env.column_log}{num_row}")
    # копируем шаблон в новый проект
    name_stage = row_project["code"] + " " + row_project["product"]
    cfd = {"Номер этапа": "",
           "Номер задачи": "",
           "Норматив часы": 0}
    for name, key in zip(env.custom_value_name, env.custom_value_key):
        cfd[name] = row_project[key]

    folder_id = env.parent_id
    my_folder = folders.get(row_project["project"])
    if my_folder:
        folder_id = my_folder

    if row_project["template_id"]:
        id_template = row_project["template_id"]
        date_start = None
        copyStatuses = "true"

    else:
        id_template = env.template_id
        copyStatuses = "false"
    resp = env.wr.copy_folder(id_template, folder_id, name_stage.upper(),
                              "", rescheduleDate=date_start,
                              rescheduleMode="Start",
                              copyStatuses=copyStatuses,
                              copyResponsibles="true")
    id_project = resp[0]["id"]
    permalink = resp[0]["permalink"]
    id_manager = env.users_name[row_project["rp"]]["id"]
    ownerIds = resp[0]["project"]["ownerIds"]
    pr = {"ownersAdd": [id_manager]}
    if ownerIds and ownerIds[0] != id_manager:
        pr["ownersRemove"] = ownerIds
    cf = env.wr.custom_field_arr(cfd)
    resp = env.wr.update_folder(id_project, customFields=cf, project=pr)
    # сохраним в таблице ID
    env.print_ss(id_project, f"{env.column_id}{num_row}", False)
    # обзначим в гугл таблице завершение работы
    env.print_ss("Finish new product:" + now_str(),
                 f"{env.column_log}{num_row}", False)
    env.print_ss(permalink, f"{env.column_link}{num_row}")
    row_project["permalink"] = permalink
    return (id_project, cfd)  # id созданного проекта , заполненные поля


def delete_product(id_project, num_row):
    ''' удаляем весь проект и стираем его ID

    '''
    env.print_ss("Del project:", f"{env.column_log}{num_row}")
    resp = env.wr.rs_del(f"folders/{id_project}")
    if resp:
        env.print_ss("Finish del project:" + now_str(),
                     f"{env.column_log}{num_row}", False)
        env.print_ss("", f"{env.column_id}{num_row}", False)
        env.print_ss("", f"{env.column_comand}{num_row}", False)
        env.print_ss("", f"{env.column_link}{num_row}")
        return True
    else:
        return False


def test_all_parametr(row_project, num_row):
    '''Проверяем все параметры в строке
        - менеджер - ок
        - технолог - ок
        - название продукта -ок
        - наличие типовой длительности для шаблона задач
        - наличие дат во всех этапах -ок
        - ????
    '''
    ok = True
    if len(row_project["id_project"]) != 0:
        ok = False
        env.db.out("Проект уже выгружен во Wrike", num_row=num_row,
                   runtime_error="y", error_type="исходные данные")
    if len(row_project["product"]) == 0:
        ok = False
        env.db.out("нет названия продукта", num_row=num_row,
                   runtime_error="y", error_type="исходные данные")
    # руководитель проекта
    id_user = env.users_name.get(row_project["rp"])
    if not id_user or not id_user.get("id"):
        ok = False
        env.db.out(f"руководителя проекта {row_project['rp']} нет во Wrike",
                   num_row=num_row,
                   runtime_error="y", error_type="исходные данные")
    # технолог
    id_user = env.users_name.get(row_project["tech"])
    if not id_user or not id_user.get("id"):
        ok = False
        env.db.out(f"технолог {row_project['tech']} нет во Wrike",
                   num_row=num_row,
                   runtime_error="y", error_type="исходные данные")
    in_all = True
    if len(row_project["dates"]) < 8:
        in_all = False
    if not in_all:
        ok = False
        env.db.out("не установленны даты этапов", num_row=num_row,
                   runtime_error="y", error_type="исходные данные")

    return ok


def write_date_to_google(num_row, project_id):
    ''' устанавливает даты из wrike Гугл таблицу
        Отмечает этапы выполеннными
    '''
    fields = ["customFields"]
    resp = env.wr.get_tasks(f"folders/{project_id}/tasks", type="Milestone",
                            fields=fields)
    lst_finish = ["" for x in range(0, 8)]
    lst_values = [["" for x in range(0, 8)]]

    for task in resp:
        resp_cf = task["customFields"]
        num_stage = env.find_cf(resp_cf, "num_stage")
        num_task = env.find_cf(resp_cf, "num_task")
        if num_task[0:2] == "00" and len(num_stage) == 1:
            due_date = task["dates"]["due"]
            y, m, d = due_date[0:10].split("-")
            value = f"{d}.{m}.{y}"
            lst_values[0][int(num_stage) - 1] = value
            if task["status"] == "Completed":
                lst_finish[int(num_stage) - 1] = "Finish"
    env.sheet_now("work_sheet")
    cells_range = (f"{env.columns_stage[0]}{num_row}:"
                   f"{env.columns_stage[-1]}{num_row}")
    env.ss.prepare_setvalues(cells_range, lst_values)
    for n, f in enumerate(lst_finish, 0):
        column = env.columns_stage[n]
        txt_field = "userEnteredFormat.textFormat"
        bcg_field = "userEnteredFormat.backgroundColor"
        cells = f"{column}{num_row}:{column}{num_row}"
        if f != "Finish":
            color_bc = env.color_bc_finish
            color_font = env.color_font_finish
        else:
            color_bc = env.color_bc
            color_font = env.color_font
        bg_color = {"backgroundColor": color_bc}
        txt_color = {"textFormat": {"foregroundColor": color_font}}
        env.ss.prepare_setcells_format(cells, bg_color, fields=bcg_field)
        env.ss.prepare_setcells_format(cells, txt_color, fields=txt_field)
    env.ss.run_prepared()
    return True


def sorted_tasks(task, all_task):
    '''ищет рекусрсивно у вех подзадачи
      расставлет  в порядке обработки вложенные вехи раньше
      родительской вехи
    '''
    sort_task = OrderedDict({})
    sub_task = task["subTaskIds"]
    for st in sub_task:
        st_param = all_task.get(st)
        if st_param:
            if st_param["dates"]["type"] == "Milestone":
                sort_task.update(sorted_tasks(st_param, all_task))
                sort_task[st_param["id"]] = st_param
    sort_task[task["id"]] = task
    return sort_task


def close_move_stage(num_row, project_id):
    ''' устанавливает статус выполненно или переносит веху на дату выполнения
    '''
    fields = ["customFields", "subTaskIds"]
    resp = env.wr.get_tasks(f"folders/{project_id}/tasks", subTasks="True",
                            fields=fields)
    sort_task = OrderedDict({})
    all_task = {task["id"]: task for task in resp}
    for task in resp:
        if task["dates"]["type"] != "Milestone":
            continue
        if sort_task.get(task["id"]):
            continue
        sort_task.update(sorted_tasks(task, all_task))

    for k, task in sort_task.items():
        # по каждой вехе смотрим выполненны ли все подчиненные
        # ищем их в all_task, если веху обновляем то менем ее в all task
        all_status = []
        task_date = make_date(task["dates"]["due"])
        max_date = task_date
        sub_task = task["subTaskIds"]
        for st in sub_task:
            st_param = all_task.get(st)
            if not st_param:
                continue
            status = st_param["status"]
            if status not in all_status:
                all_status.append(status)
            due = st_param["dates"].get("due")
            if due:
                due = make_date(due)
                if due > max_date:
                    max_date = due

        set_status = None
        if len(all_status) == 1:
            if all_status[0] != task["status"]:
                set_status = all_status[0]
        elif len(all_status) == 2:
            if "Completed" in all_status and "Cancelled" in all_status:
                set_status = "Completed"
        else:
            if "Active" in all_status:
                set_status = "Active"

        if set_status == task["status"]:
            set_status = None
        dt = None
        now_date = make_date()
        if task["status"] == "Active" or set_status == "Active":
            if now_date > task_date:
                dt = env.wr.dates_arr(type_="Milestone",
                                      due=now_date.isoformat())

        if set_status == "Completed":
            dt = env.wr.dates_arr(type_="Milestone",
                                  due=max_date.isoformat())
        if set_status or dt:
            resp = env.wr.update_task(task["id"], dates=dt, status=set_status)
            if len(resp) == 0:
                env.db.out(f"не обновляется {task['title']}",
                           num_row=num_row,
                           runtime_error="y", error_type="ошибка записи Wrike")
            task_in_all = all_task.get(task["id"])
            if task_in_all:
                task_in_all["status"] = resp[0]["status"]


def new_tech(id_project, resp_cf, technologist):
    ''' проверка на замену технолога
    '''

    wrike_technologist = env.find_cf(resp_cf, "Технолог")
    wrike_id = env.users_name.get(wrike_technologist)
    if wrike_id:
        wrike_id = wrike_id["id"]
    technologist_id = env.users_name.get(technologist)
    if technologist_id:
        technologist_id = technologist_id["id"]
        if technologist_id != wrike_id:
            return technologist_id
    return None


def change_tech_in_task(id_project, new_tech_id, num_row, cf):
    ''' Меняем технолога в задачах и пользовательских полях
    '''
    env.print_ss("update sub task:", f"{env.column_log}{num_row}")
    # получим список технологов кроме нового технолога
    tech_group = []
    for key, item in env.users_id.items():
        if item["group"] == "Технолог":
            if key == new_tech_id:
                continue
            tech_group.append(key)
    tech_group = set(tech_group)
    #  читаем из `Wrike задачи
    fields = ["responsibleIds"]
    resp = env.wr.get_tasks(f"folders/{id_project}/tasks", subTasks="true",
                            fields=fields)
    # перебираем задачи и обновляем
    n = 0
    pred_proc = 0
    len_sub = len(resp)
    for task in resp:
        n += 1
        percent = n / len_sub
        env.progress(percent)
        if pred_proc + 1 == int(percent * 10):
            pred_proc += 1
            env.print_ss(f"{num_row}:{pred_proc*10}%",
                         env.cell_progress, True, 0)
        # определям список пользователей которых нужно исключить из задачи
        ownersid = set(task["responsibleIds"])
        remove_r_bles = list(tech_group)
        add_r_bles = [new_tech_id]

        if new_tech_id in ownersid or not new_tech_id:
            # новый технолог уже установлен
            remove_r_bles = None
            add_r_bles = None
        if len(tech_group.intersection(ownersid)) == 0:
            # в задаче нет технологов менять некого
            remove_r_bles = None
            add_r_bles = None
        #  обновляем задачу
        if remove_r_bles or add_r_bles or cf:
            resp_upd = env.wr.update_task(task["id"],
                                          removeResponsibles=remove_r_bles,
                                          addResponsibles=add_r_bles,
                                          customFields=cf)
            if len(resp_upd) == 0:
                env.db.out(task["title"] + " ошибка обновления",
                           num_row=num_row,
                           runtime_error="y", error_type="обновление задачи")
                break
    else:
        #  если обработали все задачи обозначим в таблице выполнение
        env.print_ss("Finish update sub task",
                     f"{env.column_log}{num_row}", False)
        return True
    return False


def find_change_cfd(row_project, resp_cf):
    cfd = {}
    for name, key in zip(env.custom_value_name, env.custom_value_key):
        value_in_wrike = env.find_cf(resp_cf, name)
        if value_in_wrike != row_project[key] or env.update_all_cv:
            cfd[name] = row_project[key]
    return cfd


def if_edit_table(num_row, row_project, folders):
    ''' проверяем поля у папки, меняем размещение в папке проект
    '''

    if not row_project["id_project"]:
        env.db.out(f"нет ID проекта", num_row=num_row,
                   runtime_error="y", error_type="ошибка данных Google")
        return False
    my_folder = folders.get(row_project["project"])
    resp = env.wr.get_folders(f"folders/{row_project['id_project']}")

    if len(resp) > 0:
        new_tech_id = new_tech(row_project["id_project"],
                               resp[0]["customFields"],
                               row_project["tech"])
        my_parents = resp[0]["parentIds"]
        permalink = resp[0]["permalink"]
        if permalink != row_project["permalink"]:
            env.print_ss(permalink, f"{env.column_link}{num_row}")
        addParents = None
        removeParents = None
        cf = None

        if my_folder and my_folder not in my_parents:
            addParents = [my_folder]
            removeParents = my_parents

        cfd = find_change_cfd(row_project, resp[0]["customFields"])
        if len(cfd) > 0:
            cf = env.wr.custom_field_arr(cfd)
        title = None
        name_stage = row_project["code"] + " " + row_project["product"]
        if resp[0]["title"] != name_stage.upper():
            title = name_stage.upper()

        if addParents or removeParents or cf or title:
            resp = env.wr.update_folder(row_project["id_project"], title=title,
                                        addParents=addParents,
                                        removeParents=removeParents,
                                        customFields=cf)
            if len(resp) == 0:
                env.db.out(f"не обновляется {row_project['product']}",
                           num_row=num_row,
                           runtime_error="y", error_type="ошибка записи Wrike")
                return False
            #  доделать обновление полей в задачах, в том числе поле проект
            if new_tech_id or cf:
                ok = change_tech_in_task(row_project["id_project"],
                                         new_tech_id, num_row, cf)
                return ok
            else:
                return True
        else:
            return True
    else:
        env.db.out(f"не читается папка {row_project['project']}",
                   num_row=num_row,
                   runtime_error="y", error_type="ошибка чтения Wrike")
        return False


def read_stage_info(num_row):
    ''' запоминаем из таблицы значение, цвет шрифта и цвет ячейки
    '''
    finish_list = []
    dates_stage = {}
    env.sheet_now("work_sheet")
    cells_range = (f"{env.columns_stage[0]}{num_row}:"
                   f"{env.columns_stage[-1]}{num_row}")
    lst_status = env.ss.read_color_cells(cells_range)[0]
    for k, cl in enumerate(env.columns_stage, 1):
        color_stage = (lst_status[k - 1][0], lst_status[k - 1][1])
        stage_value = lst_status[k - 1][2]
        if str(color_stage) == str(env.COLOR_FINISH):
            # запоминаем у каких этапов нужно установить Finish
            finish_list.append(str(k))
        dates_stage[str(k)] = make_date(stage_value)
    return finish_list, dates_stage


def update_sub_task(id_and_cfd, rp, tech, num_row,
                    finish_status, dates_stage, personal_template):
    ''' Обновление задач в проекте. Установка исполнителей,
        пользовательских
        полей, статуса выполненно
    '''
    # обозначим в таблице что начали этап
    env.print_ss("update sub task:", f"{env.column_log}{num_row}")

    #  читаем из `Wrike задачи
    fields = ["responsibleIds", "customFields"]
    resp = env.wr.get_tasks(f"folders/{id_and_cfd[0]}/tasks",
                            subTasks="true", fields=fields)
    # перебираем задачи и обновляем
    n = 0
    pred_proc = 0
    len_sub = len(resp)
    for task in resp:
        n += 1
        percent = n / len_sub
        env.progress(percent)
        if pred_proc + 1 == int(percent * 10):
            pred_proc += 1
            env.print_ss(f"{num_row}:{pred_proc*10}%",
                         env.cell_progress, True, 0)
        # пользовательские поля
        resp_cf = task["customFields"]
        cfd = id_and_cfd[1]
        cfd["Номер этапа"] = env.find_cf(resp_cf, "Номер этапа")
        cfd["Номер задачи"] = env.find_cf(resp_cf, "Номер задачи")
        cfd["Норматив часы"] = env.find_cf(resp_cf, "Норматив часы")
        cf = env.wr.custom_field_arr(cfd)
        # определям список пользователей которых нужно исключить из задачи
        ownersid = task["responsibleIds"]
        num_stage = env.find_cf(resp_cf, "num_stage")
        num_task = env.find_cf(resp_cf, "num_task")
        remove_users = env.find_task_user(ownersid, rp, tech,
                                          num_task)
        if personal_template:
            add_users = env.find_task_user(ownersid, rp, tech, num_task,
                                           False)
        else:
            add_users = None

        type_task = task["dates"]["type"]
        dt = None
        status = None
        if not personal_template:
            #  проверяем на статус выполненно
            if num_stage in finish_status:
                status = "Completed"
            else:
                status = "Active"
            # у 'этапов' устанавливаем дату из таблицы
            if type_task == "Milestone" and num_task[0:2] == "00":
                if dates_stage.get(num_stage):
                    due = dates_stage[num_stage].isoformat()
                    dt = env.wr.dates_arr(type_="Milestone", due=due)
        #  обновляем задачу
        resp_upd = env.wr.update_task(task["id"],
                                      removeResponsibles=remove_users,
                                      addResponsibles=add_users,
                                      customFields=cf,
                                      status=status, dates=dt)
        if len(resp_upd) == 0:
            env.db.out(task["id"] + " ошибка обновления",
                       num_row=num_row,
                       runtime_error="y", error_type="обновление задачи")
            break
    else:
        env.print_ss("Finish update sub task",
                     f"{env.column_log}{num_row}", False)
        # Устанавливаем в таблицу W  вместо G
        env.print_ss("W", f"{env.column_comand}{num_row}")
        return True
    return False

def sync_google_wrike(folders):

    env.db.out("Синхронизация Гугл и Wrike", prn_time=True)
    table = env.get_work_table()
    for num_row, row_project in table.items():
        chek_old_session(row_project, num_row)
        if row_project["comand"] == "G" and env.compare_param("G"):
            ok = test_all_parametr(row_project, num_row)
            if not ok:
                continue
            m = (f"Создаем продукт #{row_project['code']} "
                 f"{row_project['product']}")
            env.print_ss(m, env.cell_log, True, 0)
            env.db.out(m, num_row=num_row, prn_time=True)
            #  определяем дату для старта проекта
            len_stage = env.get_len_stage("1")
            date_start = read_date_for_project(row_project["dates"][0],
                                               len_stage, env.HOLYDAY)
            # проверим персональный шаблон
            if row_project["template_id"]:
                m = f"Шаблон из строки {row_project['template_str']}"
                env.db.out(m, num_row=num_row)
            #  создаем новый проект
            id_and_cfd = new_product(row_project, num_row, date_start,
                                     folders)
            # установим исполнителей, пользовательские поля,
            # признак выполненно, перенесем вехи на нужные даты
            env.db.out("Обновление задач в проекте", num_row=num_row)

            # считываем статусы и даты этапов из таблицы
            finish_status = None
            dates_stage = None
            if not row_project["template_id"]:
                finish_status, dates_stage = read_stage_info(num_row)
            # обновляем задачи с учетом всех статусов
            ok = update_sub_task(id_and_cfd, row_project["rp"],
                                 row_project["tech"], num_row,
                                 finish_status, dates_stage,
                                 row_project["template_id"])
            if not ok:
                env.db.out("!Выполнение прервано!", num_row=num_row,
                           runtime_error="y",
                           error_type="ошибка обновления задач")
                return False
            ok = write_date_to_google(num_row, id_and_cfd[0])

        if row_project["comand"] == "W" and env.compare_param("M"):
            # изменение статуса и даты у вех
            m = (f"Проверяем вехи #{num_row} {row_project['code']}"
                 f" {row_project['product']}")
            env.print_ss(m, env.cell_log)
            env.db.out(m, num_row=num_row, prn_time=True)
            close_move_stage(num_row, row_project["id_project"])

        if row_project["comand"] == "W" and env.compare_param("W"):
            # обновление дат в гугл и признака выполенно
            m = (f"Обновляем даты #{num_row} {row_project['code']}"
                 f" {row_project['product']}")
            env.print_ss(m, env.cell_log)
            env.db.out(m, num_row=num_row, prn_time=True)
            ok = write_date_to_google(num_row, row_project["id_project"])

        if row_project["comand"] == "W" and env.compare_param("F"):
            # проверка на изменение полей в таблице
            m = (f"Проверяем колонки #{num_row} {row_project['code']}"
                 f" {row_project['product']}")
            env.db.out(m, num_row=num_row, prn_time=True)
            ok = if_edit_table(num_row, row_project, folders)

        if row_project["comand"] == "A" and env.compare_param("A"):
            # удаляем проект из Wrike если он там есть
            if row_project["id_project"]:
                m = (f"Удаляем проект #{num_row} {row_project['code']}"
                     f" {row_project['product']}")
                env.print_ss(m, env.cell_log)
                env.db.out(m, num_row=num_row, prn_time=True)
                ok = delete_product(row_project["id_project"], num_row)
                if not ok:
                    m = f"Ошибка удаления проекта {row_project['id_project']}"
                    env.db.out(m, num_row=num_row, runtime_error="y",
                               error_type="ошибка удаления")


def create_folder_in_parent():
    ''' создает папки для  проектов в основной папке
        возвращает словарь  ключ - имя проекта значание ID
    '''
    env.db.out("Создание папок для проектов")
    filter_project = env.get_project_on_row()
    table = env.get_project_list()
    num_str = 29
    return_dict = {}
    for project in table:
        num_str += 1
        if len(project) == 0:
            continue
        name_project = project[0]
        if env.row_filter and filter_project != name_project:
            continue
        if len(name_project) > 0:
            name_project = name_project
        wrike_link = ""
        if len(project) == 4:
            wrike_link = project[3]
        if wrike_link:
            return_dict[name_project] = env.update_project_name(wrike_link,
                                                                name_project,
                                                                num_str)
        if not wrike_link or not return_dict.get(name_project):
            return_dict[name_project] = env.make_folder(name_project, num_str)
    return return_dict


def main():
    '''Создает шаблон во Wrike на основе Гугл таблицы
    '''
    t_start = time.time()
    global env

    env = Projectenv.Projectenv(sys.argv)
    if not env:
        return

    if not env.connect_db():
        return

    if not env.user_in_base():
        return

    if not env.connect_spreadsheet():
        return

    if not env.connect_wrike():
        return

    if not env.read_global_env():
        return

    if env.db.log == "terminal":
        print(f"Запуск скрипта версии {VERSION}")

    env.sheet_now("work_sheet")

    env.print_ss(f"Обновление началось {now_str()}", env.cell_log)

    if env.compare_param("GAWFM"):
        folders = create_folder_in_parent()
        sync_google_wrike(folders)

    if env.compare_param("R"):
        report.create_report_table(Projectenv, sys.argv)
    if env.compare_param("S"):
        strategy_transfer.start_reflect(env)

    env.sheet_now("work_sheet")
    t_finish = time.time()
    m = f"Выполненно за: {int(t_finish - t_start)} с. {now_str()}"
    env.db.out(m)
    env.print_ss(m, env.cell_log)


if __name__ == '__main__':
    main()
