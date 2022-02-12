# модуль класса доступа к API Wrike
# в модуле реализованна минимальная часть работы с API
# достаточная для  "Управление продуктами и задачами"
# полная документация на API
# https://developers.wrike.com/overview/


from pprint import pprint
from datetime import date

import requests as rs


class OopsException(Exception):
    pass


class Wrike():
    def __init__(self, wrike_token, debugMode=False):
        self.debugMode = debugMode
        self.headers = {"Authorization": "bearer " + wrike_token}
        self.params = {}
        self.data = {}
        self.connect = "https://www.wrike.com/api/v4/"
        self.customfields = self.custom_field_dict()
        if debugMode:
            resp = rs.get(self.connect + "account",
                          self.params, headers=self.headers)
            pprint(resp.json())

    def test(self, resp):
        ''' выводит значение ответа от Api при debugMode=True
        '''
        if self.debugMode:
            print(resp.status_code)
            try:
                js = resp.json()
                print(type(js))
                for key, value in js.items():
                    print(f"Ключ:{key} Тип: ", end="")
                    print(f" {type(value)} Размер:{len(value)}")
                    if key == "data":
                        print("----- 0 значение-----")
                        pprint(value[0])
                        if (len(value) - 1) > 0:
                            print(f"----- {len(value)-1} значение-----")
                            pprint(value[-1])
                    elif key == "errorDescription" or key == "error":
                        pprint(value)
                    else:
                        pprint(value)
            except OopsException as e:
                print("Тест сломался", e)

    def make_params(self, local_dict, exclude_list):
        ''' создает строку параметров для передачи в запрос
        '''
        params = {}
        for arg_name, arg_value in local_dict.items():
            if arg_name not in exclude_list and arg_value:
                if isinstance(arg_value, dict) or isinstance(arg_value, list):
                    params[arg_name] = str(arg_value)
                else:
                    params[arg_name] = arg_value
        if self.debugMode:
            pprint(params)
        return params

    def manage_return(self, resp, arr="data"):
        ''' определяет что вернуть из response на основании arr

            arr="data" or "json" or key in resp Wrike
            При arr == "data" всегда возвращает список
            При arr == "resp" возворащает объект response

        '''
        if arr == "data":
            ls = resp.json().get("data")
            if ls is None:
                ls = list()
            return ls
        elif arr == "resp":
            return resp
        elif arr == "json":
            return resp.json()
        else:
            return resp.json()[arr]

    def rs_get(self, get_str, arr="data", **kwargs):
        '''выполнение любого запроса get

        возвращает status_code, resp.json()[arr]
        '''
        self.params = kwargs.copy()
        try:
            resp = rs.get(self.connect + get_str, self.params,
                          headers=self.headers)
        finally:
            self.params = {}
        self.test(resp)
        return self.manage_return(resp, arr)

    def rs_post(self, post_str, **kwargs):
        'выполнение любого запроса post'
        self.data = kwargs.copy()
        try:
            resp = rs.post(self.connect + post_str, self.data,
                           headers=self.headers)
        finally:
            self.data = {}
        self.test(resp)
        return self.manage_return(resp)

    def rs_put(self, put_str, **kwargs):
        'выполнение любого запроса put'
        self.data = kwargs.copy()
        try:
            resp = rs.put(self.connect + put_str, self.data,
                          headers=self.headers)
        finally:
            self.data = {}
        self.test(resp)
        return self.manage_return(resp)

    def rs_del(self, del_str, **kwargs):
        'выполнение любого запроса del'
        resp = rs.delete(self.connect + del_str, headers=self.headers)
        self.test(resp)
        return self.manage_return(resp)

    def get_tasks(self, task_area, descendants=None, title=None, status=None,
                  importance=None, startDate=None, dueDate=None,
                  scheduledDate=None, completedDate=None, responsibles=None,
                  permalink=None, type=None, limit=0,
                  sortField=None, sortOrder=None, subTasks=None,
                  customField=None, fields=None):
        '''Получить задачи с параметрами поиска

        https://developers.wrike.com/api/v4/tasks/
            task_area - tasks
                        folders/{folderId}/tasks
                        spaces/{spaceId}/tasks
                        tasks/{taskId},{taskId},... - up to 100 IDs
        '''
        task_params = self.make_params(locals(),
                                       ["self", "task_area", "task_params"])
        if task_area[0:6] != "tasks/":
            task_params["pageSize"] = 1000
        resp = self.rs_get(task_area, arr="json", **task_params)
        if task_area[0:6] != "tasks/":
            del task_params['pageSize']
        data = []
        responseSize = resp.get("responseSize")
        nextPageToken = resp.get("nextPageToken")
        data = resp.get("data")
        if responseSize:
            while nextPageToken:
                task_params["nextPageToken"] = nextPageToken
                resp = self.rs_get(task_area, arr="json", **task_params)
                data.extend(resp["data"])
                nextPageToken = resp.get("nextPageToken")
            return data
        else:
            if data:
                return data
            else:
                return []

    def create_task(self, folderid, title, description=None, status="Active",
                    importance="Normal", dates=None, shareds=None,
                    parents=None, responsibles=None, priorityAfter=None,
                    superTasks=None, customFields=None, fields=None,
                    follow="False", followers=None):
        '''Создать задачу

        https://developers.wrike.com/api/v4/tasks/
        '''
        task_dates = self.make_params(locals(),
                                      ["self", "folderid", "task_dates"])
        resp = self.rs_post(f"folders/{folderid}/tasks", **task_dates)
        return resp

    def update_task(self, taskid, title=None, description=None, status=None,
                    importance=None, dates=None, addParents=None,
                    removeParents=None, addResponsibles=None,
                    removeResponsibles=None, priorityAfter=None,
                    addSuperTasks=None, removeSuperTasks=None,
                    customFields=None, customStatus=None):
        '''Обновить задачу
        '''
        task_dates = self.make_params(locals(),
                                      ["self", "taskid", "task_dates"])
        resp = self.rs_put(f"tasks/{taskid}", **task_dates)
        return resp

    def create_folder(self, folderid, title, description=None,
                      shareds=None, project=None, fields=None):
        '''Создать папку

        '''
        folder_dates = self.make_params(locals(),
                                        ["self", "folderid", "folder_dates"])
        resp = self.rs_post(f"folders/{folderid}/folders", **folder_dates)
        return resp

    def get_folders(self, folderarea, permalink=None, descendants=None,
                    customField=None, updatedDate=None, project=None,
                    deleted=None, fields=None):
        '''Получить папку/проект
           folderarea - folders
                        folders/{folderId}/folders
                        spaces/{spaceId}/folders
        '''
        task_params = self.make_params(locals(),
                                       ["self", "folderarea", "task_params"])
        resp = self.rs_get(folderarea, **task_params)
        return resp

    def copy_folder(self, folderid, parent, title, titlePrefix,
                    copyDescriptions=None, copyResponsibles=None,
                    rescheduleDate=None, rescheduleMode=None,
                    copyStatuses=None):
        ''' копироуем папку/проект
        '''

        task_dates = self.make_params(locals(),
                                      ["self", "folderid", "task_dates"])
        resp = self.rs_post(f"copy_folder/{folderid}", **task_dates)
        return resp

    def update_folder(self, folderid, title=None, description=None,
                      addParents=None, removeParents=None, addShareds=None,
                      removeShareds=None, restore=None, customFields=None,
                      customColumns=None, project=None, fields=None):
        ''' обновить папку
        '''
        task_dates = self.make_params(locals(),
                                      ["self", "folderid", "task_dates"])
        resp = self.rs_put(f"folders/{folderid}", **task_dates)
        return resp

    def update_milestone_date(self, folderid, root_task_id):
        '''Обновляет дату у всех вех внутри задачи с переносом не текущую
        '''
        fld = str(["subTaskIds"])
        resp = self.get_tasks(f"folders/{folderid}/tasks",
                              subTasks="True", fields=fld)
        # перебираем все вехи кроме root, находим максимальную дату
        # завершения у подзадач и переноим веху на эту дату
        root_date = date.today()
        root_need_change = False
        for task in resp:
            if task["id"] == root_task_id:
                continue
            type_task = task.get("dates")["type"]
            if type_task == "Milestone" and task["status"] == "Active":
                sub_task_ids = task.get("subTaskIds")
                max_date = date.today()
                for sub_task_id in sub_task_ids:
                    for f_task in resp:
                        arg_1 = f_task["id"] == sub_task_id
                        arg_2 = f_task["status"] == "Active"
                        if arg_1 and arg_2:
                            my_date = self.get_date_from_task(f_task,
                                                              max_date)
                            if max_date < my_date:
                                max_date = my_date
                            break
                # заменим дату у вехи если она меньше чем max_date
                my_date = self.get_date_from_task(task, max_date)
                if max_date > my_date:
                    root_need_change = True
                    root_date = max_date
                    dt = self.dates_arr(type_="Milestone",
                                        due=max_date.isoformat())
                    self.update_task(task["id"], dates=dt)
        if root_need_change:
            dt = self.dates_arr(type_="Milestone", due=root_date.isoformat())
            self.update_task(root_task_id, dates=dt)

    def create_dependency(self, taskid, predecessorId=None, successorId=None,
                          relationType=None):
        '''Установить связи у задачи - предыдущие задачи
        '''
        task_dates = self.make_params(locals(),
                                      ["self", "taskid", "task_dates"])
        resp = self.rs_post(f"tasks/{taskid}/dependencies", **task_dates)
        return resp

    def custom_field_dict(self):
        '''Словарь пользовательских полей для управления продуктами.
        по названию поля находит его id
        словарь хранится как атрибут объекта
        '''
        resp = self.rs_get("customfields")
        list_template = ["Норматив часы", "Номер этапа", "Номер задачи",
                         "Стратегическая группа", "Руководитель проекта",
                         "Клиент", "Бренд", "Код-1С", "Название рабочее",
                         "Группа", "Линейка", "Проект", "Технолог",
                         "num_stage", "num_task", "Технология",
                         "Ссылка на проект НП", "М1", "М1 Ф", "М2",
                         "Последний этап", "М2 Ф"]

        return_dict = {}
        for field in resp:
            if field["title"] in list_template:
                return_dict[field["title"]] = field["id"]
        if self.debugMode:
            print(len(return_dict) == len(list_template))
            pprint(return_dict)
        return return_dict

    def custom_field_arr(self, name_value_dict):
        '''Возращает массив для custom field
           не найденные элементы не использует

        '''
        return_list = []
        for name, value in name_value_dict.items():
            d = self.customfields.get(name)
            if d:
                return_list.append({"id": d, "value": value})
        return return_list

    def dates_arr(self, type_="Planned", duration=480, start="", due="",
                  workOnWeekends="False"):
        ''' возвращает объект - словарь для поля dates у задач
        '''
        if type_ == "Milestone":
            duration = 0
        r_d = {"type": type_, "workOnWeekends": workOnWeekends}
        if start:
            r_d["start"] = start
        if due:
            r_d["due"] = due
        if duration:
            r_d["duration"] = duration
        return r_d

    def get_date_from_task(selff, task, max_date, s_d="due"):
        ''' из полей due или start у задачи возвращает дату
        '''
        try:
            n_date = task.get("dates")[s_d]
            y, m, d = map(int, n_date[0:10].split('-'))
            my_date = date(y, m, d)
            return my_date
        except TypeError:
            return max_date
        except ValueError:
            return max_date

    def deskr_from_str(self, deskr_str):
        ''' если в описание задачи нужно поместить список
        '''
        descr_lst = deskr_str.split(";")
        return_descr = ""
        for one_str in descr_lst:
            return_descr += "<li>" + one_str + "</li>"
        return return_descr

    def id_contacts(self):
        '''возвращает id всех пользователей Wrike
        '''
        resp = self.rs_get("contacts")
        id_dict = {}
        for user in resp:
            if user['type'] == 'Person' and not(user['deleted']):
                user_mail = user["profiles"][0]["email"]
                id_dict[user_mail] = {}
                id_dict[user_mail]['id'] = user["id"]
                id_dict[user_mail]['name'] = (user["firstName"] + " "
                                              + user["lastName"])
        if self.debugMode:
            pprint(id_dict)
        return id_dict

    def id_folders_on_name(self, foldersname_list):
        ' по имени папки или проекта возвращает id проекта'
        resp = self.rs_get("folders")
        if isinstance(foldersname_list, str):
            ls = []
            ls.append(foldersname_list)
            foldersname_list = ls
        folders_dict = {name: None for name in foldersname_list}
        for folder in resp:
            if folder["title"] in foldersname_list:
                folders_dict[folder["title"]] = folder["id"]
        if self.debugMode:
            pprint(folders_dict)
        return folders_dict


# Тестирование класса
'''def test_connect():
    token = os.getenv("wriketoken")
    wr = Wrike(token, debugMode=True)



def test_rs_get():
    token = os.getenv("wriketoken")
    wr = Wrike(token, debugMode=True)
    # js = wr.rs_get("concts")
    js = wr.rs_get("contacts")


def test_id_contacts_on_email():
    token = os.getenv("wriketoken")
    wr = Wrike(token, debugMode=True)
    wr.id_contacts_on_email(["dtitkin@alpintech.ru",
                             "noname@pcht.ru", "echernova@alpintech.ru"])


def test_get_tasks():
    token = os.getenv("wriketoken")
    wr = Wrike(token, debugMode=True)
    resp = wr.get_tasks("tasks", limit=20)


def test_id_folders_on_name():
    token = os.getenv("wriketoken")
    wr = Wrike(token, debugMode=True)
    wr.id_folders_on_name(["ТДВ ТЕСТ НОВЫЕ ПРОДУКТЫ", "Нет такой папки"])


def test_create_task():
    token = os.getenv("wriketoken")
    wr = Wrike(token, debugMode=True)
    wr.create_task("IEAEARAJI4S3SHMJ", "Тестовая задача")


if __name__ == '__main__':
    test_connect()
    # test_rs_get()
    # test_id_contacts_on_email()
    # test_get_tasks()
    # test_id_folders_on_name()
    # test_create_task()'''
