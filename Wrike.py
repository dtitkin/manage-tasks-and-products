# модуль класса доступа к API Wrike
# в модуле реализованна минимальная часть работы с API
# достаточная для  "Управление продуктами и задачами"
# полная документация на API
# https://developers.google.com/sheets/api
# https://developers.wrike.com/overview/


from pprint import pprint
import requests as rs
import os


class OopsException(Exception):
    pass


class Wrike():
    def __init__(self, wrike_token, debugMode=False):
        self.debugMode = debugMode
        self.headers = {"Authorization": "bearer " + wrike_token}
        self.params = {}
        self.data = {}
        self.connect = "https://www.wrike.com/api/v4/"
        if debugMode:
            resp = rs.get(self.connect + "account",
                          self.params, headers=self.headers)
            pprint(resp.json())

    def test(self, resp):
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
                        print(f"----- {len(value)-1} значение-----")
                        pprint(value[-1])
                    if key == "errorDescription" or key == "error":
                        pprint(value)
            except OopsException as e:
                print("Тест сломался", e)

    def make_params(self, local_dict, exclude_list):
        params = {}
        for arg_name, arg_value in local_dict.items():
            if arg_name not in exclude_list and arg_value:
                params[arg_name] = arg_value
        if self.debugMode:
            pprint(params)
        return params

    def rs_get(self, get_str, **kwargs):
        'выполнение любого запроса get'
        self.params = kwargs.copy()
        try:
            resp = rs.get(self.connect + get_str, self.params,
                          headers=self.headers)
        finally:
            self.params = {}
        self.test(resp)
        return resp

    def rs_post(self, post_str, **kwargs):
        self.data = kwargs.copy()
        try:
            resp = rs.post(self.connect + post_str, self.data,
                           headers=self.headers)
        finally:
            self.data = {}
        self.test(resp)
        return resp

    def get_tasks(self, task_area, descendants=None, title=None, status=None,
                  importance=None, startDate=None, dueDate=None,
                  scheduledDate=None, completedDate=None, type=None, limit=0,
                  sortField=None, sortOrder=None, subTasks=None,
                  customField=None, fields=None):
        '''Получить задачи с параметрами поиска

        https://developers.wrike.com/api/v4/tasks/
        Аргументы
            task_area - tasks
                        folders/{folderId}/tasks
                        spaces/{spaceId}/tasks
                        tasks/{taskId},{taskId},... - up to 100 IDs
            descendants - bool
        '''
        task_params = self.make_params(locals(),
                                       ["self", "task_area", "task_params"])
        resp = self.rs_get(task_area, **task_params)
        return resp

    def create_task(self, folderid, title, status="Active",
                    importance="Normal", dates=None, shareds=None,
                    parents=None, responsibles=None, superTasks=None,
                    customFields=None, fields=None):
        '''Создать задачу
        '''
        task_dates = self.make_params(locals(),
                                      ["self", "folderid", "task_dates"])
        resp = self.rs_post("folders/" + folderid + "/tasks", **task_dates)
        return resp

    def id_contacts_on_email(self, email_list):
        resp = self.rs_get("contacts")
        data = resp.json()["data"]
        id_dict = {mail: None for mail in email_list}
        for user in data:
            if user['type'] == 'Person':
                user_mail = user["profiles"][0]["email"]
                if user_mail in email_list:
                    id_dict[user_mail] = user["id"]
        if self.debugMode:
            pprint(id_dict)
        return id_dict

    def id_folders_on_name(self, foldersname_list):
        resp = self.rs_get("folders")
        data = resp.json()["data"]
        folders_dict = {name: None for name in foldersname_list}
        for folder in data:
            if folder["title"] in foldersname_list:
                folders_dict[folder["title"]] = folder["id"]
        if self.debugMode:
            pprint(folders_dict)
        return folders_dict


# Тестирование класса
def test_connect():
    token = os.getenv("wriketoken")
    wr = Wrike(token, debugMode=True)
    data = resp.json()["data"]


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
    resp = wr.get_tasks("tasks", limit=10)


def test_id_folders_on_name():
    token = os.getenv("wriketoken")
    wr = Wrike(token, debugMode=True)
    wr.id_folders_on_name(["ТДВ ТЕСТ НОВЫЕ ПРОДУКТЫ", "Нет такой папки"])


def test_create_task():
    token = os.getenv("wriketoken")
    wr = Wrike(token, debugMode=True)
    wr.create_task("IEAEARAJI4S3SHMJ", "Тестовая задача")


if __name__ == '__main__':
    # test_connect()
    # test_rs_get()
    # test_id_contacts_on_email()
    #test_get_tasks()
    # test_id_folders_on_name()
    #test_create_task()

