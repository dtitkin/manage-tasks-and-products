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
        self.params = {"access_token": wrike_token}
        self.connect = "https://www.wrike.com/api/v4/"
        if debugMode:
            resp = rs.get(self.connect + "account", self.params)
            pprint(resp.json())

    def rs_get(self, get_str, **kwargs):
        get_params = self.params.copy()
        get_params.update(kwargs)
        resp = rs.get(self.connect + get_str, get_params)
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
            except OopsException as e:
                print("Тест сломался", e)
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






def test_connect():
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

if __name__ == '__main__':
    # test_connect()
    # test_rs_get()
    # test_id_contacts_on_email()
