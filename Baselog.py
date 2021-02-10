# модуль класса для создания лога в MySql

# from pprint import pprint


import mysql.connector as mysql
from mysql.connector import errorcode


class Baselog():
    def __init__(self, user, password, host, database, debugMode=False):
        try:
            self.cnx = mysql.connect(user=user, password=password,
                                     host=host, database=database)
            self.sql_add_user = ("INSERT INTO users "
                                 "(name, email, token) "
                                 "VALUES (%(name)s, %(email)s, %(token)s)")
            self.sql_add_log = ("INSERT INTO log_Google_Wrike "
                                "(id_user, runtime_error, "
                                "error_type, num_row, code_1c, "
                                "work_name, message) "
                                "VALUES (%(id_user)s, %(runtime_error)s, "
                                "%(error_type)s, %(num_row)s, %(code_1c)s, "
                                "%(work_name)s, %(message)s)")
        except mysql.Error as err:
            if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
                print("Ошибка логина или пароля")
            elif err.errno == errorcode.ER_BAD_DB_ERROR:
                print("База данных не обнаружена")
            else:
                print(err)

    def add_user(self, name, email, token):
        curs = self.cnx.cursor()
        data_user = {"name": name, "email": email, "token": token}

        try:
            curs.execute(self.sql_add_user, data_user)
            self.cnx.commit()
            return True
        except mysql.Error as err:
            print(err)
            return False

    def add_log(self, **kwargs):
        ''' id_user, runtime_error, error_type, num_row, code_1c
        work_name, message
        '''
        curs = self.cnx.cursor()
        try:
            curs.execute(self.sql_add_log, kwargs)
            self.cnx.commit()
            return True
        except mysql.Error as err:
            print(err)
            return False


if __name__ == '__main__':
    pass
