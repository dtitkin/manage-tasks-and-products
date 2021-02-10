# модуль класса для создания лога в MySql

# from pprint import pprint

from datetime import datetime

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
                                "(date_time, id_user, runtime_error, "
                                "error_type, num_row, message) "
                                "VALUES (%(date_time)s,%(id_user)s, "
                                "%(runtime_error)s, %(error_type)s, "
                                "%(num_row)s, %(message)s)")
            self.sql_upd_log = (f"UPDATE log_Google_Wrike SET "
                                "runtime_error={runtime_error:s}, "
                                "date_time={date_time:s}, "
                                "error_type={error_type:s}, "
                                "num_row={num_row}, "
                                "message={message:s} "
                                "WHERE id={id_row}")
            self.id_logrow = ""
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
            curs.close()
            return True
        except mysql.Error as err:
            print(err)
            return False

    def add_log(self, **kwargs):
        ''' id_user, runtime_error, error_type, num_row, message
        '''
        date_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        kwargs["date_time"] = date_time
        curs = self.cnx.cursor()
        try:
            curs.execute(self.sql_add_log, kwargs)
            self.id_logrow = curs.lastrowid
            self.cnx.commit()
            curs.close()
            return True
        except mysql.Error as err:
            print(err)
            return False

    def update_log(self, **kwargs):
        '''runtime_error, error_type,message
        '''
        curs = self.cnx.cursor()
        date_time = "'" + \
                    str(datetime.now().strftime("%Y-%m-%d %H:%M:%S")) + "'"
        kwargs["date_time"] = date_time
        for k, v in kwargs.items():
            kwargs[k] = "'" + v + "'"
        kwargs["id_row"] = self.id_logrow
        run_str = self.sql_upd_log.format(**kwargs)
        print(run_str)
        if not self.id_logrow:
            return False
        try:
            curs.execute(run_str)
            self.id_logrow = curs.lastrowid
            self.cnx.commit()
            curs.close()
            return True
        except mysql.Error as err:
            print(err)
            return False


if __name__ == '__main__':
    pass
