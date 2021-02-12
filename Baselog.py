# модуль класса для создания лога в MySql

from datetime import datetime

import mysql.connector as mysql
from mysql.connector import errorcode


class Baselog():
    def __init__(self, user, password, host, database, log="all",
                 debugMode=False):
        ''' выводит лог на экран        - log = "terminal"
                        в базу данных   - log = "base"
                        в оба канала    - log = "all"

        '''
        self.log = log
        self.cnx = None
        if log == "base" or log == "all":
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
                self.id_logrow = None
                self.id_user = None
                self.name_user = None
                self.email_user = None
            except mysql.Error as err:
                if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
                    print("Ошибка логина или пароля")
                elif err.errno == errorcode.ER_BAD_DB_ERROR:
                    print("База данных не обнаружена")
                else:
                    print(err)

    def close(self):
        self.cnx.close()

    def is_connected(self):
        if self.cnx is not None:
            return self.cnx.is_connected()
        else:
            return False

    def add_user(self, name, email, token):
        if self.cnx is None:
            return True
        try:
            curs = self.cnx.cursor()
            data_user = {"name": name, "email": email, "token": token}
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
        if self.cnx is None:
            return True

        date_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        kwargs["date_time"] = date_time
        try:
            curs = self.cnx.cursor()
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
        if self.cnx is None:
            return True
        date_time = str(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        kwargs["date_time"] = date_time
        for k, v in kwargs.items():
            if k != "num_row":
                kwargs[k] = "'" + v + "'"
        kwargs["id_row"] = self.id_logrow
        run_str = self.sql_upd_log.format(**kwargs)
        if not self.id_logrow:
            return False
        try:
            curs = self.cnx.cursor()
            curs.execute(run_str)
            self.id_logrow = curs.lastrowid
            self.cnx.commit()
            curs.close()
            return True
        except mysql.Error as err:
            print(err)
            return False

    def out(self, msg, num_row=0, runtime_error="", error_type="",
            prn_time=False, one_str=False):

        def terminal_log(msg, num_row, prn_time=False, one_str=False):
            now = datetime.now()
            fmt = '%d.%m.%Y:%H:%M:%S'
            str_now = now.strftime(fmt)
            num_row_str = ""
            if num_row:
                num_row_str = f"#{num_row}"
            if prn_time:
                print(str_now, end=" ")

            if one_str:
                print('\r', num_row_str, msg, error_type, sep=' ', end='',
                      flush=True)
            else:
                print(num_row_str, msg, error_type)

        id_user = self.id_user
        if id_user is None:
            print("Не установлен пользователь")
            return False
        if self.log == "all" or self.log == "terminal":
            terminal_log(msg, num_row, prn_time, one_str)
            rt = True
        if self.log == "all" or self.log == "base":
            if not self.id_logrow:
                rt = self.add_log(id_user=id_user, runtime_error=runtime_error,
                                  error_type=error_type, num_row=num_row,
                                  message=msg)
            else:
                rt = self.update_log(runtime_error="", error_type=error_type,
                                     num_row=num_row, message=msg)
        return rt

    def get_user(self, user_token):
        try:
            curs = self.cnx.cursor()
            curs.execute(f"SELECT * FROM users WHERE token='{user_token}'")
            rows = curs.fetchone()
        except mysql.Error as err:
            print(err)
            return False
        if rows is None:
            return False
        else:
            self.id_user = rows[0]
            self.name_user = rows[1]
            self.email_user = rows[2]
            return True


if __name__ == '__main__':
    pass
