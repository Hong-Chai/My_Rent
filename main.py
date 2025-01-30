import sqlite3
import sys
import maps
import datetime
import compliments
import random

from docxtpl import DocxTemplate

import PyQt6
from PyQt6 import uic
from PyQt6.QtWidgets import QApplication, QFileDialog
from PyQt6.QtWidgets import QMainWindow, QTableWidgetItem, QWidget, QInputDialog
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtCore import QUrl
from PyQt6.QtGui import QIcon


def init_db():
    file_db = open("db.txt", "r", encoding="utf-8")
    global connection
    try:
        path = file_db.readline().strip()
        file_test = open(path, "r", encoding="utf-8")
        file_test.close()
        connection = sqlite3.connect(path)
    except FileNotFoundError:
        file_db.close()
        file_db = open("db.txt", "w", encoding="utf-8")
        file_db.write("test_new.db3")
        file_db.close()
        file_db = open("db.txt", "r", encoding="utf-8")
        connection = sqlite3.connect(file_db.readline().strip())
    file_db.close()


def upd_db():
    reply = []
    query = "SELECT user_name, clients.full_name, clients.out_date, clients.last_gen, clients.balance, clients.hist, MSU_hist, last_MSU, clients.in_date, clients.price,hot,cold,EL1,EL2  FROM main LEFT JOIN clients ON main.link_client = clients.id"
    res = connection.cursor().execute(query).fetchall()
    for row in res:
        if row[1] == None:
            continue

        # повышение читабельности
        user_name = row[0]
        full_name = row[1]
        out_date = row[2]
        last_gen = row[3]
        balance = row[4]
        hist = row[5]
        MSU_hist = row[6]
        last_MSU = row[7]
        in_date = row[8]
        price = row[9]
        hot = row[10]
        cold = row[11]
        EL1 = row[12]
        EL2 = row[13]

        # выезд клиента
        if datetime.datetime.strptime(row[2], "%d.%m.%Y") < datetime.datetime.now():
            reply.append(
                f"{row[1]} должен БЫЛ покинуть обьект {row[0]} еще {row[2]}, если у вас нет претензий, отвяжите его от обьекта!"
            )
        elif datetime.datetime.strptime(row[2], "%d.%m.%Y") == datetime.date.today():
            reply.append(
                f"{row[1]} должен покинуть обьект {row[0]} сегодня, если у вас нет претензий, отвяжите его от обьекта!"
            )
        elif datetime.datetime.strptime(
            row[2], "%d.%m.%Y"
        ) == datetime.date.today() + datetime.timedelta(days=1):
            reply.append(
                f"{row[1]} должен покинуть обьект {row[0]} завтра, не забудтьте передать показания для последнего платежа!"
            )
        elif datetime.datetime.strptime(
            row[2], "%d.%m.%Y"
        ) == datetime.date.today() + datetime.timedelta(days=2):
            reply.append(
                f"{row[1]} должен покинуть обьект {row[0]} послезавтра, не забудтьте передать показания для последнего платежа!"
            )
        elif datetime.datetime.strptime(
            row[2], "%d.%m.%Y"
        ) == datetime.date.today() + datetime.timedelta(days=3):
            reply.append(
                f"{row[1]} должен покинуть обьект {row[0]} через 3 дня, не забудтьте передать показания для последнего платежа!"
            )

        # исправление ошибок
        if last_gen == None:
            last_gen = in_date

        if balance == None:
            balance = 0
            query = f"UPDATE clients SET balance = '{balance}' WHERE id = (SELECT link_client FROM main WHERE user_name = '{user_name}')"
            connection.cursor().execute(query)
            connection.commit()
        else:
            balance = float(balance)

        if last_MSU == None:
            last_MSU = f"{in_date} 0 0 0 0"
        last_MSU_date = last_MSU.split()[0]
        if MSU_hist != None:
            last_new_MSU_date = MSU_hist.split("`")[-1].split()[0]
        else:
            last_new_MSU_date = last_MSU_date

        # генерация оплат
        while (
            datetime.datetime.strptime(last_gen, "%d.%m.%Y")
            + datetime.timedelta(days=30)
            < datetime.datetime.now()
        ):
            last_gen = datetime.datetime.strptime(
                last_gen, "%d.%m.%Y"
            ) + datetime.timedelta(days=30)
            last_gen = last_gen.strftime("%d.%m.%Y")
            total_price = int(price)
            MSU_text = f"Счет за {last_gen}:\nЕжемесячная плата -- {price}"
            if datetime.datetime.strptime(
                last_MSU_date, "%d.%m.%Y"
            ) != datetime.datetime.strptime(last_new_MSU_date, "%d.%m.%Y"):
                print("new MSU")
                last_MSU1 = last_MSU.split()
                now_MSU = MSU_hist.split("`")[-1].split()
                total_price += (float(now_MSU[1]) - float(last_MSU1[1])) * float(hot)
                total_price += (float(now_MSU[2]) - float(last_MSU1[2])) * float(cold)
                total_price += (float(now_MSU[3]) - float(last_MSU1[3])) * float(EL1)
                total_price += (float(now_MSU[4]) - float(last_MSU1[4])) * float(EL2)

                last_MSU_date = last_new_MSU_date
                last_MSU = MSU_hist.split("`")[-1]
                query = f"UPDATE main SET last_MSU = '{last_MSU}' WHERE user_name = '{user_name}'"
                connection.cursor().execute(query)
                connection.commit()
                MSU_text += (
                    f"\nУчтены показания {last_MSU} -- {total_price - int(price)}"
                )
                reply.append(
                    f"Показания от {last_MSU_date} учтены в счете за {user_name}"
                )

            balance -= total_price
            query = f"UPDATE clients SET balance = '{balance}' WHERE id = (SELECT link_client FROM main WHERE user_name = '{user_name}')"
            connection.cursor().execute(query)
            connection.commit()

            if hist == None:
                hist = f"{last_gen} {total_price} {MSU_text}"
            else:
                hist += f"`{last_gen} {total_price} {MSU_text}"
            query = f"UPDATE clients SET hist = '{hist}' WHERE id = (SELECT link_client FROM main WHERE user_name = '{user_name}')"
            connection.cursor().execute(query)
            connection.commit()

            reply.append(
                f"Cгенерирован платеж для {full_name} за {last_gen} на сумму {total_price}"
            )
        query = f"UPDATE clients SET last_gen = '{last_gen}' WHERE id = (SELECT link_client FROM main WHERE user_name = '{user_name}')"
        connection.cursor().execute(query)
        connection.commit()

        # напоминание о показаниях
        if datetime.datetime.strptime(
            last_MSU_date, "%d.%m.%Y"
        ) == datetime.datetime.strptime(
            last_new_MSU_date, "%d.%m.%Y"
        ) and datetime.datetime.strptime(
            last_gen, "%d.%m.%Y"
        ) + datetime.timedelta(
            days=30
        ) <= datetime.datetime.now() + datetime.timedelta(
            days=5
        ):
            reply.append(
                f"Передайте показания за обьект {user_name} до {datetime.datetime.strptime(last_gen, "%d.%m.%Y") + datetime.timedelta(days=30)}!"
            )

        # отрицательный баланс
        if balance < 0:
            reply.append(
                f"У клиента {full_name} на обьекте {user_name} отрицательный баланс!"
            )

    return reply


init_db()

reply = upd_db()


class Main_WIN(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi("main.ui", self)
        self.setWindowTitle("МОЯ аренда")
        self.setWindowIcon(QIcon("icons/edit.png"))
        hi_text = f"Добро пожаловать!\n\nСегодня\n\n{datetime.datetime.now().strftime('%d.%m.%Y')}\n\n{compliments.get_compliment()}\n\nХорошего дня!"
        self.hi_label.setText(hi_text)
        self.gt_edit.clicked.connect(self.open_edit)
        self.gt_map.clicked.connect(self.open_map)
        self.gt_settings.clicked.connect(self.open_settings)
        self.gt_fastMSU.clicked.connect(self.open_fastMSU)
        self.gt_fastPay.clicked.connect(self.open_fastPay)
        for el in reply:
            self.list_events.addItem(el)

    def load_data(self):
        pass

    def open_edit(self):
        self.second_form = Edit_WIN(self)
        self.second_form.show()

    def open_map(self):
        self.second_form = Map_WIN(self)
        self.second_form.show()

    def open_settings(self):
        self.second_form = Settings_WIN(self)
        self.second_form.show()

    def open_fastMSU(self):
        self.second_form = addMSU_WIN(self)
        self.second_form.show()

    def open_fastPay(self):
        self.second_form = addPay_WIN(self)
        self.second_form.show()

    def closeEvent(self, event):
        app.closeAllWindows()
        connection.close()


class Map_WIN(QWidget):
    def __init__(self, par):
        super().__init__()
        self.setWindowTitle("Интерактивная карта")
        self.setGeometry(100, 100, 800, 600)
        self.browser = QWebEngineView(self)
        self.browser.setGeometry(0, 0, 800, 600)
        self.load_data()

    def load_data(self):
        query = "SELECT * FROM main"
        res = connection.cursor().execute(query).fetchall()
        coords = []
        popups = []
        colors = []
        for row in res:
            coords.append(row[2].split(","))
            popups.append(row[0])
            if row[1] != None:
                date_to_end = (
                    connection.cursor()
                    .execute(f"SELECT out_date FROM clients WHERE {row[1]} IS id;")
                    .fetchall()[0][0]
                )
                date_to_end = datetime.datetime.strptime(date_to_end, "%d.%m.%Y")
                if date_to_end - datetime.datetime.now() < datetime.timedelta(days=7):
                    colors.append("purple")
                else:
                    colors.append("red")
            else:
                colors.append("green")

            print(colors)

        maps.create_map_with_markers(coords, popups, colors)
        self.browser.setHtml(open("map1.html", "r", encoding="utf-8").read())


class Edit_WIN(QWidget):
    def __init__(self, par):
        super().__init__()
        uic.loadUi("edit.ui", self)
        self.setWindowTitle("Моя недвижимость")
        self.load_data()
        self.table.doubleClicked.connect(self.open_WIN)
        self.add.clicked.connect(self.add_data)

    def load_data(self):
        query = "SELECT user_name, short_adr, clients.full_name, clients.balance FROM main LEFT JOIN clients ON main.link_client = clients.id"
        res = connection.cursor().execute(query).fetchall()
        for i in range(len(res)):
            res[i] += ("Подробнее...",)
        self.table.setColumnCount(5)
        self.table.setRowCount(0)
        for i, row in enumerate(res):
            self.table.setRowCount(self.table.rowCount() + 1)
            for j, elem in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(str(elem)))
        self.table.resizeColumnsToContents()

    def add_data(self):
        self.second_form = Adr_WIN(self, None)
        self.second_form.show()

    def open_WIN(self):
        i, j = self.table.currentColumn(), self.table.currentRow()
        if i == 1:
            self.second_form = Adr_WIN(self, self.table.item(j, 0).text())
            self.second_form.show()
        elif i == 2:
            self.second_form = Client_WIN(self, self.table.item(j, 0).text())
            self.second_form.show()
        elif i == 3:
            self.second_form = Pay_WIN(self, self.table.item(j, 0).text())
            self.second_form.show()
        elif i == 4:
            self.second_form = MSU_WIN(self, self.table.item(j, 0).text())
            self.second_form.show()


class Adr_WIN(QWidget):
    def __init__(self, par, id):
        super().__init__()
        uic.loadUi("edit_adr.ui", self)
        self.setWindowTitle("Адрес")
        self.id = id
        self.par = par
        self.flag = "old"
        self.load_data()
        self.error_dialog = PyQt6.QtWidgets.QErrorMessage()
        self.del1.clicked.connect(self.delete)
        self.save.clicked.connect(self.save_data)

    def load_data(self):
        query = "SELECT * FROM main WHERE user_name = ?"
        res = connection.cursor().execute(query, (self.id,)).fetchall()
        if res == []:
            self.flag = "new"
            return
        self.coords.setText(res[0][2])
        self.country.setText(res[0][3])
        self.region.setText(res[0][4])
        self.city.setText(res[0][5])
        self.street.setText(res[0][6])
        self.building.setText(res[0][7])
        self.flat.setText(res[0][8])
        self.hot.setValue(float(res[0][10]))
        self.cold.setValue(float(res[0][11]))
        self.EL1.setValue(float(res[0][12]))
        self.EL2.setValue(float(res[0][13]))

    def save_data(self):
        try:
            test = self.coords.text().split(", ")
            if (
                len(test) != 2
                or float(test[0]) > 90
                or float(test[0]) < -90
                or float(test[1]) > 180
                or float(test[1]) < -180
            ):
                raise BaseException
        except BaseException:
            self.error_dialog.showMessage("Неверный формат координат")
            return
        if self.flag == "new":
            name, ok_pressed = QInputDialog.getText(
                self, "Введите имя обьекта", "Введите имя обьекта"
            )
            if not ok_pressed:
                name = self.street.text() + str(random.randint(1, 100))
            query = "INSERT INTO main (user_name, coords, country, region, city, street, building, flat, hot, cold, EL1, EL2, short_adr) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            connection.cursor().execute(
                query,
                (
                    str(name),
                    self.coords.text(),
                    self.country.text(),
                    self.region.text(),
                    self.city.text(),
                    self.street.text(),
                    self.building.text(),
                    self.flat.text(),
                    self.hot.value(),
                    self.cold.value(),
                    self.EL1.value(),
                    self.EL2.value(),
                    self.street.text()
                    + " "
                    + self.building.text()
                    + " "
                    + self.flat.text(),
                ),
            )
            connection.commit()
            self.par.load_data()
            self.close()
            return
        query = "UPDATE main SET coords = ?, country = ?, region = ?, city = ?, street = ?, building = ?, flat = ?, hot = ?, cold = ?, EL1 = ?, EL2 = ?, short_adr = ? WHERE user_name = ?"
        connection.cursor().execute(
            query,
            (
                self.coords.text(),
                self.country.text(),
                self.region.text(),
                self.city.text(),
                self.street.text(),
                self.building.text(),
                self.flat.text(),
                self.hot.value(),
                self.cold.value(),
                self.EL1.value(),
                self.EL2.value(),
                self.street.text()
                + " "
                + self.building.text()
                + " "
                + self.flat.text(),
                self.id,
            ),
        )
        connection.commit()
        self.par.load_data()
        self.close()

    def delete(self):
        query = "DELETE FROM main WHERE user_name = ?"
        connection.cursor().execute(query, (self.id,))
        connection.commit()
        self.par.load_data()
        self.close()


class Settings_WIN(QWidget):
    def __init__(self, par):
        super().__init__()
        uic.loadUi("settings.ui", self)
        self.setWindowTitle("Настройки")
        self.load_data()
        self.change.clicked.connect(self.change_db)
        self.save.clicked.connect(self.save_data)

    def load_data(self):
        file = open("settings.txt", "r", encoding="utf-8")
        l = list(file.read().split("`"))
        file.close()
        self.name.setText(l[0])
        self.document.setText(l[1])
        self.phone.setText(l[2])
        self.card.setText(l[3])

    def save_data(self):
        file = open("settings.txt", "w", encoding="utf-8")
        file.write(self.name.text() + "`")
        file.write(self.document.text() + "`")
        file.write(self.phone.text() + "`")
        file.write(self.card.text())
        file.close()
        self.close()

    def change_db(self):
        file = open("db.txt", "w", encoding="utf-8")
        fname = QFileDialog.getOpenFileName(
            self, "Выбрать БД", "", "База данных (*.db3);"
        )[0]
        if fname == "":
            return
        file.write(fname)
        file.close()
        connection.close()
        init_db()


class Client_WIN(QWidget):
    def __init__(self, par, id):
        super().__init__()
        uic.loadUi("edit_cl.ui", self)
        self.setWindowTitle("Клиент")
        self.flag = "old"
        self.id = id
        self.par = par
        self.error_dialog = PyQt6.QtWidgets.QErrorMessage()
        self.del1.clicked.connect(self.delete)
        self.save.clicked.connect(self.save_data)
        self.gen_doc.clicked.connect(self.generate_doc)
        self.load_data()

    def load_data(self):
        query = "SELECT * FROM clients WHERE id = (SELECT link_client FROM main WHERE user_name = ?)"
        res = connection.cursor().execute(query, (self.id,)).fetchall()
        self.res = res
        if res == []:
            self.flag = "new"
            self.length.setEnabled(True)
            self.length.setStyleSheet("")
            return
        self.name.setText(res[0][1])
        self.document.setText(res[0][2])
        self.phone.setText(res[0][3])
        self.l1 = res[0][4]
        self.length.setText("до " + res[0][6])
        self.price.setText(res[0][7])

    def generate_doc(self):
        doc = DocxTemplate("templates\\arenda.docx")
        context = {
            "client_name": self.name.text(),
            "client_doc": self.document.text(),
            "length": self.l1,
            "price": self.price.text(),
            "client_phone": self.phone.text(),
            "in_date": self.res[0][5],
            "out_date": self.res[0][6],
            "cur_date": datetime.datetime.now().strftime("%d.%m.%Y"),
            "my_doc": open("settings.txt", "r", encoding="utf-8").read().split("`")[1],
            "my_phone": open("settings.txt", "r", encoding="utf-8")
            .read()
            .split("`")[2],
            "my_name": open("settings.txt", "r", encoding="utf-8").read().split("`")[0],
            "adr": connection.cursor()
            .execute("SELECT short_adr FROM main WHERE user_name = ?", (self.id,))
            .fetchall()[0][0],
            "city": connection.cursor()
            .execute("SELECT city FROM main WHERE user_name = ?", (self.id,))
            .fetchall()[0][0],
        }
        doc.render(context)
        doc.save(f"{self.name.text()} договор аренды.docx")

    def save_data(self):
        try:
            test = self.price.text()
            if float(test) < 0:
                raise BaseException
        except BaseException:
            self.error_dialog.showMessage("Неверный формат цены")
            return
        if self.flag == "new":
            try:
                test = self.length.text()
                if int(test) < 0:
                    raise BaseException
            except BaseException:
                self.error_dialog.showMessage("Неверный формат длительности")
                return
            query = "INSERT INTO clients (full_name, document, phone, length, in_date, out_date, price) VALUES (?, ?, ?, ?, ?, ?, ?)"
            connection.cursor().execute(
                query,
                (
                    self.name.text(),
                    self.document.text(),
                    self.phone.text(),
                    self.length.text(),
                    datetime.datetime.now().strftime("%d.%m.%Y"),
                    (
                        datetime.datetime.now()
                        + datetime.timedelta(days=int(self.length.text()) * 30)
                    ).strftime("%d.%m.%Y"),
                    self.price.text(),
                ),
            )
            connection.commit()
            query = "UPDATE main SET link_client = (SELECT id FROM clients WHERE full_name = ?) WHERE user_name = ?"
            connection.cursor().execute(query, (self.name.text(), self.id))
            connection.commit()
            self.par.load_data()
            self.close()
            return
        query = "UPDATE clients SET full_name = ?, document = ?, phone = ?, price = ? WHERE id = (SELECT link_client FROM main WHERE user_name = ?)"
        connection.cursor().execute(
            query,
            (
                self.name.text(),
                self.document.text(),
                self.phone.text(),
                self.price.text(),
                self.id,
            ),
        )
        connection.commit()
        self.par.load_data()
        self.close()

    def delete(self):
        query = "UPDATE main SET link_client = NULL WHERE link_client = (SELECT id FROM clients WHERE id = (SELECT link_client FROM main WHERE user_name = ?))"
        connection.cursor().execute(query, (self.id,))
        connection.commit()
        query = "DELETE FROM clients WHERE id = (SELECT link_client FROM main WHERE user_name = ?)"
        connection.cursor().execute(query, (self.id,))
        connection.commit()
        self.par.load_data()
        self.close()


class MSU_WIN(QWidget):
    def __init__(self, par, id):
        super().__init__()
        uic.loadUi("view_MSU.ui", self)
        self.setWindowTitle("MSU")
        self.id = id
        self.load_data()
        self.add.clicked.connect(self.add_data)

    def load_data(self):
        self.list.clear()
        query = "SELECT MSU_hist FROM main WHERE user_name = ?"
        res = connection.cursor().execute(query, (self.id,)).fetchall()
        data = []
        if res[0][0] != None:
            data = res[0][0].split("`")
        for i in reversed(data):
            i = i.split(" ")
            self.list.addItem(
                i[0]
                + "\nГВС: "
                + i[1]
                + " ХВС: "
                + i[2]
                + "\nЭл1: "
                + i[3]
                + " Эл2: "
                + i[4]
            )

    def add_data(self):
        self.secondWindow = addMSU_WIN(self, self.id)
        self.secondWindow.show()


class addMSU_WIN(QWidget):
    def __init__(self, par, iter=None):
        super().__init__()
        uic.loadUi("parce_MSU.ui", self)
        self.setWindowTitle("MSU")
        self.par = par
        self.iter = iter
        self.save.clicked.connect(self.save_data)
        self.load_data()

    def load_data(self):
        query = "SELECT user_name FROM main"
        res = connection.cursor().execute(query).fetchall()
        for i in res:
            self.choose.addItem(i[0])
        if self.iter != None:
            self.choose.setCurrentText(self.iter)
            self.choose.setEnabled(False)
            self.choose.setStyleSheet("background:lightgrey")

    def save_data(self):
        id = self.choose.currentText()
        query = "SELECT MSU_hist FROM main WHERE user_name = ?"
        res = connection.cursor().execute(query, (id,)).fetchall()
        data = []
        if res[0][0] != None:
            data = res[0][0].split("`")
        add_text = (
            datetime.datetime.now().strftime("%d.%m.%Y")
            + " "
            + str(self.hot.value())
            + " "
            + str(self.cold.value())
            + " "
            + str(self.EL1.value())
            + " "
            + str(self.EL2.value())
        )
        data.append(add_text)
        data = "`".join(data)
        query = "UPDATE main SET MSU_hist = ? WHERE user_name = ?"
        connection.cursor().execute(query, (data, id))
        connection.commit()
        self.par.load_data()
        self.close()


class Pay_WIN(QWidget):
    def __init__(self, par, id):
        super().__init__()
        uic.loadUi("view_MSU.ui", self)
        self.setWindowTitle("PAY")
        self.id = id
        self.load_data()
        self.add.clicked.connect(self.add_data)

    def load_data(self):
        self.list.clear()
        query = "SELECT hist FROM clients WHERE id = (SELECT link_client FROM main WHERE user_name = ?)"
        res = connection.cursor().execute(query, (self.id,)).fetchall()
        data = []
        if res == []:
            self.close()
            self.add.setEnabled(0)
            return
        if res[0][0] != None:
            data = res[0][0].split("`")
        for i in reversed(data):
            self.list.addItem(i)

    def add_data(self):
        self.secondWindow = addPay_WIN(self, self.id)
        self.secondWindow.show()


class addPay_WIN(QWidget):
    def __init__(self, par, iter=None):
        super().__init__()
        uic.loadUi("pay.ui", self)
        self.setWindowTitle("PAY")
        self.par = par
        self.iter = iter
        self.pay.setChecked(1)
        self.save.clicked.connect(self.save_data)
        self.load_data()

    def load_data(self):
        query = "SELECT user_name FROM main WHERE link_client NOT NULL"
        res = connection.cursor().execute(query).fetchall()
        for i in res:
            self.choose.addItem(i[0])
        if self.iter != None:
            self.choose.setCurrentText(self.iter)
            self.choose.setEnabled(False)
            self.choose.setStyleSheet("background:lightgrey")

    def save_data(self):
        id = self.choose.currentText()
        summ = self.total.value()
        comm = ""
        if self.pay.isChecked():
            comm = "Оплата"
        else:
            comm = "Пользовательский счет"
            summ = -summ

        comm += (
            f" от {datetime.date.today()}: \n{self.comm.text()} -- {self.total.value()}"
        )
        query = "SELECT balance FROM clients WHERE id = (SELECT link_client FROM main WHERE user_name = ?)"
        res = connection.cursor().execute(query, (id,)).fetchall()
        balance = float(res[0][0])

        query = f"UPDATE clients SET balance = '{balance + summ}' WHERE id = (SELECT link_client FROM main WHERE user_name = '{id}')"
        connection.cursor().execute(query)
        connection.commit()

        query = "SELECT hist FROM clients WHERE id = (SELECT link_client FROM main WHERE user_name = ?)"
        res = connection.cursor().execute(query, (id,)).fetchall()
        hist = res[0][0]

        if hist == None:
            hist = comm
        else:
            hist += f"`{comm}"
        query = f"UPDATE clients SET hist = '{hist}' WHERE id = (SELECT link_client FROM main WHERE user_name = '{id}')"
        connection.cursor().execute(query)
        connection.commit()
        self.par.load_data()
        self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = Main_WIN()
    ex.show()
    sys.exit(app.exec())
