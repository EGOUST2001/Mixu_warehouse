import json, sys, time
import os
import PyQt5
# import checkText
import simtrader
import dlMixu
import xres_rs
from PyQt5 import QtCore, QtWidgets, uic
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import *
from PyQt5.QtWidgets import QLabel, QDialog, QFileDialog, QMessageBox

# Корректное масштабирование интерфейса
if hasattr(QtCore.Qt, 'AA_EnableHighDpiScaling'):
    PyQt5.QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)

if hasattr(QtCore.Qt, 'AA_UseHighDpiPixmaps'):
    PyQt5.QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)

QT_AUTO_SCREEN_SCALE_FACTOR = 1

# Поток для работы прогресс бара
class Thread(QThread):
    _signal = pyqtSignal(int)

    def __init__(self):
        super(Thread, self).__init__()

    def __del__(self):
        self.wait()

    def run(self):
        for i in range(100):
            # time.sleep(0.1)
            self._signal.emit(i)

# Диалог окно для функций
class CustomDialog(QDialog):
    path = ""
    text_input = ""
    path_capha = ""
    path_google = 0

    def __init__(self, path_capha='not_captcha.png', message='Message'):
        self.path_capha = path_capha
        super(CustomDialog, self).__init__()
        uic.loadUi('дизайн/mixu_dialog.ui', self)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
        self.pixmap = QPixmap(self.path_capha)
        self.pixmap = self.pixmap.scaled(188, 106)
        # self.lineEdit_text.setText(checkText.check())
        label = QLabel(self)
        label.setPixmap(self.pixmap)
        form_layout = QtWidgets.QFormLayout(self.image_frame)
        form_layout.addWidget(label)
        self.btn_accept.clicked.connect(self.getText)
        self.btn_getFile.clicked.connect(self.getFileDialog)
        self.lineEdit_text_2.textEdited.connect(self.check)
        self.label_13.setText(message)
        self.setModal(True)
        # self.exec_()

    # Функции для перетаскивания окна
    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.old_pos = event.pos()

    def mouseReleaseEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.old_pos = None

    def mouseMoveEvent(self, event):
        if not self.old_pos:
            return
        delta = event.pos() - self.old_pos
        self.move(self.pos() + delta)

    def getFileDialog(self):
        self.path = QFileDialog.getOpenFileName(caption='Open file', filter="Excel (*.xls *.xlsx)")
        self.btn_getFile.setText(f"Файл: {self.path[0].split('/')[-1]}")
        self.path = self.path[0]

    def getText(self):
        text_input = self.lineEdit_text.text()

    def check(self):
        try:
            if len(self.lineEdit_text_2.text()) > 0:
                self.btn_getFile.setEnabled(False)
                self.btn_accept.setEnabled(True)
                self.path_google = self.lineEdit_text_2.text()
            elif len(self.lineEdit_text_2.text()) == 0:
                self.btn_getFile.setEnabled(True)
                self.btn_accept.setEnabled(False)
        except Exception as e:
            print(e)


class ErrorDialog(QtWidgets.QErrorMessage):
    text_error = ""

    def __init__(self, text_error):
        self.text_error = f"Ошибка: {text_error}"
        super(ErrorDialog, self).__init__()
        uic.loadUi('дизайн/ErrorDialog.ui', self)
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
        self.label_18.setText(text_error)
        self.setModal(True)


    # Функции для перетаскивания окна
    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.old_pos = event.pos()

    def mouseReleaseEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.old_pos = None

    def mouseMoveEvent(self, event):
        if not self.old_pos:
            return
        delta = event.pos() - self.old_pos
        self.move(self.pos() + delta)


class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()  # Call the inherited classes __init__ method
        uic.loadUi('дизайн/mixu_gui.ui', self)  # Load the .ui file
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
        self.show()

        self.btn_state_1.clicked.connect(self.entrance_1)
        self.btn_state_3.clicked.connect(self.refund_1)
        self.btn_state_2.clicked.connect(self.extradition_1)
        self.btn_save_log.clicked.connect(self.save_log)
        self.btn_close.clicked.connect(self.close_app)
        self.btn_trey.clicked.connect(self.trey_app)

        file2 = open("доп/log_file.json")
        data = json.load(file2)

        self.lineEdit_login_1.setText(data.get("simtrader").get("login"))
        self.lineEdit_password_1.setText(data.get("simtrader").get("password"))
        self.lineEdit_login_2.setText(data.get("programm2").get("login"))
        self.lineEdit_password_2.setText(data.get("programm2").get("password"))

    def btnFunc(self):
        self.thread = Thread()
        self.thread.start()
        self.btn.setEnabled(False)

    # Функции для перетаскивания окна
    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.old_pos = event.pos()

    def mouseReleaseEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.old_pos = None

    def mouseMoveEvent(self, event):
        if not self.old_pos:
            return
        delta = event.pos() - self.old_pos
        self.move(self.pos() + delta)

    def entrance_1(self):
        try:
            path_capha = "not_captcha.png"
            driver = 0
            if self.checkBox_1.isChecked():
                driver = simtrader.get_c()
                path_capha = "crop_captcha.png"
            dlg = CustomDialog(path_capha=path_capha, message="Поступление")
            if dlg.exec():
                if not dlg.path_google == 0:
                    # google_disk.download_from_disk(dlg.path_google)
                    dlg.path = os.getcwd() + "/from_disk.xlsx"
                if self.checkBox_1.isChecked():
                    login = self.lineEdit_login_1.text()
                    password = self.lineEdit_password_1.text()
                    capcha = dlg.lineEdit_text.text()
                    simtrader.log_in_to_the_system(driver, login, password, capcha)
                    simtrader.entrance(driver, dlg.path)  # rates,

                if self.checkBox_2.isChecked():
                    login2 = self.lineEdit_login_2.text()
                    password2 = self.lineEdit_password_2.text()
                    driver = dlMixu.log_in(login2, password2, driver=driver)
                    dlMixu.entrance_lc(driver, dlg.path)
                # if self.checkBox_3.isChecked():
                #     print("запуск 3")
                # if not dlg.path_google == 0:
                #     google_disk.load_to_disk(dlg.path_google)
            else:
                print(123)
        except Exception as e:
            er = ErrorDialog(str(e))
            er.exec_()

    def extradition_1(self):
        try:
            path_capha = "not_captcha.png"
            driver = 0
            if self.checkBox_1.isChecked():
                driver = simtrader.get_c()
                path_capha = "crop_captcha.png"
            dlg = CustomDialog(path_capha=path_capha, message="Выдача")
            if dlg.exec():
                if not dlg.path_google == 0:
                    # google_disk.download_from_disk(dlg.path_google)
                    dlg.path = os.getcwd() + "/from_disk.xlsx"
                if self.checkBox_1.isChecked():
                    login = self.lineEdit_login_1.text()
                    password = self.lineEdit_password_1.text()
                    capcha = dlg.lineEdit_text.text()
                    simtrader.log_in_to_the_system(driver, login, password, capcha)
                    simtrader.issuing_sim(driver, dlg.path, self.checkBox_4.isChecked())

                if self.checkBox_2.isChecked():
                    login2 = self.lineEdit_login_2.text()
                    password2 = self.lineEdit_password_2.text()
                    driver = dlMixu.log_in(login2, password2, driver)
                    dlMixu.extradition_lc(driver, dlg.path)
                # if self.checkBox_3.isChecked():
                #     print("запуск 3")
                # if not dlg.path_google == 0:
                #     google_disk.load_to_disk(dlg.path_google)
        except Exception as e:
            er = ErrorDialog(str(e))
            er.exec()

    def refund_1(self):
        try:
            path_capha = "not_captcha.png"
            driver = 0
            if self.checkBox_1.isChecked():
                driver = simtrader.get_c()
                path_capha = "crop_captcha.png"
            dlg = CustomDialog(path_capha=path_capha, message="Возврат")
            if dlg.exec():
                # if not dlg.path_google == 0:
                #     google_disk.download_from_disk(dlg.path_google)
                #     dlg.path = os.getcwd() + "/from_disk.xlsx"
                if self.checkBox_1.isChecked():
                    login = self.lineEdit_login_1.text()
                    password = self.lineEdit_password_1.text()
                    capcha = dlg.lineEdit_text.text()
                    simtrader.log_in_to_the_system(driver, login, password, capcha)

                if self.checkBox_2.isChecked():
                    login2 = self.lineEdit_login_2.text()
                    password2 = self.lineEdit_password_2.text()
                    driver = dlMixu.log_in(login2, password2, driver)
                    dlMixu.refund_lc(driver, dlg.path)
                # if self.checkBox_3.isChecked():
                #     print("запуск 3")
                # if not dlg.path_google == 0:
                #     google_disk.load_to_disk(dlg.path_google)
            else:
                print(123)
        except Exception as e:
            print(e)

    def save_log(self):
        try:
            file2 = open("доп/log_file.json")
            data = json.load(file2)
            data.update(
                {"simtrader": {"login": self.lineEdit_login_1.text(), "password": self.lineEdit_password_1.text()}})
            data.update(
                {"programm2": {"login": self.lineEdit_login_2.text(), "password": self.lineEdit_password_2.text()}})
            data.update(
                {"programm3": {"login": self.lineEdit_login_3.text(), "password": self.lineEdit_password_3.text()}})
            with open("доп/log_file.json", "w", encoding="utf-8") as file:
                json.dump(data, file)
        except Exception as e:
            er = ErrorDialog(str(e))
            er.exec()

    def close_app(self):
        sys.exit()

    def trey_app(self):
        self.showMinimized()


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = Ui()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
