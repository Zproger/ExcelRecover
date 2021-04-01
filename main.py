import os
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from handler import ThreadHandler
from des import *


# Реализовывает логику интерфейса программы
class Interface(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Данные для запуска подбора
        self.excel_file = None

        # Отключаем стандартные границы окна программы
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.center()

        # Обработчики кнопок
        self.ui.pushButton_3.clicked.connect(self.close)
        self.ui.pushButton.clicked.connect(self.choose_file)
        self.ui.pushButton_2.clicked.connect(self.start_process)

        # Создаем экземпляр класса с потоком и его обработчик
        self.handler = ThreadHandler()
        self.handler.signal.connect(self.signal_handler)


    # Перетаскивание безрамочного окна
    def center(self):
        qr = self.frameGeometry()
        cp = QtWidgets.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def mousePressEvent(self, event):
        self.oldPos = event.globalPos()

    def mouseMoveEvent(self, event):
        try:
            delta = QtCore.QPoint(event.globalPos() - self.oldPos)
            self.move(self.x() + delta.x(), self.y() + delta.y())
            self.oldPos = event.globalPos()
        except AttributeError:
            pass


    # Выбрать файл для обработки
    def choose_file(self):
        file = QtWidgets.QFileDialog.getOpenFileName(self)[0]
        if file:
            self.excel_file = file
            self.ui.label_4.setText(os.path.basename(self.excel_file))


    # Запустить процесс подбора
    def start_process(self):
        # Проверяем выбранные пользователем галочки
        conf_list = [
            self.ui.checkBox.isChecked(),
            self.ui.checkBox_2.isChecked(),
            self.ui.checkBox_3.isChecked(),
        ]

        # Если выбран файл - запускаем процесс
        if self.excel_file:
            # Инициализируем атрибуты класса
            self.handler.filepath = self.excel_file
            self.handler.config = conf_list

            if any(conf_list):
                self.handler.start()
                self.ui.pushButton.setDisabled(True)
                self.ui.pushButton_2.setDisabled(True)
            else:
                message = "Необходимо настроить конфигурацию"
                QtWidgets.QMessageBox.warning(self, "Ошибка", message)
        else:
            message = "Необходимо выбрать файл"
            QtWidgets.QMessageBox.warning(self, "Ошибка", message)


    # Обработчик сигналов
    def signal_handler(self, value):
        if value[0] == "result":
            self.ui.label_3.setText(value[1])
            message = f"Пароль восстановлен: {value[1]}\n"
            message += "Также создан архив без пароля: decrypted.xlsx"
            QtWidgets.QMessageBox.about(self, "Результат", message)
            self.ui.pushButton.setDisabled(False)
            self.ui.pushButton_2.setDisabled(False)
        elif value[0] == "fail":
            self.ui.label_3.setText(value[1])


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = Interface()
    window.show()
    sys.exit(app.exec_())
