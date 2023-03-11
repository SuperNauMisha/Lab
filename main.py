import sys
import datetime
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow, QDateTimeEdit, QPushButton
from PyQt5.QtSerialPort import QSerialPort, QSerialPortInfo
from PyQt5.QtCore import QIODevice
from pyqtgraph import PlotWidget
import pyqtgraph as pg
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from PyQt5.QtWidgets import QFileDialog


class MyWidget(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('main.ui', self)
        self.setWindowTitle("Коагулограф")
        self.dt_now = datetime.datetime.today()
        self.dateTimeEdit.setDateTime(self.dt_now)
        self.interferences = 0
        self.connectButton.clicked.connect(self.buttonConDis)
        self.saveButton.clicked.connect(self.save)
        self.clearButton.clicked.connect(self.onClear)
        self.serial = QSerialPort()
        self.serial.setBaudRate(9600)
        self.isConnected = False

        self.interferences = 0
        self.data_list = []
        self.time = []
        self.now_time = 0

        self.ports_name_list = []
        self.ports_num_list = []
        self.strok_data = ''
        self.oldstrok_data = ''

        ports = QSerialPortInfo().availablePorts()
        for port in ports:
            full_name = port.portName() + " " + port.description()
            print(full_name)
            self.ports_name_list.append(full_name)
            self.ports_num_list.append(port.portName())

        self.ports.addItems(self.ports_name_list)
        self.serial.readyRead.connect(self.onRead)

    def buttonConDis(self):
        if self.connectButton.text() == "Начать":
            self.connectButton.setText("Остановить")
            self.onConnect()
        elif self.connectButton.text() == "Остановить":
            self.connectButton.setText("Начать")
            self.onDisconnect()

    def onConnect(self):
        print("connect")
        choose_index = self.ports_name_list.index(self.ports.currentText())
        choose_com_port = self.ports_num_list[choose_index]
        print(choose_com_port)
        self.serial.setPortName(choose_com_port)
        self.serial.open(QIODevice.ReadOnly)
        self.serial.readyRead.connect(self.onRead)

    def onDisconnect(self):
        print("disconnect")
        self.serial.close()

    def onClear(self):
        self.interferences = 0
        self.oldstrok_data = ''
        self.strok_data = ''
        self.data_list = []
        self.time = []
        self.now_time = 0
        self.nameEdit.setText('')
        self.numEdit.setText('')
        self.diagnosisEdit.clear()
        self.conditionEdit.clear()
        self.dt_now = datetime.datetime.today()
        self.dateTimeEdit.setDateTime(self.dt_now)
        self.graph.clear()

    def onRead(self):
        try:
            data = self.serial.readLine()
            self.strok_data = str(data)[2:-1]
            if r"\n" not in self.strok_data:
                self.oldstrok_data += self.strok_data
            else:
                print(str(self.oldstrok_data + self.strok_data)[0:-4])
                if self.interferences < 10:
                    self.interferences += 1
                    self.oldstrok_data = ''
                else:
                    self.data_list.append(int(str(self.oldstrok_data + self.strok_data)[0:-4]))
                    self.oldstrok_data = ''
                    self.time.append(self.now_time)
                    self.graph.plot(self.time, self.data_list)
                    self.graph.disableAutoRange()
                    self.graph.setLimits(yMin=0, yMax=600, xMin=0, xMax=1800)
                    self.now_time += 0.5
        except Exception as err:
            print("err", err)

    def save(self):
        named_data_patient = ["Дата и время", "ФИО", "№ Истории болезни", "Диагноз", "Обстоятельства"]
        data_patient = [self.dateTimeEdit.dateTime().toString('dd.MM.yyyy hh:mm'), self.nameEdit.text(),\
                        self.numEdit.text(), self.diagnosisEdit.toPlainText(), self.conditionEdit.toPlainText()]
        wb = Workbook()
        wb.create_sheet(title='Первый лист', index=0)
        sheet = wb['Первый лист']
        print(data_patient[0])
        for col in range(len(self.time)):
            cell = sheet.cell(row=col + 1, column=1)
            cell.value = self.time[col]

        for col in range(len(self.data_list)):
            cell = sheet.cell(row=col + 1, column=2)
            cell.value = self.data_list[col]

        for col in range(len(named_data_patient)):
            cell = sheet.cell(row=col + 1, column=3)
            cell.value = named_data_patient[col]

        for col in range(len(data_patient)):
            cell = sheet.cell(row=col + 1, column=4)
            cell.value = data_patient[col]
        try:
            filename = QFileDialog.getSaveFileName(self, "Сохранить в таблицу", str(data_patient[1].split()[0]) + '_' +\
                                                   self.dateTimeEdit.dateTime().toString('dd-MM-yyyy_hh-mm'), "*.xlsx")
        except:
            filename = QFileDialog.getSaveFileName(self, "Сохранить в таблицу", '', "*.xlsx")
        print(filename)
        try:
            wb.save(filename[0])
        except:
            print('Save error')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyWidget()
    ex.show()
    sys.exit(app.exec_())
