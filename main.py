import sys
import datetime
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow, QDateTimeEdit, QPushButton
from PyQt5.QtSerialPort import QSerialPort, QSerialPortInfo
from PyQt5.QtCore import QIODevice
from pyqtgraph import PlotWidget
import pyqtgraph as pg
import openpyxl
from openpyxl.chart import BarChart, Reference
from PyQt5.QtWidgets import QFileDialog


class MyWidget(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('main.ui', self)
        self.setWindowTitle("Коагулограф")
        self.dt_now = datetime.datetime.today()
        self.dateTimeEdit.setDateTime(self.dt_now)
        self.dateTimeEdit_2.setDateTime(self.dt_now)
        self.interferences = 0
        self.connectButton.clicked.connect(self.buttonConDis)
        self.saveButton.clicked.connect(self.save)
        self.clearButton.clicked.connect(self.onClear)
        self.importButton.clicked.connect(self.onImport)
        self.serial = QSerialPort()
        self.serial.setBaudRate(9600)
        self.isConnected = False
        self.graph.disableAutoRange()
        self.graph.setLimits(yMin=0, yMax=250, xMin=0, xMax=1600)
        self.graph.setBackground('w')
        self.pen = pg.mkPen(color=(255, 0, 0))

        self.named_data_patient = ["Дата и время", "Дополнительные дата и время", "ФИО", "№ Истории болезни", "Диагноз", "Обстоятельства", \
                              "Фибриноген", "ПТИ", "МНО", "АЧТВ", "ACT", "Д-Димер"]
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
        self.dateTimeEdit_2.setDateTime(self.dt_now)
        self.graph.clear()
        self.fibrinogenEdit.setValue(0)
        self.ptiEdit.setValue(0)
        self.mnoEdit.setValue(0)
        self.actvEdit.setValue(0)
        self.actEdit.setValue(0)
        self.ddimerEdit.setValue(0)

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

                    self.graph.plot(self.time, self.data_listб, pen=self.pen)

                    self.now_time += 0.5
        except Exception as err:
            print("err", err)

    def save(self):
        data_patient = [self.dateTimeEdit.dateTime().toString('dd.MM.yyyy hh:mm'), \
                        self.dateTimeEdit_2.dateTime().toString('dd.MM.yyyy hh:mm'), self.nameEdit.text(),\
                        self.numEdit.text(), self.diagnosisEdit.toPlainText(), self.conditionEdit.toPlainText(), \
                        self.fibrinogenEdit.value(), self.ptiEdit.value(), self.mnoEdit.value(), self.actvEdit.value(),\
                        self.actEdit.value(), self.ddimerEdit.value()]
        wb = openpyxl.Workbook()
        wb.create_sheet(title='Первый лист', index=0)
        sheet = wb['Первый лист']
        print(data_patient[0])
        for row in range(len(self.time)):
            cell = sheet.cell(row=row + 1, column=1)
            cell.value = self.time[row]

        for row in range(len(self.data_list)):
            cell = sheet.cell(row=row + 1, column=2)
            cell.value = self.data_list[row]

        for row in range(len(self.named_data_patient)):
            cell = sheet.cell(row=row + 1, column=3)
            cell.value = self.named_data_patient[row]

        for row in range(len(data_patient)):
            cell = sheet.cell(row=row + 1, column=4)
            cell.value = data_patient[row]
        print(data_patient)
        try:
            filename = QFileDialog.getSaveFileName(self, "Сохранить в таблицу", \
                                                   str(self.nameEdit.text().split()[0]) + '_' +\
                                                   self.dateTimeEdit.dateTime().toString('dd-MM-yyyy_hh-mm'), "*.xlsx")
        except:
            filename = QFileDialog.getSaveFileName(self, "Сохранить в таблицу", '', "*.xlsx")
        print(filename)
        try:
            wb.save(filename[0])
        except:
            print('Save error')

    def onImport(self):
        self.onClear()
        filename = QFileDialog.getOpenFileName(self, "Импортировать из таблицы", '', "*.xlsx")
        wb = openpyxl.load_workbook(filename[0])
        sh_names = wb.sheetnames
        ind = self.numListEdit.value()
        print(ind)
        sheet = wb[sh_names[ind - 1]]
        row = 1
        while True:
            row += 1
            cell = sheet.cell(row=row, column=1)
            val = cell.value
            if val == None:
                print("end")
                break
            else:
                self.time.append(val)

                cell = sheet.cell(row=row, column=2)
                val = cell.value
                self.data_list.append(val)
        self.graph.plot(self.time, self.data_list, pen=self.pen)
        self.graph.disableAutoRange()
        self.graph.setLimits(yMin=0, yMax=250, xMin=0, xMax=1600)
        try:
            data_patient = []
            for row in range(len(self.named_data_patient)):
                cell = sheet.cell(row=row + 1, column=4)
                val = cell.value
                data_patient.append(val)
            datetime_str1 = data_patient[0]
            datetime1 = datetime.datetime.strptime(datetime_str1, '%d.%m.%Y %H:%M')
            datetime_str2 = data_patient[1]
            datetime2 = datetime.datetime.strptime(datetime_str2, '%d.%m.%Y %H:%M')
            self.dateTimeEdit.setDateTime(datetime1)
            self.dateTimeEdit_2.setDateTime(datetime2)
            self.nameEdit.setText(data_patient[2])
            self.numEdit.setText(data_patient[3])
            self.diagnosisEdit.setPlainText(data_patient[4])
            self.conditionEdit.setPlainText(data_patient[5])
            self.fibrinogenEdit.setValue(data_patient[6])
            self.ptiEdit.setValue(data_patient[7])
            self.mnoEdit.setValue(data_patient[8])
            self.actvEdit.setValue(data_patient[9])
            self.actEdit.setValue(data_patient[10])
            self.ddimerEdit.setValue(data_patient[11])
        except:
            pass





if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyWidget()
    ex.show()
    sys.exit(app.exec_())
