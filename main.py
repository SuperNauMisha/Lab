import sys
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton
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
        self.interferences = 0
        self.connectButton.clicked.connect(self.onConnect)
        self.disconnectButton.clicked.connect(self.onDisconnect)
        self.saveButton.clicked.connect(self.save)
        self.serial = QSerialPort()
        self.serial.setBaudRate(9600)

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



    def onConnect(self):
        print("connect")
        choose_index = self.ports_name_list.index(self.ports.currentText())
        choose_com_port = self.ports_num_list[choose_index]
        print(choose_com_port)
        self.serial.setPortName(choose_com_port)
        self.serial.open(QIODevice.ReadOnly)
        self.serial.readyRead.connect(self.onRead)

    def onDisconnect(self):
        self.serial.close()


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
        wb = Workbook()
        wb.create_sheet(title='Первый лист', index=0)
        sheet = wb['Первый лист']
        for col in range(len(self.data_list)):
            cell = sheet.cell(row=col + 1, column=2)
            cell.value = self.data_list[col]

        for col in range(len(self.time)):
            cell = sheet.cell(row=col + 1, column=1)
            cell.value = self.time[col]

        filename = QFileDialog.getSaveFileName(self, "Сохранить в таблицу", "", "*.xlsx")
        print(filename)
        wb.save(filename[0])



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyWidget()
    ex.show()
    sys.exit(app.exec_())
