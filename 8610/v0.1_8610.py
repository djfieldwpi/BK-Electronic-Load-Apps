# Old tester compilation command
# CD to .py file location
# pyinstaller --onefile --windowed --name="BatteryTester" guimin.py

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QLineEdit, QPushButton, QStyle, QMessageBox)
from PyQt6.QtCore import QSize
from PyQt6.QtGui import QDoubleValidator

from openpyxl import Workbook, load_workbook
from datetime import datetime

from pathlib import Path

import pyvisa, time, json

# ====================================
CONFIG_FILE = Path.home() / "BatteryTester" / "8610_settings.json"
CONFIG_FILE.parent.mkdir(exist_ok=True)

def load_settings():
    global defaultPath, defaultCurrLimit, defaultConstPower, defaultStopVoltage

    if not CONFIG_FILE.exists():
        return

    with open(CONFIG_FILE, "r") as f:
        data = json.load(f)

    defaultPath = data.get("defaultPath", "")
    defaultCurrLimit = data.get("defaultCurrlimit", 0.0)
    defaultConstPower = data.get("defaultConstPower", 0.0)
    defaultStopVoltage = data.get("defaultStopVoltage", 0.0)

def save_settings():
    data = {
        "defaultPath": defaultPath,
        "defaultCurrLimit": defaultCurrLimit,
        "defaultConstPower": defaultConstPower,
        "defaultStopVoltage": defaultStopVoltage
    }

    with open(CONFIG_FILE, "w") as f:
        json.dump(data, f, indent=4)

defaultPath = ""
defaultCurrLimit = 0.0
defaultConstPower = 0.0
defaultStopVoltage = 0.0


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Flags
        self.started = False

        # Connection
        self.rm = None
        self.inst = None

        # Monitor
        self.monitorVoltage = 0.0
        self.monitorCurrent = 0.0
        self.monitorPower = 0.0

        # Record
        self.recordVoltage = 0.0
        self.recordCurrent = 0.0
        self.recordPower = 0.0
        self.recordStatus = ""

        # Excel
        self.wb = None
        self.filePath = ""

        self.setWindowTitle("Battery Load Tester")
        self.setFixedSize(QSize(600, 900))

        # Overall Layout
        vLayout = QVBoxLayout()
        fileLayout = QHBoxLayout()
        limitLayout = QHBoxLayout()
        constLayout = QHBoxLayout()
        stopLayout = QHBoxLayout()
        buttonLayout = QHBoxLayout()
        vLayout.addLayout(fileLayout)
        vLayout.addLayout(limitLayout)
        vLayout.addLayout(constLayout)
        vLayout.addLayout(stopLayout)
        vLayout.addLayout(buttonLayout)

        # File Layout
        self.fileLabel = QLabel("<b style=\"font-size: 20px\">File Name:</b>")
        self.fileEntry = QLineEdit()
        self.fileEntry.setStyleSheet("""
                                    QLineEdit{
                                    font-size: 20px
                                    }
                                    """)
        self.fileEntry.setText(defaultPath)
        self.fileEntry.setEnabled(not self.started)
        self.testLabel = QLabel("<b style=\"font-size: 20px\">Lot Number:</b>")
        self.testEntry = QLineEdit()
        self.testEntry.setStyleSheet("""
                                    QLineEdit{
                                    font-size: 20px
                                    }
                                    """)
        self.testEntry.setPlaceholderText("Test ID...")
        self.testEntry.setMaximumWidth(200)
        self.testEntry.setEnabled(not self.started)

        fileLayout.addWidget(self.fileLabel)
        fileLayout.addWidget(self.fileEntry)
        fileLayout.addWidget(self.testLabel)
        fileLayout.addWidget(self.testEntry)

        # Current Limit Layout
        self.limitLabel = QLabel("<b style=\"font-size: 20px\">Current Limit:</b>")
        self.limitEntry = QLineEdit()
        self.limitEntry.setStyleSheet("""
                                    QLineEdit{
                                    font-size: 20px
                                    }
                                    """)
        self.limitEntry.setPlaceholderText(str(defaultCurrLimit) + " A")
        self.limitEntry.setValidator(QDoubleValidator())
        self.limitEntry.setEnabled(not self.started)

        limitLayout.addWidget(self.limitLabel)
        limitLayout.addWidget(self.limitEntry)

        # Constant Power Layout
        self.constLabel = QLabel("<b style=\"font-size: 20px\">Constant Power:</b>")
        self.constEntry = QLineEdit()
        self.constEntry.setStyleSheet("""
                                    QLineEdit{
                                    font-size: 20px
                                    }
                                    """)
        self.constEntry.setPlaceholderText(str(defaultConstPower) + " W")
        self.constEntry.setValidator(QDoubleValidator())
        self.constEntry.setEnabled(not self.started)

        constLayout.addWidget(self.constLabel)
        constLayout.addWidget(self.constEntry)

        # Stop Voltage Layout
        self.stopLabel = QLabel("<b style=\"font-size: 20px\">Stop Voltage:</b>")
        self.stopEntry = QLineEdit()
        self.stopEntry.setStyleSheet("""
                                    QLineEdit{
                                    font-size: 20px
                                    }
                                    """)
        self.stopEntry.setPlaceholderText(str(defaultStopVoltage) + " V")
        self.stopEntry.setValidator(QDoubleValidator())
        self.stopEntry.setEnabled(not self.started)

        stopLayout.addWidget(self.stopLabel)
        stopLayout.addWidget(self.stopEntry)

        # Button Layout
        self.startButton = QPushButton("START TEST")
        self.startButton.setEnabled(not self.started)
        #self.startButton.clicked.connect(self.startLot)
        self.startButton.setStyleSheet("""
                                            QPushButton {
                                            background-color: #00BB00;
                                            font-size: 35px;
                                            }
                                            QPushButton:hover {
                                            background-color: #00CB00;
                                            }
                                            QPushButton:pressed {
                                            background-color: #009B00;
                                            }
                                            """)
        self.endButton = QPushButton("END TEST")
        self.endButton.setEnabled(self.started)
        #self.endButton.clicked.connect(self.endLot)
        self.endButton.setStyleSheet("""
                                        QPushButton {
                                        background-color: #BB0000;
                                        font-size: 35px;
                                        }
                                        QPushButton:hover {
                                        background-color: #CB0000;
                                        }
                                        QPushButton:pressed {
                                        background-color: #9B0000;
                                        }
                                        """)

        buttonLayout.addWidget(self.startButton)
        buttonLayout.addWidget(self.endButton)

        window = QWidget()
        window.setLayout(vLayout)
        self.setCentralWidget(window)

    def updateUI(self):
        self.fileEntry.setEnabled(not self.started)
        self.testEntry.setEnabled(not self.started)
        self.limitEntry.setEnabled(not self.started)
        self.constEntry.setEnabled(not self.started)
        self.stopEntry.setEnabled(not self.started)

        self.startButton.setEnabled(not self.started)
        self.endButton.setEnabled(self.started)

    def connect(self):
        self.rm = pyvisa.ResourceManager()
        resources = self.rm.list_resources()
        if resources:
            self.inst = self.rm.open_resource(resources[0])
            return True
        else:
            return False

    def monitor(self, record):
        status = int(self.inst.query("STAT:QUES:COND?"))

        oc = bool(status & (1 << 1))
        ps = bool(status & (1 << 4))

        if oc and ps:
            self.recordVoltage = monitorVoltage
            self.recordCurrent = monitorCurrent
            self.recordPower = monitorPower
            return False

        else:
            v = float(self.inst.query("MEAS:VOLT?"))
            c = float(self.inst.query("MEAS:CURR?"))
            p = float(self.inst.query("FETC:POW?"))

            if v < self.stopVoltage:
                return False

            monitorVoltage = v
            monitorCurrent = c
            monitorPower = p

            if record:
                self.recordVoltage = v
                self.recordCurrent = c
                self.recordPower = p
        
        return True

    def initializeWorkbook(self):
        self.wb = Workbook()
        sheet = self.wb.active
        sheet.title = self.testEntry.text()
        sheet["A1"] = "Test ID"
        sheet["B1"] = "Current Limit"
        sheet["C1"] = "Constant Power"
        sheet["D1"] = "Stop Voltage"
        sheet["E1"] = "Time"
        sheet["F1"] = "Voltage"
        sheet["G1"] = "Current"
        sheet["H1"] = "Power"
        sheet["I1"] = "Status"
        self.wb.save(self.filePath)

    def record(self):
        sheet = self.wb.active
        sheet.append([
            self.testEntry.text(),
            float(self.currLimit),
            float(self.constPower),
            float(self.stopVoltage),
            datetime.now(),
            float(self.recordVoltage),
            float(self.recordCurrent),
            float(self.recordPower),
            str(self.recordStatus)
        ])
        self.wb.save(self.filePath)

    def start(self):
        self.started = True
        self.updateUI(self)
        
        if not self.connect(self):
            # Error Message
            self.started = False
            self.updateUI()
            return

        self.initializeWorkbook(self)

        self.inst.write("CURR:PROT:STAT ON")
        self.inst.write("CURR:PROT:LEV " + str(defaultCurrLimit))
        self.inst.write("CURR:PROT:DEL 0") # Delay TBD

        self.inst.write("FUNC POW")
        self.inst.write("POW " + str(defaultConstPower))

        self.inst.write("INP ON")
        flag = True
        last_monitor = time.monotonic()
        monitorCount = 0
        while flag:
            if time.monotonic() - last_monitor >= 30:
                flag = self.monitor(self, monitorCount % 4 == 0)
                monitorCount += 1
        self.end()
            

    def end(self):
        self.wb.close()
        self.inst.write("POW 0")
        self.inst.write("INP OFF")
        self.inst.write("SYST:LOC")
        self.inst.close()
        self.rm.close()
        self.started = False

    def closeEvent(self, event):
        if self.started:
            msg = QMessageBox(self)
            msg.setWindowTitle("End Test?")
            msg.setText("Close and save test?")

            saveButton = msg.addButton("Save Test", QMessageBox.ButtonRole.AcceptRole)
            discardButton = msg.addButton("Delete Test", QMessageBox.ButtonRole.DestructiveRole)
            cancelButton = msg.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)

            msg.exec()

            if msg.clickedButton() == saveButton:
                self.end()
                save_settings()
                event.accept()
            elif msg.clickedButton() == discardButton:
                save_settings()
                event.accept()
            elif msg.clickedButton() == cancelButton:
                event.ignore()
        else:
            save_settings()
            event.accept()

load_settings()
app = QApplication([])
window = MainWindow()
window.show()
app.exec()