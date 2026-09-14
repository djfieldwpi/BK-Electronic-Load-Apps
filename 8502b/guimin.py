
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QWidget,
                             QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QSizePolicy,
                             QFileDialog, QStyle, QDialog, QMessageBox, QCheckBox)
from PyQt6.QtCore import QSize, Qt
from PyQt6.QtGui import QDoubleValidator, QIntValidator
from pathlib import Path
from openpyxl import Workbook, load_workbook
import json, serial, time, threading

CONFIG_FILE = Path.home() / "BatteryTester" / "settings.json"
CONFIG_FILE.parent.mkdir(exist_ok=True)

def load_settings():
    global loadAmp, initialVoltageMin, avgLoadVoltageMin, defaultPath, comPort, ccvTime, debugCheck

    if not CONFIG_FILE.exists():
        return

    with open(CONFIG_FILE, "r") as f:
        data = json.load(f)

    loadAmp = data.get("loadAmp", 0.0)
    initialVoltageMin = data.get("initialVoltageMin", 0.0)
    avgLoadVoltageMin = data.get("avgLoadVoltageMin", 0.0)
    defaultPath = data.get("defaultPath", "")
    comPort = data.get("comPort", 0)
    ccvTime = data.get("ccvTime", 3.0)
    debugCheck = data.get("debugCheck", False)

def save_settings():
    data = {
        "loadAmp": loadAmp,
        "initialVoltageMin": initialVoltageMin,
        "avgLoadVoltageMin": avgLoadVoltageMin,
        "defaultPath": defaultPath,
        "comPort": comPort,
        "ccvTime": ccvTime,
        "debugCheck": debugCheck
    }

    with open(CONFIG_FILE, "w") as f:
        json.dump(data, f, indent=4)

password = "password"
loadAmp = 0.0
initialVoltageMin = 0.0
avgLoadVoltageMin = 0.0
comPort = 0
defaultPath = ""
ccvTime = 0.0
debugCheck = False

class SetOption(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Set Parameters")
        self.setFixedSize(QSize(400,300))

        formLayout = QVBoxLayout()
        portLayout = QHBoxLayout()
        pathLayout = QHBoxLayout()
        ampLayout = QHBoxLayout()
        initialLayout = QHBoxLayout()
        averageLayout = QHBoxLayout()
        timeLayout = QHBoxLayout()
        saveLayout = QHBoxLayout()
        debugLayout = QHBoxLayout()
        formLayout.addLayout(pathLayout)
        formLayout.addLayout(portLayout)
        formLayout.addLayout(ampLayout)
        formLayout.addLayout(initialLayout)
        formLayout.addLayout(averageLayout)
        formLayout.addLayout(timeLayout)
        formLayout.addLayout(debugLayout)
        formLayout.addLayout(saveLayout)

        self.pathLabel = QLabel("Default File Path:")
        self.pathEntry = QLineEdit()
        self.pathEntry.setPlaceholderText(str(defaultPath))
        pathLayout.addWidget(self.pathLabel)
        pathLayout.addWidget(self.pathEntry)

        self.portLabel = QLabel("Com Port Number:")
        self.portEntry = QLineEdit()
        self.portEntry.setValidator(QIntValidator())
        self.portEntry.setPlaceholderText(str(comPort))
        portLayout.addWidget(self.portLabel)
        portLayout.addWidget(self.portEntry)

        self.ampLabel = QLabel("Constant Current Amps:")
        self.ampEntry = QLineEdit()
        self.ampEntry.setMaximumWidth(100)
        self.ampEntry.setPlaceholderText(str(loadAmp))
        self.ampEntry.setValidator(QDoubleValidator())
        ampLayout.addWidget(self.ampLabel)
        ampLayout.addWidget(self.ampEntry)

        self.initialLabel = QLabel("OCV Min:")
        self.initialEntry = QLineEdit()
        self.initialEntry.setMaximumWidth(100)
        self.initialEntry.setPlaceholderText(str(initialVoltageMin))
        self.initialEntry.setValidator(QDoubleValidator())
        initialLayout.addWidget(self.initialLabel)
        initialLayout.addWidget(self.initialEntry)

        self.averageLabel = QLabel("CCV Min:")
        self.averageEntry = QLineEdit()
        self.averageEntry.setMaximumWidth(100)
        self.averageEntry.setPlaceholderText(str(avgLoadVoltageMin))
        self.averageEntry.setValidator(QDoubleValidator())
        averageLayout.addWidget(self.averageLabel)
        averageLayout.addWidget(self.averageEntry)

        self.timeLabel = QLabel("CCV Time")
        self.timeEntry = QLineEdit()
        self.timeEntry.setMaximumWidth(100)
        self.timeEntry.setPlaceholderText(str(ccvTime))
        self.timeEntry.setValidator(QDoubleValidator())
        timeLayout.addWidget(self.timeLabel)
        timeLayout.addWidget(self.timeEntry)

        self.debugLabel = QLabel("Enable CCV Dump?")
        self.debugEntry = QCheckBox()
        self.debugEntry.setChecked(debugCheck)
        self.debugLabel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        debugLayout.addWidget(self.debugLabel)
        debugLayout.addWidget(self.debugEntry)

        self.passwordEntry = QLineEdit()
        self.passwordEntry.setPlaceholderText("Enter Password")
        self.passwordEntry.setEchoMode(QLineEdit.EchoMode.Password)
        self.saveButton = QPushButton("Save")
        self.saveButton.clicked.connect(self.on_save)
        saveLayout.addWidget(self.passwordEntry)
        saveLayout.addWidget(self.saveButton)

        self.setLayout(formLayout)

    def on_save(self):
        global loadAmp, initialVoltageMin, avgLoadVoltageMin, defaultPath, comPort, ccvTime, debugCheck

        if self.passwordEntry.text() != password:
            self.passwordEntry.clear()
            self.passwordEntry.setPlaceholderText("Password Incorrect")
            return

        
        if self.timeEntry.text().strip():
            temp = float(self.timeEntry.text())
            if temp >= 3.0:
                ccvTime = temp
            else:
                self.passwordEntry.clear()
                self.passwordEntry.setPlaceholderText("CCV Time set too low.")
                return

        if self.pathEntry.text().strip():
            defaultPath = self.pathEntry.text()
        if self.ampEntry.text().strip():
            loadAmp = float(self.ampEntry.text())
        if self.initialEntry.text().strip():
            initialVoltageMin = float(self.initialEntry.text())
        if self.averageEntry.text().strip():
            avgLoadVoltageMin = float(self.averageEntry.text())
        if self.portEntry.text().strip():
            comPort = int(self.portEntry.text())
        
        debugCheck = self.debugEntry.isChecked()

        save_settings()
        self.accept()

        

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.LotStarted = False
        self.BatteryTesting = False

        self.wb = None
        self.filePath = ""
        
        self.wbDUMP = None
        self.filePathDUMP = ""

        self.initialVoltageRec = 0.0
        self.avgLoadVoltageRec = 0.0
        self.batteryNum = 1
        self.goodCount = 0
        self.badCount = 0
        self.testPassed = False
        self.polls = 0

        self.ser = None
        self.confirm = 0x00

        self.setWindowTitle("Battery Load Tester")
        self.setFixedSize(QSize(800,250))

        # Overall layout
        vLayout = QVBoxLayout()

        # Horizontal Boxes
        fileLayout = QHBoxLayout()
        buttonLayout = QHBoxLayout()
        signalLayout = QHBoxLayout()
        tallyLayout = QHBoxLayout()
        vLayout.addLayout(fileLayout)
        vLayout.addLayout(buttonLayout)
        vLayout.addLayout(signalLayout)
        vLayout.addLayout(tallyLayout)

        # File Layout
        self.fileLabel = QLabel("<b style=\"font-size: 20px\">File Name:</b>")
        self.fileEntry = QLineEdit()
        self.fileEntry.setStyleSheet("""
                                    QLineEdit{
                                    font-size: 20px
                                    }
                                    """)
        self.fileEntry.setText(defaultPath)
        self.fileEntry.setEnabled(not self.LotStarted)
        self.fileButton = QPushButton()
        self.fileButton.setFixedHeight(40)
        self.fileButton.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton))
        self.fileButton.setMaximumWidth(40)
        self.fileButton.setEnabled(not self.LotStarted)
        self.fileButton.clicked.connect(self.findExistingFile)
        self.pollLabel = QLabel("<b style=\"font-size: 20px\">0</b>")
        self.pollLabel.setMaximumWidth(40)

        self.lotLabel = QLabel("<b style=\"font-size: 20px\">Lot Number:</b>")
        self.lotEntry = QLineEdit()
        self.lotEntry.setStyleSheet("""
                                    QLineEdit{
                                    font-size: 20px
                                    }
                                    """)
        self.lotEntry.setPlaceholderText("Lot #...")
        self.lotEntry.setMaximumWidth(200)
        self.lotEntry.setEnabled(not self.LotStarted)

        self.settingsButton = QPushButton()
        self.settingsButton.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon))
        self.settingsButton.setFixedHeight(40)
        self.settingsButton.clicked.connect(self.openSettings)
        self.settingsButton.setEnabled(not self.LotStarted)

        fileLayout.addWidget(self.fileLabel)
        fileLayout.addWidget(self.fileEntry)
        fileLayout.addWidget(self.fileButton)
        fileLayout.addWidget(self.lotLabel)
        fileLayout.addWidget(self.lotEntry)
        fileLayout.addWidget(self.settingsButton)
        fileLayout.addWidget(self.pollLabel)

        # Button Layout
        self.startLotButton = QPushButton("START LOT")
        self.startLotButton.setEnabled(not self.LotStarted)
        self.startLotButton.clicked.connect(self.startLot)
        self.startLotButton.setStyleSheet("""
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
        self.testButton = QPushButton("TEST BATTERY")
        self.testButton.setEnabled(self.LotStarted and not self.BatteryTesting)
        self.testButton.clicked.connect(self.testBattery)
        self.testButton.setStyleSheet("""
                                      font-size: 35px;
                                      """)
        self.endLotButton = QPushButton("END LOT")
        self.endLotButton.setEnabled(self.LotStarted)
        self.endLotButton.clicked.connect(self.endLot)
        self.endLotButton.setStyleSheet("""
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

        buttonLayout.addWidget(self.startLotButton)
        buttonLayout.addWidget(self.testButton)
        buttonLayout.addWidget(self.endLotButton)

        # Signal Layout
        self.signalLabel = QLabel("<h1 style=\"text-align: center\">Awaiting Lot Start...</h1>")
        signalLayout.addWidget(self.signalLabel)

        # Tally Layout
        self.lotNumberLabel = QLabel(f"<b style=\"font-size: 20px; text-align: center\">Lot Number: {str(self.lotEntry.text())}</b></div>")
        self.goodLabel = QLabel(f"<div style=\"font-size: 20px; text-align: center; width: 100%;\"><b>Good Count: {str(self.goodCount)}<br>OCV: {str(self.initialVoltageRec)}</b></div>")
        self.badLabel = QLabel(f"<div style=\"font-size: 20px; text-align: center; width: 100%;\"><b>Bad Count: {str(self.badCount)}<br>CCV: {str(self.avgLoadVoltageRec)}</b></div>")

        tallyLayout.addWidget(self.lotNumberLabel)
        tallyLayout.addWidget(self.goodLabel)
        tallyLayout.addWidget(self.badLabel)
        #, alignment=Qt.AlignmentFlag.AlignCenter
        window = QWidget()
        window.setLayout(vLayout)
        self.setCentralWidget(window)

    def findExistingFile(self):
        file_path= QFileDialog.getExistingDirectory(
        self,
        "Select or Create Folder",
        ""
        )

        if not file_path:
            return  # user cancelled

        # Show selected path in UI
        self.fileEntry.setText(file_path)
    
    def openSettings(self):
        dialog = SetOption(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.updateUI()

    def updateUI(self):
        self.fileEntry.setEnabled(not self.LotStarted)
        self.fileButton.setEnabled(not self.LotStarted)
        self.lotEntry.setEnabled(not self.LotStarted)
        self.settingsButton.setEnabled(not self.LotStarted)

        self.startLotButton.setEnabled(not self.LotStarted)
        self.testButton.setEnabled(self.LotStarted and not self.BatteryTesting)
        self.endLotButton.setEnabled(self.LotStarted and not self.BatteryTesting)

        self.fileEntry.setText(defaultPath)
        self.lotNumberLabel.setText(f"<div style=\"font-size: 20px; text-align: center; width: 100%;\"><b>Lot Number: {str(self.lotEntry.text())}</b></div>")
        self.goodLabel.setText(f"<div style=\"font-size: 20px; text-align: center; width: 100%;\"><b>Good Count: {str(self.goodCount)}<br>OCV: {str(self.initialVoltageRec)}</b></div>")
        self.badLabel.setText(f"<div style=\"font-size: 20px; text-align: center; width: 100%;\"><b>Bad Count: {str(self.badCount)}<br>CCV: {str(self.avgLoadVoltageRec)}</b></div>")
        self.pollLabel.setText(f"<div style=\"font-size: 20px;\">{str(self.polls)}</div>")
    
    def csum(self, command):                          # this function automatically calculates the checksum
        sum = 0
        for i in range((len(command) - 1)):     
            sum+= command[i]                    # this loop sums the all the bytes of the input except for the last one
        return (0xFF & sum)                     # because some sums may overflow two bytes, we shorten the sum to two bytes

    def Printer(self, read):                          # this function makes the 26 byte format more readable
        x = " "        
        for y in range(len(read)):
            x+=" "
            x+=hex(read[y]).replace('0x','')    # replaces the 0x with a spacing instead
        print(x)                                # and prints it

    def Command(self, com):                           # We will use this function to send commands and receive replies    
        com[0] = 0xAA                           # Automatically sets the start bit
        com[25] = self.csum(com)                     # Automatically places the checksum
        print("Command Sent:\t\t",end=' ')
        self.Printer(com)                            # Prints the command in a readable format
        self.ser.write(com)                          # Writes the command
        print("Reponse Received:\t",end=' ')     
        resp = self.ser.read(26)
        if(resp == self.confirm):
            print("\tNo Error")                 # If we receive the no error string, print "No Error"
        else:
            self.Printer(resp)                       # Else Print the String 
        print("\n")
    def read_display(self):
        cmd = [0] * 26
        cmd[0] = 0xAA
        cmd[2] = 0x5F
        cmd[25] = self.csum(cmd)

        self.ser.reset_input_buffer()  # important for debugging stability
        self.ser.write(bytes(cmd))

        # --- SYNC TO START BYTE ---
        while True:
            b = self.ser.read(1)
            if b == b'\xAA':
                break

        rest = self.ser.read(25)
        resp = b'\xAA' + rest

        voltage = (
            resp[3]
            | (resp[4] << 8)
            | (resp[5] << 16)
            | (resp[6] << 24)
        ) / 1000.0

        return voltage

    def startLot(self):
        self.updateUI()

        self.filePath = self.fileEntry.text().strip() + "\\lot" + self.lotEntry.text() + ".xlsx"
        folder = Path(self.fileEntry.text().strip())

        if debugCheck:
            self.filePathDUMP = self.fileEntry.text().strip() + "\\lot" + self.lotEntry.text() + "CCVdump.xlsx"

        if not folder.exists():
            folder.mkdir(parents = True, exist_ok=True)

        if not self.filePath:
            self.signalLabel.setText("<h1 style=\"text-align: center\">No workbook selected.</h1>")
            self.updateUI()
            return
        
        if Path(self.filePath).exists():
            self.signalLabel.setText("<h1 style=\"text-align: center\">Lot Number Already Exists.</h1>")
            self.updateUI()
            return
        else:
            self.wb = Workbook()
            sheet = self.wb.active
            sheet.title = "Lot " + str(self.lotEntry.text())
            sheet["A1"] = "Lot Number"
            sheet["B1"] = "Battery Number"
            sheet["C1"] = "OCV"
            sheet["D1"] = "OCV Min"
            sheet["E1"] = "CCV Low"
            sheet["F1"] = "CCV Low Min"
            sheet["G1"] = "Load Amps"
            sheet["H1"] = "Status"

            if debugCheck:
                self.wbDUMP = Workbook()
                sheet = self.wbDUMP.active
                sheet.title = "Lot " + str(self.lotEntry.text()) + " CCV Dump"
                sheet["A1"] = "Lot Number"
                sheet["B1"] = "Battery Number"
                sheet["C1"] = "CCV Time"
                sheet["D1"] = "CCV Read Count"
                sheet["E1"] = "CCV Reads..."
        
        self.signalLabel.setText("<h1 style=\"text-align: center\">Connecting...</h1>")
        self.updateUI()

        self.ser = serial.Serial('COM' + str(comPort), 9600, timeout=1)
        #self.ser.flush()

        cmd=[0xAA,0,0x20,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0xCB]
        
        self.ser.write(cmd)
        self.confirm = self.ser.read()
        self.ser.reset_input_buffer()

        if self.confirm != b'\xAA':
            self.signalLabel.setText("<h1 style=\"text-align: center\">Could not connect.</h1>")
            self.updateUI()
            return
        
        self.LotStarted = True
        self.signalLabel.setText("<h1 style=\"text-align: center\">Opening Excel Sheet.</h1>")
        self.updateUI()
        
        self.signalLabel.setText("<h1 style=\"text-align: center\">Attach Battery.</h1>")
        self.updateUI()

    def bkOperations(self):
        self.initialVoltageRec = 0.0
        self.avgLoadVoltageRec = 0.0
        self.BatteryTesting = True
        self.updateUI()

        # CC mode
        cmd = [0] * 26
        cmd[2] = 0x28
        cmd[3] = 0
        self.Command(cmd)

        # Set Current
        current = int(loadAmp * 10000)

        cmd = [0] * 26
        cmd[2] = 0x2A
        cmd[3] = current & 0xFF
        cmd[4] = (current >> 8) & 0xFF
        cmd[5] = (current >> 16) & 0xFF
        cmd[6] = (current >> 24) & 0xFF
        self.Command(cmd)

        self.signalLabel.setText("<h1 style=\"text-align: center\">Reading Initial Voltage.</h1>")
        self.updateUI()

        # Read before load on
        self.initialVoltageRec = self.read_display()

        
        samples = []

        if self.initialVoltageRec >= initialVoltageMin:

            # Turn load on
            cmd = [0] * 26
            cmd[2] = 0x21
            cmd[3] = 1
            self.Command(cmd)

            self.signalLabel.setText("<h1 style=\"text-align: center\">Testing CCV.</h1>")
            self.updateUI()

            """
            time.sleep(ccvTime - 2.0)
            avg1 = self.read_display()
            print(f"11 sec: {avg1:.3f} V")

            time.sleep(1)
            avg2 = self.read_display()
            print(f"11 sec: {avg2:.3f} V")

            time.sleep(1)
            avg3 = self.read_display()
            print(f"11 sec: {avg3:.3f} V")

            self.avgLoadVoltageRec = (avg1 + avg2 + avg3) / 3.0
            self.avgLoadVoltageRec = round(self.avgLoadVoltageRec, 3)
            """
            end_time = 0.0
            if debugCheck:
                end_time = time.monotonic() + ccvTime
            else:
                time.sleep(ccvTime - 3.0)

                end_time = time.monotonic() + 3.0

            while time.monotonic() < end_time:
                try:
                    voltage = self.read_display()
                    samples.append(voltage)
                except Exception as e:
                    print(e)

            if samples:
                self.avgLoadVoltageRec = round(min(samples), 3)
                # self.avgLoadVoltageRec = round(sum(samples) / len(samples), 3)
                print(len(samples))
                print(samples)
                self.polls = len(samples)
                self.updateUI()
            else:
                self.avgLoadVoltageRec = 0.0
                self.signalLabel.setText("<h1 style=\"text-align: center\">Testing CCV Failed.</h1>")
                self.updateUI()

            # Turn load off
            cmd = [0] * 26
            cmd[2] = 0x21
            cmd[3] = 0
            self.Command(cmd)

        if self.initialVoltageRec >= initialVoltageMin and self.avgLoadVoltageRec >= avgLoadVoltageMin:
            self.testPassed = True
        
        sheet = self.wb.active
        sheet.append([
            self.lotEntry.text(),
            int(self.batteryNum),
            float(self.initialVoltageRec),
            float(initialVoltageMin),
            float(self.avgLoadVoltageRec),
            float(avgLoadVoltageMin),
            float(loadAmp),
            self.testPassed
        ])

        if debugCheck:
            sheet = self.wbDUMP.active
            sheet.append([
                self.lotEntry.text(),
                int(self.batteryNum), 
                ccvTime, len(samples), 
                *samples
            ])

        if self.testPassed:
            self.signalLabel.setText(f"<h1 style=\"text-align: center; color: #00BB00\">{self.batteryNum}: Passed - Attach New Battery.</h1>")
        else:
            if self.avgLoadVoltageRec == 0.0:
                self.signalLabel.setText(f"<h1 style=\"text-align: center; color: #BB0000\">{self.batteryNum}: Failed OCV - Attach New Battery.</h1>")
            else:
                self.signalLabel.setText(f"<h1 style=\"text-align: center; color: #BB0000\">{self.batteryNum}: Failed CCV - Attach New Battery.</h1>")
        self.updateUI()

        if self.testPassed:
            self.goodCount = self.goodCount + 1
        else:
            self.badCount = self.badCount + 1
        self.batteryNum = self.batteryNum + 1
        self.BatteryTesting = False
        self.updateUI()
    
    def testBattery(self):
        if self.BatteryTesting:
            return
        
        self.testPassed = False

        threading.Thread(target=self.bkOperations, daemon=True).start()

    def endLot(self):
        self.wb.save(self.filePath)
        self.wb = None
        if debugCheck:
            self.wbDUMP.save(self.filePathDUMP)
            self.wbDUMP = None
        self.LotStarted = False
        self.signalLabel.setText("<h1 style=\"text-align: center\">Awaiting Lot Start...</h1>")
        self.fileEntry.clear()
        self.lotEntry.clear()
        self.batteryNum = 1
        self.goodCount = 0
        self.badCount = 0
        self.ser = None
        self.updateUI()
    
    def closeEvent(self, event):
        if self.LotStarted:
            # Create the message box
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("End Lot?")
            msg_box.setText("Save to End Lot")
            
            # Add buttons
            save_button = msg_box.addButton("Save", QMessageBox.ButtonRole.AcceptRole)
            discard_button = msg_box.addButton("Delete Lot", QMessageBox.ButtonRole.DestructiveRole)
            cancel_button = msg_box.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
            
            msg_box.exec()

            # Handle the user's choice
            if msg_box.clickedButton() == save_button:
                self.endLot()
                event.accept() # Close the window
            elif msg_box.clickedButton() == discard_button:
                event.accept() # Close the window without saving
            elif msg_box.clickedButton() == cancel_button:
                event.ignore() # Cancel the close event and return to the app
        else:
            event.accept() # No changes, just close

    

load_settings()
app = QApplication([])
window = MainWindow()
window.show()

app.exec()
