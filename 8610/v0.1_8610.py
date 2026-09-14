
# Old tester compilation command
# pyinstaller --onefile --windowed --name="BatteryTester" guimin.py

import pyvisa
import time


currLimit = 4.0
constPower = 62.0
stopVoltage = 0.1

rm = pyvisa.ResourceManager()

resources = rm.list_resources()
print(resources)

inst = rm.open_resource(resources[0])

print("Instrument:", inst.query("*IDN?"))

# Set current protection limit

inst.write("CURR:PROT:STAT ON")
inst.write("CURR:PROT:LEV " + str(currLimit))

# Verify the setting
ocp_state = inst.query("CURR:PROT:STAT?")
ocp_level = inst.query("CURR:PROT:LEV?")

print("OCP state:", ocp_state)
print("OCP level:", ocp_level)

# Set constant-power mode

inst.write("FUNC POW")
inst.write("POW " + str(constPower))

# Turn load on

inst.write("INP ON")

time.sleep(3)

# Monitor 

voltage = float(inst.query("MEAS:VOLT?"))
current = float(inst.query("MEAS:CURR?"))
power = float(inst.query("FETC:POW?"))

print(f"Voltage: {voltage:.3f} V | "f"Current: {current:.3f} A | "f"Power: {power:.3f} W")

inst.write("POW 0")
inst.write("INP OFF")
inst.write("SYST:LOC")
inst.close()
rm.close()



# ====================================

currLimit = 4.0
constPower = 62.0
stopVoltage = 0.1

monitorVoltage = 0.0
monitorCurrent = 0.0
monitorPower = 0.0

recordVoltage = 0.0
recordCurrent = 0.0
recordPower = 0.0

def monitor(record):
    status = int(inst.query("STAT:QUES:COND?"))

    oc = bool(status & (1 << 1))
    ps = bool(status & (1 << 4))

    if oc and ps:
        recordVoltage = monitorVoltage
        recordCurrent = monitorCurrent
        recordPower = monitorPower

    else:
        v = float(inst.query("MEAS:VOLT?"))
        c = float(inst.query("MEAS:CURR?"))
        p = float(inst.query("FETC:POW?"))

        monitorVoltage = v
        monitorCurrent = c
        monitorPower = p

        if record:
            recordVoltage = v
            recordCurrent = c
            recordPower = p

def loadOFF():
    inst.write("POW 0")
    inst.write("INP OFF")
    inst.write("SYST:LOC")
    inst.close()
    rm.close()