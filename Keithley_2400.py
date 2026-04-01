import csv
import matplotlib.pyplot as plt
import serial
import time
import pandas as pd

from pathlib import Path
from serial.tools import list_ports
from typing import Optional


def list_serial_ports():
    ports = sorted(list_ports.comports(), key=lambda p: p.device)
    for p in ports:
        print(f"{p.device} | {p.description} | {p.hwid}")

def keithley_port():
    ports = sorted(list_ports.comports(), key=lambda p: p.device)
    try:
        for p in ports:
            #print(p.description)
            if 'ATEN' in p.description:
                name = str(p.device)
        return name
    except: print("Chopper don't connected")

def plot_iv_curve(
    data: list[dict[str, float]],
    title: str = "I(V)",
    xlabel: str = "Voltage, V",
    ylabel: str = "Current, A",
    show: bool = True,
    save_path: str | None = None,
    marker: str = "o",
    linestyle: str = "-",
    grid: bool = True,
    ):
    
    if not data:
        raise ValueError("Пустой массив данных: нечего строить.")

    x = [float(point["voltage"]) for point in data]
    y = [float(point["current"]) for point in data]

    plt.figure(figsize=(7, 5))
    plt.plot(x, y, marker=marker, linestyle=linestyle)
    plt.title(title)
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)

    if grid:
        plt.grid(True)

    plt.tight_layout()



class Keithley2400:
    # -----------------------------
    # -----LOW LEWEL OPERATIONS----
    # -----------------------------
    def __enter__(self):
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def __init__(
            self, port, 
            baudrate = 9600, 
            timeout = 2.0, 
            bytesize = serial.EIGHTBITS, 
            parity = serial.PARITY_NONE, 
            stopbits = serial.STOPBITS_ONE, 
            xonxoff = False, 
            rtscts = False, 
            dsrdtr = False, 
            write_termination = "\r", 
            read_termination = "\r", 
            encoding = "ascii"):

        self.port = port
        self.baudrate = baudrate
        self.timeout = timeout
        self.bytesize = bytesize
        self.parity = parity
        self.stopbits = stopbits
        self.xonxoff = xonxoff
        self.rtscts = rtscts
        self.dsrdtr = dsrdtr
        self.write_termination = write_termination
        self.read_termination = read_termination
        self.encoding = encoding
        self._ser: Optional[serial.Serial] = None

    def _ensure_connected(self):
        if self._ser is None or not self._ser.is_open:
            raise ConnectionError("Порт не открыт. Сначала вызови connect().")
        return self._ser

    def close(self):
        if self._ser is not None and self._ser.is_open:
            self._ser.close()

    def connect(self):

        if self._ser is not None and self._ser.is_open:
            return

        self._ser = serial.Serial(
            port=self.port,
            baudrate=self.baudrate,
            timeout=self.timeout,
            bytesize=self.bytesize,
            parity=self.parity,
            stopbits=self.stopbits,
            xonxoff=self.xonxoff,
            rtscts=self.rtscts,
            dsrdtr=self.dsrdtr)

        time.sleep(0.2)

        self._ser.reset_input_buffer()
        self._ser.reset_output_buffer()

    def query(self, command: str, delay = 0.1):
        self.write(command)
        if delay > 0:
            time.sleep(delay)
        return self.read_line()

    def _query_float(self, command: str, delay: float = 0.1):
        response = self.query(command, delay=delay).strip()
        first_value = response.split(",")[0].strip()
        return float(first_value)

    def read_line(self):
        ser = self._ensure_connected()
        raw = ser.read_until(self.read_termination.encode(self.encoding))
        if not raw:
            raise TimeoutError("Timeout from Keithley 2400")
        return raw.decode(self.encoding, errors="replace").strip()

    def write(self, command: str):
        ser = self._ensure_connected()
        full_command = f"{command}{self.write_termination}"
        ser.write(full_command.encode(self.encoding))
        ser.flush()


    # -----------------------------
    # -------BASIC OPERATIONS------
    # -----------------------------
    def check_source_status(self, res = True):
        if self.query(':OUTP?') == '0': res = False
        elif self.query(':OUTP?') == '1': res = True
        else: res = 'SOMETHING WRONG'
        return res
    
    def clear_status(self):
        self.write("*CLS")

    def identify(self):
        return self.query("*IDN?")
    
    def reset(self):
        self.write("*RST")

    def output_on(self):
        self.write(':OUTP ON')

    def output_off(self):
        self.write(':OUTP OFF')

    def check_source(self):
        return self.query(':ROUT:TERM?')
    
    def select_source(self, source='front'):
        if (self.query(':ROUT:TERM?') == 'FRON') and (source == 'rear'):
            self.write(':ROUT:TERM REAR')

        elif (self.query(':ROUT:TERM?') == 'REAR') and (source == 'front'):
            self.write(':ROUT:TERM FRON')

        else: self.write(':ROUT:TERM FRON')


    # -----------------------------
    # ------SOURCE OPERATIONS------
    # -----------------------------
    def set_current_source(self, current, compliance_voltage = 10.0, turn_on = True):
        self.write(":SOUR:FUNC CURR")
        self.write(":SOUR:CURR:MODE FIX")
        self.write(":SOUR:CURR:RANG:AUTO ON")

        self.write(':SENS:FUNC "VOLT"')
        self.write(":SENS:VOLT:RANG:AUTO ON")
        self.write(f":SENS:VOLT:PROT {compliance_voltage}")

        self.write(f":SOUR:CURR:LEV {current}")

        if turn_on:
            self.output_on()

    def set_voltage_source(self, voltage, compliance_current = 0.1, turn_on = True):
        self.write(':SOUR:FUNC VOLT')
        self.write(":SOUR:VOLT:MODE FIX")
        self.write(":SOUR:VOLT:RANG:AUTO ON")

        self.write(':SENS:FUNC "CURR"')
        self.write(":SENS:CURR:RANG:AUTO ON")
        self.write(f":SENS:CURR:PROT {compliance_current}")

        self.write(f":SOUR:VOLT:LEV {voltage}")

        if turn_on:
            self.output_on()

    def set_current_point(self, current, compliance = 10.0, settle_time = 0.2):
        
        self.set_current_source(current=current, compliance=compliance, turn_on=True)
        
        time.sleep(settle_time)

        result = self.measure_iv(delay=0.05)
        result["source_mode"] = "current"
        result["set_value"] = current
        return result

    def set_voltage_point(self, voltage, compliance_current = 0.01, settle_time = 0.2):
        
        self.set_voltage_source(voltage=voltage, compliance_current=compliance_current, turn_on=True)

        time.sleep(settle_time)
        result = self.measure_iv(delay=0.05)
        result["source_mode"] = "voltage"
        result["set_value"] = voltage
        return result

    def measure_current(self, delay: float = 0.1):
        self.write(':SENS:FUNC "CURR"')
        self.write(":FORM:ELEM CURR")
        return self._query_float(":READ?", delay=delay)
    
    def measure_voltage(self, delay: float = 0.1):
        self.write(':SENS:FUNC "VOLT"')
        self.write(":FORM:ELEM VOLT")
        return self._query_float(":READ?", delay=delay)

    def measure_iv(self, delay: float = 0.1):
        voltage = self.measure_voltage(delay=delay)
        current = self.measure_current(delay=delay)
        return {
            "voltage": voltage,
            "current": current}


    # -----------------------------
    # -------SWEEP OPERATIONS------
    # -----------------------------
    @staticmethod
    def _build_sweep_values(start, stop, step):
        if step == 0:
            raise ValueError("step не может быть равен 0.")

        values: list[float] = []
        x = start
        tol = abs(step) * 1e-9 + 1e-15

        if step > 0:
            while x <= stop + tol:
                values.append(round(x, 12))
                x += step
        else:
            while x >= stop - tol:
                values.append(round(x, 12))
                x += step

        return values
    
    def save_iv_to_csv(self, data, file_path, delimiter = ","):
    
        if not data:
            raise ValueError("Пустой массив данных: нечего сохранять в CSV.")

        path = Path(file_path)
        path.parent.mkdir(parents=True, exist_ok=True)

        # Собираем все ключи, чтобы не потерять поля
        fieldnames: list[str] = []
        for row in data:
            for key in row.keys():
                if key not in fieldnames:
                    fieldnames.append(key)

        with path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=delimiter)
            writer.writeheader()
            writer.writerows(data)

    def sweep_iv_by_voltage(self, start, stop, step, compliance_current = 0.01, settle_time = 1e-3, output_off_after = True):
        values = self._build_sweep_values(start, stop, step)
        data: list[dict[str, float]] = []

        self.set_voltage_source(voltage=values[0], compliance_current=compliance_current, turn_on=True)

        for set_voltage in values:
            self.write(f":SOUR:VOLT:LEV {set_voltage}")
            time.sleep(settle_time)

            point = self.measure_iv(delay=0.05)
            point["source_mode"] = "voltage"
            point["set_value"] = set_voltage
            data.append(point)

        if output_off_after:
            self.output_off()

        return data

    
    def sweep_iv_by_current(self, start, stop, step, compliance_voltage = 10.0, settle_time = 1e-3, output_off_after = True):
        
        values = self._build_sweep_values(start, stop, step)
        data: list[dict[str, float]] = []

        self.set_current_source(
            current=values[0],
            compliance_voltage=compliance_voltage,
            turn_on=True,
        )

        for set_current in values:
            self.write(f":SOUR:CURR:LEV {set_current}")
            time.sleep(settle_time)

            point = self.measure_iv(delay=0.05)
            point["source_mode"] = "current"
            point["set_value"] = set_current
            data.append(point)

        if output_off_after:
            self.output_off()

        return data
    


