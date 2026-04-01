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


class Keithley2400:

    def __init__(
        self,
        port: str,
        baudrate: int = 9600,
        timeout: float = 2.0,
        bytesize: int = serial.EIGHTBITS,
        parity: str = serial.PARITY_NONE,
        stopbits: float = serial.STOPBITS_ONE,
        xonxoff: bool = False,
        rtscts: bool = False,
        dsrdtr: bool = False,
        write_termination: str = "\r",
        encoding: str = "ascii",
        open_delay: float = 0.5,
        query_delay: float = 0.05):
        
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
        self.encoding = encoding
        self.open_delay = open_delay
        self.query_delay = query_delay

        self._ser: Optional[serial.Serial] = None

    # ------------------------------------------------------------------
    # БАЗОВАЯ РАБОТА С ПОРТОМ
    # ------------------------------------------------------------------
    def connect(self, do_sync: bool = True, verify: bool = False):

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
            dsrdtr=self.dsrdtr,
        )

        time.sleep(self.open_delay)
        self.clear_buffers()

        if do_sync:
            self.sync()

        if verify:
            _ = self.identify()

    def close(self, try_output_off: bool = False):

        if self._ser is None or not self._ser.is_open:
            return

        if try_output_off:
            try:
                self.output_off()
                time.sleep(0.05)
            except Exception:
                pass

        try:
            self.clear_buffers()
        except Exception:
            pass

        self._ser.close()

    def __enter__(self):
        self.connect(do_sync=True, verify=False)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close(try_output_off=False)

    def _ensure_connected(self):
        
        if self._ser is None or not self._ser.is_open:
            raise ConnectionError("Port is close. Before use connect().")
        return self._ser

    def clear_buffers(self):

        ser = self._ensure_connected()
        ser.reset_input_buffer()
        ser.reset_output_buffer()

    def sync(self) -> None:

        ser = self._ensure_connected()

        ser.reset_input_buffer()
        ser.reset_output_buffer()

        # Send Ctrl-C
        ser.write(b"\x03")
        ser.flush()
        time.sleep(0.2)

        ser.reset_input_buffer()
        ser.reset_output_buffer()


    # ------------------------------------------------------------------
    # НИЗКОУРОВНЕВЫЙ ОБМЕН
    # ------------------------------------------------------------------
    def write(self, command: str):
        ser = self._ensure_connected()
        full_command = f"{command}{self.write_termination}"
        ser.write(full_command.encode(self.encoding))
        ser.flush()

    def _read_response(
        self,
        timeout: Optional[float] = None,
        terminators: tuple[bytes, ...] = (b"\r\n", b"\n", b"\r")):

        ser = self._ensure_connected()
        effective_timeout = self.timeout if timeout is None else timeout

        deadline = time.monotonic() + effective_timeout
        data = bytearray()

        while True:
            if time.monotonic() > deadline:
                if data:
                    break
                raise TimeoutError("Таймаут ожидания ответа от Keithley 2400.")

            chunk = ser.read(1)
            if chunk:
                data += chunk
                deadline = time.monotonic() + effective_timeout

                if any(data.endswith(term) for term in terminators):
                    break
            else:

                if data:
                    break

        text = data.decode(self.encoding, errors="replace").strip()
        if not text:
            raise TimeoutError("Answer is empty")
        return text

    def query(
        self,
        command: str,
        delay: Optional[float] = None,
        timeout: Optional[float] = None,
        retries: int = 1,
        do_sync_on_fail: bool = True):
       
        actual_delay = self.query_delay if delay is None else delay
        last_exc: Optional[Exception] = None

        for attempt in range(retries + 1):
            try:
                self.write(command)
                if actual_delay > 0:
                    time.sleep(actual_delay)
                return self._read_response(timeout=timeout)
            except Exception as exc:
                last_exc = exc
                if attempt < retries:
                    if do_sync_on_fail:
                        try:
                            self.sync()
                        except Exception:
                            pass
                    time.sleep(0.2)
                else:
                    break

        if last_exc is not None:
            raise last_exc
        raise RuntimeError("Unexpected error query().")

    def ask_float(
        self,
        command: str,
        delay: Optional[float] = None,
        timeout: Optional[float] = None,
        retries: int = 1):

        response = self.query(
            command=command,
            delay=delay,
            timeout=timeout,
            retries=retries)
        
        first_value = response.split(",")[0].strip()
        return float(first_value)


    # ------------------------------------------------------------------
    # СЕРВИСНЫЕ КОМАНДЫ
    # ------------------------------------------------------------------
    def identify(self):
        return self.query("*IDN?", retries=2)

    def reset(self):
        self.write("*RST")
        time.sleep(0.2)

    def clear_status(self):
        self.write("*CLS")
        time.sleep(0.05)

    def get_error(self):
        return self.query(":SYST:ERR?", retries=1)


    # ------------------------------------------------------------------
    # УПРАВЛЕНИЕ ВЫХОДОМ
    # ------------------------------------------------------------------
    def output_on(self):
        self.write(":OUTP ON")
        time.sleep(0.05)

    def output_off(self):
        self.write(":OUTP OFF")
        time.sleep(0.05)

    def is_output_on(self):
        response = self.query(":OUTP?", retries=2).strip().upper()

        if response in {"1", "ON"}:
            return True
        if response in {"0", "OFF"}:
            return False

        raise ValueError(f"Unexpected answer for :OUTP?: {response!r}")

    def check_source(self):
        response = self.query(':ROUT:TERM?')

        if response == None:
            ValueError(f"Unexpected answer for :ROUT:TERM?: {response!r}")
        else:    
            return self.query(':ROUT:TERM?')
    
    def select_source(self, source='front'):
        if (self.query(':ROUT:TERM?') == 'FRON') and (source == 'rear'):
            self.write(':ROUT:TERM REAR')
            time.sleep(0.05)

        elif (self.query(':ROUT:TERM?') == 'REAR') and (source == 'front'):
            self.write(':ROUT:TERM FRON')
            time.sleep(0.05)

        else: 
            self.write(':ROUT:TERM FRON')
            time.sleep(0.05)


    # ------------------------------------------------------------------
    # КОНФИГУРАЦИЯ ИСТОЧНИКА
    # ------------------------------------------------------------------
    def set_source_voltage(
        self,
        voltage: float,
        compliance_current: float = 0.01,
        turn_on: bool = True,
    ) -> None:
        self.write(":SOUR:FUNC VOLT")
        self.write(":SOUR:VOLT:MODE FIX")
        self.write(":SOUR:VOLT:RANG:AUTO ON")

        self.write(':SENS:FUNC "CURR"')
        self.write(":SENS:CURR:RANG:AUTO ON")
        self.write(f":SENS:CURR:PROT {compliance_current}")

        self.write(f":SOUR:VOLT:LEV {voltage}")

        if turn_on:
            self.output_on()

    def set_source_current(
        self,
        current: float,
        compliance_voltage: float = 10.0,
        turn_on: bool = True,
    ) -> None:
        self.write(":SOUR:FUNC CURR")
        self.write(":SOUR:CURR:MODE FIX")
        self.write(":SOUR:CURR:RANG:AUTO ON")

        self.write(':SENS:FUNC "VOLT"')
        self.write(":SENS:VOLT:RANG:AUTO ON")
        self.write(f":SENS:VOLT:PROT {compliance_voltage}")

        self.write(f":SOUR:CURR:LEV {current}")

        if turn_on:
            self.output_on()


    # ------------------------------------------------------------------
    # ИЗМЕРЕНИЯ
    # ------------------------------------------------------------------
    def measure_current(self, delay: float = 0.05):

        self.write(':SENS:FUNC "CURR"')
        self.write(":FORM:ELEM CURR")
        return self.ask_float(":READ?", delay=delay, retries=2)

    def measure_voltage(self, delay: float = 0.05):

        self.write(':SENS:FUNC "VOLT"')
        self.write(":FORM:ELEM VOLT")
        return self.ask_float(":READ?", delay=delay, retries=2)

    def measure_iv(self, delay: float = 0.05):
        
        voltage = self.measure_voltage(delay=delay)
        current = self.measure_current(delay=delay)
        return {
            "voltage": voltage,
            "current": current,
        }


    # ------------------------------------------------------------------
    # УСТАНОВКА В ТОЧКУ
    # ------------------------------------------------------------------
    def set_voltage_point(
        self,
        voltage: float,
        compliance_current: float = 0.01,
        settle_time: float = 0.2):

        self.set_source_voltage(
            voltage=voltage,
            compliance_current=compliance_current,
            turn_on=True)
        
        time.sleep(settle_time)
        result = self.measure_iv(delay=0.05)
        result["source_mode"] = "voltage"
        result["set_value"] = voltage
        return result

    def set_current_point(
        self,
        current: float,
        compliance_voltage: float = 10.0,
        settle_time: float = 0.2):

        self.set_source_current(
            current=current,
            compliance_voltage=compliance_voltage,
            turn_on=True)
        
        time.sleep(settle_time)
        result = self.measure_iv(delay=0.05)
        result["source_mode"] = "current"
        result["set_value"] = current
        return result

    # ------------------------------------------------------------------
    # SWEEP
    # ------------------------------------------------------------------
    @staticmethod
    def _build_sweep_values(start: float, stop: float, step: float) -> list[float]:
        if step == 0:
            raise ValueError("step не может быть равен 0.")

        if start < stop and step < 0:
            raise ValueError("Для возрастающего диапазона step должен быть > 0.")

        if start > stop and step > 0:
            raise ValueError("Для убывающего диапазона step должен быть < 0.")

        values: list[float] = []
        x = start
        tol = abs(step) * 1e-12 + 1e-15

        if step > 0:
            while x <= stop + tol:
                values.append(round(x, 12))
                x += step
        else:
            while x >= stop - tol:
                values.append(round(x, 12))
                x += step

        return values

    def sweep_iv_by_voltage(
        self,
        start: float,
        stop: float,
        step: float,
        compliance_current: float = 0.01,
        settle_time: float = 0.2,
        output_off_after: bool = True):
        
        values = self._build_sweep_values(start, stop, step)
        data: list[dict[str, float]] = []

        self.set_source_voltage(
            voltage=values[0],
            compliance_current=compliance_current,
            turn_on=True,
        )

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

    def sweep_iv_by_current(
        self,
        start: float,
        stop: float,
        step: float,
        compliance_voltage: float = 10.0,
        settle_time: float = 0.2,
        output_off_after: bool = True):
        
        values = self._build_sweep_values(start, stop, step)
        data: list[dict[str, float]] = []

        self.set_source_current(
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

    # ------------------------------------------------------------------
    # -------------------------------CSV--------------------------------
    # ------------------------------------------------------------------
    def save_iv_to_csv(
        self,
        data: list[dict[str, float]],
        file_path: str,
        delimiter: str = ","):
        
        if not data:
            raise ValueError("Пустой массив данных: нечего сохранять в CSV.")

        path = Path(file_path)
        path.parent.mkdir(parents=True, exist_ok=True)

        fieldnames: list[str] = []
        for row in data:
            for key in row.keys():
                if key not in fieldnames:
                    fieldnames.append(key)

        with path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=delimiter)
            writer.writeheader()
            writer.writerows(data)

    # ------------------------------------------------------------------
    # ГРАФИКИ
    # ------------------------------------------------------------------
    def plot_iv_curve(
        self,
        data: list[dict[str, float]],
        title: str = "I(V)",
        xlabel: str = "Voltage, V",
        ylabel: str = "Current, A",
        show: bool = True,
        save_path: str | None = None,
        marker: str = "o",
        linestyle: str = "-",
        grid: bool = True,
    ) -> None:
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

        if save_path is not None:
            Path(save_path).parent.mkdir(parents=True, exist_ok=True)
            plt.savefig(save_path, dpi=300)

        if show:
            plt.show()
        else:
            plt.close()

    def plot_iv_curve_semilogy(
        self,
        data: list[dict[str, float]],
        title: str = "I(V) semilogy",
        xlabel: str = "Voltage, V",
        ylabel: str = "|Current|, A",
        show: bool = True,
        save_path: str | None = None,
        marker: str = "o",
        linestyle: str = "-",
        grid: bool = True):

        if not data:
            raise ValueError("Пустой массив данных: нечего строить.")

        x = [float(point["voltage"]) for point in data]
        y = [abs(float(point["current"])) for point in data]

        plt.figure(figsize=(7, 5))
        plt.semilogy(x, y, marker=marker, linestyle=linestyle)
        plt.title(title)
        plt.xlabel(xlabel)
        plt.ylabel(ylabel)

        if grid:
            plt.grid(True, which="both")

        plt.tight_layout()

        if save_path is not None:
            Path(save_path).parent.mkdir(parents=True, exist_ok=True)
            plt.savefig(save_path, dpi=300)

        if show:
            plt.show()
        else:
            plt.close()


