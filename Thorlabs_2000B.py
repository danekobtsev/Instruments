import serial
import time

from serial.tools import list_ports

def list_serial_ports():
    ports = sorted(list_ports.comports(), key=lambda p: p.device)
    for p in ports:
        print(f"{p.device} | {p.description} | {p.hwid}")

def chopper_port():
    ports = sorted(list_ports.comports(), key=lambda p: p.device)
    try:
        for p in ports:
            if p.description == 'MC2000B':
                name = str(p.device)
        return name
    except: print("Chopper don't connected")

class Chopper:

    BLADE_FREQUENCY_LIMITS = {
        0: (4, 200),      # MC1F2
        1: (20, 1000),    # MC1F10
        2: (30, 1500),    # MC1F15
        3: (60, 3000),    # MC1F30
        4: (120, 6000),   # MC1F60
        5: (200, 10000),  # MC1F100
        6: (20, 1000),    # MC1F10HP
        7: (4, 200),      # MC1F2P10
        8: (12, 600),     # MC1F6P10
        9: (20, 1000),    # MC1F10A
        10: None,         # MC2F330
        11: None,         # MC2F47
        12: None,         # MC2F57B
        13: None,         # MC2F860
        14: None,         # MC2F5360
    }

    BLADE_NAMES = {
        0: 'MC1F2',
        1: 'MC1F10',
        2: 'MC1F15',
        3: 'MC1F30',
        4: 'MC1F60',
        5: 'MC1F100',
        6: 'MC1F10HP',
        7: 'MC1F2P10',
        8: 'MC1F6P10',
        9: 'MC1F10A',
        10: 'MC2F330',
        11: 'MC2F47',
        12: 'MC2F57B',
        13: 'MC2F860',
        14: 'MC2F5360'
    }


    #______________________#
    #____INNER_COMMANDS____#
    #______________________#

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def __init__(self, port, baudrate=115200, timeout=1):
        self.port = port
        self.baudrate = baudrate
        self.timeout = timeout
        self.ser = None

    def close(self):
        if self.ser is not None and self.ser.is_open: 
            self.ser.close()

    def open(self):
        if self.ser is not None and self.ser.is_open:
            return
        self.ser = serial.Serial(
            port=self.port,
            baudrate=self.baudrate,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=self.timeout,
            write_timeout=self.timeout,
            xonxoff=False,
            rtscts=False,
            dsrdtr=False,
        )
        time.sleep(0.2)
        self.ser.reset_input_buffer()
        self.ser.reset_output_buffer()

    def restore_defaults(self):
        return self.send_raw("restore")

    def _clean_response(self, text):
        lines = []
        text = text.replace("\r", "\n")
        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue
            if line == ">":
                continue
            lines.append(line)
        return "\n".join(lines)

    def _command(self, keyword, value):
            return self._send_raw(f"{keyword}={value}")

    def _ensure_open(self):
        if self.ser is None or not self.ser.is_open:
            raise RuntimeError("Порт не открыт. Сначала вызови open().")

    def _extract_value(self, reply: str) -> str:
        lines = reply.replace("\r", "\n").split("\n")
        lines = [line.strip() for line in lines if line.strip() and line.strip() != ">"]
        return lines[-1]

    def _query(self, keyword):
        return self._send_raw(f"{keyword}?")

    def _read_until_prompt(self):
        self._ensure_open()
        reply = self.ser.read_until(b">")
        return reply.decode("ascii", errors="replace")
    
    def _send_raw(self, text):
        self._ensure_open()
        message = (text + "\r").encode("ascii")
        self.ser.write(message)
        self.ser.flush()
        raw_reply = self._read_until_prompt()
        clean_reply = self._clean_response(raw_reply)
        return clean_reply

    def _validate_frequency_for_current_blade(self, freq):
        limits = self.get_blade_limits()
        if limits is None:
            # Для гармонических blade не делаем жесткую проверку здесь
            return

        f_min, f_max = limits
        if not (f_min <= freq <= f_max):
            blade_name = self.get_blade_name()
            raise ValueError(
                f"Частота {freq} Гц вне диапазона для {blade_name}: {f_min}..{f_max} Гц"
            )
    

    #______________________#
    #____QUERY_COMMANDS____#
    #______________________#
    def get_blade(self):
        return self._extract_value(self._query('blade'))
    
    def get_blade_name(self):
        raw = self.get_blade()
        try:
            idx = int(raw)
            return self.BLADE_NAMES.get(idx, f"UNKNOWN({idx})")
        except ValueError:
            return raw
    
    def get_blade_limits(self):
        raw = self.get_blade()
        idx = int(raw)
        return self.BLADE_FREQUENCY_LIMITS.get(idx)

    def get_dharmonic(self):
        return self._extract_value(self._query("dharmonic"))

    def get_enable(self):
        return self._extract_value(self._query('enable'))

    def get_frequency(self):
        return self._extract_value(self._query('freq'))

    def get_help(self):
        return self._extract_value(self._query(''))

    def get_id(self):
        return self._extract_value(self._query('id'))

    def get_input_frequency(self):
        return self._extract_value(self._query("input"))
    
    def get_nharmonic(self):
        return self._extract_value(self._query("nharmonic"))

    def get_oncycle(self):
        return self._extract_value(self._query("oncycle"))

    def get_output_mode(self):
        return self._extract_value(self._query("output"))

    def get_phase(self):
        return self._extract_value(self._query('phase'))

    def get_reference_mode(self):
        return self._extract_value(self._query("ref"))

    def get_refout_frequency(self):
        return self._extract_value(self._query('freq'))
    
    def get_verbose(self):
        return self._extract_value(self._query("verbose"))


    #______________________#
    #___CONTROL_COMMANDS___#
    #______________________#
    
    def set_blade(self, blade_index):
        blade_index = int(blade_index)
        if blade_index not in self.BLADE_NAMES:
            raise ValueError(f"Неизвестный blade index: {blade_index}")
        return self._command("blade", blade_index)
    
    def set_dharmonic(self, value):
        value = int(value)
        if not (1 <= value <= 15):
            raise ValueError("dharmonic должен быть в диапазоне 1..15")
        return self._command("dharmonic", value)

    def set_frequency(self, freq_hz):
        freq_hz = float(freq_hz)
        if freq_hz <= 0:
            raise ValueError("Частота должна быть > 0")
        self._validate_frequency_for_current_blade(freq_hz)

        # Если частота целая, отправим как int, чтобы команда выглядела аккуратно
        value = int(freq_hz) if freq_hz.is_integer() else freq_hz
        return self._command("freq", value)

    def set_oncycle(self, value):
        value = int(value)
        if not (1 <= value <= 50):
            raise ValueError("oncycle должен быть в диапазоне 1..50")
        return self._command("oncycle", value)

    def set_nharmonic(self, value):
        value = int(value)
        if not (1 <= value <= 15):
            raise ValueError("nharmonic должен быть в диапазоне 1..15")
        return self._command("nharmonic", value)

    def set_output_mode(self, mode):
        return self._command("output", int(mode))

    def set_phase(self, phase_deg):
        phase_deg = float(phase_deg)
        if not (0 <= phase_deg <= 360):
            raise ValueError("Фаза должна быть в диапазоне 0..360 градусов")

        value = int(phase_deg) if phase_deg.is_integer() else phase_deg
        return self._command("phase", value)

    def set_reference_mode(self, mode):
        return self._command("ref", int(mode))
    
    def set_verbose(self, enabled=True):
        return self._command("verbose", 1 if enabled else 0)

    
    #_______________________#
    #____ON/OFF_COMMANDS____#
    #_______________________#
    def enable(self):
        return self._command("enable", 1)

    def disable(self):
        return self._command("enable", 0)
    

    #______________________#
    #__FREQUENCY_SWEEPING__#
    #______________________#
    def sweep_frequencies(self, freqs, dwell_s=1.0, enable_before=True,
                          disable_after=False, callback=None):
        if dwell_s < 0:
            raise ValueError("dwell_s должен быть >= 0")
        
        results = []

        if enable_before:
            self.enable()
            time.sleep(0.3)

        try:
            for freq in freqs:
                self.set_frequency(freq)
                time.sleep(dwell_s)

                point = {
                    "set_freq": freq,
                    "read_freq": self.get_frequency(),
                    "timestamp": time.time(),
                }

                if callback is not None:
                    callback_result = callback(freq, self)
                    point["callback_result"] = callback_result

                results.append(point)

        finally:
            if disable_after:
                self.disable()

        return results

    def sweep_frequency_range(self, start, stop, step, dwell_s=1.0,
                              enable_before=True, disable_after=False,
                              callback=None, include_stop=True):

        start = float(start)
        stop = float(stop)
        step = float(step)

        if step == 0:
            raise ValueError("step не может быть 0")

        if start < stop and step < 0:
            raise ValueError("Для возрастания частоты step должен быть > 0")

        if start > stop and step > 0:
            raise ValueError("Для убывания частоты step должен быть < 0")

        freqs = []
        current = start

        if step > 0:
            while current < stop or (include_stop and current <= stop):
                freqs.append(round(current, 10))
                current += step
        else:
            while current > stop or (include_stop and current >= stop):
                freqs.append(round(current, 10))
                current += step

        return self.sweep_frequencies(
            freqs=freqs,
            dwell_s=dwell_s,
            enable_before=enable_before,
            disable_after=disable_after,
            callback=callback,
        )



PORT = chopper_port()

with Chopper(PORT) as chopper:
    #chopper.open()

    print(chopper.get_id())
    chopper.sweep_frequency_range(100, 1000, 50, 5, True, True)

    chopper.close()
