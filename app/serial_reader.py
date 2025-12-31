import serial
import serial.tools.list_ports
import time
import threading

from app.constants import BAUD_RATE
from app.storage import load_data, save_data, CARDS_FILE


def normalize_uid(uid):
    return uid.strip().upper().replace(" ", "")


def auto_detect_port():
    ports = list(serial.tools.list_ports.comports())
    for port in ports:
        if any(keyword in port.description for keyword in ["Arduino", "CH340", "USB Serial"]):
            return port.device

    if ports:
        return ports[0].device

    return None

class CardReader:
    def __init__(self, selected_port, on_card_callback):
        self.selected_port = selected_port
        self.on_card_callback = on_card_callback

        self.card_mode_running = False
        self.reader_enabled = False
        self.card_thread = None
        self.ser = None

    def start(self):
        if self.card_mode_running:
            return

        self.reader_enabled = True
        self.card_mode_running = True
        self.card_thread = threading.Thread(
            target=self._card_mode_worker,
            daemon=True
        )
        self.card_thread.start()

    def stop(self):
        self.reader_enabled = False
        self.card_mode_running = False
        if self.ser and self.ser.is_open:
            self.ser.close()

    def _card_mode_worker(self):
        try:
            self.ser = serial.Serial(self.selected_port, BAUD_RATE, timeout=1)
        except Exception:
            self.stop()
            return

        while self.card_mode_running and self.reader_enabled:
            try:
                if self.ser.in_waiting:
                    uid_line = self.ser.readline().decode("utf-8").strip()
                    if uid_line:
                        uid = normalize_uid(uid_line)
                        self.on_card_callback(uid)
                time.sleep(0.1)
            except Exception:
                break

        if self.ser and self.ser.is_open:
            self.ser.close()

def get_card_students():
    return load_data(CARDS_FILE, {})


def save_card_students(card_students):
    return save_data(CARDS_FILE, card_students)
