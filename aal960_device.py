# -*- coding: utf-8 -*-
"""
aal960_device.py — абстракция устройств STMP-960:
- RealDevice960: реальный калибратор по COM
- SimDevice960: эмуляция с ручным управлением (отдельное окно)
"""

import threading
import time
from typing import Callable, Optional, Any

import serial
import tkinter as tk
from tkinter import ttk, messagebox

from aal960_protocol import FrameParser960, START_FRAME, Measurement


CallbackType = Callable[[Measurement], Any]


# ============= РЕАЛЬНОЕ УСТРОЙСТВО =============

class RealDevice960:
    """Работа с реальным калибратором через COM-порт."""

    def __init__(self, port: str, baud: int, callback: CallbackType):
        self.port = port
        self.baud = baud
        self.callback = callback

        self.ser: Optional[serial.Serial] = None
        self.running = False
        self.thread: Optional[threading.Thread] = None

    def start(self):
        """Открываем порт, посылаем стартовый кадр и запускаем поток опроса."""
        # "Пинок" калибратора
        try:
            tmp = serial.Serial(self.port, self.baud, timeout=0.2)
            time.sleep(1.0)
            tmp.write(START_FRAME)
            tmp.close()
        except Exception as e:
            print("Start-frame warning:", e)

        time.sleep(0.3)

        # Основное соединение
        try:
            self.ser = serial.Serial(self.port, self.baud, timeout=0.2)
        except Exception as e:
            raise RuntimeError(f"Не удалось открыть порт: {e}")

        self.running = True
        self.thread = threading.Thread(target=self._poll_loop, daemon=True)
        self.thread.start()

    def stop(self):
        """Остановить опрос и закрыть порт."""
        self.running = False
        if self.ser:
            try:
                self.ser.close()
            except:
                pass
        self.ser = None
        self.thread = None

    def _poll_loop(self):
        while self.running:
            try:
                if not self.ser:
                    time.sleep(0.05)
                    continue

                payload = FrameParser960.read_frame(self.ser)
                if payload:
                    meas = FrameParser960.parse_payload(payload)
                    if meas and self.callback:
                        self.callback(meas)

                time.sleep(0.003)

            except serial.SerialException:
                print("Потеряно соединение!")
                self.running = False
                break
            except Exception as e:
                print("poll_loop error:", e)
                time.sleep(0.1)


# ============= ЭМУЛЯТОР С ОТДЕЛЬНЫМ ОКНОМ =============

P_UNITS = [
    'мм вод.ст.',
    'мм рт.ст.',
    'мбар',
    'бар',
    'psi',
    'Па',
    'МПа',
    'кПа',
    'inHg',
    'inH2O',
    'kg/cm²',
]


class SimDevice960:
    """
    Эмулятор STMP-960:
    - отдельное окно, где можно задать режим, давление, сигнал/реле;
    - по кнопке «Отправить» вызывает callback(Measurement).
    """

    def __init__(self, master: tk.Misc, callback: CallbackType):
        self.master = master
        self.callback = callback
        self.top: Optional[tk.Toplevel] = None

        # UI state
        self.mode_var = tk.StringVar(value="I/P")
        self.p_val_var = tk.StringVar(value="0.0")
        self.p_unit_var = tk.StringVar(value="кПа")
        self.signal_val_var = tk.StringVar(value="4.0")
        self.relay_state_var = tk.StringVar(value="open")  # open/closed

    # Публичные методы, похожие на RealDevice960

    def start(self):
        """Показать окно эмулятора."""
        if self.top is not None and tk.Toplevel.winfo_exists(self.top):
            self.top.lift()
            return
        self._build_window()

    def stop(self):
        """Закрыть окно эмулятора."""
        if self.top is not None and tk.Toplevel.winfo_exists(self.top):
            self.top.destroy()
        self.top = None

    # UI

    def _build_window(self):
        self.top = tk.Toplevel(self.master)
        self.top.title("Эмулятор STMP-960")
        self.top.geometry("380x260")
        self.top.resizable(False, False)

        frm = ttk.Frame(self.top, padding=10)
        frm.pack(fill="both", expand=True)

        # Режим
        ttk.Label(frm, text="Режим:").grid(row=0, column=0, sticky="w")
        modes_frame = ttk.Frame(frm)
        modes_frame.grid(row=0, column=1, columnspan=2, sticky="w", pady=5)
        ttk.Radiobutton(modes_frame, text="I/P (мА)", value="I/P", variable=self.mode_var,
                        command=self._update_mode_widgets).pack(side="left", padx=5)
        ttk.Radiobutton(modes_frame, text="V/P (В)", value="V/P", variable=self.mode_var,
                        command=self._update_mode_widgets).pack(side="left", padx=5)
        ttk.Radiobutton(modes_frame, text="Реле", value="Реле", variable=self.mode_var,
                        command=self._update_mode_widgets).pack(side="left", padx=5)

        # Давление
        ttk.Label(frm, text="Давление:").grid(row=1, column=0, sticky="e", pady=5)
        ttk.Entry(frm, textvariable=self.p_val_var, width=12).grid(row=1, column=1, sticky="w")
        cb_pu = ttk.Combobox(frm, values=P_UNITS, textvariable=self.p_unit_var, width=10, state="readonly")
        cb_pu.grid(row=1, column=2, padx=5, sticky="w")

        # Сигнал / реле
        self.lbl_signal = ttk.Label(frm, text="Ток, мА:")
        self.lbl_signal.grid(row=2, column=0, sticky="e", pady=5)

        self.ent_signal = ttk.Entry(frm, textvariable=self.signal_val_var, width=12)
        self.ent_signal.grid(row=2, column=1, sticky="w")

        self.lbl_signal_unit = ttk.Label(frm, text="мА")
        self.lbl_signal_unit.grid(row=2, column=2, sticky="w")

        # Для режима "Реле" — переключатель замкнут/разомкнут
        self.frm_relay = ttk.Frame(frm)
        ttk.Radiobutton(self.frm_relay, text="замкнут", value="closed", variable=self.relay_state_var).pack(
            side="left", padx=5
        )
        ttk.Radiobutton(self.frm_relay, text="разомкнут", value="open", variable=self.relay_state_var).pack(
            side="left", padx=5
        )

        # Кнопки
        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=4, column=0, columnspan=3, pady=15)

        ttk.Button(btn_frame, text="Отправить в GUI", command=self._send).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Закрыть", command=self.stop).pack(side="left", padx=5)

        self._update_mode_widgets()
        self.top.protocol("WM_DELETE_WINDOW", self.stop)

    def _update_mode_widgets(self):
        """Спрятать/показать поля в зависимости от режима."""
        mode = self.mode_var.get()
        if mode in ("I/P", "V/P"):
            # показать поле сигнала
            self.lbl_signal.grid()
            self.ent_signal.grid()
            self.lbl_signal_unit.grid()
            self.frm_relay.grid_forget()

            if mode == "I/P":
                self.lbl_signal.config(text="Ток:")
                self.lbl_signal_unit.config(text="мА")
                if not self.signal_val_var.get():
                    self.signal_val_var.set("4.0")
            else:
                self.lbl_signal.config(text="Напряжение:")
                self.lbl_signal_unit.config(text="В")
                if not self.signal_val_var.get():
                    self.signal_val_var.set("0.0")
        else:
            # режим Реле
            self.lbl_signal.grid_remove()
            self.ent_signal.grid_remove()
            self.lbl_signal_unit.grid_remove()
            self.frm_relay.grid(row=2, column=1, columnspan=2, sticky="w", pady=5)

    def _send(self):
        """Считать значения из полей и отправить Measurement в callback."""
        mode = self.mode_var.get()

        try:
            p = float(self.p_val_var.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Ошибка", "Неверное значение давления")
            return

        p_unit = self.p_unit_var.get() or "кПа"

        if mode in ("I/P", "V/P"):
            try:
                sig_val = float(self.signal_val_var.get().replace(",", "."))
            except ValueError:
                messagebox.showerror("Ошибка", "Неверное значение сигнала")
                return

            if mode == "I/P":
                sig_unit = "мА"
            else:
                sig_unit = "В"

            meas = Measurement(
                mode=mode,
                pressure=p,
                pressure_unit=p_unit,
                signal=sig_val,
                signal_unit=sig_unit,
                relay_state=None,
                raw=b"",
            )
        else:
            # Реле
            state = self.relay_state_var.get()
            if state == "closed":
                relay = True
                sig_unit = "контакт замкнут"
            else:
                relay = False
                sig_unit = "контакт разомкнут"

            meas = Measurement(
                mode="Реле",
                pressure=p,
                pressure_unit=p_unit,
                signal=1.0 if relay else 0.0,
                signal_unit=sig_unit,
                relay_state=relay,
                raw=b"",
            )

        if self.callback:
            self.callback(meas)
