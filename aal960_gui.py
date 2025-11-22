# -*- coding: utf-8 -*-
"""
aal960_gui.py — современный GUI для AAL960 Metrology Monitoring

Использует:
- aal960_device.RealDevice960 / SimDevice960
- aal960_protocol.Measurement

Функции:
- Подключение к реальному STMP-960 или запуск эмулятора
- Онлайн отображение текущих значений
- Вкладка "Поверка" с фиксацией точек
- Вкладка "Мониторинг" с графиком давления по времени
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from dataclasses import dataclass
from typing import List, Optional
import time
from datetime import datetime

import serial.tools.list_ports

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure

from aal960_device import RealDevice960, SimDevice960
from aal960_protocol import Measurement


# =============== ВСПОМОГАТЕЛЬНЫЕ СТРУКТУРЫ ===============

@dataclass
class CalibPoint:
    timestamp: str
    mode: str
    pressure: float
    pressure_unit: str
    signal: float
    signal_unit: str


# =============== ОСНОВНОЕ ПРИЛОЖЕНИЕ ===============

class AAL960App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("AAL960 Metrology Monitoring")
        self.root.geometry("1100x700")

        # текущее устройство (реальное или симулятор)
        self.device: Optional[RealDevice960 | SimDevice960] = None

        # переменные состояния
        self.simulation_mode = tk.BooleanVar(value=False)

        self.current_mode = tk.StringVar(value="—")
        self.current_p = tk.DoubleVar(value=0.0)
        self.current_p_unit = tk.StringVar(value="—")
        self.current_signal = tk.DoubleVar(value=0.0)
        self.current_signal_unit = tk.StringVar(value="—")

        self.status_var = tk.StringVar(value="Отключено")

        # для поверки
        self.calib_points: List[CalibPoint] = []

        # для мониторинга
        self.monitor_running = False
        self.monitor_data: List[tuple[float, float]] = []  # (t_rel, pressure)
        self.monitor_t0 = 0.0

        # UI
        self._build_ui()

    # ---------- СТРОИМ ИНТЕРФЕЙС ----------

    def _build_ui(self):
        # Верхняя панель подключения
        frm_conn = ttk.LabelFrame(self.root, text=" Подключение ")
        frm_conn.pack(fill="x", padx=10, pady=8)

        ttk.Label(frm_conn, text="Порт:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.port_combo = ttk.Combobox(frm_conn, width=15, state="readonly")
        self.port_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Button(frm_conn, text="Обновить", command=self.refresh_ports).grid(
            row=0, column=2, padx=5, pady=5
        )

        ttk.Label(frm_conn, text="Скорость:").grid(row=0, column=3, padx=5, pady=5, sticky="e")
        self.baud_combo = ttk.Combobox(
            frm_conn, width=8, values=["9600", "19200", "38400", "57600", "115200"], state="readonly"
        )
        self.baud_combo.set("9600")
        self.baud_combo.grid(row=0, column=4, padx=5, pady=5, sticky="w")

        ttk.Checkbutton(
            frm_conn,
            text="Эмуляция",
            variable=self.simulation_mode
        ).grid(row=0, column=5, padx=5, pady=5, sticky="w")

        self.btn_connect = ttk.Button(frm_conn, text="Подключиться", command=self.connect)
        self.btn_connect.grid(row=0, column=6, padx=10, pady=5)

        self.btn_disconnect = ttk.Button(frm_conn, text="Отключиться", command=self.disconnect, state="disabled")
        self.btn_disconnect.grid(row=0, column=7, padx=5, pady=5)

        ttk.Label(frm_conn, text="Статус:").grid(row=0, column=8, padx=10, pady=5, sticky="e")
        self.lbl_status = ttk.Label(frm_conn, textvariable=self.status_var, foreground="red")
        self.lbl_status.grid(row=0, column=9, padx=5, pady=5, sticky="w")

        # Средняя панель: текущие значения
        frm_current = ttk.LabelFrame(self.root, text=" Текущее значение ")
        frm_current.pack(fill="x", padx=10, pady=5)

        ttk.Label(frm_current, text="Режим:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        ttk.Label(frm_current, textvariable=self.current_mode, font=("", 11, "bold")).grid(
            row=0, column=1, padx=5, pady=5, sticky="w"
        )

        ttk.Label(frm_current, text="Давление:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        ttk.Label(frm_current, textvariable=self.current_p, font=("", 11, "bold")).grid(
            row=0, column=3, padx=2, pady=5, sticky="w"
        )
        ttk.Label(frm_current, textvariable=self.current_p_unit).grid(
            row=0, column=4, padx=2, pady=5, sticky="w"
        )

        ttk.Label(frm_current, text="Сигнал:").grid(row=0, column=5, padx=5, pady=5, sticky="e")
        ttk.Label(frm_current, textvariable=self.current_signal, font=("", 11, "bold")).grid(
            row=0, column=6, padx=2, pady=5, sticky="w"
        )
        ttk.Label(frm_current, textvariable=self.current_signal_unit).grid(
            row=0, column=7, padx=2, pady=5, sticky="w"
        )

        # Notebook с вкладками
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self._build_tab_calibration()
        self._build_tab_monitor()
        self._build_tab_settings()

        # инициализируем список портов
        self.refresh_ports()

    def _build_tab_calibration(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Поверка")

        # Кнопки
        frm_btn = ttk.Frame(tab)
        frm_btn.pack(fill="x", padx=5, pady=5)

        ttk.Button(frm_btn, text="Добавить точку (фиксировать)", command=self.add_calib_point).pack(
            side="left", padx=5
        )
        ttk.Button(frm_btn, text="Удалить выбранную", command=self.remove_selected_calib_point).pack(
            side="left", padx=5
        )
        ttk.Button(frm_btn, text="Очистить все", command=self.clear_calib_points).pack(
            side="left", padx=5
        )
        ttk.Button(frm_btn, text="Экспорт в CSV", command=self.export_calib_to_csv).pack(
            side="left", padx=5
        )

        # Таблица точек
        columns = ("time", "mode", "pressure", "p_unit", "signal", "s_unit")
        self.tree_calib = ttk.Treeview(tab, columns=columns, show="headings", height=15)
        self.tree_calib.pack(fill="both", expand=True, padx=5, pady=5)

        self.tree_calib.heading("time", text="Время")
        self.tree_calib.heading("mode", text="Режим")
        self.tree_calib.heading("pressure", text="Давление")
        self.tree_calib.heading("p_unit", text="Ед. давл.")
        self.tree_calib.heading("signal", text="Сигнал")
        self.tree_calib.heading("s_unit", text="Ед. сигнала")

        self.tree_calib.column("time", width=130)
        self.tree_calib.column("mode", width=60)
        self.tree_calib.column("pressure", width=100)
        self.tree_calib.column("p_unit", width=90)
        self.tree_calib.column("signal", width=100)
        self.tree_calib.column("s_unit", width=90)

    def _build_tab_monitor(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Мониторинг")

        frm_top = ttk.Frame(tab)
        frm_top.pack(fill="x", padx=5, pady=5)

        self.btn_mon_start = ttk.Button(frm_top, text="Старт мониторинга", command=self.start_monitor)
        self.btn_mon_start.pack(side="left", padx=5)

        self.btn_mon_stop = ttk.Button(frm_top, text="Стоп", command=self.stop_monitor, state="disabled")
        self.btn_mon_stop.pack(side="left", padx=5)

        self.btn_mon_clear = ttk.Button(frm_top, text="Очистить", command=self.clear_monitor)
        self.btn_mon_clear.pack(side="left", padx=5)

        # Фигура matplotlib
        frm_plot = ttk.Frame(tab)
        frm_plot.pack(fill="both", expand=True, padx=5, pady=5)

        self.fig = Figure(figsize=(6, 4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_title("Давление по времени")
        self.ax.set_xlabel("Время, с")
        self.ax.set_ylabel("Давление")

        self.canvas = FigureCanvasTkAgg(self.fig, master=frm_plot)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        toolbar = NavigationToolbar2Tk(self.canvas, frm_plot)
        toolbar.update()
        toolbar.pack(fill="x")

    def _build_tab_settings(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Настройки")

        ttk.Label(tab, text="Здесь позже появятся настройки экспорта, путей, форматов и т.п.").pack(
            padx=10, pady=10, anchor="w"
        )

    # ---------- ПОДКЛЮЧЕНИЕ / ОТКЛЮЧЕНИЕ ----------

    def refresh_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self.port_combo["values"] = ports
        if ports and not self.port_combo.get():
            self.port_combo.set(ports[0])

    def connect(self):
        if self.device is not None:
            return

        if self.simulation_mode.get():
            # эмулятор
            self.device = SimDevice960(self.root, self._device_callback)
            self.device.start()

            self.btn_connect.config(state="disabled")
            self.btn_disconnect.config(state="normal")

            self.status_var.set("Эмуляция")
            self.lbl_status.config(foreground="blue")
        else:
            port = self.port_combo.get()
            if not port:
                messagebox.showerror("Ошибка", "Выберите COM-порт")
                return

            try:
                baud = int(self.baud_combo.get())
            except Exception:
                messagebox.showerror("Ошибка", "Неверная скорость")
                return

            try:
                dev = RealDevice960(port, baud, self._device_callback)
                dev.start()
            except Exception as e:
                messagebox.showerror("Ошибка подключения", str(e))
                return

            self.device = dev

            self.btn_connect.config(state="disabled")
            self.btn_disconnect.config(state="normal")

            self.status_var.set("Подключено")
            self.lbl_status.config(foreground="green")

    def disconnect(self):
        if self.device is not None:
            try:
                self.device.stop()
            except Exception as e:
                print("device.stop error:", e)
            self.device = None

        self.btn_connect.config(state="normal")
        self.btn_disconnect.config(state="disabled")

        self.status_var.set("Отключено")
        self.lbl_status.config(foreground="red")

        # сброс текущих значений
        self.current_mode.set("—")
        self.current_p.set(0.0)
        self.current_p_unit.set("—")
        self.current_signal.set(0.0)
        self.current_signal_unit.set("—")

        # останавливаем мониторинг
        self.stop_monitor()
        self.clear_monitor()

    # ---------- КОЛБЭК ОТ УСТРОЙСТВА ----------

    def _device_callback(self, meas: Measurement):
        """Вызывается из потока устройства — кидаем в основной поток Tk."""
        self.root.after(0, self.on_measurement, meas)

    def on_measurement(self, meas: Measurement):
        """Обновляем текущие значения и мониторинг."""
        self.current_mode.set(meas.mode)
        self.current_p.set(meas.pressure)
        self.current_p_unit.set(meas.pressure_unit)

        # для реле отображаем текст
        if meas.relay_state is not None:
            self.current_signal.set(1.0 if meas.relay_state else 0.0)
            self.current_signal_unit.set(meas.signal_unit)
        else:
            try:
                self.current_signal.set(float(meas.signal))
            except Exception:
                self.current_signal.set(0.0)
            self.current_signal_unit.set(meas.signal_unit)

        # мониторинг
        if self.monitor_running:
            if self.monitor_t0 == 0.0:
                self.monitor_t0 = time.time()
            t_rel = time.time() - self.monitor_t0
            self.monitor_data.append((t_rel, meas.pressure))
            self._update_monitor_plot()

    # ---------- ПОВЕРКА ----------

    def add_calib_point(self):
        if self.device is None:
            messagebox.showwarning("Нет подключения", "Сначала подключите калибратор или эмулятор.")
            return

        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        point = CalibPoint(
            timestamp=ts,
            mode=self.current_mode.get(),
            pressure=self.current_p.get(),
            pressure_unit=self.current_p_unit.get(),
            signal=self.current_signal.get(),
            signal_unit=self.current_signal_unit.get(),
        )
        self.calib_points.append(point)

        self.tree_calib.insert(
            "", "end",
            values=(
                point.timestamp,
                point.mode,
                f"{point.pressure:.6f}",
                point.pressure_unit,
                f"{point.signal:.6f}",
                point.signal_unit,
            ),
        )

    def remove_selected_calib_point(self):
        sel = self.tree_calib.selection()
        if not sel:
            return
        for item in sel:
            self.tree_calib.delete(item)
        # для простоты пока не синхронизируем с self.calib_points по индексу
        # (в будущем можно хранить id и очищать список)

    def clear_calib_points(self):
        self.calib_points.clear()
        for item in self.tree_calib.get_children():
            self.tree_calib.delete(item)

    def export_calib_to_csv(self):
        if not self.calib_points:
            messagebox.showinfo("Экспорт", "Нет точек для экспорта.")
            return

        fname = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV файлы", "*.csv"), ("Все файлы", "*.*")],
            title="Сохранить точки поверки",
        )
        if not fname:
            return

        try:
            with open(fname, "w", encoding="utf-8") as f:
                f.write("timestamp;mode;pressure;pressure_unit;signal;signal_unit\n")
                for p in self.calib_points:
                    f.write(
                        f"{p.timestamp};{p.mode};{p.pressure:.6f};{p.pressure_unit};"
                        f"{p.signal:.6f};{p.signal_unit}\n"
                    )
            messagebox.showinfo("Экспорт", f"Данные сохранены в {fname}")
        except Exception as e:
            messagebox.showerror("Ошибка экспорта", str(e))

    # ---------- МОНИТОРИНГ ----------

    def start_monitor(self):
        if not self.device:
            messagebox.showwarning("Нет подключения", "Сначала подключите калибратор или эмулятор.")
            return
        self.monitor_running = True
        self.monitor_t0 = 0.0
        self.btn_mon_start.config(state="disabled")
        self.btn_mon_stop.config(state="normal")

    def stop_monitor(self):
        self.monitor_running = False
        self.btn_mon_start.config(state="normal")
        self.btn_mon_stop.config(state="disabled")

    def clear_monitor(self):
        self.monitor_data.clear()
        self.monitor_t0 = 0.0
        self.ax.cla()
        self.ax.set_title("Давление по времени")
        self.ax.set_xlabel("Время, с")
        self.ax.set_ylabel("Давление")
        self.canvas.draw()

    def _update_monitor_plot(self):
        if not self.monitor_data:
            return
        xs, ys = zip(*self.monitor_data)
        self.ax.cla()
        self.ax.plot(xs, ys)
        self.ax.set_title("Давление по времени")
        self.ax.set_xlabel("Время, с")
        self.ax.set_ylabel(f"Давление, {self.current_p_unit.get() or ''}")
        self.canvas.draw()


# =============== ТОЧКА ВХОДА ===============

def run_app():
    root = tk.Tk()
    app = AAL960App(root)
    root.mainloop()


if __name__ == "__main__":
    run_app()
