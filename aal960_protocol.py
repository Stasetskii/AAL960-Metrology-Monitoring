# -*- coding: utf-8 -*-
"""
aal960_protocol.py — протокол STMP-960 (только чтение):
- контрольная сумма
- разбор кадров 0x30 15/16/17
- константы единиц измерения
- объект Measurement для дальнейшей работы в GUI/логике
"""

import struct
from dataclasses import dataclass
from functools import reduce
from typing import Optional


# === Константы протокола ===

ADDR = 0x01

START_FRAME = b"\x55\x55\x01\x03\x20\x23\x00\x01\xAA\xAA"

# Таблица единиц давления (НОВАЯ, как в твоём 999.py)
UNITS_P = {
    0x00: "мм вод.ст.",
    0x01: "мм рт.ст.",
    0x02: "мбар",
    0x03: "бар",
    0x04: "psi",
    0x05: "Па",
    0x06: "МПа",
    0x07: "кПа",
    0x0A: "inHg",
    0x0B: "inH2O",
    0x0C: "kg/cm²",
}

# Таблица единиц сигнала
UNITS_E = {
    0x08: "мА",
    0x09: "В",
}


# === Вспомогательная логика ===

def cs(data: list[int]) -> int:
    """XOR-контрольная сумма (как в твоём 999.py)."""
    return reduce(lambda x, y: x ^ y, data, 0)


@dataclass
class Measurement:
    """
    Нормализованное представление измерения от STMP-960.

    mode:
        "I/P"  — режим ток/давление
        "V/P"  — режим напряжение/давление
        "Реле" — режим реле

    pressure:
        значение давления (float)
    pressure_unit:
        единица давления (строка, берётся из UNITS_P)

    signal:
        для I/P, V/P — float (ток мА или напряжение В)
        для Реле      — bool (True = замкнуто, False = разомкнуто) либо int для “неизвестных” кодов

    signal_unit:
        строка единицы сигнала:
            "мА", "В",
            "замкнут"/"разомкнут"/"реле 0x.."

    relay_state:
        None  — если режим не реле
        True  — контакт замкнут
        False — контакт разомкнут

    raw:
        сырые 13 байт payload на всякий случай
    """
    mode: str
    pressure: float
    pressure_unit: str
    signal: float | bool | int
    signal_unit: str
    relay_state: Optional[bool] = None
    raw: bytes = b""


class FrameParser960:
    """
    Класс-обёртка над разбором кадров STMP-960.

    read_frame(ser):
        читает ОДИН корректный payload (13 байт) из serial.Serial
    parse_payload(data):
        превращает payload в Measurement
    """

    @staticmethod
    def read_frame(ser) -> Optional[bytes]:
        """
        Читает один валидный payload (data) из последовательного порта.

        Логика 1-в-1 с твоей функцией read_frame из 999.py:
        - ждём, пока в буфере >= 20 байт
        - ищем префикс 0x55 0x55
        - проверяем addr, length
        - длину payload
        - окончание 0xAA 0xAA
        - XOR-контрольную сумму
        """
        while ser.in_waiting >= 20:
            # префикс
            if ser.read(2) != b"\x55\x55":
                continue

            # заголовок
            hdr = ser.read(2)
            if len(hdr) < 2:
                return None

            addr, length = hdr
            if addr != ADDR:
                continue

            # payload + cs + конец
            payload = ser.read(length + 3)
            if len(payload) < length + 3:
                return None

            data = payload[:length]
            recv_cs = payload[length]
            end = payload[length + 1:]

            # окончание
            if end != b"\xAA\xAA":
                continue

            # контрольная сумма
            if recv_cs != cs([ADDR, length] + list(data)):
                continue

            return data

        return None

    @staticmethod
    def parse_payload(data: bytes) -> Optional[Measurement]:
        """
        Разбор 13-байтного payload в Measurement.

        Полностью повторяет логику poll_loop из твоего 999.py:

        header = data[0:3]
        - 30 15 01 → режим I/P
        - 30 16 01 → режим V/P
        - 30 17 01 → режим Реле
        """

        if len(data) != 13:
            return None

        header = data[0:3]

        # ===== I/P или V/P =====
        if header in (b"\x30\x15\x01", b"\x30\x16\x01"):
            p = struct.unpack(">f", data[3:7])[0]
            sec = struct.unpack(">f", data[8:12])[0]

            p_unit = UNITS_P.get(data[7], f"0x{data[7]:02X}")
            sec_unit = UNITS_E.get(data[12], f"0x{data[12]:02X}")

            mode = "I/P" if header == b"\x30\x15\x01" else "V/P"

            return Measurement(
                mode=mode,
                pressure=p,
                pressure_unit=p_unit,
                signal=sec,
                signal_unit=sec_unit,
                relay_state=None,
                raw=data,
            )

        # ===== Реле =====
        if header == b"\x30\x17\x01":
            p = struct.unpack(">f", data[3:7])[0]
            p_unit = UNITS_P.get(data[7], f"0x{data[7]:02X}")

            code = data[12]

            if code == 0x03:
                sec: bool | int = True
                sec_unit = "замкнут"
                relay_state: Optional[bool] = True
            elif code == 0x04:
                sec = False
                sec_unit = "разомкнут"
                relay_state = False
            else:
                # неизвестный код реле — как в твоём коде:
                # sec = code, sec_unit = f"реле 0x{code:02X}"
                sec = int(code)
                sec_unit = f"реле 0x{code:02X}"
                relay_state = None

            return Measurement(
                mode="Реле",
                pressure=p,
                pressure_unit=p_unit,
                signal=sec,
                signal_unit=sec_unit,
                relay_state=relay_state,
                raw=data,
            )

        # неизвестный header
        return None
