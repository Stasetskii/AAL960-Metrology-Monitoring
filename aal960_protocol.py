# -*- coding: utf-8 -*-
"""
aal960_protocol.py — модуль протокола STMP-960
"""

import struct
from dataclasses import dataclass
from functools import reduce

ADDR = 0x01

UNITS_P = {
    0x00: 'мм вод.ст.',
    0x01: 'мм рт.ст.',
    0x02: 'мбар',
    0x03: 'бар',
    0x04: 'psi',
    0x05: 'Па',
    0x06: 'МПа',
    0x07: 'кПа',
    0x0A: 'inHg',
    0x0B: 'inH2O',
    0x0C: 'kg/cm²',
}

UNITS_E = {
    0x08: 'мА',
    0x09: 'В',
}

START_FRAME = b"\x55\x55\x01\x03\x20\x23\x00\x01\xAA\xAA"


def cs(data):
    """XOR-контрольная сумма"""
    return reduce(lambda a, b: a ^ b, data, 0)


@dataclass
class Measurement:
    mode: str
    pressure: float
    pressure_unit: str
    signal: float
    signal_unit: str
    relay_state: bool | None = None
    raw: bytes = b""
    


class FrameParser960:
    """Основной парсер кадров STMP-960."""

    @staticmethod
    def read_frame(ser):
        """Читает один корректный кадр (20 байт). Возвращает payload."""
        if ser.in_waiting < 20:
            return None

        if ser.read(2) != b"\x55\x55":
            return None

        hdr = ser.read(2)
        if len(hdr) < 2:
            return None

        addr, length = hdr
        if addr != ADDR:
            return None

        payload = ser.read(length + 3)
        if len(payload) < length + 3:
            return None

        data = payload[:length]
        recv_cs = payload[length]
        end = payload[length + 1:]

        if end != b"\xAA\xAA":
            return None

        if recv_cs != cs([ADDR, length] + list(data)):
            return None

        return data

    @staticmethod
    def parse_payload(data: bytes) -> Measurement | None:
        """Разбор payload, возвращает Measurement."""
        if len(data) != 13:
            return None

        header = data[0:3]

        # ======== I/P ========
        if header == b"\x30\x15\x01":
            p = struct.unpack(">f", data[3:7])[0]
            p_unit = UNITS_P.get(data[7], f"0x{data[7]:02X}")

            signal = struct.unpack(">f", data[8:12])[0]
            sig_unit = UNITS_E.get(data[12], f"0x{data[12]:02X}")

            return Measurement(
                mode="I/P",
                pressure=p,
                pressure_unit=p_unit,
                signal=signal,
                signal_unit=sig_unit,
                relay_state=None,
                raw=data
            )

        # ======== V/P ========
        if header == b"\x30\x16\x01":
            p = struct.unpack(">f", data[3:7])[0]
            p_unit = UNITS_P.get(data[7], f"0x{data[7]:02X}")

            signal = struct.unpack(">f", data[8:12])[0]
            sig_unit = UNITS_E.get(data[12], f"0x{data[12]:02X}")

            return Measurement(
                mode="V/P",
                pressure=p,
                pressure_unit=p_unit,
                signal=signal,
                signal_unit=sig_unit,
                relay_state=None,
                raw=data
            )

        # ======== Реле ========
        if header == b"\x30\x17\x01":
            p = struct.unpack(">f", data[3:7])[0]
            p_unit = UNITS_P.get(data[7], f"0x{data[7]:02X}")

            code = data[12]
            if code == 0x03:
                relay = True
                sig_unit = "замкнут"
                signal = 1.0
            elif code == 0x04:
                relay = False
                sig_unit = "разомкнут"
                signal = 0.0
            else:
                relay = None
                sig_unit = f"реле 0x{code:02X}"
                signal = float(code)

            return Measurement(
                mode="Реле",
                pressure=p,
                pressure_unit=p_unit,
                signal=signal,
                signal_unit=sig_unit,
                relay_state=relay,
                raw=data
            )

        # неизвестный пакет
        return None
