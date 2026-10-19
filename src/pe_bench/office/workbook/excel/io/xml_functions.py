import binascii
from base64 import b64decode
from datetime import datetime
from typing import TypeAlias

import lxml.etree

safe_parser = lxml.etree.XMLParser(resolve_entities=False)

Element: TypeAlias = lxml.etree._Element


def from_string(string: bytes) -> Element:  # type: ignore
    return lxml.etree.fromstring(string, parser=safe_parser)


def is_true(value: str | None) -> bool:
    return value in ("1", "true")


def safe_int(value: str | None, default: int = 0) -> int:
    if value is None:
        return default
    try:
        return int(value)
    except ValueError:
        return default


def safe_float(value: str | None, default: float = 0.0) -> float:
    if value is None:
        return default
    try:
        return float(value)
    except ValueError:
        return default


def safe_datetime(value: str | None) -> datetime | None:
    if value is None:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


def decode_hex_binary(hex_value: str) -> bytes | None:
    try:
        return bytes.fromhex(hex_value)
    except ValueError:
        return None


def decode_base64_binary(base64_value: str) -> bytes | None:
    try:
        return b64decode(base64_value)
    except (binascii.Error, ValueError, TypeError):
        return None
