from __future__ import annotations

import re
from datetime import datetime
from decimal import Decimal, InvalidOperation


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_digits(value: object) -> str:
    return re.sub(r"\D+", "", normalize_text(value))


def normalize_ruc(value: object) -> str:
    digits = normalize_digits(value)
    return digits if len(digits) == 13 else ""


def validate_ruc(value: object) -> bool:
    return bool(normalize_ruc(value))


def normalize_numero_comprobante(value: object) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    match = re.search(r"(\d{3})\D?(\d{3})\D?(\d{1,9})", text)
    if not match:
        return ""
    estab, punto, sec = match.groups()
    return f"{estab}-{punto}-{sec.zfill(9)}"


def validate_numero_comprobante(value: object) -> bool:
    return bool(normalize_numero_comprobante(value))


def normalize_establecimiento(value: object) -> str:
    digits = normalize_digits(value)
    return digits[:3].zfill(3) if 1 <= len(digits) <= 3 else ""


def normalize_punto_emision(value: object) -> str:
    digits = normalize_digits(value)
    return digits[:3].zfill(3) if 1 <= len(digits) <= 3 else ""


def normalize_secuencial(value: object) -> str:
    digits = normalize_digits(value)
    return digits.zfill(9) if 1 <= len(digits) <= 9 else ""


def _clave_mod11_valida(clave48: str, digito_verificador: str) -> bool:
    factores = [2, 3, 4, 5, 6, 7]
    total = 0
    idx_factor = 0
    for digito in reversed(clave48):
        total += int(digito) * factores[idx_factor]
        idx_factor = (idx_factor + 1) % len(factores)
    modulo = 11 - (total % 11)
    if modulo == 11:
        esperado = 0
    elif modulo == 10:
        esperado = 1
    else:
        esperado = modulo
    return str(esperado) == str(digito_verificador)


def normalize_clave_acceso(value: object) -> str:
    digits = normalize_digits(value)
    if len(digits) != 49:
        return ""
    if not _clave_mod11_valida(digits[:48], digits[48]):
        return ""
    return digits


def validate_clave_acceso(value: object) -> bool:
    return bool(normalize_clave_acceso(value))


def normalize_obligado_contabilidad(value: object) -> str:
    text = normalize_text(value).upper()
    if not text:
        return ""
    if "SI" in text:
        return "SI"
    if "NO" in text:
        return "NO"
    return ""


def validate_obligado_contabilidad(value: object) -> bool:
    return bool(normalize_obligado_contabilidad(value))


def normalize_fecha(value: object) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    formatos = (
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%d/%m/%Y %H:%M:%S",
        "%d-%m-%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d-%m-%Y %H:%M",
    )
    for fmt in formatos:
        try:
            dt = datetime.strptime(text, fmt)
            return dt.strftime("%d/%m/%Y" if "H" not in fmt else "%d/%m/%Y %H:%M:%S")
        except ValueError:
            continue
    match = re.search(r"(\d{2}[/-]\d{2}[/-]\d{4}(?:\s+\d{2}:\d{2}(?::\d{2})?)?)", text)
    if not match:
        return ""
    raw = match.group(1).replace("-", "/")
    if len(raw.split()) == 2 and len(raw.split()[1].split(":")) == 2:
        raw = f"{raw}:00"
    return normalize_fecha(raw)


def validate_fecha(value: object) -> bool:
    return bool(normalize_fecha(value))


def normalize_decimal(value: object) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = normalize_text(value)
    if not text:
        return None
    candidates = [
        text,
        text.replace(".", "").replace(",", "."),
        text.replace(",", ""),
    ]
    for cand in candidates:
        try:
            return float(Decimal(cand))
        except (InvalidOperation, ValueError):
            continue
    return None


def validate_total(value: object) -> bool:
    decimal = normalize_decimal(value)
    return decimal is not None
