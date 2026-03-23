from __future__ import annotations

from dataclasses import dataclass, field
import re
import unicodedata

from .pdf_reader import DocumentLayout, LineBox
from .validators import (
    normalize_clave_acceso,
    normalize_decimal,
    normalize_establecimiento,
    normalize_fecha,
    normalize_numero_comprobante,
    normalize_obligado_contabilidad,
    normalize_punto_emision,
    normalize_ruc,
    normalize_secuencial,
)


def _norm(text: str) -> str:
    base = unicodedata.normalize("NFKD", text or "").encode("ascii", "ignore").decode("ascii")
    base = re.sub(r"[^A-Za-z0-9 ]+", " ", base).upper()
    return re.sub(r"\s+", " ", base).strip()


@dataclass
class FieldEvidence:
    value: object = ""
    source: str = ""
    confidence: float = 0.0


@dataclass
class ExtractionResult:
    fields: dict[str, FieldEvidence] = field(default_factory=dict)
    used_ocr: bool = False

    def set_if_better(self, name: str, value: object, source: str, confidence: float) -> None:
        if value in (None, ""):
            return
        current = self.fields.get(name)
        if current is None or confidence > current.confidence:
            self.fields[name] = FieldEvidence(value=value, source=source, confidence=confidence)

    def values_dict(self) -> dict:
        return {name: evidence.value for name, evidence in self.fields.items()}


FIELD_ALIASES = {
    "rucEmisor": ["R U C", "R.U.C", "RUC"],
    "razonSocialEmisor": ["RAZON SOCIAL", "RAZON SOCIAL EMISOR", "NOMBRES Y APELLIDOS EMISOR"],
    "nombreComercial": ["NOMBRE COMERCIAL"],
    "direccionMatrizEmisor": ["DIRECCION MATRIZ"],
    "direccionSucursalEmisor": ["DIRECCION SUCURSAL", "DIRECCION ESTABLECIMIENTO"],
    "obligadoContabilidad": ["OBLIGADO A LLEVAR CONTABILIDAD"],
    "claveAcceso": ["CLAVE DE ACCESO"],
    "numeroComprobante": ["NO", "NO.", "NUMERO", "NUMERO DE DOCUMENTO"],
    "agenteRetencion": ["AGENTE DE RETENCION RESOLUCION NO", "AGENTE DE RETENCION RESOLUCION", "AGENTE RETENCION"],
    "contribuyenteEspecial": ["CONTRIBUYENTE ESPECIAL", "CONTRIBUYENTE"],
    "tipoContribuyenteRIMPE": ["RIMPE", "CONTRIBUYENTE RIMPE", "TIPO CONTRIBUYENTE RIMPE"],
    "fechaEmision": ["FECHA EMISION", "FECHA DE EMISION"],
    "razonSocialComprador": ["RAZON SOCIAL NOMBRES Y APELLIDOS", "RAZON SOCIAL COMPRADOR", "NOMBRES Y APELLIDOS"],
    "identificacionComprador": ["IDENTIFICACION COMPRADOR", "IDENTIFICACION", "RUC C I"],
    "valorTotal": ["IMPORTE TOTAL", "VALOR TOTAL", "TOTAL"],
}

DOC_TYPE_PATTERNS = (
    ("FACTURA", "Factura"),
    ("NOTA DE CREDITO", "Nota de Credito"),
    ("NOTA DE DEBITO", "Nota de Debito"),
    ("COMPROBANTE DE RETENCION", "Retencion"),
    ("LIQUIDACION DE COMPRA", "Liquidacion de compra"),
    ("GUIA DE REMISION", "Guia de Remision"),
)


def _all_lines(layout: DocumentLayout) -> list[LineBox]:
    lines: list[LineBox] = []
    for page in layout.pages:
        lines.extend(page.lines)
    return lines


def _first_page_lines(layout: DocumentLayout) -> list[LineBox]:
    return layout.pages[0].lines if layout.pages else []


def _line_value_after_label(line: LineBox, label_aliases: list[str]) -> str:
    text = line.text.strip()
    norm = _norm(text)
    for alias in label_aliases:
        alias_norm = _norm(alias)
        if not alias_norm or alias_norm not in norm:
            continue
        if ":" in text:
            izquierda, derecha = text.split(":", 1)
            if alias_norm in _norm(izquierda):
                suffix = derecha.strip(" .:-")
                if suffix:
                    return suffix
        match = re.search(rf"(?i){re.escape(alias.strip())}\s+(.+)$", text)
        if match:
            suffix = match.group(1).strip(" .:-")
            if suffix:
                return suffix
    return ""


def _find_label_lines(lines: list[LineBox], aliases: list[str], regions: tuple[str, ...] | None = None) -> list[LineBox]:
    found: list[LineBox] = []
    aliases_norm = [_norm(alias) for alias in aliases]
    for line in lines:
        if regions and line.region not in regions:
            continue
        line_norm = _norm(line.text)
        if any(alias in line_norm for alias in aliases_norm):
            found.append(line)
    return found


def _capture_nearby_value(lines: list[LineBox], label_line: LineBox) -> tuple[str, float, str]:
    same_page = [ln for ln in lines if ln.page_index == label_line.page_index]
    line_height = max(8.0, label_line.y1 - label_line.y0)

    inline_value = _line_value_after_label(label_line, [label_line.text])
    if inline_value:
        return inline_value, 0.89, label_line.source

    right_candidates = [
        ln for ln in same_page
        if ln.x0 >= label_line.x1 - 2
        and abs(ln.y0 - label_line.y0) <= max(3.5, line_height * 0.55)
    ]
    if right_candidates:
        right_candidates.sort(key=lambda ln: (abs(ln.y0 - label_line.y0), ln.x0))
        return right_candidates[0].text.strip(), 0.92, right_candidates[0].source

    below_candidates = [
        ln for ln in same_page
        if 0 < (ln.y0 - label_line.y1) <= max(22.0, line_height * 2.2)
        and abs(ln.x0 - label_line.x0) <= 80
    ]
    if below_candidates:
        below_candidates.sort(key=lambda ln: (ln.y0, abs(ln.x0 - label_line.x0)))
        return below_candidates[0].text.strip(), 0.82, below_candidates[0].source

    region_candidates = [
        ln for ln in same_page
        if ln.region == label_line.region
        and abs(ln.y0 - label_line.y0) <= max(28.0, line_height * 2.8)
        and ln != label_line
    ]
    if region_candidates:
        region_candidates.sort(key=lambda ln: (abs(ln.y0 - label_line.y0), abs(ln.x0 - label_line.x1)))
        return region_candidates[0].text.strip(), 0.68, region_candidates[0].source
    return "", 0.0, ""


def _extract_with_label(result: ExtractionResult, lines: list[LineBox], field_name: str, aliases: list[str], regions: tuple[str, ...] | None = None) -> None:
    for label_line in _find_label_lines(lines, aliases, regions=regions):
        raw_value = _line_value_after_label(label_line, aliases)
        confidence = 0.86
        source = label_line.source
        if not raw_value:
            raw_value, confidence, source = _capture_nearby_value(lines, label_line)
        if not raw_value:
            continue
        _normalize_and_set(result, field_name, raw_value, source or label_line.source, confidence)


def _normalize_and_set(result: ExtractionResult, field_name: str, raw_value: object, source: str, confidence: float) -> None:
    value = raw_value
    if field_name == "rucEmisor":
        value = normalize_ruc(raw_value)
    elif field_name == "numeroComprobante":
        value = normalize_numero_comprobante(raw_value)
    elif field_name == "establecimiento":
        value = normalize_establecimiento(raw_value)
    elif field_name == "puntoEmision":
        value = normalize_punto_emision(raw_value)
    elif field_name == "secuencial":
        value = normalize_secuencial(raw_value)
    elif field_name == "claveAcceso":
        value = normalize_clave_acceso(raw_value)
    elif field_name == "obligadoContabilidad":
        value = normalize_obligado_contabilidad(raw_value)
    elif field_name == "fechaEmision":
        value = normalize_fecha(raw_value)
    elif field_name == "valorTotal":
        value = normalize_decimal(raw_value)
    else:
        value = str(raw_value).strip()
    if value in ("", None):
        return
    result.set_if_better(field_name, value, source, confidence)


def _extract_tipo_documento(result: ExtractionResult, lines: list[LineBox]) -> None:
    top_lines = [ln for ln in lines if ln.region in {"top_left", "top_right"}][:20]
    for line in top_lines:
        line_norm = _norm(line.text)
        for token, label in DOC_TYPE_PATTERNS:
            if token in line_norm:
                result.set_if_better("tipoDocumento", label, line.source, 0.96)
                return


def _extract_numero_comprobante(result: ExtractionResult, lines: list[LineBox]) -> None:
    _extract_with_label(result, lines, "numeroComprobante", FIELD_ALIASES["numeroComprobante"], regions=("top_right", "top_left"))
    value = result.fields.get("numeroComprobante")
    if not value:
        for line in lines[:30]:
            match = re.search(r"(\d{3}\D?\d{3}\D?\d{9})", line.text)
            if match:
                numero = normalize_numero_comprobante(match.group(1))
                if numero:
                    result.set_if_better("numeroComprobante", numero, line.source, 0.72)
                    break
    numero_ev = result.fields.get("numeroComprobante")
    if numero_ev:
        partes = str(numero_ev.value).split("-")
        if len(partes) == 3:
            result.set_if_better("establecimiento", partes[0], numero_ev.source, numero_ev.confidence)
            result.set_if_better("puntoEmision", partes[1], numero_ev.source, numero_ev.confidence)
            result.set_if_better("secuencial", partes[2], numero_ev.source, numero_ev.confidence)


def _extract_contribuyente_fields(result: ExtractionResult, lines: list[LineBox]) -> None:
    for label_line in _find_label_lines(lines, FIELD_ALIASES["contribuyenteEspecial"], regions=("top_left", "top_right")):
        raw, conf, source = _capture_nearby_value(lines, label_line)
        if not raw:
            continue
        digits = re.sub(r"\D+", "", raw)
        if digits:
            result.set_if_better("contribuyenteEspecial", digits, source, conf)
        elif "RIMPE" in _norm(raw):
            result.set_if_better("tipoContribuyenteRIMPE", raw.strip(), source, conf)
    _extract_with_label(result, lines, "tipoContribuyenteRIMPE", FIELD_ALIASES["tipoContribuyenteRIMPE"], regions=("top_left", "top_right"))


def _extract_clave(result: ExtractionResult, lines: list[LineBox]) -> None:
    _extract_with_label(result, lines, "claveAcceso", FIELD_ALIASES["claveAcceso"], regions=("top_right", "center"))
    if "claveAcceso" in result.fields:
        return
    for line in lines:
        match = re.search(r"(\d{49})", line.text)
        if match:
            clave = normalize_clave_acceso(match.group(1))
            if clave:
                result.set_if_better("claveAcceso", clave, line.source, 0.55)
                return


def extract_fields_from_layout(layout: DocumentLayout) -> ExtractionResult:
    result = ExtractionResult(used_ocr=layout.used_ocr)
    lines = _all_lines(layout)
    first_page_lines = _first_page_lines(layout)
    if not lines:
        return result

    _extract_tipo_documento(result, first_page_lines)
    _extract_with_label(result, first_page_lines, "rucEmisor", FIELD_ALIASES["rucEmisor"], regions=("top_left", "top_right"))
    _extract_with_label(result, first_page_lines, "razonSocialEmisor", FIELD_ALIASES["razonSocialEmisor"], regions=("top_left",))
    _extract_with_label(result, first_page_lines, "nombreComercial", FIELD_ALIASES["nombreComercial"], regions=("top_left",))
    _extract_with_label(result, first_page_lines, "direccionMatrizEmisor", FIELD_ALIASES["direccionMatrizEmisor"], regions=("top_left",))
    _extract_with_label(result, first_page_lines, "direccionSucursalEmisor", FIELD_ALIASES["direccionSucursalEmisor"], regions=("top_left",))
    _extract_with_label(result, first_page_lines, "obligadoContabilidad", FIELD_ALIASES["obligadoContabilidad"], regions=("top_left", "top_right"))
    _extract_with_label(result, first_page_lines, "agenteRetencion", FIELD_ALIASES["agenteRetencion"], regions=("top_left", "top_right"))
    _extract_contribuyente_fields(result, first_page_lines)
    _extract_numero_comprobante(result, first_page_lines)
    _extract_clave(result, first_page_lines)
    _extract_with_label(result, first_page_lines, "fechaEmision", FIELD_ALIASES["fechaEmision"], regions=("center", "top_right", "top_left"))
    _extract_with_label(result, first_page_lines, "razonSocialComprador", FIELD_ALIASES["razonSocialComprador"], regions=("center",))
    _extract_with_label(result, first_page_lines, "identificacionComprador", FIELD_ALIASES["identificacionComprador"], regions=("center",))
    _extract_with_label(result, first_page_lines, "valorTotal", FIELD_ALIASES["valorTotal"], regions=("bottom", "center"))

    return result
