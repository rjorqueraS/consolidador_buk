# app.py
import os
import re
from io import BytesIO
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st
from unidecode import unidecode
from dateutil import parser as dtparser
from datetime import datetime, time


# ------- Equivalencias de encabezados -------
EQUIV: Dict[str, List[str]] = {
    "rut": [
        "rut", "run", "r.u.t", "r.u.n", "documento", "doc", "id", "dni",
        "rut trabajador", "run trabajador", "cod. pers", "cod pers", "codigo persona", "codigo trabajador"
    ],
    "dv": ["dv", "digito verificador", "d.v", "digito", "dígito"],
    "nombre": ["nombre", "nombres", "trabajador", "colaborador", "empleado", "persona", "nombre trabajador"],

    "especialidad": [
        "especialidad", "especialización", "especialidad laboral", "oficio", "perfil", "cargo",
        "puesto", "ocupacion", "ocupación"
    ],
    "contrato": [
        "contrato", "tipo contrato", "tipocontrato", "modalidad", "tipo de contrato", "relacion laboral",
        "relación laboral"
    ],
    "supervisor": [
        "supervisor", "jefe directo", "jefe", "encargado", "capataz", "mandante supervisor", "line manager"
    ],
    "rut_empleador": [
        "rut empleador", "rut empresa", "rut contratista", "rut contratante", "rut razon social",
        "rut razón social", "rut de la empresa", "rut del empleador"
    ],

    # (no se exportan, solo por robustez)
    "fecha": ["fecha", "dia", "día", "fecha jornada", "f asistencia"],
    "entrada": ["entrada", "hora entrada", "ingreso", "checkin", "h.entrada"],
    "salida": ["salida", "hora salida", "egreso", "checkout", "h.salida"],
}

OUT_COLS = ["rut", "nombre", "especialidad", "contrato", "supervisor", "rut_empleador", "marcas"]


def norm(s: str) -> str:
    if s is None:
        return ""
    s = unidecode(str(s)).lower()
    s = re.sub(r"[\s\.\-_/:;]+", "", s)
    return s


def map_columns(cols: List[str]) -> Dict[str, str]:
    nc = {norm(c): c for c in cols}
    mapping: Dict[str, str] = {}
    for canon, variants in EQUIV.items():
        for v in variants:
            if norm(v) in nc:
                mapping[canon] = nc[norm(v)]
                break
    return mapping


def clean_text(x) -> Optional[str]:
    if pd.isna(x):
        return None
    s = str(x).strip()
    return s if s else None


def parse_fecha(x) -> Optional[pd.Timestamp]:
    if pd.isna(x):
        return None
    if isinstance(x, (pd.Timestamp, datetime)):
        return pd.to_datetime(x).normalize()
    try:
        dt = dtparser.parse(str(x), dayfirst=True, fuzzy=True)
        return pd.Timestamp(dt.date())
    except Exception:
        return None


def parse_hora(x) -> Optional[str]:
    if pd.isna(x):
        return None
    if isinstance(x, (pd.Timestamp, datetime)):
        return f"{x.hour:02d}:{x.minute:02d}"
    if isinstance(x, time):
        return f"{x.hour:02d}:{x.minute:02d}"
    try:
        v = float(x)
        if 0 <= v < 2:
            total_min = int(round(v * 24 * 60))
            hh, mm = divmod(total_min, 60)
            return f"{hh%24:02d}:{mm:02d}"
    except Exception:
        pass
    s = str(x).strip().replace(".", ":").replace(";", ":")
    m = re.match(r"^(\d{1,2})[:h]?(\d{2})$", s)
    if m:
        hh, mm = int(m.group(1)), int(m.group(2))
        if 0 <= hh < 24 and 0 <= mm < 60:
            return f"{hh:02d}:{mm:02d}"
    digits = re.sub(r"\D", "", s)
    if len(digits) in (3, 4):
        if len(digits) == 3:
            digits = "0" + digits
        hh, mm = int(digits[:2]), int(digits[2:])
        if 0 <= hh < 24 and 0 <= mm < 60:
            return f"{hh:02d}:{mm:02d}"
    return None


# --- RUT helpers ---
_RUT_DIGITS = re.compile(r"[.\-]")


def normalize_rut_str(rut_str: Optional[str], dv_str: Optional[str] = None) -> Optional[str]:
    if rut_str is None and dv_str is None:
        return None
    if rut_str is None:
        return None

    s = str(rut_str).strip().upper().replace(" ", "")
    s_no_sep = _RUT_DIGITS.sub("", s)

    if dv_str is not None and str(dv_str).strip():
        dv = str(dv_str).strip().upper()[-1:]
        cuerpo = re.sub(r"\D", "", s_no_sep)
        if not cuerpo:
            return None
        return f"{cuerpo}-{dv}"

    if "-" in s:
        p = s.split("-")
        cuerpo = re.sub(r"\D", "", p[0])
        dv = p[1][-1:].upper() if len(p) > 1 else ""
        if not cuerpo or not dv:
            return None
        return f"{cuerpo}-{dv}"

    cuerpo = re.sub(r"\D", "", s_no_sep)
    return cuerpo or None


def rut_group_key(rut_normalizado: Optional[str], nombre: Optional[str]) -> Optional[str]:
    if rut_normalizado:
        return rut_normalizado.split("-")[0]
    if nombre:
        return unidecode(nombre).strip().lower()
    return None


def process_sheet(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.dropna(how="all")
    if df.empty:
        return pd.DataFrame()

    mapping = map_columns(list(df.columns))

    rut_col = mapping.get("rut")
    dv_col = mapping.get("dv")
    nombre_col = mapping.get("nombre")
    esp_col = mapping.get("especialidad")
    contrato_col = mapping.get("contrato")
    sup_col = mapping.get("supervisor")
    rut_emp_col = mapping.get("rut_empleador")

    out = pd.DataFrame()
    out["rut"] = None

    if rut_col:
        out["rut"] = [
            normalize_rut_str(r, df[dv_col].iloc[i] if dv_col else None)
            for i, r in enumerate(df[rut_col])
        ]

    out["nombre"] = df[nombre_col].apply(clean_text) if nombre_col else None
    out["especialidad"] = df[esp_col].apply(clean_text) if esp_col else None
    out["contrato"] = df[contrato_col].apply(clean_text) if contrato_col else None
    out["supervisor"] = df[sup_col].apply(clean_text) if sup_col else None
    out["rut_empleador"] = [normalize_rut_str(x) for x in df[rut_emp_col]] if rut_emp_col else None

    out = out[(out["rut"].notna()) | (out["nombre"].notna())]
    if out.empty:
        return pd.DataFrame()

    out["__marca"] = 1
    out["__key"] = [
        rut_group_key(out["rut"].iloc[i], out["nombre"].iloc[i] if "nombre" in out else None)
        for i in range(len(out))
    ]
    out = out[out["__key"].notna()]
    return out


def _sniff_excel_kind(file_bytes: bytes, filename: str) -> str:
    """
    Devuelve 'xlsx' o 'xls' intentando adivinar por firma.
    - xlsx es zip: comienza con b'PK'
    - xls clásico es OLE2: comienza con D0 CF 11 E0
    """
    head = file_bytes[:8]
    if head.startswith(b"PK"):
        return "xlsx"
    if head.startswith(b"\xD0\xCF\x11\xE0"):
        return "xls"

    # fallback por extensión si la firma no ayudó
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".xlsx":
        return "xlsx"
    if ext == ".xls":
        return "xls"
    return "unknown"


def process_excel_bytes(file_bytes: bytes, filename: str) -> Tuple[pd.DataFrame, List[str]]:
    logs: List[str] = []
    kind = _sniff_excel_kind(file_bytes, filename)

    # Elegimos engine por tipo real
    if kind == "xlsx":
        engine = "openpyxl"
    elif kind == "xls":
        engine = "xlrd"
    else:
        return pd.DataFrame(), [f"[ERROR] Formato no soportado: {filename}"]

    # Leer workbook
    try:
        xls = pd.ExcelFile(BytesIO(file_bytes), engine=engine)
        sheets = xls.sheet_names
    except Exception as e:
        # fallback: si venía mal renombrado, intentamos el otro engine
        try:
            alt_engine = "xlrd" if engine == "openpyxl" else "openpyxl"
            xls = pd.ExcelFile(BytesIO(file_bytes), engine=alt_engine)
            sheets = xls.sheet_names
            engine = alt_engine
            logs.append(f"[AVISO] {filename}: se detectó extensión engañosa, se leyó con '{engine}'.")
        except Exception as e2:
            return pd.DataFrame(), [f"[ERROR] No se pudo abrir {filename}: {e2}"]

    parts = []
    for sh in sheets:
        try:
            df = pd.read_excel(BytesIO(file_bytes), sheet_name=sh, engine=engine)
            part = process_sheet(df)
            if not part.empty:
                parts.append(part)
            else:
                logs.append(f"[AVISO] Hoja sin filas válidas: {filename} :: {sh}")
        except Exception as e:
            logs.append(f"[ERROR] Falló al leer hoja '{sh}' en {filename}: {e}")

    if parts:
        return pd.concat(parts, ignore_index=True), logs
    return pd.DataFrame(), logs


def first_non_null(series: pd.Series) -> Optional[str]:
    for v in series:
        if pd.notna(v) and str(v).strip():
            return v
    return None


def build_output_excel(agg: pd.DataFrame, logs: List[str]) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        agg[OUT_COLS].to_excel(xw, index=False, sheet_name="Consolidado")
        if logs:
            pd.DataFrame({"log": logs}).to_excel(xw, index=False, sheet_name="Log")
    out.seek(0)
    return out.getvalue()


def run_app():
    st.set_page_config(page_title="Consolidado Asistencia", layout="wide")
    st.title("📋 Consolidado de Asistencia (Excel)")
    st.write("Sube uno o varios archivos **.xls** o **.xlsx**. La app consolida a 1 fila por RUT (o nombre si no hay RUT) y cuenta marcas.")

    uploaded = st.file_uploader(
        "Subir archivos de asistencia",
        type=["xls", "xlsx"],
        accept_multiple_files=True
    )

    if not uploaded:
        st.info("Carga archivos para comenzar.")
        return

    if st.button("Procesar", type="primary"):
        logs: List[str] = []
        chunks = []

        with st.spinner("Procesando archivos..."):
            for uf in uploaded:
                df, l = process_excel_bytes(uf.getvalue(), uf.name)
                logs.extend(l)
                if not df.empty:
                    chunks.append(df)
                else:
                    logs.append(f"[AVISO] Sin datos consolidados en: {uf.name}")

        if not chunks:
            st.error("No se consolidó ninguna fila. Revisa el log.")
            st.text_area("Log", "\n".join(logs), height=280)
            return

        raw = pd.concat(chunks, ignore_index=True)

        grp = raw.groupby("__key", dropna=False)
        agg = pd.DataFrame({
            "rut": grp["rut"].agg(first_non_null),
            "nombre": grp["nombre"].agg(first_non_null),
            "especialidad": grp["especialidad"].agg(first_non_null),
            "contrato": grp["contrato"].agg(first_non_null),
            "supervisor": grp["supervisor"].agg(first_non_null),
            "rut_empleador": grp["rut_empleador"].agg(first_non_null),
            "marcas": grp["__marca"].sum()
        }).reset_index(drop=True)

        if agg["rut"].notna().any():
            agg = agg.sort_values(["rut", "nombre"], kind="mergesort")
        else:
            agg = agg.sort_values(["nombre"], kind="mergesort")

        st.success(f"✅ Listo: {len(agg)} trabajadores; {raw.shape[0]} marcas")
        st.dataframe(agg[OUT_COLS], use_container_width=True)

        excel_bytes = build_output_excel(agg, logs)
        st.download_button(
            "⬇️ Descargar consolidado_asistencia.xlsx",
            data=excel_bytes,
            file_name="consolidado_asistencia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if logs:
            st.text_area("Log", "\n".join(logs), height=280)


if __name__ == "__main__":
    run_app()
