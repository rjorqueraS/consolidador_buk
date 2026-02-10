import streamlit as st
import pandas as pd
import re
from unidecode import unidecode
import io

# --- Configuración de la Página ---
st.set_page_config(page_title="Consolidador de Asistencia", page_icon="📊")

# --- Funciones de Lógica Interna ---
EQUIV = {
    "rut": ["rut", "run", "r.u.t", "r.u.n", "documento", "id", "dni", "codigo persona", "codigo trabajador"],
    "dv": ["dv", "digito verificador", "d.v", "digito"],
    "nombre": ["nombre", "nombres", "trabajador", "colaborador", "empleado", "persona"],
    "especialidad": ["especialidad", "oficio", "perfil", "cargo", "puesto", "ocupacion"],
    "contrato": ["contrato", "tipo contrato", "modalidad", "relacion laboral"],
    "supervisor": ["supervisor", "jefe directo", "jefe", "encargado", "capataz"],
    "rut_empleador": ["rut empleador", "rut empresa", "rut contratista", "rut razon social"],
}

OUT_COLS = ["rut", "nombre", "especialidad", "contrato", "supervisor", "rut_empleador", "marcas"]

def norm(s: str) -> str:
    if not s: return ""
    return re.sub(r"[\s\.\-_/:;]+", "", unidecode(str(s)).lower())

def limpiar_rut(rut_val, dv_val=None):
    if pd.isna(rut_val): return None
    s = str(rut_val).upper().replace(".", "").replace("-", "").replace(" ", "").strip()
    cuerpo = "".join(filter(str.isdigit, s))
    if not cuerpo: return None
    
    if dv_val and not pd.isna(dv_val):
        dv = str(dv_val).strip().upper()[-1]
    else:
        dv = s[-1]
        cuerpo = cuerpo[:-1] if len(cuerpo) > 1 else cuerpo
    return f"{cuerpo}-{dv}"

# --- Interfaz de Streamlit ---
st.title("📊 Consolidador de Asistencia")
st.markdown("Sube tus archivos Excel (.xlsx o .xls) para unificarlos en un solo reporte detallado.")

uploaded_files = st.file_uploader("Arrastra aquí tus archivos", accept_multiple_files=True, type=['xlsx', 'xls'])

if uploaded_files:
    datos_acumulados = []
    
    with st.spinner('Procesando archivos...'):
        for uploaded_file in uploaded_files:
            try:
                # Determinar motor
                engine = "openpyxl" if uploaded_file.name.endswith('x') else "xlrd"
                xls = pd.ExcelFile(uploaded_file, engine=engine)
                
                for hoja in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=hoja)
                    if df.empty: continue

                    # Normalización de columnas
                    cols_actuales = {norm(c): c for c in df.columns}
                    mapeo = {}
                    for canon, variantes in EQUIV.items():
                        for v in variantes:
                            if norm(v) in cols_actuales:
                                mapeo[canon] = cols_actuales[norm(v)]
                                break
                    
                    if "rut" in mapeo or "nombre" in mapeo:
                        temp = pd.DataFrame()
                        r_col = mapeo.get("rut")
                        d_col = mapeo.get("dv")
                        
                        if r_col:
                            temp["rut"] = [limpiar_rut(row[r_col], row[d_col] if d_col else None) for _, row in df.iterrows()]
                        else:
                            temp["rut"] = None

                        for campo in ["nombre", "especialidad", "contrato", "supervisor"]:
                            col_real = mapeo.get(campo)
                            temp[campo] = df[col_real].astype(str).str.strip().replace(['nan', 'None', 'NaN'], None) if col_real else None
                        
                        re_col = mapeo.get("rut_empleador")
                        temp["rut_empleador"] = df[re_col].apply(limpiar_rut) if re_col else None
                        
                        temp["__marca"] = 1
                        temp["__key"] = temp["rut"].fillna(temp["nombre"].apply(lambda x: norm(x) if x else None))
                        
                        datos_acumulados.append(temp.dropna(subset=["__key"]))
            except Exception as e:
                st.error(f"Error en {uploaded_file.name}: {e}")

    if datos_acumulados:
        full_df = pd.concat(datos_acumulados, ignore_index=True)
        
        # Agrupación
        def pick_first(s):
            valid = s.dropna()
            return valid.iloc[0] if not valid.empty else None

        res = full_df.groupby("__key").agg({
            "rut": pick_first,
            "nombre": pick_first,
            "especialidad": pick_first,
            "contrato": pick_first,
            "supervisor": pick_first,
            "rut_empleador": pick_first,
            "__marca": "sum"
        }).reset_index(drop=True)

        res = res.rename(columns={"__marca": "marcas"})
        res = res.sort_values(by=["nombre"])
        
        # Mostrar vista previa
        st.success(f"✅ ¡Procesado con éxito! {len(res)} trabajadores encontrados.")
        st.dataframe(res[OUT_COLS].head(10)) # Muestra los primeros 10

        # Botón de Descarga
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            res[OUT_COLS].to_excel(writer, index=False, sheet_name='Consolidado')
        
        st.download_button(
            label="📥 Descargar Excel Consolidado",
            data=output.getvalue(),
            file_name="consolidado_asistencia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No se pudo extraer información de los archivos subidos.")
