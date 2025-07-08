# schedule_app.py – HORARIO ENLACES CAV (versión final completa)
# ---------------------------------------------------------------------------
# Ejecutar:
#     streamlit run schedule_app.py
# Requiere Streamlit ≥ 1.29
# ---------------------------------------------------------------------------
"""
Aplicación Streamlit para gestionar reservas de salas de enlaces y recursos
tecnológicos del Colegio Antonio Varas (2° semestre 2025).
Secciones:
▶ Registrar   – Añadir nueva reserva.
📂 Base datos  – Ver, editar, eliminar y descargar informe.
📅 Semana     – Vista semanal (lunes-viernes, 6 bloques).
"""
from __future__ import annotations
import hashlib
# Helper para parsear fechas en todo el script

def parse_date(val: str | dt.date | dt.datetime) -> dt.date:
    """
    Convierte un string, date o datetime a dt.date.
    Acepta formatos: 'DD/MM/YYYY', 'YYYY-MM-DD', con o sin hora.
    """
    if isinstance(val, dt.datetime): return val.date()
    if isinstance(val, dt.date): return val
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%d/%m/%Y %H:%M", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"): 
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Formato de fecha inválido: {s}")
import datetime as dt
import time
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
import xlsxwriter  # type: ignore

# Configuración global
EXCEL = "Recursos.xlsx"
SHEET = "Reservas"
DATE_FMT = "%d/%m/%Y"
BURGUNDY = "#800000"

# Bloques horarios (6 bloques)
BLOQUES = [
    (dt.time(8, 0),  dt.time(9, 30)),   # Bloque 1
    (dt.time(9, 45), dt.time(11, 15)),  # Bloque 2
    (dt.time(11, 30),dt.time(13, 0)),   # Bloque 3
    (dt.time(14, 0), dt.time(15, 30)),  # Bloque 4
    (dt.time(15, 45),dt.time(16, 30)),  # Bloque 5
    (dt.time(16, 30),dt.time(18, 30)),  # Bloque 6
]

# Helpers
def as_time(val: str | dt.time) -> dt.time:
    """
    Convierte un string o dt.time a dt.time.
    Acepta formatos 'HH:MM:SS', 'HH:MM' o 'H'.
    """
    if isinstance(val, dt.time):
        return val
    s = str(val).strip()
    for fmt in ("%H:%M:%S", "%H:%M", "%H"):
        try:
            return dt.datetime.strptime(s, fmt).time()
        except ValueError:
            continue
    raise ValueError(f"Formato de hora inválido: {s}")

def overlap(hi1: dt.time, hf1: dt.time, hi2: dt.time, hf2: dt.time) -> bool:
    return max(hi1, hi2) < min(hf1, hf2)

# Inicializar archivo Excel si no existe
if not Path(EXCEL).exists():
    cols = ["Fecha", "Hora inicio", "Hora fin", "Profesor", "Curso", "Recurso", "Observaciones"]
    pd.DataFrame(columns=cols).to_excel(EXCEL, sheet_name=SHEET, index=False)

# Cargar datos
df = pd.read_excel(EXCEL, sheet_name=SHEET, dtype=str).fillna("")
xl = pd.ExcelFile(EXCEL)

# Dinámicas para formularios
def recalc_lists(df: pd.DataFrame) -> tuple[list[str], list[str], list[str]]:
    get_col = lambda s: xl.parse(s).iloc[:, 0].dropna().astype(str).str.strip().tolist() if s in xl.sheet_names else []
    profs = sorted(set(get_col("Profesores")) | set(df["Profesor"].unique()))
    cursos = sorted(set(get_col("Cursos")) | set(df["Curso"].unique()))
    recs = sorted(set(get_col("Recursos")) | set(df["Recurso"].str.split(", ").explode().dropna().unique()))
    return profs, cursos, recs

PROFESORES, CURSOS, RECURSOS = recalc_lists(df)

# Guardado atómico
def atomic_save(df_save: pd.DataFrame) -> None:
    for _ in range(5):
        try:
            wb = load_workbook(EXCEL)
            if SHEET in wb.sheetnames:
                wb.remove(wb[SHEET])
            ws = wb.create_sheet(SHEET, 0)
            for c, col in enumerate(df_save.columns, 1):
                ws.cell(1, c, col)
            for r, row in enumerate(df_save.itertuples(index=False), 2):
                for c, val in enumerate(row, 1):
                    ws.cell(r, c, val)
            wb.save(EXCEL)
            break
        except PermissionError:
            time.sleep(0.2)

# Informe Excel + gráficos
def build_report(df_report: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df_report.to_excel(writer, sheet_name="Reservas", index=False)
        recurso = df_report["Recurso"].str.split(", ").explode()
        recurso.value_counts().rename("Reservas").to_excel(writer, sheet_name="Uso recurso")
        df_report["Profesor"].value_counts().rename("Reservas").to_excel(writer, sheet_name="Uso profesor")
        wb = writer.book
        sh = wb.add_worksheet("Gráficos")
        # Gráfico uso recurso
        uc = recurso.value_counts()
        pie1 = wb.add_chart({"type": "pie"})
        pie1.add_series({
            "categories": f"='Uso recurso'!$A$2:$A${uc.size+1}",
            "values":     f"='Uso recurso'!$B$2:$B${uc.size+1}",
            "data_labels": {"percentage": True},
        })
        pie1.set_title({"name": "Uso recurso"})
        sh.insert_chart("A1", pie1)
        # Gráfico uso profesor
        up = df_report["Profesor"].value_counts()
        pie2 = wb.add_chart({"type": "pie"})
        pie2.add_series({
            "categories": f"='Uso profesor'!$A$2:$A${up.size+1}",
            "values":     f"='Uso profesor'!$B$2:$B${up.size+1}",
            "data_labels": {"percentage": True},
        })
        pie2.set_title({"name": "Uso profesor"})
        sh.insert_chart("H1", pie2)
    return buf.getvalue()

# Estilos y cabecera
st.set_page_config("HORARIO ENLACES CAV", page_icon="📅", layout="wide")
CSS = f"""
html, body, [data-testid='stApp'] {{ background:#fff; color:#000; }}
.stButton>button, .stDownloadButton>button {{ background:{BURGUNDY}; color:#fff; }}
section[data-testid='stSidebar'] {{ display:none; }}
[data-testid='stAppViewContainer']>div {{ padding:0 3rem; }}
label {{ color:#000!important; font-weight:600; }}
.stTableContainer .stTable th, .stTableContainer .stTable td {{ border:1px solid #000!important; color:#000!important; }}
[data-testid='stDataEditorContainer'] * {{ color:#000!important; }}
div[role='tablist'] > button[role='tab'] {{ color:#000!important; }}
/* Estilo para listas desplegables en data editor */
[data-testid='stDataEditorContainer'] select, [data-testid='stDataEditorContainer'] option {{ color:#fff !important; background:#000 !important; }}
/* Ajuste de duración para toasts */
.stToast > div {{ animation-duration: 8s !important; }}
/* Dark mode */
@media (prefers-color-scheme: dark) {{
  html, body, [data-testid='stApp'] {{ background:#000; color:#fff; }}
  .stButton>button, .stDownloadButton>button {{ background:#5c0000; color:#fff; }}
  label {{ color:#fff!important; }}
  .stTableContainer .stTable th, .stTableContainer .stTable td {{ border:1px solid #fff!important; color:#fff!important; }}
  [data-testid='stDataEditorContainer'] * {{ color:#fff!important; }}
  div[role='tablist'] > button[role='tab'] {{ color:#fff!important; }}
}}
"""
st.markdown(f"<style>{CSS}</style>", unsafe_allow_html=True)
st.markdown(f"<h1 style='text-align:center;color:{BURGUNDY};'>📅 HORARIO ENLACES CAV 💻</h1>", unsafe_allow_html=True)

# Pestañas
# Agregar pestaña de Mantenimientos
tab_reg, tab_db, tab_week, tab_maint = st.tabs([
    "▶ Registrar",
    "📂 Base datos",
    "📅 Semana",
    "🔧 Mantenimiento"
])

# Función toast
def toast(msg: str, kind: str = "info") -> None:
    icons = {"success": "✅", "error": "❌", "warning": "⚠️", "info": "ℹ️"}
    try:
        st.toast(msg, icon=icons.get(kind))
    except Exception:
        getattr(st, kind)(msg)

# ▶ Registrar
with tab_reg:
    st.markdown("<h2 style='color:#000;'>▶ Registrar nueva reserva</h2>", unsafe_allow_html=True)
    # Actualizar listas dinámicas
    PROFESORES, CURSOS, RECURSOS = recalc_lists(df)
    # Ordenar cursos: Básico, Medio, Dif
    def course_key(c: str) -> tuple[int, str]:
        cu = c.upper()
        if "BÁSIC" in cu or "BASIC" in cu: return (0, cu)
        if "MEDIO" in cu: return (1, cu)
        if "DIF" in cu: return (2, cu)
        return (3, cu)
    CURSOS = sorted(CURSOS, key=course_key)

    # Formulario: fecha, hora, observaciones
    c1, c2 = st.columns(2)
    with c1:
        fecha = st.date_input(
            "Fecha", dt.date.today(), format="DD/MM/YYYY",
            help="Selecciona la fecha de la reserva"
        )
        hi = st.time_input("Hora inicio", BLOQUES[0][0])
        hf = st.time_input("Hora fin",    BLOQUES[0][1])
    with c2:
        prof = st.selectbox("Profesor", PROFESORES)
        curso = st.selectbox("Curso", CURSOS)
        # Recursos en mantenimiento (si hay hoja Mantenimientos)
        unavail = []
        if "Mantenimientos" in xl.sheet_names:
            mant = xl.parse("Mantenimientos").fillna("")
            col_start = next((c for c in mant.columns if 'fecha' in c.lower() and 'inicio' in c.lower()), None)
            col_end   = next((c for c in mant.columns if 'fecha' in c.lower() and 'fin' in c.lower()), None)
            if col_start and col_end:
                mant['Inicio'] = mant[col_start].apply(parse_date)
                mant['Fin']    = mant[col_end].apply(parse_date)
                unavail = mant[(mant['Inicio'] <= fecha) & (mant['Fin'] >= fecha)]['Recurso'].astype(str).tolist()
        if unavail:
            st.warning(f"Recursos en mantenimiento: {', '.join(unavail)}")
        # Seleccionar recursos disponibles junto a Hora fin("Curso", CURSOS)
        # Seleccionar recursos disponibles junto a Hora fin
        available_recs = [r for r in RECURSOS if r not in unavail]
        recs = st.multiselect("Recursos", available_recs)
    # Observaciones
    obs = st.text_area("Observaciones", height=80, key="obs_reg")
    

    # Guardar reserva
    if st.button("💾 Guardar reserva", use_container_width=True):
        # Validaciones básicas
        if hi >= hf:
            toast("Hora inicio debe ser antes de fin.", "warning"); st.stop()
        if not recs:
            toast("Selecciona al menos un recurso.", "warning"); st.stop()
        fstr = fecha.strftime(DATE_FMT)
        # Detectar conflictos por recurso
        conflict_list: list[str] = []
        rows_today = df[df["Fecha"] == fstr]
        for rsrc in recs:
            for _, old in rows_today.iterrows():
                old_recs = [x.strip() for x in str(old["Recurso"]).split(",")]
                if rsrc in old_recs and overlap(
                    hi, hf,
                    as_time(old["Hora inicio"]),
                    as_time(old["Hora fin"])
                ):
                    conflict_list.append(rsrc)
                    break
        if conflict_list:
            toast(f"Choque: recursos ocupados: {', '.join(conflict_list)}.", "error")
            # Ofrecer bloques alternativos
            alt: list[str] = []
            for idx, (h1, h2) in enumerate(BLOQUES, start=1):
                libre = True
                for rsrc in recs:
                    for _, old in rows_today.iterrows():
                        old_recs = [x.strip() for x in str(old["Recurso"]).split(",")]
                        if rsrc in old_recs and overlap(h1, h2,
                                        as_time(old["Hora inicio"]),
                                        as_time(old["Hora fin"])):
                            libre = False
                            break
                    if not libre:
                        break
                if libre:
                    alt.append(f"Bloque {idx} ({h1.strftime('%H:%M')}-{h2.strftime('%H:%M')})")
            if alt:
                st.info(f"Bloques alternativos disponibles: {', '.join(alt)}")
            st.stop()
        # Crear nueva reserva
        new = pd.DataFrame([{
            "Fecha": fstr,
            "Hora inicio": hi.strftime("%H:%M"),
            "Hora fin": hf.strftime("%H:%M"),
            "Profesor": prof,
            "Curso": curso,
            "Recurso": ", ".join(recs),
            "Observaciones": obs.strip(),
        }])
        df = pd.concat([df, new], ignore_index=True)
        atomic_save(df)
        toast("✅ Reserva guardada.", "success")

# Generar archivo .ics para calendario
def build_ics(df_events: pd.DataFrame, prof: str) -> str:
    """
    Genera un archivo iCalendar (.ics) con las reservas filtradas por profesor.
    """
    now = dt.datetime.utcnow()
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//CAV//Horario Enlaces//ES",
        "CALSCALE:GREGORIAN",
    ]
    for idx, row in df_events.iterrows():
        # Asegurar fecha como dt.date y horas como dt.time
        date_obj = parse_date(row['Fecha'])
        start_time = as_time(row['Hora inicio'])
        end_time   = as_time(row['Hora fin'])
        # Combinar fecha y hora
        start_dt = dt.datetime.combine(date_obj, start_time)
        end_dt   = dt.datetime.combine(date_obj, end_time)
        uid = f"{prof}-{idx}@cav.cl"
        summary = f"Reserva {row['Recurso']} ({row['Curso']})"
        description = row.get('Observaciones', '')
        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{now.strftime('%Y%m%dT%H%M%SZ')}",
            f"DTSTART;TZID=America/Santiago:{start_dt.strftime('%Y%m%dT%H%M%S')}",
            f"DTEND;TZID=America/Santiago:{end_dt.strftime('%Y%m%dT%H%M%S')}",
            f"SUMMARY:{summary}",
            f"DESCRIPTION:{description}",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    return "\n".join(lines)
# 📂 Base datos
with tab_db:
    st.markdown("<h2 style='color:#000;'>📂 Base datos de reservas</h2>", unsafe_allow_html=True)
    # Copia de datos para edición
    editor = df.copy()
    # Convertir Tipo de datos para editor
    editor['Fecha'] = pd.to_datetime(
        editor['Fecha'], dayfirst=True, infer_datetime_format=True
    ).dt.date
    editor['Hora inicio'] = editor['Hora inicio'].apply(as_time)
    editor['Hora fin']    = editor['Hora fin'].apply(as_time)
    # Configuración de columnas
    cfg = {
        "Fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
        "Hora inicio": st.column_config.TimeColumn("Hora inicio", format="HH:mm"),
        "Hora fin": st.column_config.TimeColumn("Hora fin", format="HH:mm"),
        "Profesor": st.column_config.SelectboxColumn("Profesor", options=PROFESORES),
        "Curso": st.column_config.SelectboxColumn("Curso", options=CURSOS),
        "Recurso": st.column_config.SelectboxColumn("Recurso", options=RECURSOS),
        "Observaciones": st.column_config.TextColumn("Observaciones"),
    }
    # Columnas para editor y botones de exportación
    col1, col2, col3 = st.columns([3, 1, 1])
    with col1:
        edited = st.data_editor(
            editor,
            hide_index=False,
            use_container_width=True,
            column_config=cfg,
        )
    with col2:
        rpt = build_report(df)
        st.download_button(
            "📥 Informe Excel",
            rpt,
            file_name="informe_reservas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col3:
        st.markdown("---")
        st.subheader("Exportar calendario (.ics)")
        prof_cal = st.selectbox("Seleccionar profesor", PROFESORES, key="cal_prof")
        cal_df = df[df["Profesor"] == prof_cal]
        ics_data = build_ics(cal_df, prof_cal)
        st.download_button(
            "📅 Descargar calendario",
            ics_data,
            file_name=f"{prof_cal}_reservas.ics",
            mime="text/calendar",
            use_container_width=True,
        )
    # Botones para guardar cambios y eliminar registros
    dcol1, dcol2 = st.columns(2)
    with dcol1:
        if st.button("💾 Guardar cambios", key="save_db_edits", use_container_width=True):
            atomic_save(edited)
            toast("Cambios guardados en Base datos.", "success")
    with dcol2:
        to_drop = st.multiselect(
            "Seleccionar registros a eliminar", options=edited.index, key="drop_db"
        )
        if st.button("🗑️ Eliminar registros", key="delete_db", use_container_width=True) and to_drop:
            new_df = edited.drop(index=to_drop).reset_index(drop=True)
            atomic_save(new_df)
            toast("Registros eliminados.", "success")

# 📅 Semana
with tab_week:
    st.markdown("<h2 style='color:#000;'>📅 Vista semanal</h2>", unsafe_allow_html=True)
    # Selecciona cualquier fecha de la semana (cualquier día)
    fecha_ref = st.date_input(
        "Selecciona fecha de la semana (cualquier día)",
        dt.date.today(),
        format="DD/MM/YYYY",
        help="Elige un día dentro de la semana que quieres visualizar"
    )
    # Calcular rango lunes-viernes
    start = fecha_ref - dt.timedelta(days=fecha_ref.weekday())  # lunes
    dias = [start + dt.timedelta(days=i) for i in range(5)]      # hasta viernes
    # Encabezados en español
    nombres = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
    cols = [f"{nombres[i]} {dias[i].strftime('%d/%m')}" for i in range(5)]
    # Filas: bloques 1-6
    rows = [f"Bloque {i+1}" for i in range(len(BLOQUES))]
    tabla = pd.DataFrame("", index=rows, columns=cols)
    # Rellenar datos en tabla
    for i, (hi, hf) in enumerate(BLOQUES):
        for j, d in enumerate(dias):
            sel = df[
                (df['Fecha'].apply(parse_date) == d)
                & df.apply(
                    lambda r: overlap(hi, hf, as_time(r['Hora inicio']), as_time(r['Hora fin'])),
                    axis=1
                )
            ]
            if not sel.empty:
                lines = []
                for _, row in sel.iterrows():
                    lines.append(f"{row['Hora inicio']}-{row['Hora fin']} | {row['Profesor']} – {row['Curso']} | {row['Recurso']} | {row['Observaciones']}")
                tabla.iat[i, j] = "\n\n".join(lines)
    # Colores únicos por reserva
    import hashlib
    def get_color(text: str) -> str:
        h = hashlib.md5(text.encode('utf-8')).hexdigest()
        return f"#{h[:6]}"
    def highlight_cell(val):
        if not isinstance(val, str) or not val:
            return ''
        color = get_color(val)
        return f'background-color: {color} !important; color: #000 !important;'
    styled = (
        tabla.style
        .set_properties(**{'white-space': 'pre-wrap'})
        .applymap(highlight_cell)
        .set_table_styles([
            {'selector': 'th', 'props': [('min-width', '200px')]},
            {'selector': 'td', 'props': [('min-width', '200px')]}  
        ])
        .set_table_attributes('style="table-layout:auto; width:100%;"')
    )
    html = styled.to_html()
    st.markdown(html, unsafe_allow_html=True)

# 🔧 Mantenimiento
with tab_maint:
    st.markdown("<h2 style='color:#000;'>🔧 Gestión de Mantenimiento</h2>", unsafe_allow_html=True)
    def_tab = "Mantenimientos"
    # Cargar o inicializar dataframe de mantenimientos
    if def_tab in xl.sheet_names:
        mant_df = xl.parse(def_tab).fillna("")
    else:
        mant_df = pd.DataFrame(columns=["Recurso", "FechaInicio", "HoraInicio", "FechaFin", "HoraFin"])
            # Formulario para agregar nuevo mantenimiento
    st.subheader("Agregar nuevo mantenimiento")
    # Recurso
    rsrc_maint = st.selectbox("Recurso", RECURSOS, key="maint_res")
    # Fechas de mantenimiento
    d1, d2 = st.columns(2)
    with d1:
        start_date = st.date_input("Fecha inicio", dt.date.today(), key="mant_start_date")
    with d2:
        end_date = st.date_input("Fecha fin", dt.date.today(), key="mant_end_date")
    # Horas de mantenimiento
    t1, t2 = st.columns(2)
    with t1:
        start_time = st.time_input("Hora inicio", BLOQUES[0][0], key="mant_start_time")
    with t2:
        end_time = st.time_input("Hora fin", BLOQUES[0][1], key="mant_end_time")
    # Guardar nuevo mantenimiento
    if st.button("Guardar mantenimiento", use_container_width=True, key="save_maint"):
        new_row = {
            "Recurso": rsrc_maint,
            "FechaInicio": start_date,
            "HoraInicio": start_time,
            "FechaFin": end_date,
            "HoraFin": end_time,
        }
        mant_df = pd.concat([mant_df, pd.DataFrame([new_row])], ignore_index=True)
        # Guardar hoja Mantenimientos
        with pd.ExcelWriter(EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            mant_df.to_excel(writer, sheet_name=def_tab, index=False)
        st.success("Mantenimiento registrado correctamente.")

    # Editor de mantenimientos existentes
    st.subheader("Editar o eliminar mantenimientos")
        # Convertir tipos para edición
    mant_editor = mant_df.copy()
    # Fecha
    if 'FechaInicio' in mant_editor.columns:
        mant_editor['FechaInicio'] = mant_editor['FechaInicio'].apply(parse_date)
    if 'FechaFin' in mant_editor.columns:
        mant_editor['FechaFin'] = mant_editor['FechaFin'].apply(parse_date)
    # Hora
    if 'HoraInicio' in mant_editor.columns:
        mant_editor['HoraInicio'] = mant_editor['HoraInicio'].apply(as_time)
    if 'HoraFin' in mant_editor.columns:
        mant_editor['HoraFin'] = mant_editor['HoraFin'].apply(as_time)
    # Configurar columnas
    cfg_maint = {
        "Recurso": st.column_config.SelectboxColumn("Recurso", options=RECURSOS),
        "FechaInicio": st.column_config.DateColumn("Fecha inicio", format="DD/MM/YYYY"),
        "HoraInicio": st.column_config.TimeColumn("Hora inicio", format="HH:mm"),
        "FechaFin": st.column_config.DateColumn("Fecha fin", format="DD/MM/YYYY"),
        "HoraFin": st.column_config.TimeColumn("Hora fin", format="HH:mm"),
    }
    edited = st.data_editor(mant_editor, hide_index=False, use_container_width=True, column_config=cfg_maint)
    col3, col4 = st.columns(2)
    with col3:
        if st.button("💾 Guardar cambios mantenimiento", key="save_maint_edits", use_container_width=True):
            # Guardar cambios
            with pd.ExcelWriter(EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                edited.to_excel(writer, sheet_name=def_tab, index=False)
            st.success("Cambios en mantenimiento guardados.")
    with col4:
        to_drop = st.multiselect("Seleccionar mantenimientos a eliminar", options=edited.index)
        if st.button("🗑️ Eliminar mantenimiento", key="del_maint", use_container_width=True) and to_drop:
            kept = edited.drop(index=to_drop).reset_index(drop=True)
            with pd.ExcelWriter(EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                kept.to_excel(writer, sheet_name=def_tab, index=False)
            st.success("Mantenimientos eliminados.")
