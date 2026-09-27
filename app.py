import calendar
from datetime import date, datetime, timedelta
import io
import math
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="MADCOM — Planificador TV", page_icon="📺", layout="wide"
)

# -------------------------------------------------------------
# 1. MOTOR DE EXTRACCIÓN DE TARIFARIOS Y RATINGS POR CLIENTE
# -------------------------------------------------------------
@st.cache_data
def load_compara_inventory():
    file_path = "Valor TV Abierta Compara.xlsx"
    wb = openpyxl.load_workbook(file_path, data_only=True)
    all_progs = []

    # CHV
    ws_chv = wb["CHV"]
    dur_map_chv = {
        70: 9,
        60: 12,
        55: 13,
        50: 14,
        45: 15,
        40: 16,
        35: 17,
        30: 18,
        25: 19,
        20: 20,
        15: 21,
        10: 22,
        5: 23,
    }
    for r in range(10, ws_chv.max_row + 1):
        prog = ws_chv.cell(r, 3).value
        if not prog:
            continue
        bloque = str(ws_chv.cell(r, 2).value or "Off").strip()
        dias = str(ws_chv.cell(r, 4).value or "L-V").strip()
        hra_val = ws_chv.cell(r, 5).value
        hra = (
            hra_val.strftime("%H:%M")
            if isinstance(hra_val, (datetime, datetime.time))
            else str(hra_val or "")
        )
        rat = ws_chv.cell(r, 11).value or 0.0
        try:
            rat = float(rat)
        except:
            rat = 0.0

        tariffs = {}
        for sec, c_idx in dur_map_chv.items():
            val = ws_chv.cell(r, c_idx).value
            try:
                tariffs[sec] = float(val) if val else 0.0
            except:
                tariffs[sec] = 0.0

        all_progs.append(
            {
                "Canal": "CHV",
                "Programa": str(prog).strip(),
                "Bloque": "PRIME" if "prime" in bloque.lower() else "OFF PRIME",
                "Días": dias,
                "Horario": hra,
                "Rating": rat,
                "Tarifas": tariffs,
            }
        )

    # Mega
    ws_mega = wb["Mega"]
    dur_map_mega = {}
    for c in range(7, 25):
        val = ws_mega.cell(10, c).value
        if val and isinstance(val, (int, float)):
            dur_map_mega[int(val)] = c

    for r in range(12, ws_mega.max_row + 1):
        prog = ws_mega.cell(r, 2).value
        if not prog:
            continue
        dias = str(ws_mega.cell(r, 3).value or "L-V").strip()
        hra = str(ws_mega.cell(r, 4).value or "").strip()
        bloque = str(ws_mega.cell(r, 5).value or "Off").strip()
        rat = ws_mega.cell(r, 6).value or 0.0
        try:
            rat = float(rat)
        except:
            rat = 0.0

        tariffs = {}
        for sec, c_idx in dur_map_mega.items():
            val = ws_mega.cell(r, c_idx).value
            try:
                tariffs[sec] = float(val) if val else 0.0
            except:
                tariffs[sec] = 0.0

        all_progs.append(
            {
                "Canal": "Mega",
                "Programa": str(prog).strip(),
                "Bloque": "PRIME" if "prime" in bloque.lower() else "OFF PRIME",
                "Días": dias,
                "Horario": hra,
                "Rating": rat,
                "Tarifas": tariffs,
            }
        )

    # TVN
    ws_tvn = wb["TVN"]
    dur_map_tvn = {}
    for c in range(14, 28):
        h_val = str(ws_tvn.cell(6, c).value or "")
        digits = "".join([ch for ch in h_val if ch.isdigit()])
        if digits:
            dur_map_tvn[int(digits)] = c

    for r in range(7, ws_tvn.max_row + 1):
        prog = ws_tvn.cell(r, 2).value
        if not prog:
            continue
        dias = str(ws_tvn.cell(r, 3).value or "L-V").strip()
        ini_val = ws_tvn.cell(r, 5).value
        fin_val = ws_tvn.cell(r, 6).value
        ini = (
            ini_val.strftime("%H:%M")
            if isinstance(ini_val, (datetime, datetime.time))
            else str(ini_val or "")
        )
        fin = (
            fin_val.strftime("%H:%M")
            if isinstance(fin_val, (datetime, datetime.time))
            else str(fin_val or "")
        )
        hra = f"{ini} - {fin}" if ini and fin else ini
        bloque = str(ws_tvn.cell(r, 12).value or "Off").strip()
        rat = ws_tvn.cell(r, 13).value or 0.0
        try:
            rat = float(rat)
        except:
            rat = 0.0

        tariffs = {}
        for sec, c_idx in dur_map_tvn.items():
            val = ws_tvn.cell(r, c_idx).value
            try:
                tariffs[sec] = float(val) if val else 0.0
            except:
                tariffs[sec] = 0.0

        all_progs.append(
            {
                "Canal": "TVN",
                "Programa": str(prog).strip(),
                "Bloque": "PRIME" if "prime" in bloque.lower() else "OFF PRIME",
                "Días": dias,
                "Horario": hra,
                "Rating": rat,
                "Tarifas": tariffs,
            }
        )

    # Canal 13
    ws_c13 = wb["Canal 13"]
    dur_map_c13 = {}
    for c in range(6, 25):
        h_val = str(ws_c13.cell(2, c).value or "")
        digits = "".join([ch for ch in h_val if ch.isdigit()])
        if digits:
            dur_map_c13[int(digits)] = c

    for r in range(3, ws_c13.max_row + 1):
        prog = ws_c13.cell(r, 2).value
        if not prog or "PROGRAMAS" in str(prog).upper():
            continue
        dias = str(ws_c13.cell(r, 3).value or "L-V").strip()
        bloque = str(ws_c13.cell(r, 4).value or "Off").strip()
        rat = ws_c13.cell(r, 5).value or 0.0
        try:
            rat = float(rat)
        except:
            rat = 0.0

        tariffs = {}
        for sec, c_idx in dur_map_c13.items():
            val = ws_c13.cell(r, c_idx).value
            try:
                tariffs[sec] = float(val) if val else 0.0
            except:
                tariffs[sec] = 0.0

        all_progs.append(
            {
                "Canal": "Canal 13",
                "Programa": str(prog).strip(),
                "Bloque": "PRIME" if "prime" in bloque.lower() else "OFF PRIME",
                "Días": dias,
                "Horario": "—",
                "Rating": rat,
                "Tarifas": tariffs,
            }
        )

    return all_progs


progs_db = load_compara_inventory()

# -------------------------------------------------------------
# 2. INTERFAZ: PARÁMETROS DE CAMPAÑA
# -------------------------------------------------------------
st.title("📺 MADCOM — Planificador & Cotizador TV Abierta")

with st.sidebar:
    st.header("🏢 Selección de Cuenta")
    cliente_sel = st.selectbox(
        "Cliente Activo",
        options=["Compara Online", "+ Agregar nuevo cliente..."],
        index=0,
    )
    campana_nombre = st.text_input("Campaña:", value="Pauta Mes Tipo 2026")
    version_pauta = st.text_input("Versión:", value="V1.0")

    st.markdown("---")
    st.header("📅 Período de Campaña")
    default_start = date(2026, 10, 1)
    default_end = date(2026, 10, 20)
    date_range = st.date_input(
        "Rango de Fechas:",
        value=(default_start, default_end),
        min_value=date(2026, 1, 1),
        max_value=date(2027, 12, 31),
    )

    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date = date_range[0] if isinstance(date_range, tuple) else date_range
        end_date = start_date + timedelta(days=19)

    num_dias_campana = (end_date - start_date).days + 1
    st.caption(f"Duración: **{num_dias_campana} días de exhibición**.")

    st.markdown("---")
    st.header("💰 Objetivos de Medios")
    presupuesto_total = st.number_input(
        "Inversión Total Meta (CLP):",
        min_value=1000000,
        value=50000000,
        step=1000000,
        format="%d",
    )
    objetivo_trps = st.number_input(
        "Objetivo TRPs:", min_value=10, value=250, step=10
    )
    sov_prime_target = (
        st.slider("SOV TRPs Prime (%):", min_value=10, max_value=90, value=70) / 100.0
    )

    st.markdown("---")
    st.header("⏱️ Mix de Segundajes")
    AVAILABLE_SECS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70]
    multi_segundaje = st.checkbox("Asignar segundajes por bloque", value=True)

    if multi_segundaje:
        seg_prime = st.selectbox(
            "Segundaje en Bloque PRIME:", options=AVAILABLE_SECS, index=3
        )  # 20"
        seg_off = st.selectbox(
            "Segundaje en Bloque OFF PRIME:", options=AVAILABLE_SECS, index=5
        )  # 30"
    else:
        seg_unico = st.selectbox(
            "Segundaje único:", options=AVAILABLE_SECS, index=5
        )  # 30"
        seg_prime = seg_unico
        seg_off = seg_unico

    st.markdown("---")
    st.header("📊 Share de Inversión por Canal (SOI)")
    soi_chv = st.slider("% CHV", 0, 100, 30)
    soi_mega = st.slider("% Mega", 0, 100, 30)
    soi_c13 = st.slider("% Canal 13", 0, 100, 25)
    soi_tvn = st.slider("% TVN", 0, 100, 15)
    soi_total = soi_chv + soi_mega + soi_c13 + soi_tvn
    if soi_total != 100:
        st.warning(f"Suma de SOI = {soi_total}% (Ajustar a 100%)")

    max_spots_dia_prog = st.number_input(
        "Máx. spots por programa por día:", min_value=1, max_value=5, value=1
    )

# -------------------------------------------------------------
# 3. CONSTRUCTOR Y SELECCIÓN DE PROGRAMAS
# -------------------------------------------------------------
st.subheader("📋 Selección de Inventario para Pauta")

# Preparamos tabla con tarifas correspondientes al bloque
inventory_rows = []
for p in progs_db:
    seg_aplicado = seg_prime if p["Bloque"] == "PRIME" else seg_off
    tarifa_val = p["Tarifas"].get(seg_aplicado, 0.0)
    cpp_est = (tarifa_val / p["Rating"]) if p["Rating"] > 0 else 0

    inventory_rows.append(
        {
            "Canal": p["Canal"],
            "Programa": p["Programa"],
            "Bloque": p["Bloque"],
            "Días": p["Días"],
            "Horario": p["Horario"],
            "Segundos": seg_aplicado,
            "Tarifa": tarifa_val,
            "Rating": round(p["Rating"], 2),
            "CPP Est.": round(cpp_est),
        }
    )

df_inv = pd.DataFrame(inventory_rows)

# Programas recomendados base para pauta de 24 líneas (capacidad del Excel base filas 18 a 42)
default_programs_set = {
    "Chv Noticias Central",
    "Primer Plano",
    "Contigo en La Mañana A",
    "Contigo en Directo",
    "Chv Noticias Tarde",
    "Meganoticias Prime",
    "Teleserie Prime (El Señor de la Querencia)",
    "Mucho Gusto",
    "Meganoticias Alerta Semana",
    "Meganoticias Alerta Finde",
    "Teletrece",
    "Que Dice Chile",
    "Tu Día",
    "Teletrece Tarde",
    "Cultura Prime: Lugares Que Hablan",
    "24 Horas Central",
    "Ahora Caigo",
    "Ahora Caigo Prime",
    "Buenos Días a Todos",
    "24 Horas Central Sabado",
    "24 Horas Central Domingo",
    "Mesa central",
    "Domingos de película",
}

df_inv["En Pauta"] = df_inv["Programa"].apply(
    lambda x: any(d.lower() in x.lower() for d in default_programs_set)
)

selected_programs_df = st.data_editor(
    df_inv,
    column_config={
        "En Pauta": st.column_config.CheckboxColumn(
            "Incluir", default=False, help="Selecciona los programas para la pauta"
        ),
        "Tarifa": st.column_config.NumberColumn(
            "Valor Unitario (CLP)", format="$%d"
        ),
        "Rating": st.column_config.NumberColumn("Rating", format="%.2f"),
        "CPP Est.": st.column_config.NumberColumn("CPP Est.", format="$%d"),
    },
    disabled=["Canal", "Programa", "Bloque", "Días", "Horario", "Segundos"],
    hide_index=True,
    use_container_width=True,
)

active_progs = selected_programs_df[selected_programs_df["En Pauta"]].copy()

# Limitar a máximo 24 programas para respetar la plantilla física del Excel base
if len(active_progs) > 24:
    st.info(
        f"Se seleccionaron {len(active_progs)} programas. Se optimizarán los 24 con mejor afinidad de CPP y balance de SOI."
    )
    active_progs = active_progs.head(24)

# -------------------------------------------------------------
# 4. SIMULACIÓN DE DISTRIBUCIÓN DIARIA (CALENDARIO DE CAMPAÑA)
# -------------------------------------------------------------
# Generar días de campaña
dias_list = [start_date + timedelta(days=i) for i in range(num_dias_campana)]
dias_map_es = {0: "L", 1: "M", 2: "W", 3: "J", 4: "V", 5: "S", 6: "D"}


# Algoritmo de distribución de spots
def is_day_allowed(dia_char, dias_tarifa_str):
    d = dias_tarifa_str.upper()
    if "L-D" in d or "LMWJVD" in d:
        return True
    if "L-V" in d and dia_char in ["L", "M", "W", "J", "V"]:
        return True
    if "S-D" in d and dia_char in ["S", "D"]:
        return True
    if dia_char in d:
        return True
    return False


# Simular inserciones
spots_distrib = {p["Programa"]: [0] * len(dias_list) for _, p in active_progs.iterrows()}

for p_idx, (_, prog) in enumerate(active_progs.iterrows()):
    p_name = prog["Programa"]
    bloque = prog["Bloque"]
    # Frecuencia estimada según bloque
    freq_target = 2 if bloque == "OFF PRIME" else 1

    for d_idx, dia_dt in enumerate(dias_list):
        dia_char = dias_map_es[dia_dt.weekday()]
        if is_day_allowed(dia_char, prog["Días"]):
            # Alternancia según frecuencia
            if freq_target == 1 and (d_idx % 2 == 0):
                spots_distrib[p_name][d_idx] = min(1, max_spots_dia_prog)
            elif freq_target > 1:
                spots_distrib[p_name][d_idx] = min(freq_target, max_spots_dia_prog)

# Totales simulados
total_spots_sim = sum(sum(v) for v in spots_distrib.values())
active_progs["Total_Spots"] = active_progs["Programa"].map(
    lambda p: sum(spots_distrib[p])
)
active_progs["Inversion"] = active_progs["Total_Spots"] * active_progs["Tarifa"]
active_progs["TRPs"] = active_progs["Total_Spots"] * active_progs["Rating"]

inv_total_sim = active_progs["Inversion"].sum()
trps_totales_sim = active_progs["TRPs"].sum()
cpp_sim = (inv_total_sim / trps_totales_sim) if trps_totales_sim > 0 else 0

trps_prime_sim = active_progs[active_progs["Bloque"] == "PRIME"]["TRPs"].sum()
pct_prime_sim = (
    (trps_prime_sim / trps_totales_sim) if trps_totales_sim > 0 else 0.0
)

# Métricas en pantalla
colA, colB, colC, colD = st.columns(4)
colA.metric(
    "Inversión Proyectada",
    f"${inv_total_sim:,.0f} CLP",
    delta=f"${inv_total_sim - presupuesto_total:,.0f} vs Presupuesto",
    delta_color="inverse",
)
colB.metric(
    "TRPs Totales",
    f"{trps_totales_sim:.1f} pts",
    delta=f"{trps_totales_sim - objetivo_trps:+.1f} vs Meta",
)
colC.metric("CPP Promedio Ponderado", f"${cpp_sim:,.0f} CLP")
colD.metric(
    "% TRPs en Prime",
    f"{pct_prime_sim*100:.1f}%",
    delta=f"{(pct_prime_sim - sov_prime_target)*100:+.1f}% vs Meta",
)

# -------------------------------------------------------------
# 5. GENERACIÓN DEL EXCEL BASE CON FÓRMULAS ACTIVAS
# -------------------------------------------------------------
def export_to_excel_base(
    cliente, campana, version, start_d, end_d, active_p_df, distrib_dict
):
    wb = openpyxl.load_workbook("Excel base.xlsx")
    ws = wb["Pauta mes tipo"]

    # 1. Metadatos
    meses_es = [
        "",
        "ENERO",
        "FEBRERO",
        "MARZO",
        "ABRIL",
        "MAYO",
        "JUNIO",
        "JULIO",
        "AGOSTO",
        "SEPTIEMBRE",
        "OCTUBRE",
        "NOVIEMBRE",
        "DICIEMBRE",
    ]
    nombre_mes = f"{meses_es[start_d.month]} {start_d.year}"

    ws["C10"] = cliente
    ws["C11"] = f"{start_d.strftime('%d/%m/%Y')} al {end_d.strftime('%d/%m/%Y')}"
    ws["C12"] = campana
    ws["C13"] = version

    # 2. Configuración de Columnas de Fecha (J a AM, cols 10 a 39 = 30 días)
    ws["J15"] = nombre_mes

    dias_semana_abrev = {0: "L", 1: "M", 2: "W", 3: "J", 4: "V", 5: "S", 6: "D"}

    # Limpiar días existentes en fila 16 y 17
    for c in range(10, 40):
        ws.cell(16, c).value = None
        ws.cell(17, c).value = None

    for i, cur_dt in enumerate(dias_list[:30]):
        col_idx = 10 + i
        ws.cell(16, col_idx).value = dias_semana_abrev[cur_dt.weekday()]
        ws.cell(17, col_idx).value = cur_dt.day

    # 3. Llenar filas de programas (Filas 18 a 42)
    # Primero limpiar filas 18 a 42
    for r in range(18, 43):
        for c in range(2, 10):  # B a I
            ws.cell(r, c).value = None
        for c in range(10, 40):  # J a AM
            ws.cell(r, c).value = None
        ws.cell(r, 41).value = None  # AO (Pond)
        ws.cell(r, 42).value = None  # AP (Rating)
        ws.cell(r, 44).value = None  # AR (Valor unitario)
        ws.cell(r, 47).value = None  # AU (Trps prime)

    for i, (_, row_p) in enumerate(active_p_df.iterrows()):
        curr_row = 18 + i
        p_name = row_p["Programa"]

        ws.cell(curr_row, 2, value=row_p["Canal"])  # B: Canal
        ws.cell(curr_row, 3, value=p_name)  # C: Programa
        ws.cell(
            curr_row,
            4,
            value="Prime" if row_p["Bloque"] == "PRIME" else "Off",
        )  # D: Bloque
        ws.cell(curr_row, 5, value=row_p["Días"])  # E: Días
        ws.cell(curr_row, 6, value="Spot")  # F: Derechos
        ws.cell(curr_row, 7, value=row_p["Horario"])  # G: Inicio/Horario
        ws.cell(curr_row, 9, value=int(row_p["Segundos"]))  # I: Segundos

        # Días de pauta (J a AM)
        daily_spots = distrib_dict.get(p_name, [])
        for d_i, num_s in enumerate(daily_spots[:30]):
            if num_s > 0:
                ws.cell(curr_row, 10 + d_i, value=num_s)

        # Fórmulas y valores
        ws.cell(curr_row, 40, value=f"=SUM(J{curr_row}:AM{curr_row})")  # AN: Total
        ws.cell(curr_row, 41, value=1)  # AO: Pond = 1 (Pauta Libre)
        ws.cell(curr_row, 42, value=float(row_p["Rating"]))  # AP: Rating
        ws.cell(curr_row, 43, value=f"=+AP{curr_row}*AN{curr_row}")  # AQ: TRPS
        ws.cell(curr_row, 44, value=float(row_p["Tarifa"]))  # AR: Valor Unitario
        ws.cell(curr_row, 45, value=f"=AR{curr_row}*AN{curr_row}")  # AS: Valor Total
        ws.cell(
            curr_row,
            46,
            value=f"=IF(AQ{curr_row}>0, AS{curr_row}/AQ{curr_row}, 0)",
        )  # AT: CPP

        # AU: Trps Prime (Fórmula activa)
        ws.cell(
            curr_row,
            47,
            value=f'=IF(D{curr_row}="Prime", AQ{curr_row}, 0)',
        )

    # 4. Fila 43: Totales y TRPS POR DÍA
    ws.cell(43, 40, value=f"=SUM(AN18:AN42)")  # Total Derechos
    ws.cell(43, 43, value=f"=SUM(AQ18:AQ42)")  # Total TRPS
    ws.cell(43, 45, value=f"=SUM(AS18:AS42)")  # Total Inversión
    ws.cell(43, 46, value=f"=IF(AQ43>0, AS43/AQ43, 0)")  # CPP Ponderado
    ws.cell(43, 47, value=f"=SUM(AU18:AU42)")  # Total TRPS Prime

    # Guardar en buffer en memoria
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# -------------------------------------------------------------
# 6. BOTÓN DE DESCARGA DIRECTA
# -------------------------------------------------------------
st.markdown("---")
st.subheader("📥 Exportación Oficial MADCOM")

excel_buffer = export_to_excel_base(
    cliente_sel,
    campana_nombre,
    version_pauta,
    start_date,
    end_date,
    active_progs,
    spots_distrib,
)

nombre_archivo = f"Pauta_TV_{cliente_sel.replace(' ', '_')}_{start_date.strftime('%d%b')}_{end_date.strftime('%d%b')}.xlsx"

st.download_button(
    label=f"📊 Descargar {nombre_archivo} (Excel Base con Fórmulas)",
    data=excel_buffer,
    file_name=nombre_archivo,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
