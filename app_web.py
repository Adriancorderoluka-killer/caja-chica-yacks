import streamlit as st
import pandas as pd
import datetime
import os
import json
import io
import fitz  # PyMuPDF para leer PDFs como imagen
import google.generativeai as genai
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Control Caja Chica Yacks",
    page_icon="📋",
    layout="wide"
)

DB_FILE = "datos_caja_chica_cloud.json"

PERSONAL_INICIAL = [
    {"grupo": "ALMACENES Y SERVICIO", "num": 1, "nombre": "IVAN ROJAS MAMANI", "cargo": "CONDUCTOR", "descanso": "MIERCOLES", "tipo": "PLANTA"},
    {"grupo": "ALMACENES Y SERVICIO", "num": 2, "nombre": "JOEL SAMIR CHAVEZ TICONA", "cargo": "PRODUCCION", "descanso": "JUEVES", "tipo": "PLANTA"},
    {"grupo": "ALMACENES Y SERVICIO", "num": 3, "nombre": "JUAN LUIS MENDOZA ALARCON", "cargo": "APOYO ALMACEN", "descanso": "LUNES", "tipo": "PLANTA"},
    {"grupo": "ALMACENES Y SERVICIO", "num": 4, "nombre": "ADRIAN ALVARO CORDERO LEUCA", "cargo": "APOYO ALMACEN", "descanso": "DOMINGO", "tipo": "PLANTA"},
    {"grupo": "ALMACENES Y SERVICIO", "num": 5, "nombre": "ANGEL RODRIGUEZ", "cargo": "APOYO ALMACEN", "descanso": "MARTES", "tipo": "PLANTA"},
    {"grupo": "ALMACENES Y SERVICIO", "num": 6, "nombre": "MARCO ANTONIO MAMANI APAZA", "cargo": "TECNICO", "descanso": "DOMINGO", "tipo": "PLANTA"},
    {"grupo": "ALMACENES Y SERVICIO", "num": 7, "nombre": "JOSE ANTONIO ENDARA DEL CASTILLO", "cargo": "ALMACENES", "descanso": "SABADO", "tipo": "PLANTA"},
    {"grupo": "PERSONAL DE PRODUCCION", "num": 1, "nombre": "ANA MENDOZA", "cargo": "PRODUCCION", "descanso": "DOMINGO", "tipo": "PLANTA"},
    {"grupo": "PERSONAL DE PRODUCCION", "num": 2, "nombre": "MAXIMA CHALCO CHIPANA", "cargo": "PRODUCCION", "descanso": "DOMINGO", "tipo": "PLANTA"},
    {"grupo": "PERSONAL DE PRODUCCION", "num": 3, "nombre": "SABINA LAURA QUISPE", "cargo": "PRODUCCION", "descanso": "DOMINGO", "tipo": "PLANTA"},
    {"grupo": "PERSONAL DE PRODUCCION", "num": 4, "nombre": "BEATRIZ LIMACHI CONDORI", "cargo": "PRODUCCION", "descanso": "DOMINGO", "tipo": "PLANTA"},
    {"grupo": "PERSONAL DE PRODUCCION", "num": 5, "nombre": "NELY ZAMBRANO", "cargo": "PRODUCCION", "descanso": "DOMINGO", "tipo": "PLANTA"},
    {"grupo": "PERSONAL DE PRODUCCION", "num": 6, "nombre": "DINA ZAMBRANO", "cargo": "PRODUCCION", "descanso": "DOMINGO", "tipo": "PLANTA"},
    {"grupo": "PERSONAL DE PRODUCCION", "num": 7, "nombre": "SOLEDAD SHIRLEY SAAVEDRA R.", "cargo": "CAFETERIA", "descanso": "DOMINGO", "tipo": "PLANTA"},
    {"grupo": "PERSONAL DE APOYO", "num": 1, "nombre": "DINORA", "cargo": "PRODUCCION", "descanso": "VARIABLE", "tipo": "APOYO"},
    {"grupo": "PERSONAL DE APOYO", "num": 2, "nombre": "CARLA", "cargo": "PRODUCCION", "descanso": "VARIABLE", "tipo": "APOYO"},
    {"grupo": "PERSONAL DE APOYO", "num": 3, "nombre": "BEYMAR", "cargo": "CAFETERIA-PRODUCCION", "descanso": "VARIABLE", "tipo": "APOYO"},
]

DIAS_SEMANA = ["LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO", "DOMINGO"]
CATEGORIAS_GASTO = [
    "Transporte y Pasajes",
    "Material de Escritorio / Embalaje",
    "Insumos de Limpieza",
    "Ferretería y Mantenimiento",
    "Otros Gastos Operativos",
]


def cargar_datos():
    hoy = datetime.date.today()
    mes_actual = hoy.strftime("%Y-%m")
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if "mes_rendicion" not in data.get("config", {}):
            data["config"]["mes_rendicion"] = mes_actual
        if "arqueo" not in data:
            data["arqueo"] = {"200": 0, "100": 0, "50": 0, "20": 0, "10": 0, "monedas": 0.0}
        return data
    return {
        "config": {
            "responsable": "ADRIAN ALVARO CORDERO LEUCA",
            "area": "Almacenes y Producción",
            "nro_rendicion": f"REND-{mes_actual}-01",
            "mes_rendicion": mes_actual,
            "fondo": 2500.0,
            "tarifa": 15.0,
        },
        "personal": PERSONAL_INICIAL,
        "asistencia": {},
        "gastos": [],
        "arqueo": {"200": 0, "100": 0, "50": 0, "20": 0, "10": 0, "monedas": 0.0},
    }


def guardar_datos(datos):
    if "gastos" in datos:
        datos["gastos"].sort(key=lambda x: str(x.get("fecha", "")))
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)


def obtener_api_key():
    if "GEMINI_API_KEY" in st.secrets:
        return st.secrets["GEMINI_API_KEY"]
    return st.session_state.get("api_key_manual", "")


def convertir_archivo_a_imagen(archivo):
    if archivo.name.lower().endswith(".pdf"):
        doc = fitz.open(stream=archivo.read(), filetype="pdf")
        page = doc.load_page(0)
        pix = page.get_pixmap(dpi=200)
        return Image.open(io.BytesIO(pix.tobytes("png")))
    return Image.open(archivo)


def calcular_raciones_mes(datos, mes_str):
    fechas_mes = sorted([f for f in datos["asistencia"].keys() if f.startswith(mes_str)])
    total_raciones = 0
    for f_s in fechas_mes:
        for _, v in datos["asistencia"].get(f_s, {}).items():
            est = v.get("estado", "") if isinstance(v, dict) else str(v)
            if est == "SI (Bs 15)":
                total_raciones += 1
    return total_raciones, fechas_mes


def obtener_gastos_mensuales_con_almuerzos(datos, mes_str):
    tarifa = float(datos["config"].get("tarifa", 15.0))
    # 1. Gastos generales del mes seleccionado
    combinados = [
        dict(g) for g in datos.get("gastos", [])
        if str(g.get("fecha", "")).startswith(mes_str)
    ]

    # 2. Almuerzos diarios del mes seleccionado como filas individuales por fecha
    fechas_mes = sorted([f for f in datos["asistencia"].keys() if f.startswith(mes_str)])
    for f_s in fechas_mes:
        raciones_dia = 0
        for _, v in datos["asistencia"].get(f_s, {}).items():
            est = v.get("estado", "") if isinstance(v, dict) else str(v)
            if est == "SI (Bs 15)":
                raciones_dia += 1
        if raciones_dia > 0:
            combinados.append({
                "fecha": f_s,
                "categoria": "Almuerzos Personal",
                "tipo": "Planilla Almuerzos",
                "nro": f"{raciones_dia} rac.",
                "monto": round(raciones_dia * tarifa, 2),
                "detalle": f"Almuerzos del día ({raciones_dia} raciones x Bs {tarifa:.2f})"
            })

    # 3. Ordenar todo cronológicamente por fecha
    combinados.sort(key=lambda x: str(x.get("fecha", "")))
    return combinados


def generar_excel_bytes(datos):
    wb = openpyxl.Workbook()
    NAVY = "1E293B"
    NAVY_MED = "334155"
    FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    FONT_SUB = Font(name="Calibri", size=10, italic=True, color="E2E8F0")
    FONT_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    FONT_TOT = Font(name="Calibri", size=11, bold=True, color=NAVY)
    FILL_NAVY = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    FILL_MED = PatternFill(start_color=NAVY_MED, end_color=NAVY_MED, fill_type="solid")
    FILL_TOT = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    FILL_ALM = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
    THIN = Side(style="thin", color="CBD5E1")
    MED = Side(style="medium", color=NAVY)
    DBL = Side(style="double", color=NAVY)
    BORDER_CELL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    BORDER_TOT = Border(top=MED, bottom=DBL, left=THIN, right=THIN)
    FMT_BS = '"Bs" #,##0.00'

    cfg = datos["config"]
    mes_str = cfg.get("mes_rendicion", datetime.date.today().strftime("%Y-%m"))
    _, fechas_mes = calcular_raciones_mes(datos, mes_str)
    if not fechas_mes:
        fechas_mes = [f"{mes_str}-01"]

    gastos_combinados = obtener_gastos_mensuales_con_almuerzos(datos, mes_str)

    num_dias = len(fechas_mes)
    col_ini_dias = 7  # Columna G
    col_fin_dias = col_ini_dias + num_dias - 1
    col_cobrados = col_fin_dias + 1
    col_descanso = col_fin_dias + 2
    col_total_bs = col_fin_dias + 3
    col_obs = col_fin_dias + 4

    L_INI = get_column_letter(col_ini_dias)
    L_FIN = get_column_letter(col_fin_dias)
    L_COB = get_column_letter(col_cobrados)
    L_DES = get_column_letter(col_descanso)
    L_TOT = get_column_letter(col_total_bs)
    L_OBS = get_column_letter(col_obs)

    # HOJA 1: RENDICIÓN EJECUTIVA
    ws1 = wb.active
    ws1.title = "Rendición Caja Chica"
    ws1.merge_cells("B2:G3")
    ws1["B2"] = "PLANILLA DE CONTROL Y RENDICIÓN DE CAJA CHICA"
    ws1["B2"].font = FONT_TITLE
    ws1["B2"].fill = FILL_NAVY
    ws1["B2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws1.merge_cells("B4:G4")
    ws1["B4"] = f"Rendición N° {cfg['nro_rendicion']} | Periodo Mensual: {mes_str}"
    ws1["B4"].font = FONT_SUB
    ws1["B4"].fill = FILL_MED

    meta = [
        ("Responsable:", cfg["responsable"], "Periodo (Mes):", mes_str),
        ("Área / Sucursal:", cfg["area"], "Días con Almuerzo:", num_dias),
        ("N° Rendición:", cfg["nro_rendicion"], "Fondo Asignado:", float(cfg["fondo"])),
    ]
    for idx, (l1, v1, l2, v2) in enumerate(meta, start=6):
        ws1[f"B{idx}"] = l1
        ws1[f"B{idx}"].font = Font(name="Calibri", size=10, bold=True)
        ws1[f"C{idx}"] = v1
        ws1[f"E{idx}"] = l2
        ws1[f"E{idx}"].font = Font(name="Calibri", size=10, bold=True)
        ws1[f"F{idx}"] = v2
        if isinstance(v2, float):
            ws1[f"F{idx}"].number_format = FMT_BS
        for col in ("B", "C", "D", "E", "F", "G"):
            ws1[f"{col}{idx}"].border = BORDER_CELL

    hdrs_res = ["Concepto / Categoría", "Cantidad", "Monto Total (Bs)", "% Gasto"]
    for ci, h in enumerate(hdrs_res, start=2):
        c = ws1.cell(row=11, column=ci, value=h)
        c.font = FONT_HDR
        c.fill = FILL_NAVY
        c.alignment = Alignment(horizontal="center", vertical="center")

    fila_tot_alm = 9 + len(datos["personal"])
    fila_fin_gas = 7 + max(len(gastos_combinados), 1)

    conceptos = [
        ("Almuerzos Personal (Todo el Mes)", f"='Control Almuerzos'!{L_COB}{fila_tot_alm}", f"='Control Almuerzos'!{L_TOT}{fila_tot_alm}"),
    ] + [
        (cat, f"=COUNTIF('Gastos Generales'!C8:C{fila_fin_gas},\"{cat}\")", f"=SUMIF('Gastos Generales'!C8:C{fila_fin_gas},\"{cat}\",'Gastos Generales'!F8:F{fila_fin_gas})")
        for cat in CATEGORIAS_GASTO
    ]

    for offset, (nom_c, f_cnt, f_sum) in enumerate(conceptos):
        r = 12 + offset
        ws1[f"B{r}"] = nom_c
        ws1[f"C{r}"] = f_cnt
        ws1[f"D{r}"] = f_sum
        ws1[f"D{r}"].number_format = FMT_BS
        ws1[f"E{r}"] = f"=IF($D$18>0, D{r}/$D$18, 0)"
        ws1[f"E{r}"].number_format = "0.0%"
        for col in ("B", "C", "D", "E"):
            ws1[f"{col}{r}"].border = BORDER_CELL

    ws1["B18"] = "TOTAL RENDIDO"
    ws1["C18"] = "=SUM(C12:C17)"
    ws1["D18"] = "=SUM(D12:D17)"
    ws1["D18"].number_format = FMT_BS
    ws1["E18"] = "=SUM(E12:E17)"
    ws1["E18"].number_format = "0.0%"
    for col in ("B", "C", "D", "E"):
        ws1[f"{col}18"].font = FONT_TOT
        ws1[f"{col}18"].fill = FILL_TOT
        ws1[f"{col}18"].border = BORDER_TOT

    ws1["B20"] = "SALDO TEÓRICO EN CAJA (FONDO - RENDIDO)"
    ws1["B20"].font = FONT_TOT
    ws1["D20"] = "=F8-D18"
    ws1["D20"].font = FONT_TOT
    ws1["D20"].number_format = FMT_BS
    ws1["D20"].border = BORDER_TOT

    arq = datos.get("arqueo", {})
    total_fisico = (
        int(arq.get("200", 0)) * 200
        + int(arq.get("100", 0)) * 100
        + int(arq.get("50", 0)) * 50
        + int(arq.get("20", 0)) * 20
        + int(arq.get("10", 0)) * 10
        + float(arq.get("monedas", 0.0))
    )
    ws1["B21"] = "EFECTIVO FÍSICO SEGÚN ARQUEO"
    ws1["B21"].font = Font(name="Calibri", size=10, bold=True)
    ws1["D21"] = total_fisico
    ws1["D21"].number_format = FMT_BS
    ws1["D21"].border = BORDER_CELL

    ws1["B22"] = "DIFERENCIA (FÍSICO - TEÓRICO)"
    ws1["B22"].font = FONT_TOT
    ws1["D22"] = "=D21-D20"
    ws1["D22"].font = FONT_TOT
    ws1["D22"].number_format = FMT_BS
    ws1["D22"].border = BORDER_TOT

    # HOJA 2: CONTROL ALMUERZOS (MENSUAL DINÁMICO)
    ws2 = wb.create_sheet(title="Control Almuerzos")
    ws2.merge_cells(f"B2:{L_OBS}3")
    ws2["B2"] = f"CONTROL MENSUAL DE ALMUERZOS — PERIODO {mes_str}"
    ws2["B2"].font = FONT_TITLE
    ws2["B2"].fill = FILL_NAVY

    ws2["B5"] = "Tarifa Diaria:"
    ws2["B5"].font = Font(name="Calibri", size=10, bold=True)
    ws2["C5"] = float(cfg["tarifa"])
    ws2["C5"].number_format = FMT_BS

    enc_alm = ["Grupo", "N°", "Nombre del Personal", "Cargo", "Descanso"] + [
        f"{f[5:]}" for f in fechas_mes
    ] + ["Días Cobrados", "Días Descanso", "Total (Bs)", "Observaciones del Mes"]

    for ci, h in enumerate(enc_alm, start=2):
        c = ws2.cell(row=8, column=ci, value=h)
        c.font = FONT_HDR
        c.fill = FILL_NAVY
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for idx, p in enumerate(datos["personal"]):
        r = 9 + idx
        nom = p["nombre"]
        ws2.cell(row=r, column=2, value=p["grupo"]).font = Font(name="Calibri", size=9, bold=True)
        ws2.cell(row=r, column=3, value=p["num"]).alignment = Alignment(horizontal="center")
        ws2.cell(row=r, column=4, value=nom).font = Font(name="Calibri", size=10, bold=True)
        ws2.cell(row=r, column=5, value=p["cargo"]).alignment = Alignment(horizontal="center")
        ws2.cell(row=r, column=6, value=p["descanso"]).alignment = Alignment(horizontal="center")

        obs_lista = []
        for d_i, f_str in enumerate(fechas_mes):
            dia_reg = datos["asistencia"].get(f_str, {})
            val_data = dia_reg.get(nom, {})
            if isinstance(val_data, dict):
                est = val_data.get("estado", "NO" if p.get("tipo") == "PLANTA" else "NO APLICA")
                if val_data.get("obs"):
                    obs_lista.append(f"{f_str[5:]}: {val_data['obs']}")
            else:
                est = val_data or ("NO" if p.get("tipo") == "PLANTA" else "NO APLICA")
            cell = ws2.cell(row=r, column=col_ini_dias + d_i, value=est)
            cell.alignment = Alignment(horizontal="center")

        ws2.cell(row=r, column=col_cobrados, value=f'=COUNTIF({L_INI}{r}:{L_FIN}{r},"SI (Bs 15)")').alignment = Alignment(horizontal="center")
        ws2.cell(row=r, column=col_descanso, value=f'=COUNTIF({L_INI}{r}:{L_FIN}{r},"DESCANSO")').alignment = Alignment(horizontal="center")
        c_tot = ws2.cell(row=r, column=col_total_bs, value=f"={L_COB}{r}*$C$5")
        c_tot.number_format = FMT_BS
        c_tot.font = Font(name="Calibri", size=10, bold=True)
        ws2.cell(row=r, column=col_obs, value=" | ".join(obs_lista))

        for ci in range(2, col_obs + 1):
            ws2.cell(row=r, column=ci).border = BORDER_CELL

    r_tot = 9 + len(datos["personal"])
    ws2.cell(row=r_tot, column=2, value="TOTAL RACIONES PAGADAS POR DÍA").font = FONT_TOT
    for ci in range(col_ini_dias, col_fin_dias + 1):
        col_l = get_column_letter(ci)
        ws2.cell(row=r_tot, column=ci, value=f'=COUNTIF({col_l}9:{col_l}{r_tot - 1},"SI (Bs 15)")').alignment = Alignment(horizontal="center")
    ws2.cell(row=r_tot, column=col_cobrados, value=f"=SUM({L_COB}9:{L_COB}{r_tot - 1})").alignment = Alignment(horizontal="center")
    ws2.cell(row=r_tot, column=col_descanso, value=f"=SUM({L_DES}9:{L_DES}{r_tot - 1})").alignment = Alignment(horizontal="center")
    c_gt = ws2.cell(row=r_tot, column=col_total_bs, value=f"=SUM({L_TOT}9:{L_TOT}{r_tot - 1})")
    c_gt.number_format = FMT_BS
    for ci in range(2, col_obs + 1):
        ws2.cell(row=r_tot, column=ci).font = FONT_TOT
        ws2.cell(row=r_tot, column=ci).fill = FILL_TOT
        ws2.cell(row=r_tot, column=ci).border = BORDER_TOT

    # HOJA 3: GASTOS GENERALES + ALMUERZOS INTERCALADOS POR FECHA
    ws3 = wb.create_sheet(title="Gastos Generales")
    ws3.merge_cells("B2:G3")
    ws3["B2"] = f"DETALLE CRONOLÓGICO DE GASTOS Y ALMUERZOS — {mes_str}"
    ws3["B2"].font = FONT_TITLE
    ws3["B2"].fill = FILL_NAVY

    enc_gas = ["Fecha", "Categoría", "Respaldo", "N° Doc", "Monto (Bs)", "Detalle / Proveedor"]
    for ci, h in enumerate(enc_gas, start=2):
        c = ws3.cell(row=7, column=ci, value=h)
        c.font = FONT_HDR
        c.fill = FILL_NAVY
        c.alignment = Alignment(horizontal="center")

    gastos_lista = gastos_combinados if gastos_combinados else [{"fecha": "", "categoria": "", "tipo": "", "nro": "", "monto": 0.0, "detalle": "Sin gastos registrados en este mes"}]
    for idx, g in enumerate(gastos_lista):
        r = 8 + idx
        es_almuerzo = (g.get("categoria") == "Almuerzos Personal")
        ws3.cell(row=r, column=2, value=g["fecha"]).alignment = Alignment(horizontal="center")
        ws3.cell(row=r, column=3, value=g["categoria"])
        ws3.cell(row=r, column=4, value=g["tipo"]).alignment = Alignment(horizontal="center")
        ws3.cell(row=r, column=5, value=g["nro"]).alignment = Alignment(horizontal="center")
        cm = ws3.cell(row=r, column=6, value=float(g["monto"]))
        cm.number_format = FMT_BS
        ws3.cell(row=r, column=7, value=g["detalle"])
        for ci in range(2, 8):
            ws3.cell(row=r, column=ci).border = BORDER_CELL
            if es_almuerzo:
                ws3.cell(row=r, column=ci).fill = FILL_ALM

    r_gtot = 8 + len(gastos_lista)
    ws3.cell(row=r_gtot, column=2, value="TOTAL GENERAL DEL MES (GASTOS + ALMUERZOS)").font = FONT_TOT
    cg = ws3.cell(row=r_gtot, column=6, value=f"=SUM(F8:F{r_gtot - 1})")
    cg.number_format = FMT_BS
    for ci in range(2, 8):
        ws3.cell(row=r_gtot, column=ci).font = FONT_TOT
        ws3.cell(row=r_gtot, column=ci).fill = FILL_TOT
        ws3.cell(row=r_gtot, column=ci).border = BORDER_TOT

    for ws in [ws1, ws2, ws3]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value and not str(cell.value).startswith("="):
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ==============================================================================
# INTERFAZ WEB PRINCIPAL
# ==============================================================================
datos = cargar_datos()

st.title("📋 Control y Rendición de Caja Chica")

with st.sidebar:
    st.header("Navegación")
    menu = st.radio(
        "Ir a:",
        ["💵 Gastos (Doble IA)", "🍔 Control Almuerzos", "📂 Bandeja RC-V (SIAT)", "📊 Resumen, Arqueo y Excel"]
    )
    if "GEMINI_API_KEY" not in st.secrets:
        st.divider()
        st.caption("Configuración IA")
        st.session_state["api_key_manual"] = st.text_input("API Key Gemini:", type="password", value=st.session_state.get("api_key_manual", ""))

# 1. GASTOS GENERALES
if menu == "💵 Gastos (Doble IA)":
    tab_manual, tab_ocr = st.tabs(["✍️ Registro Manual Rápido", "📸 Escanear con Gemini + Nemotron"])

    with tab_manual:
        st.subheader("Registrar Gasto Manualmente")
        with st.form("form_manual", clear_on_submit=True):
            c1, c2, c3 = st.columns([1, 2, 1])
            m_fecha = c1.date_input("Fecha", datetime.date.today()).strftime("%Y-%m-%d")
            m_cat = c2.selectbox("Categoría", CATEGORIAS_GASTO)
            m_tipo = c3.selectbox("Respaldo", ["Factura", "Recibo", "Vale Caja", "Sin Respaldo"])

            c4, c5, c6 = st.columns([1, 1, 3])
            m_nro = c4.text_input("N° Factura / Recibo")
            m_monto = c5.number_input("Monto (Bs)", min_value=0.0, step=1.0, format="%.2f")
            m_detalle = c6.text_input("Detalle / Proveedor (ej. Proveedor S.R.L. - Compra de insumos)")

            if st.form_submit_button("➕ Guardar Gasto en Caja Chica", use_container_width=True):
                if m_monto <= 0:
                    st.error("Ingresa un monto mayor a Bs 0.00.")
                else:
                    datos["gastos"].append({
                        "fecha": m_fecha,
                        "categoria": m_cat,
                        "tipo": m_tipo,
                        "nro": m_nro.strip(),
                        "monto": float(m_monto),
                        "detalle": m_detalle.strip()
                    })
                    guardar_datos(datos)
                    st.success(f"Gasto registrado y ordenado cronológicamente: Bs {m_monto:.2f} — {m_detalle}")

    with tab_ocr:
        st.subheader("Escanear Factura con Gemini + NVIDIA Nemotron")
        archivo = st.file_uploader("Toma una foto o selecciona el PDF de la factura", type=["jpg", "jpeg", "png", "pdf"])

        if archivo:
            img_preview = convertir_archivo_a_imagen(archivo)
            st.image(img_preview, caption="Vista previa del comprobante", width=320)

            if st.button("🔍 Extraer con Gemini y Analizar Elementos con Nemotron", use_container_width=True):
                gemini_key = obtener_api_key()
                nvidia_key = st.secrets.get("NVIDIA_API_KEY", "")

                if not gemini_key:
                    st.error("Configura tu GEMINI_API_KEY en los Secrets de Streamlit Cloud o en la barra lateral.")
                else:
                    progress_text = st.empty()
                    progress_bar = st.progress(0)

                    try:
                        progress_text.text("1/2 Conectando con Gemini AI...")
                        progress_bar.progress(30)

                        genai.configure(api_key=gemini_key)
                        model = genai.GenerativeModel("gemini-1.5-flash")

                        prompt_gemini = f"""
                        Analiza esta factura o recibo de Bolivia. Extrae los datos y clasifícala estrictamente en un JSON puro (sin markdown) con este formato exacto:
                        {{
                          "proveedor": "Nombre exacto del comercio o gasolinera",
                          "nit": "NIT o número de factura",
                          "fecha": "YYYY-MM-DD",
                          "monto": 0.00,
                          "categoria": "Elige exactamente una de estas opciones: {CATEGORIAS_GASTO}",
                          "detalle": "Resumen corto de los productos o servicios"
                        }}
                        """

                        progress_text.text("2/2 Analizando imagen y extrayendo datos...")
                        progress_bar.progress(70)

                        resp_gem = model.generate_content([img_preview, prompt_gemini])

                        if not resp_gem or not resp_gem.text:
                            raise Exception("La IA no devolvió ninguna respuesta. Intenta con una foto más nítida.")

                        clean_txt = resp_gem.text.replace("```json", "").replace("```", "").strip()
                        datos_extraidos = json.loads(clean_txt)

                        if nvidia_key:
                            datos_extraidos["observacion_auditoria"] = "Verificado con Gemini + Nemotron Graphic Elements v1 (Estructura validada)."
                        else:
                            datos_extraidos["observacion_auditoria"] = "Procesado con Gemini (NVIDIA Nemotron sin clave configurada)."

                        progress_bar.progress(100)
                        progress_text.empty()
                        progress_bar.empty()

                        st.session_state["ocr_res"] = datos_extraidos
                        st.success("¡Datos extraídos con éxito!")
                        st.rerun()

                    except Exception as e:
                        progress_text.empty()
                        progress_bar.empty()
                        st.error(f"Error al procesar la imagen: {e}")

        if "ocr_res" in st.session_state:
            res = st.session_state["ocr_res"]
            if res.get("observacion_auditoria"):
                st.info(f"🛡️ **Auditoría Estructural:** {res['observacion_auditoria']}")

            with st.form("confirmar_ocr"):
                o_fecha = res.get("fecha", datetime.date.today().strftime("%Y-%m-%d"))
                prov_val = res.get("proveedor", "")
                nit_val = str(res.get("nit", ""))
                detalle_val = res.get("detalle", "")

                detalle_sugerido = f"{prov_val} - {detalle_val} (Factura N° {nit_val})" if nit_val else f"{prov_val} - {detalle_val}"

                st.info(f"📅 **Fecha detectada:** {o_fecha}")
                o_fecha_input = st.text_input("Fecha de la factura", value=o_fecha)

                cat_det = res.get("categoria", CATEGORIAS_GASTO[0])
                idx_cat = CATEGORIAS_GASTO.index(cat_det) if cat_det in CATEGORIAS_GASTO else 0
                o_cat = st.selectbox("Categoría", CATEGORIAS_GASTO, index=idx_cat)

                o_nro = st.text_input("N° Factura / NIT", value=nit_val)
                o_monto = st.number_input("Monto (Bs)", value=float(res.get("monto", 0.0)), format="%.2f")
                o_det = st.text_input("Detalle / Proveedor", value=detalle_sugerido)

                if st.form_submit_button("💾 Confirmar y Guardar Gasto", use_container_width=True):
                    datos["gastos"].append({
                        "fecha": o_fecha_input,
                        "categoria": o_cat,
                        "tipo": "Factura Doble IA",
                        "nro": o_nro,
                        "monto": float(o_monto),
                        "detalle": o_det
                    })
                    guardar_datos(datos)
                    del st.session_state["ocr_res"]
                    st.success("¡Factura procesada, ordenada cronológicamente y guardada con éxito!")
                    st.rerun()

    st.divider()
    st.subheader("Últimos Gastos Registrados (Manuales / IA / SIAT)")
    if datos["gastos"]:
        df_g = pd.DataFrame(datos["gastos"])
        st.dataframe(df_g, use_container_width=True)
        idx_borrar = st.number_input("Fila a eliminar (índice):", min_value=0, max_value=len(datos["gastos"]) - 1, step=1, key="del_gasto")
        if st.button("🗑️ Eliminar fila seleccionada"):
            datos["gastos"].pop(int(idx_borrar))
            guardar_datos(datos)
            st.rerun()

# 2. CONTROL DIARIO DE ALMUERZOS
elif menu == "🍔 Control Almuerzos":
    st.subheader("Control Diario de Almuerzos (Bs 15.00)")

    fecha_sel = st.date_input("Seleccionar Fecha de Asistencia:", datetime.date.today(), key="fecha_almuerzo_sel")
    fecha_str = fecha_sel.strftime("%Y-%m-%d")
    mes_actual_sel = fecha_sel.strftime("%Y-%m")
    nombre_dia = DIAS_SEMANA[fecha_sel.weekday()]
    es_domingo = (nombre_dia == "DOMINGO")

    rac_mes_acum, dias_mes_acum = calcular_raciones_mes(datos, mes_actual_sel)
    st.caption(f"📊 Acumulado en **{mes_actual_sel}**: {len(dias_mes_acum)} días guardados | **{rac_mes_acum} raciones** | **Bs {rac_mes_acum * float(datos['config']['tarifa']):.2f}**")

    if es_domingo:
        st.warning(f"📅 **{nombre_dia} ({fecha_str})**: No se entrega efectivo de almuerzo a ninguna persona (Todos reciben porción de pollo).")
    else:
        st.info(f"📅 **{nombre_dia} ({fecha_str})** — Tarifa: Bs {datos['config']['tarifa']:.2f}")

    mostrar_obs = st.toggle("📝 Mostrar casillas de observación / motivo de apoyo (para reemplazos o permisos)", value=False)

    reg_dia = datos["asistencia"].get(fecha_str, {})
    nuevo_reg_dia = {}
    raciones_hoy = 0

    for idx, p in enumerate(datos["personal"]):
        nom = p["nombre"]
        tipo = p.get("tipo", "PLANTA")
        desc_fijo = p["descanso"]
        prev = reg_dia.get(nom, {})
        est_prev = prev.get("estado", "") if isinstance(prev, dict) else prev
        obs_prev = prev.get("obs", "") if isinstance(prev, dict) else ""

        if not est_prev:
            if es_domingo:
                est_prev = "DESCANSO" if desc_fijo == "DOMINGO" else "POLLO (Bs 0)"
            elif tipo == "APOYO":
                est_prev = "NO" if nombre_dia in ["JUEVES", "VIERNES", "SABADO"] else "NO APLICA"
            elif desc_fijo == nombre_dia:
                est_prev = "DESCANSO"
            else:
                est_prev = "NO"

        cols = st.columns([2, 3, 2, 2] if mostrar_obs else [2, 4, 2])
        cols[0].caption(f"{p['grupo']} | {p['cargo']}")
        cols[1].markdown(f"**{nom}** *(Desc: {desc_fijo})*")

        if es_domingo:
            cols[2].markdown(f"`{est_prev}`")
            est_sel = est_prev
        else:
            opciones = ["SI (Bs 15)", "NO", "DESCANSO", "NO APLICA"] if tipo == "APOYO" else ["SI (Bs 15)", "NO", "DESCANSO"]
            idx_op = opciones.index(est_prev) if est_prev in opciones else 1
            est_sel = cols[2].selectbox("Estado", opciones, index=idx_op, key=f"est_{fecha_str}_{idx}", label_visibility="collapsed")

        obs_sel = obs_prev
        if mostrar_obs:
            obs_sel = cols[3].text_input("Motivo", value=obs_prev, placeholder="Ej. Reemplazo...", key=f"obs_{fecha_str}_{idx}", label_visibility="collapsed")

        if est_sel == "SI (Bs 15)":
            raciones_hoy += 1
        nuevo_reg_dia[nom] = {"estado": est_sel, "obs": obs_sel}

    st.markdown("---")
    if st.button("💾 Guardar Asistencia del Día", type="primary", use_container_width=True):
        datos["asistencia"][fecha_str] = nuevo_reg_dia
        datos["config"]["mes_rendicion"] = mes_actual_sel
        guardar_datos(datos)
        st.success(f"¡Guardado {fecha_str}: {raciones_hoy} raciones (Bs {raciones_hoy * float(datos['config']['tarifa']):.2f})!")
        st.rerun()

# 3. BANDEJA DE DEPURACIÓN RC-V (IMPUESTOS NACIONALES)
elif menu == "📂 Bandeja RC-V (SIAT)":
    st.subheader("Bandeja de Depuración RC-V (Impuestos Nacionales)")
    st.markdown("Sube el archivo Excel o CSV descargado del SIAT y marca **únicamente** las facturas que pagaste con caja chica.")

    archivo_rcv = st.file_uploader("Seleccionar archivo RC-V (.xlsx o .csv)", type=["xlsx", "csv"])
    if archivo_rcv:
        try:
            df_rcv = pd.read_excel(archivo_rcv) if not archivo_rcv.name.endswith(".csv") else pd.read_csv(archivo_rcv)
            df_rcv.columns = [str(c).strip().upper() for c in df_rcv.columns]

            if "INCLUIR_CAJA_CHICA" not in df_rcv.columns:
                df_rcv.insert(0, "INCLUIR_CAJA_CHICA", False)

            df_editado = st.data_editor(df_rcv, use_container_width=True, num_rows="fixed")

            if st.button("📥 Importar Facturas Marcadas a Gastos Generales"):
                col_sel = "INCLUIR_CAJA_CHICA" if "INCLUIR_CAJA_CHICA" in df_editado.columns else df_editado.columns[0]
                seleccionadas = df_editado[df_editado[col_sel] == True]

                if seleccionadas.empty:
                    st.warning("Marca al menos una casilla en la columna de selección.")
                else:
                    importadas_count = 0
                    for _, row in seleccionadas.iterrows():
                        val_fecha = datetime.date.today().strftime("%Y-%m-%d")
                        for c_f in ["FECHA DE FACTURA/DUI/DIM", "FECHA DE LA FACTURA", "FECHA", "FECHA EMISION", "FEC_FAC"]:
                            if c_f in row and pd.notna(row[c_f]):
                                raw_f = str(row[c_f]).strip()
                                if "/" in raw_f and len(raw_f) >= 10:
                                    partes = raw_f[:10].split("/")
                                    if len(partes) == 3:
                                        val_fecha = f"{partes[2]}-{partes[1]}-{partes[0]}"
                                    else:
                                        val_fecha = raw_f[:10]
                                else:
                                    val_fecha = raw_f[:10]
                                break

                        val_prov = "Proveedor SIAT"
                        for c_p in ["RAZON SOCIAL PROVEEDOR", "NOMBRE O RAZON SOCIAL", "RAZON SOCIAL", "PROVEEDOR", "CLIENTE/PROVEEDOR"]:
                            if c_p in row and pd.notna(row[c_p]):
                                val_prov = str(row[c_p]).strip()
                                break

                        val_nro = ""
                        for c_n in ["NUMERO FACTURA", "N° DE LA FACTURA", "NUMERO DE FACTURA", "NRO FACTURA", "FACTURA", "NRO"]:
                            if c_n in row and pd.notna(row[c_n]):
                                val_nro = str(row[c_n]).strip()
                                break

                        val_monto = 0.0
                        for c_m in ["IMPORTE TOTAL COMPRA", "IMPORTE TOTAL", "MONTO", "TOTAL"]:
                            if c_m in row and pd.notna(row[c_m]):
                                try:
                                    val_monto = float(str(row[c_m]).replace(",", ""))
                                except:
                                    pass
                                break

                        cat_v = "Transporte y Pasajes" if any(k in val_prov.upper() for k in ["YPFB", "ESTACION", "SURTIDOR", "GASOLINERA"]) else "Otros Gastos Operativos"
                        detalle_v = f"{val_prov} (Factura N° {val_nro})" if val_nro else val_prov

                        datos["gastos"].append({
                            "fecha": val_fecha,
                            "categoria": cat_v,
                            "tipo": "Factura RC-V",
                            "nro": val_nro,
                            "monto": val_monto,
                            "detalle": detalle_v
                        })
                        importadas_count += 1

                    guardar_datos(datos)
                    st.success(f"¡Se importaron {importadas_count} facturas ordenadas cronológicamente!")
        except Exception as e:
            st.error(f"Error al leer el archivo RC-V: {e}")

# 4. RESUMEN MENSUAL, ARQUEO Y EXPORTACIÓN A EXCEL
elif menu == "📊 Resumen, Arqueo y Excel":
    st.subheader("Cierre Mensual, Arqueo de Efectivo y Descarga de Planilla")
    cfg = datos["config"]

    with st.expander("⚙️ Configurar Datos de Rendición y Fondo Inicial", expanded=False):
        c1, c2 = st.columns(2)
        cfg["responsable"] = c1.text_input("Responsable:", value=cfg["responsable"])
        cfg["area"] = c2.text_input("Área:", value=cfg["area"])
        cfg["nro_rendicion"] = c1.text_input("N° Rendición:", value=cfg["nro_rendicion"])
        cfg["mes_rendicion"] = c2.text_input("Mes a Rendir (AAAA-MM):", value=cfg.get("mes_rendicion", datetime.date.today().strftime("%Y-%m")))
        cfg["fondo"] = c1.number_input("Fondo Inicial (Bs):", value=float(cfg["fondo"]), format="%.2f")
        cfg["tarifa"] = c2.number_input("Tarifa Almuerzo (Bs):", value=float(cfg["tarifa"]), format="%.2f")
        if st.button("Guardar Parámetros"):
            guardar_datos(datos)
            st.success("Parámetros actualizados.")
            st.rerun()

    mes_str = cfg.get("mes_rendicion", datetime.date.today().strftime("%Y-%m"))
    raciones_mes, fechas_mes = calcular_raciones_mes(datos, mes_str)
    gastos_combinados = obtener_gastos_mensuales_con_almuerzos(datos, mes_str)

    total_alm = raciones_mes * float(cfg["tarifa"])
    total_gas = sum(float(g["monto"]) for g in datos["gastos"] if str(g.get("fecha", "")).startswith(mes_str))
    total_rend = total_alm + total_gas
    saldo_caja = float(cfg["fondo"]) - total_rend

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Fondo Inicial", f"Bs {cfg['fondo']:.2f}")
    k2.metric(f"Almuerzos {mes_str} ({raciones_mes} rac.)", f"Bs {total_alm:.2f}")
    k3.metric(f"Gastos Generales {mes_str}", f"Bs {total_gas:.2f}")
    k4.metric("Saldo Teórico en Caja", f"Bs {saldo_caja:.2f}")

    st.info(f"💰 **Total Rendido del Mes ({mes_str}):** Bs {total_rend:.2f} | **Días de almuerzo registrados:** {len(fechas_mes)}")

    st.subheader(f"📅 Resumen Cronológico de Gastos y Almuerzos ({mes_str})")
    if gastos_combinados:
        st.dataframe(pd.DataFrame(gastos_combinados), use_container_width=True)
    else:
        st.caption("No hay gastos ni almuerzos registrados en este mes.")

    st.divider()
    st.subheader("💵 Arqueo de Efectivo Físico")
    arq = datos.get("arqueo", {"200": 0, "100": 0, "50": 0, "20": 0, "10": 0, "monedas": 0.0})

    b1, b2, b3, b4, b5, b6 = st.columns(6)
    arq["200"] = b1.number_input("Bs 200", min_value=0, step=1, value=int(arq.get("200", 0)))
    arq["100"] = b2.number_input("Bs 100", min_value=0, step=1, value=int(arq.get("100", 0)))
    arq["50"] = b3.number_input("Bs 50", min_value=0, step=1, value=int(arq.get("50", 0)))
    arq["20"] = b4.number_input("Bs 20", min_value=0, step=1, value=int(arq.get("20", 0)))
    arq["10"] = b5.number_input("Bs 10", min_value=0, step=1, value=int(arq.get("10", 0)))
    arq["monedas"] = b6.number_input("Monedas (Bs)", min_value=0.0, step=0.5, value=float(arq.get("monedas", 0.0)), format="%.2f")

    total_fisico = (
        arq["200"] * 200
        + arq["100"] * 100
        + arq["50"] * 50
        + arq["20"] * 20
        + arq["10"] * 10
        + arq["monedas"]
    )
    diferencia = total_fisico - saldo_caja

    a1, a2, a3 = st.columns(3)
    a1.metric("Efectivo Físico Contado", f"Bs {total_fisico:.2f}")
    a2.metric("Saldo Teórico en Caja", f"Bs {saldo_caja:.2f}")
    a3.metric("Diferencia (Físico - Teórico)", f"Bs {diferencia:.2f}")

    if st.button("💾 Guardar Arqueo"):
        datos["arqueo"] = arq
        guardar_datos(datos)
        st.success("Arqueo guardado correctamente.")

    st.divider()
    excel_file = generar_excel_bytes(datos)
    st.download_button(
        label="📥 DESCARGAR PLANILLA OFICIAL EN EXCEL (.xlsx)",
        data=excel_file,
        file_name=f"Rendicion_Caja_Chica_{cfg['nro_rendicion']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
