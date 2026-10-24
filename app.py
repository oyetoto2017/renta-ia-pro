import io
import os
import re
import json
import hmac
import hashlib
import datetime
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, Any, List, Tuple

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pypdf
import streamlit as st
from docx import Document
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
from google.genai.errors import APIError

try:
    from pdf2image import convert_from_bytes
    import pytesseract
    OCR_DISPONIBLE = True
except Exception:
    OCR_DISPONIBLE = False

st.set_page_config(
    page_title="RentaIA Pro | Declaración de Renta DIAN 210",
    page_icon="📊",
    layout="wide",
)

# ==============================================================================
# 1. PARÁMETROS OFICIALES DIAN MULTIAÑO (PREPARADO PARA CUALQUIER AÑO FUTURO)
# ==============================================================================
UVT_OFICIAL = {
    2022: 38_004,
    2023: 42_412,  # Resolución DIAN 001264 de 2022
    2024: 47_065,  # Resolución DIAN 000187 de 2023
    2025: 49_799,  # Resolución DIAN 000193 de 2024 (AG 2025 se presenta en 2026)
    2026: 52_374,  # Resolución DIAN 000238 de 2025 (AG 2026 se presenta en 2027)
}

COMP_INFLACIONARIO_HISTORICO = {
    2023: 0.6671,
    2024: 0.5088,
    2025: 0.5544,
    2026: 0.5544,
}

SECCIONALES_DIAN = [
    "32 - Bogotá D.C.",
    "11 - Medellín",
    "05 - Cali",
    "02 - Barranquilla",
    "04 - Bucaramanga",
    "06 - Cartagena",
    "09 - Ibagué",
    "10 - Manizales",
    "14 - Pasto",
    "16 - Pereira",
    "20 - Tunja",
    "26 - Villavicencio",
    "43 - Yopal",
    "13 - Neiva",
    "07 - Cúcuta",
    "19 - Santa Marta",
    "01 - Armenia",
    "12 - Montería",
    "15 - Popayán",
    "21 - Valledupar",
    "00 - Otra Dirección Seccional",
]

ACTIVIDADES_CIIU_COMUNES = [
    "0010 - Asalariado / Empleado",
    "8211 - Actividades combinadas de servicios administrativos de oficina",
    "0090 - Rentista de capital (CDTs, intereses, arrendamientos)",
    "0020 - Pensionado",
    "7020 - Actividades de consultoría de gestión / administración",
    "6910 - Actividades jurídicas",
    "6920 - Actividades de contabilidad, teneduría de libros y auditoría",
    "6201 - Desarrollo de sistemas informáticos y tecnología",
    "7112 - Actividades de ingeniería y consultoría técnica",
    "4711 - Comercio al por menor en establecimientos no especializados",
    "6810 - Actividades inmobiliarias realizadas con bienes propios o arrendados",
    "8559 - Otros tipos de educación y capacitación",
    "OTRO - Escribir otro código CIIU manualmente",
]

LIMITES_LEY_2277 = {
    "pct_bolsa_global": 0.40,
    "tope_bolsa_uvt": 1_340,
    "exenta_25_uvt": 790,
    "vivienda_uvt": 1_200,
    "prepagada_mes_uvt": 16,
    "dep_10pct_mes_uvt": 32,
    "dep_extra_uvt": 72,
    "max_dep_extra": 4,
    "icetex_uvt": 100,
    "factura_elec_pct": 0.01,
    "factura_elec_uvt": 240,
    "afc_fpv_uvt": 3_800,
    "pensiones_exentas_uvt": 12_000,
}

TABLA_ART_241 = [
    (0, 1_090, 0.00, 0),
    (1_090, 1_700, 0.19, 0),
    (1_700, 4_100, 0.28, 116),
    (4_100, 8_670, 0.33, 788),
    (8_670, 18_970, 0.35, 2_296),
    (18_970, 31_000, 0.37, 5_901),
    (31_000, float("inf"), 0.39, 10_352),
]


def uvt_sugerida_por_anio(anio: int) -> int:
    if anio in UVT_OFICIAL:
        return UVT_OFICIAL[anio]
    return UVT_OFICIAL[max(UVT_OFICIAL.keys())]


def calcular_dv_dian(nit_str: str) -> str:
    """Calcula el Dígito de Verificación oficial DIAN (Casilla 6) con módulo 11."""
    limpio = "".join(ch for ch in str(nit_str) if ch.isdigit())
    if not limpio:
        return ""
    primos = [3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71]
    suma = sum(int(dig) * primos[i] for i, dig in enumerate(reversed(limpio)) if i < len(primos))
    residuo = suma % 11
    return str(residuo if residuo in (0, 1) else 11 - residuo)


def orientacion_vencimiento_dian(nit_str: str, anio_presentacion: int) -> str:
    """Devuelve orientación de mes estimado de vencimiento según los 2 últimos dígitos."""
    limpio = "".join(ch for ch in str(nit_str) if ch.isdigit())
    if len(limpio) < 2:
        return ""
    ultimos_dos = int(limpio[-2:])
    if 1 <= ultimos_dos <= 26:
        mes_est = "Agosto"
    elif 27 <= ultimos_dos <= 66:
        mes_est = "Septiembre"
    else:
        mes_est = "Octubre"
    return f"Últimos 2 dígitos: **{limpio[-2:]}** (Vencimiento habitual DIAN: **{mes_est} de {anio_presentacion}**)"


# ==============================================================================
# 2. SEGURIDAD, LIMPIEZA DE API KEY Y MANEJO ANTIFALLOS DE GEMINI
# ==============================================================================
def limpiar_api_key(raw_key: str) -> str:
    if not raw_key:
        return ""
    texto = str(raw_key).strip()
    match = re.search(r"(AIza[0-9A-Za-z_\-]{25,})", texto)
    if match:
        return match.group(1)
    return re.sub(r'[\s"\']', "", texto)


def obtener_secreto(nombre: str) -> str:
    try:
        val = str(st.secrets.get(nombre, "")).strip()
    except Exception:
        val = os.getenv(nombre, "").strip()
    return limpiar_api_key(val) if nombre == "GEMINI_API_KEY" else val


def generar_licencia_cliente(identificador: str, semilla: str) -> str:
    limpio = identificador.strip().upper()
    firma = hmac.new(semilla.encode("utf-8"), limpio.encode("utf-8"), hashlib.sha256).hexdigest()[:12].upper()
    return f"RENTA-{firma}"


def verificar_acceso() -> bool:
    clave_maestra = obtener_secreto("CLAVE_MAESTRA")
    if not clave_maestra:
        st.session_state["es_admin"] = True
        return True
    if st.session_state.get("autenticado", False):
        return True

    st.title("🔐 RentaIA Pro — Acceso Protegido")
    st.markdown("Ingresa tu **Clave Maestra de Administrador** o una **Licencia Activa**.")
    c1, c2 = st.columns(2)
    doc_usuario = c1.text_input("Documento / Correo del Titular (solo para clientes)")
    codigo_ingresado = c2.text_input("Clave Maestra o Código de Licencia", type="password")

    if st.button("Ingresar a la Plataforma", type="primary"):
        if hmac.compare_digest(codigo_ingresado.strip(), clave_maestra):
            st.session_state["autenticado"] = True
            st.session_state["es_admin"] = True
            st.rerun()
        elif doc_usuario.strip():
            lic_esperada = generar_licencia_cliente(doc_usuario, clave_maestra)
            if hmac.compare_digest(codigo_ingresado.strip().upper(), lic_esperada):
                st.session_state["autenticado"] = True
                st.session_state["es_admin"] = False
                st.rerun()
            else:
                st.error("Licencia inválida para el documento ingresado.")
        else:
            st.error("Credenciales incorrectas.")
    return False


def ejecutar_gemini_seguro(api_key: str, contents: Any, config: types.GenerateContentConfig) -> str:
    llave = limpiar_api_key(api_key)
    if not llave or not llave.startswith("AIza"):
        raise ValueError(
            "Configura una Gemini API Key válida (inicia por 'AIza...') en Secrets o en la barra lateral."
        )

    client = genai.Client(api_key=llave)
    ultimo_error = None
    for mod in ["gemini-2.5-flash", "gemini-2.0-flash"]:
        try:
            resp = client.models.generate_content(model=mod, contents=contents, config=config)
            return resp.text
        except APIError as e:
            ultimo_error = e
            codigo = getattr(e, "code", None)
            if codigo in (400, 401, 403):
                raise ValueError(
                    f"Google rechazó la API Key (Código {codigo}). Verifica en Google AI Studio que tu llave esté activa."
                )
            continue
        except Exception as e:
            ultimo_error = e
            continue

    raise RuntimeError(
        f"El servicio de IA alcanzó temporalmente el límite por minuto. Espera unos segundos e intenta de nuevo ({type(ultimo_error).__name__})."
    )


# ==============================================================================
# 3. MOTOR MATEMÁTICO DETERMINISTA UNIVERSAL (AYUDA RENTA DIAN + LEY 2277)
# ==============================================================================
def redondear_dian(valor: float) -> int:
    """Redondeo oficial DIAN al múltiplo de 1.000 más cercano (Art. 577 E.T.)."""
    if valor <= 0:
        return 0
    return int(
        (Decimal(str(valor)) / Decimal("1000")).quantize(Decimal("1"), rounding=ROUND_HALF_UP) * 1000
    )


def estado_inicial_datos() -> Dict[str, Any]:
    return {
        # Datos de Identificación Formulario 210 (Casillas 5 a 24)
        "nombre_declarante": "",
        "nit_declarante": "",
        "seccional_dian": "32 - Bogotá D.C.",
        "cod_actividad": "8211",
        # Patrimonio y Topes DIAN (Casillas 29 a 31)
        "efectivo": 0.0,
        "otros_activos": 0.0,
        "inmuebles": 0.0,
        "vehiculos": 0.0,
        "deudas": 0.0,
        "consignaciones": 0.0,
        "tarjetas_credito": 0.0,
        "compras_totales": 0.0,
        "responsable_iva": False,
        "patrimonio_liquido_anterior": 0.0,
        # Subcédula Trabajo Asalariado (Casillas 32 a 36)
        "ing_brutos_trabajo": 0.0,
        "salud_pension_trabajo": 0.0,
        "cesantias_exentas": 0.0,
        # Subcédula Independientes: Honorarios y Servicios (Simula C.32 vs C.43-46)
        "ing_brutos_honorarios": 0.0,
        "salud_pension_honorarios": 0.0,
        "arl_honorarios": 0.0,
        "costos_honorarios": 0.0,
        # Subcédula Rentas de Capital: CDTs, Intereses, Arriendos (Casillas 58 a 61)
        "ing_brutos_capital": 0.0,
        "incrngo_capital": 0.0,
        "costos_capital": 0.0,
        # Subcédula Rentas No Laborales: Ventas / Comercio (Casillas 74 a 78)
        "ing_brutos_no_laborales": 0.0,
        "incrngo_no_laborales": 0.0,
        "costos_no_laborales": 0.0,
        # Pensiones, Dividendos y Ganancias Ocasionales
        "ing_brutos_pensiones": 0.0,
        "salud_pensiones": 0.0,
        "dividendos_gravables": 0.0,
        "ing_ganancia_ocasional": 0.0,
        "ing_loterias_rifas": 0.0,
        "costos_ganancia_ocasional": 0.0,
        "exenta_ganancia_ocasional": 0.0,
        # Beneficios Tributarios Ley 2277 y Descuentos
        "num_dependientes": 0,
        "compras_factura_electronica": 0.0,
        "aplicar_beneficio_fe": False,
        "intereses_vivienda": 0.0,
        "medicina_prepagada": 0.0,
        "meses_prepagada": 12,
        "intereses_icetex": 0.0,
        "gmf_certificado_100": 0.0,
        "aportes_afc_fpv": 0.0,
        "donaciones_entidades": 0.0,
        # Liquidación Privada: Retenciones, Saldos, Anticipo y Extemporaneidad
        "retenciones_anio": 0.0,
        "saldo_favor_anterior": 0.0,
        "anio_declara_num": 0,
        "impuesto_neto_anterior": 0.0,
        "anticipo_anio_anterior": 0.0,
        "meses_extemporaneidad": 0,
        "aplica_reduccion_640": False,
    }


def calcular_impuesto_art241(rlg: float, uvt: float) -> float:
    if rlg <= 0 or uvt <= 0:
        return 0.0
    rlg_uvt = rlg / uvt
    for desde, hasta, tarifa, base_uvt in TABLA_ART_241:
        if desde < rlg_uvt <= hasta:
            return (base_uvt + ((rlg_uvt - desde) * tarifa)) * uvt
    return 0.0


def liquidar_declaracion_210(
    d: Dict[str, Any], uvt: float, uvt_pat: float, uvt_ing: float, uvt_sancion: float
) -> Dict[str, Any]:
    L = LIMITES_LEY_2277

    # 1. PATRIMONIO (Casillas 29, 30, 31)
    pat_bruto_raw = d["efectivo"] + d["otros_activos"] + d["inmuebles"] + d["vehiculos"]
    c29_patrimonio_bruto = redondear_dian(pat_bruto_raw)
    c30_deudas = redondear_dian(d["deudas"])
    c31_patrimonio_liquido = max(0, c29_patrimonio_bruto - c30_deudas)

    # 2. SUBCÉDULAS CAPITAL Y NO LABORALES
    c58_ing_capital = redondear_dian(d["ing_brutos_capital"])
    c59_inc_capital = min(c58_ing_capital, redondear_dian(d["incrngo_capital"]))
    neto_capital = max(0, c58_ing_capital - c59_inc_capital)
    c60_costos_capital = min(neto_capital, redondear_dian(d["costos_capital"]))
    c61_renta_liq_capital = max(0, neto_capital - c60_costos_capital)

    c74_ing_nolab = redondear_dian(d["ing_brutos_no_laborales"])
    c76_inc_nolab = min(c74_ing_nolab, redondear_dian(d["incrngo_no_laborales"]))
    neto_nolab = max(0, c74_ing_nolab - c76_inc_nolab)
    c77_costos_nolab = min(neto_nolab, redondear_dian(d["costos_no_laborales"]))
    c78_renta_liq_nolab = max(0, neto_nolab - c77_costos_nolab)

    # 3. SIMULADOR INTELIGENTE MUISCA: RUTA A (25% Exento C.32) vs RUTA B (Honorarios con Costos C.43-45)
    ing_trab_total_bruto = redondear_dian(d["ing_brutos_trabajo"]) + redondear_dian(d["ing_brutos_honorarios"])
    inc_trab_total = min(
        ing_trab_total_bruto,
        redondear_dian(d["salud_pension_trabajo"]) + redondear_dian(d["salud_pension_honorarios"]),
    )
    ing_brutos_cg = ing_trab_total_bruto + c58_ing_capital + c74_ing_nolab
    ing_netos_cg = max(0, ing_trab_total_bruto - inc_trab_total) + neto_capital + neto_nolab

    hay_trabajo_o_hon = ing_trab_total_bruto > 0
    c38_vivienda = redondear_dian(min(d["intereses_vivienda"], L["vivienda_uvt"] * uvt))
    ded_prep = redondear_dian(min(d["medicina_prepagada"], d["meses_prepagada"] * L["prepagada_mes_uvt"] * uvt)) if hay_trabajo_o_hon else 0
    ded_dep10 = redondear_dian(min(ing_trab_total_bruto * 0.10, 12 * L["dep_10pct_mes_uvt"] * uvt)) if (d["num_dependientes"] > 0 and hay_trabajo_o_hon) else 0
    ded_icetex = redondear_dian(min(d.get("intereses_icetex", 0.0), L["icetex_uvt"] * uvt))
    ded_gmf = redondear_dian(d["gmf_certificado_100"] * 0.50)
    c39_otras_ded = ded_prep + ded_dep10 + ded_icetex + ded_gmf
    c35_afc = redondear_dian(min(d["aportes_afc_fpv"], ing_brutos_cg * 0.30, L["afc_fpv_uvt"] * uvt))
    cesantias_ex = redondear_dian(d.get("cesantias_exentas", 0.0))

    deducciones_comunes = c38_vivienda + c39_otras_ded + c35_afc + cesantias_ex
    limite_40_cg = redondear_dian(min(ing_netos_cg * L["pct_bolsa_global"], L["tope_bolsa_uvt"] * uvt))
    tope_790_pesos = L["exenta_25_uvt"] * uvt

    # Ruta A: Trabajo + Servicios se declaran en Casilla 32 aplicando el 25% exento
    c32_a = ing_trab_total_bruto
    c33_a = inc_trab_total
    c34_a = max(0, c32_a - c33_a)
    c91_a = c34_a + c61_renta_liq_capital + c78_renta_liq_nolab
    base_25_a = max(0, c34_a - deducciones_comunes)
    exenta_25_a = redondear_dian(min(base_25_a * 0.25, tope_790_pesos))
    solicitados_a = deducciones_comunes + exenta_25_a
    aceptados_40_a = min(solicitados_a, limite_40_cg, c91_a)
    remanente_a = max(0, c91_a - aceptados_40_a)

    # Ruta B: Honorarios en Casillas 43-46 imputando costos reales + ARL (Casilla 45)
    costos_hon_total = redondear_dian(d["costos_honorarios"] + d.get("arl_honorarios", 0.0))
    c32_b = redondear_dian(d["ing_brutos_trabajo"])
    c33_b = min(c32_b, redondear_dian(d["salud_pension_trabajo"]))
    c34_b = max(0, c32_b - c33_b)
    c43_b = redondear_dian(d["ing_brutos_honorarios"])
    pila_disp_hon = redondear_dian(d["salud_pension_honorarios"]) + max(0, redondear_dian(d["salud_pension_trabajo"]) - c33_b)
    c44_b = min(c43_b, pila_disp_hon)
    neto_hon_b = max(0, c43_b - c44_b)
    c45_b = min(neto_hon_b, costos_hon_total)
    c46_b = max(0, neto_hon_b - c45_b)
    c91_b = c34_b + c46_b + c61_renta_liq_capital + c78_renta_liq_nolab
    base_25_b = max(0, c34_b - deducciones_comunes)
    exenta_25_b = redondear_dian(min(base_25_b * 0.25, tope_790_pesos))
    solicitados_b = deducciones_comunes + exenta_25_b
    aceptados_40_b = min(solicitados_b, limite_40_cg, c91_b)
    remanente_b = max(0, c91_b - aceptados_40_b)

    if costos_hon_total > 0 and c43_b > 0 and remanente_b < remanente_a:
        estrategia = "Honorarios con Costos y ARL (Casillas 43 a 46 — menor impuesto)"
        c32_ing_trabajo, c33_inc_trabajo, c34_renta_liq_trabajo = c32_b, c33_b, c34_b
        c43_ing_honorarios, c44_inc_honorarios, c45_costos_hon, c46_renta_liq_hon = c43_b, c44_b, c45_b, c46_b
        c91_renta_liq_cg = c91_b
        c36_exenta_25 = exenta_25_b + cesantias_ex
        beneficios_40_limitados = aceptados_40_b
        solicitados_bolsa_40 = solicitados_b
        remanente_cg = remanente_b
    else:
        estrategia = "Rentas de Trabajo con 25% Exento (Casillas 32 a 36 — óptima Ley 2277)"
        c32_ing_trabajo, c33_inc_trabajo, c34_renta_liq_trabajo = c32_a, c33_a, c34_a
        c43_ing_honorarios, c44_inc_honorarios, c45_costos_hon, c46_renta_liq_hon = 0, 0, 0, 0
        c91_renta_liq_cg = c91_a
        c36_exenta_25 = exenta_25_a + cesantias_ex
        beneficios_40_limitados = aceptados_40_a
        solicitados_bolsa_40 = solicitados_a
        remanente_cg = remanente_a

    # 4. BENEFICIOS FUERA DEL LÍMITE DEL 40% (Casilla 28 Factura Electrónica + Casilla 139 Dependientes)
    ded_fe_potencial = redondear_dian(
        min(d["compras_factura_electronica"] * L["factura_elec_pct"], L["factura_elec_uvt"] * uvt)
    )
    c28_ded_factura_1pct = ded_fe_potencial if d.get("aplicar_beneficio_fe", False) else 0
    c139_adicion_dependientes = (
        redondear_dian(min(int(d["num_dependientes"]), L["max_dep_extra"]) * L["dep_extra_uvt"] * uvt)
        if hay_trabajo_o_hon else 0
    )
    beneficios_fuera_40 = min(remanente_cg, c28_ded_factura_1pct + c139_adicion_dependientes)

    c92_exentas_y_ded_limitadas = beneficios_40_limitados + beneficios_fuera_40
    c97_rlg_cedula_general = max(0, c91_renta_liq_cg - c92_exentas_y_ded_limitadas)

    # 5. PENSIONES, DIVIDENDOS Y GANANCIAS OCASIONALES
    c99_pen_brutos = redondear_dian(d["ing_brutos_pensiones"])
    c100_pen_salud = min(c99_pen_brutos, redondear_dian(d["salud_pensiones"]))
    c101_pen_liquida = max(0, c99_pen_brutos - c100_pen_salud)
    c102_pen_exenta = min(c101_pen_liquida, redondear_dian(L["pensiones_exentas_uvt"] * uvt))
    c103_rlg_pensiones = max(0, c101_pen_liquida - c102_pen_exenta)

    c107_dividendos = redondear_dian(d["dividendos_gravables"])
    c111_rlg_total = c97_rlg_cedula_general + c103_rlg_pensiones + c107_dividendos

    go_activos_bruto = redondear_dian(d["ing_ganancia_ocasional"])
    go_loterias_bruto = redondear_dian(d.get("ing_loterias_rifas", 0.0))
    c112_go_ingresos = go_activos_bruto + go_loterias_bruto
    c113_go_costos = redondear_dian(d["costos_ganancia_ocasional"])
    c114_go_exentas = redondear_dian(d["exenta_ganancia_ocasional"])
    go_activos_grav = max(0, go_activos_bruto - c113_go_costos - c114_go_exentas)
    c115_go_gravable = go_activos_grav + go_loterias_bruto

    # 6. EVALUACIÓN DE LOS 6 TOPES DIAN
    ing_brutos_totales = ing_brutos_cg + c99_pen_brutos + c107_dividendos + c112_go_ingresos
    tope_pat_pesos = uvt_pat * uvt
    tope_ing_pesos = uvt_ing * uvt

    causales: List[str] = []
    if pat_bruto_raw >= tope_pat_pesos:
        causales.append(f"Patrimonio Bruto >= {uvt_pat:,.0f} UVT (${tope_pat_pesos:,.0f})")
    if ing_brutos_totales >= tope_ing_pesos:
        causales.append(f"Ingresos Brutos >= {uvt_ing:,.0f} UVT (${tope_ing_pesos:,.0f})")
    if d["consignaciones"] >= tope_ing_pesos:
        causales.append(f"Consignaciones/Inversiones >= {uvt_ing:,.0f} UVT (${tope_ing_pesos:,.0f})")
    if d["tarjetas_credito"] >= tope_ing_pesos:
        causales.append(f"Tarjetas de Crédito >= {uvt_ing:,.0f} UVT")
    if d["compras_totales"] >= tope_ing_pesos:
        causales.append(f"Compras Totales >= {uvt_ing:,.0f} UVT")
    if d["responsable_iva"]:
        causales.append("Responsable de IVA a 31 de diciembre")

    obligado_bool = len(causales) > 0
    estado_obligacion = "Sí obligado a declarar" if obligado_bool else "No obligado a declarar"

    # 7. LIQUIDACIÓN DE IMPUESTO, DESCUENTOS, ANTICIPO Y SANCIONES (Casillas 116 a 137)
    c116_impuesto_rlg = redondear_dian(calcular_impuesto_art241(c111_rlg_total, uvt))
    desc_donaciones = redondear_dian(d.get("donaciones_entidades", 0.0) * 0.25)
    desc_div = redondear_dian(max(0.0, (c107_dividendos - (1_090 * uvt)) * 0.19)) if c107_dividendos > (1_090 * uvt) else 0
    c125_descuentos_trib = min(c116_impuesto_rlg, desc_donaciones + desc_div)
    c126_impuesto_neto_renta = max(0, c116_impuesto_rlg - c125_descuentos_trib)

    c127_impuesto_go = redondear_dian((go_activos_grav * 0.15) + (go_loterias_bruto * 0.20))
    c129_total_impuesto_cargo = c126_impuesto_neto_renta + c127_impuesto_go

    c130_anticipo_anterior = redondear_dian(d["anticipo_anio_anterior"])
    c131_saldo_favor_anterior = redondear_dian(d["saldo_favor_anterior"])
    c132_retenciones_anio = redondear_dian(d["retenciones_anio"])
    imp_ant = redondear_dian(d["impuesto_neto_anterior"])

    anio_n = int(d["anio_declara_num"])
    if anio_n == 0 or c126_impuesto_neto_renta == 0:
        c133_anticipo_siguiente = 0
        ahorro_ant = 0
    else:
        pct = {1: 0.25, 2: 0.50}.get(anio_n, 0.75)
        op1 = max(0.0, (c126_impuesto_neto_renta * pct) - c132_retenciones_anio)
        if anio_n > 1:
            op2 = max(0.0, (((c126_impuesto_neto_renta + imp_ant) / 2.0) * pct) - c132_retenciones_anio)
            c133_anticipo_siguiente = redondear_dian(min(op1, op2))
            ahorro_ant = redondear_dian(abs(op1 - op2))
        else:
            c133_anticipo_siguiente = redondear_dian(op1)
            ahorro_ant = 0

    c135_sanciones = 0
    meses_mora = int(d["meses_extemporaneidad"])
    if meses_mora > 0:
        sancion_minima = redondear_dian(10 * uvt_sancion)
        if c129_total_impuesto_cargo > 0:
            bruta = min(c129_total_impuesto_cargo, c129_total_impuesto_cargo * 0.05 * meses_mora)
        else:
            bruta = min(ing_brutos_totales * 0.05, ing_brutos_totales * 0.005 * meses_mora)
        if d["aplica_reduccion_640"]:
            bruta *= 0.50
        c135_sanciones = max(sancion_minima, redondear_dian(bruta))

    alivios_totales = c130_anticipo_anterior + c131_saldo_favor_anterior + c132_retenciones_anio
    c134_saldo_pagar_impuesto = max(0, c129_total_impuesto_cargo + c133_anticipo_siguiente - alivios_totales)
    balance_total = c129_total_impuesto_cargo + c133_anticipo_siguiente + c135_sanciones - alivios_totales
    c136_total_saldo_pagar = max(0, balance_total)
    c137_total_saldo_favor = abs(min(0, balance_total))

    # 8. AUDITORÍA PREVENTIVA Y OPORTUNIDADES
    alertas: List[str] = []
    pat_ant = redondear_dian(d["patrimonio_liquido_anterior"])
    if pat_ant > 0:
        delta_pat = c31_patrimonio_liquido - pat_ant
        justificable = ing_netos_cg + c101_pen_liquida + c115_go_gravable - imp_ant
        if delta_pat > justificable:
            alertas.append(
                f"Riesgo de Comparación Patrimonial (Art. 236 E.T.): Tu patrimonio líquido creció "
                f"${redondear_dian(delta_pat - justificable):,.0f} por encima de tus ingresos netos justificables."
            )

    sugerencias: List[str] = []
    if hay_trabajo_o_hon and inc_trab_total == 0:
        sugerencias.append(
            "Tienes ingresos de trabajo/servicios pero $0 en Salud y Pensión: sube tus planillas PILA en el Paso 1 o escríbelos en el Paso 3."
        )
    if ded_fe_potencial > 0 and not d.get("aplicar_beneficio_fe", False):
        sugerencias.append(
            f"Tienes compras con Factura Electrónica que permiten restar ${ded_fe_potencial:,.0f} adicionales en Casilla 28. Actívalo en el Paso 4 si pagaste con medios electrónicos."
        )
    cupo_40 = max(0, limite_40_cg - solicitados_bolsa_40)
    if cupo_40 > 300_000 and c126_impuesto_neto_renta > 0:
        sugerencias.append(
            f"Aún tienes ${cupo_40:,.0f} de cupo libre en la bolsa del 40% / 1.340 UVT para reducir tu impuesto (50% del 4x1000, intereses de vivienda, ICETEX, medicina prepagada o aportes AFC/FPV)."
        )

    return {
        "obligado_bool": obligado_bool,
        "estado_obligacion": estado_obligacion,
        "causales_obligacion": causales,
        "estrategia": estrategia,
        "remanente_ruta_25": remanente_a,
        "remanente_ruta_costos": remanente_b,
        "limite_40_cg": limite_40_cg,
        "usado_40_cg": beneficios_40_limitados,
        "cupo_libre_40": cupo_40,
        "ing_brutos_totales": ing_brutos_totales,
        "tope_pat_pesos": tope_pat_pesos,
        "tope_ing_pesos": tope_ing_pesos,
        # Casillas oficiales exactas Formulario 210
        "c28_factura_electronica": c28_ded_factura_1pct,
        "c29_patrimonio_bruto": c29_patrimonio_bruto,
        "c30_deudas": c30_deudas,
        "c31_patrimonio_liquido": c31_patrimonio_liquido,
        "c32_ing_trabajo": c32_ing_trabajo,
        "c33_inc_trabajo": c33_inc_trabajo,
        "c34_renta_liq_trabajo": c34_renta_liq_trabajo,
        "c35_afc_fpv": c35_afc,
        "c36_exenta_25": c36_exenta_25,
        "c38_vivienda": c38_vivienda,
        "c39_prepagada": c39_otras_ded,
        "c43_ing_honorarios": c43_ing_honorarios,
        "c44_inc_honorarios": c44_inc_honorarios,
        "c45_costos_honorarios": c45_costos_hon,
        "c46_renta_liq_honorarios": c46_renta_liq_hon,
        "c58_ing_capital": c58_ing_capital,
        "c59_inc_capital": c59_inc_capital,
        "c60_costos_capital": c60_costos_capital,
        "c61_renta_liq_capital": c61_renta_liq_capital,
        "c74_ing_nolab": c74_ing_nolab,
        "c76_inc_nolab": c76_inc_nolab,
        "c77_costos_nolab": c77_costos_nolab,
        "c78_renta_liq_nolab": c78_renta_liq_nolab,
        "c91_renta_liq_cg": c91_renta_liq_cg,
        "c92_exentas_ded_limitadas": c92_exentas_y_ded_limitadas,
        "c97_rlg_cedula_general": c97_rlg_cedula_general,
        "c99_pen_brutos": c99_pen_brutos,
        "c100_pen_salud": c100_pen_salud,
        "c102_pen_exenta": c102_pen_exenta,
        "c103_rlg_pensiones": c103_rlg_pensiones,
        "c107_dividendos": c107_dividendos,
        "c111_rlg_total": c111_rlg_total,
        "rlg_en_uvt": round(c111_rlg_total / uvt, 2) if uvt > 0 else 0,
        "c112_go_ingresos": c112_go_ingresos,
        "c113_go_costos": c113_go_costos,
        "c114_go_exentas": c114_go_exentas,
        "c115_go_gravable": c115_go_gravable,
        "c116_impuesto_rlg": c116_impuesto_rlg,
        "c125_descuentos_trib": c125_descuentos_trib,
        "c126_impuesto_neto_renta": c126_impuesto_neto_renta,
        "c127_impuesto_go": c127_impuesto_go,
        "c129_total_impuesto_cargo": c129_total_impuesto_cargo,
        "c130_anticipo_anterior": c130_anticipo_anterior,
        "c131_saldo_favor_anterior": c131_saldo_favor_anterior,
        "c132_retenciones_anio": c132_retenciones_anio,
        "c133_anticipo_siguiente": c133_anticipo_siguiente,
        "c134_saldo_pagar_impuesto": c134_saldo_pagar_impuesto,
        "c135_sanciones": c135_sanciones,
        "c136_total_saldo_pagar": c136_total_saldo_pagar,
        "c137_total_saldo_favor": c137_total_saldo_favor,
        "c138_num_dependientes": int(d["num_dependientes"]),
        "c139_adicion_dependientes": c139_adicion_dependientes,
        "alivios_totales": alivios_totales,
        "ahorro_anticipo": ahorro_ant,
        "alertas_auditoria": alertas,
        "sugerencias_optimizacion": sugerencias,
    }


# ==============================================================================
# 4. LECTOR MASIVO PILA MULTI-PDF + EXÓGENA DIAN + CERTIFICADOS IA EN LOTE
# ==============================================================================
class CertificadoExtraido(BaseModel):
    tipo_documento: str = Field(description="Ej: Certificado 220, Certificado CDT, Extracto Bancario, Predial, Prepagada")
    entidad: str = Field(description="Banco o empresa emisora")
    efectivo_bancos_dic31: float = Field(default=0.0, description="Saldo en cuentas a 31 de diciembre")
    valor_inmueble_o_cdt_dic31: float = Field(default=0.0, description="Avalúo predial o capital CDT vigente a dic 31")
    saldo_deuda_dic31: float = Field(default=0.0, description="Saldo capital de deudas/hipotecas a dic 31")
    ingresos_brutos_trabajo: float = Field(default=0.0, description="Salarios o pagos laborales (Casilla 32)")
    cesantias_exentas: float = Field(default=0.0, description="Cesantías e intereses de cesantías consignadas/pagadas")
    ingresos_brutos_honorarios: float = Field(default=0.0, description="Honorarios o servicios independientes (Casilla 43)")
    salud_pension_obligatoria: float = Field(default=0.0, description="Suma de salud obligatoria + pensión obligatoria")
    ingresos_capital_cdt: float = Field(default=0.0, description="Intereses o rendimientos financieros pagados (Casilla 58)")
    componente_inflacionario_cdt: float = Field(default=0.0, description="Valor no gravado / componente inflacionario CDT (Casilla 59)")
    intereses_vivienda: float = Field(default=0.0, description="Intereses crédito hipotecario o leasing")
    medicina_prepagada: float = Field(default=0.0, description="Pagos anuales medicina prepagada o póliza")
    gmf_4x1000_certificado: float = Field(default=0.0, description="100% del GMF 4x1000 certificado")
    aportes_afc_fpv: float = Field(default=0.0, description="Aportes voluntarios AFC o pensión voluntaria")
    retenciones_practicadas: float = Field(default=0.0, description="Retención en la fuente practicada en el año (Casilla 132)")


def _nvl(val) -> float:
    try:
        if val is None or isinstance(val, bool) or pd.isna(val):
            return 0.0
        if isinstance(val, str):
            limpio = re.sub(r"[^\d.\-]", "", val.replace(",", ""))
            return float(limpio) if limpio else 0.0
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def _parse_cop_texto(s: str) -> float:
    s = str(s).strip().replace("$", "").replace(" ", "")
    if not s:
        return 0.0
    if re.match(r"^\d{1,3}(\.\d{3})+(,\d+)?$", s):
        s = s.replace(".", "").replace(",", ".")
    elif re.match(r"^\d{1,3}(,\d{3})+(\.\d+)?$", s):
        s = s.replace(",", "")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def extraer_pila_pdf_nativo(nombre_archivo: str, archivo_bytes: bytes) -> dict:
    """Lee automáticamente Salud, Pensión y ARL de planillas PDF (MiPlanilla, SOI, Aportes en Línea, Asopagos) sin requerir IA."""
    try:
        reader = pypdf.PdfReader(io.BytesIO(archivo_bytes))
        texto_completo = "\n".join((p.extract_text() or "") for p in reader.pages)
    except Exception:
        return {"es_pila": False, "archivo": nombre_archivo}

    t_low = texto_completo.lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    es_pila = any(k in t_low for k in [
        "planilla integrada", "autoliquidacion de aportes", "miplanilla", "soi",
        "aportesenlinea", "asopagos", "totales por subsistema", "detalle por cotizante"
    ])
    if not es_pila:
        return {"es_pila": False, "archivo": nombre_archivo}

    lineas = [re.sub(r"\s+", " ", l).strip() for l in texto_completo.splitlines() if l.strip()]

    num_planilla = ""
    periodo = ""
    m_per = re.search(r"\b(20\d{2}-\d{2})\b", texto_completo)
    if m_per:
        periodo = m_per.group(1)

    m_nom = re.search(r"(\d{7,12})", nombre_archivo)
    for idx, l in enumerate(lineas):
        if re.match(r"^20\d{2}-\d{2}$", l) and idx + 1 < len(lineas):
            for k in range(idx + 1, min(idx + 5, len(lineas))):
                if re.match(r"^\d{7,12}$", lineas[k]):
                    num_planilla = lineas[k]
                    break
        if num_planilla:
            break
    if not num_planilla and m_nom:
        num_planilla = m_nom.group(1)

    salud, pension, arl, total_pagado = 0.0, 0.0, 0.0, 0.0

    # Estrategia 1: Celdas verticales en "TOTALES POR SUBSISTEMA" (toma el aporte neto antes de mora)
    for idx, l in enumerate(lineas):
        ll = l.lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
        ventana = lineas[idx + 1 : min(idx + 5, len(lineas))]
        nums_ventana = []
        for v in ventana:
            if re.match(r"^\$?\s*\d{1,3}(?:[.,]\d{3})+$", v) or re.match(r"^\$?\s*\d{4,8}$", v):
                val_n = _parse_cop_texto(v)
                if 1_000 <= val_n <= 35_000_000:
                    nums_ventana.append(val_n)
            elif any(stop in v.lower() for stop in ["salud", "pension", "pensión", "riesgos", "ccf", "totales"]):
                break

        if ll == "salud" and nums_ventana and salud == 0:
            salud = nums_ventana[0]
        elif ll == "pension" and nums_ventana and pension == 0:
            pension = nums_ventana[0]
        elif ll in ("riesgos laborales", "riesgos profesionales", "arl") and nums_ventana and arl == 0:
            arl = nums_ventana[0]
        elif ll == "totales" and nums_ventana and total_pagado == 0:
            total_pagado = nums_ventana[-1]

    # Estrategia 2: Formato horizontal en una sola línea
    if salud == 0 or pension == 0:
        for l in lineas:
            ll = l.lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
            montos = [
                _parse_cop_texto(m)
                for m in re.findall(r"\$\s*\d{1,3}(?:[.,]\d{3})+|\b\d{1,3}(?:\.\d{3})+\b", l)
            ]
            montos = [n for n in montos if 2_000 <= n <= 35_000_000]
            if not montos:
                continue
            if salud == 0 and ("salud" in ll or "eps" in ll) and "periodo" not in ll:
                salud = montos[-1]
            elif pension == 0 and ("pension" in ll or "afp" in ll or "colpensiones" in ll) and "periodo" not in ll:
                pension = montos[-1]
            elif arl == 0 and ("riesgos laborales" in ll or "positiva" in ll or "sura" in ll):
                arl = montos[-1]

    if total_pagado == 0:
        for l in lineas:
            if l.startswith("$"):
                val_p = _parse_cop_texto(l)
                if val_p > 0:
                    total_pagado = val_p
                    break

    return {
        "es_pila": (salud + pension) > 0,
        "archivo": nombre_archivo,
        "planilla": num_planilla or nombre_archivo,
        "periodo": periodo or "Detectado",
        "salud": salud,
        "pension": pension,
        "salud_mas_pension": salud + pension,
        "arl": arl,
        "total_pagado": total_pagado if total_pagado > 0 else (salud + pension + arl),
    }


def limpiar_y_clasificar_exogena_dian(archivo_bytes: bytes, anio_gravable: int) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    xls = pd.ExcelFile(io.BytesIO(archivo_bytes))
    hoja = xls.sheet_names[0]
    for h in xls.sheet_names:
        if any(p in h.lower() for p in ["report", "tercer", "detalle", "informac", "exogen"]):
            hoja = h
            break

    df_raw = pd.read_excel(xls, sheet_name=hoja, header=None)

    nit_meta, nombre_meta, anio_meta = "", "", None
    for idx in range(min(15, len(df_raw))):
        fila_vals = [str(x).strip() if pd.notna(x) else "" for x in df_raw.iloc[idx].values]
        etiqueta = fila_vals[0].lower() if fila_vals else ""
        valor_celda = fila_vals[2] if len(fila_vals) > 2 else ""
        if "identificacion:" in etiqueta.replace("ó", "o") and valor_celda:
            nit_meta = valor_celda.split(".")[0].strip()
        elif ("nombres" in etiqueta or "razon social" in etiqueta.replace("ó", "o")) and valor_celda:
            nombre_meta = valor_celda.strip()
        elif "año al que se refiere" in etiqueta and valor_celda.isdigit():
            anio_meta = int(valor_celda)

    df_sin_vacios = df_raw.dropna(how="all").reset_index(drop=True)

    fila_header = None
    claves_header = ["nit", "razon", "razón", "detalle", "valor", "uso declaracion", "uso declaración", "adicional"]
    mejor_puntaje = 0
    for idx in range(min(30, len(df_sin_vacios))):
        textos_fila = [str(x).strip().lower() for x in df_sin_vacios.iloc[idx].values if pd.notna(x)]
        puntaje = sum(1 for celda in textos_fila if any(k in celda for k in claves_header))
        if puntaje > mejor_puntaje and puntaje >= 3:
            mejor_puntaje = puntaje
            fila_header = idx

    if fila_header is not None:
        raw_cols = [
            str(c).strip() if pd.notna(c) and str(c).strip() else f"Columna_{i+1}"
            for i, c in enumerate(df_sin_vacios.iloc[fila_header].values)
        ]
        nuevas_cols, vistos, contador_nit = [], {}, 0
        for c in raw_cols:
            nom = re.sub(r"\s+", " ", c).strip()
            if nom.upper() == "NIT":
                contador_nit += 1
                nom = "NIT Reportante" if contador_nit == 1 else ("NIT Declarante" if contador_nit == 2 else f"NIT_{contador_nit}")
            elif nom.lower().startswith("nombre /"):
                nom = "Entidad / Persona que Reporta"
            elif nom.lower().startswith("nombre/raz"):
                nom = "Nombre Reportado"
            if nom in vistos:
                vistos[nom] += 1
                nom = f"{nom}_{vistos[nom]}"
            else:
                vistos[nom] = 1
            nuevas_cols.append(nom)
        df = df_sin_vacios.iloc[fila_header + 1 :].copy()
        df.columns = nuevas_cols
    else:
        df = df_sin_vacios.copy()
        df.columns = [f"Columna_{i+1}" for i in range(df.shape[1])]

    df = df.dropna(how="all").reset_index(drop=True)

    col_nit_rep = next((c for c in df.columns if "nit reportante" in c.lower() or c.lower() == "columna_1"), df.columns[0])
    col_detalle = next((c for c in df.columns if "detalle" in c.lower() or "concepto" in c.lower()), df.columns[min(4, len(df.columns) - 1)])
    col_valor = next((c for c in df.columns if any(p in c.lower() for p in ["valor", "monto", "cuantia", "cuantía"])), df.columns[min(5, len(df.columns) - 1)])
    col_uso = next((c for c in df.columns if "uso" in c.lower() or "sugerida" in c.lower()), None)

    mask_tope = df[col_nit_rep].isna() & df[col_detalle].astype(str).str.strip().str.lower().str.startswith("tope")
    df_topes_dian = df[mask_tope].copy()
    df_terceros = df[~mask_tope].copy()

    topes_oficiales = {}
    for _, r_t in df_topes_dian.iterrows():
        k_t = str(r_t[col_detalle]).strip().lower()
        v_t = _nvl(r_t[col_valor])
        if "tope 1" in k_t:
            topes_oficiales["tope_1_ingresos"] = v_t
        elif "tope 2" in k_t:
            topes_oficiales["tope_2_patrimonio"] = v_t
        elif "tope 3" in k_t:
            topes_oficiales["tope_3_tc"] = v_t
        elif "tope 4" in k_t:
            topes_oficiales["tope_4_movimientos"] = v_t
        elif "tope 5" in k_t:
            topes_oficiales["tope_5_compras"] = v_t

    resumen = {
        "nombre_detectado": nombre_meta,
        "nit_detectado": nit_meta,
        "anio_detectado": anio_meta,
        "efectivo_bancos": 0.0,
        "inversiones_cdt_saldo": 0.0,
        "inmuebles_vehiculos": 0.0,
        "deudas": 0.0,
        "consignaciones_tope4": 0.0,
        "tarjetas_credito": 0.0,
        "compras_totales": 0.0,
        "factura_electronica_1pct": 0.0,
        "ing_trabajo": 0.0,
        "ing_honorarios": 0.0,
        "ing_no_laborales": 0.0,
        "salud_pension": 0.0,
        "ing_capital": 0.0,
        "incrngo_capital_sugerido": 0.0,
        "retenciones": 0.0,
        "intereses_vivienda": 0.0,
        "medicina_prepagada": 0.0,
        "gmf_100": 0.0,
    }

    ing_por_nit = {}
    for _, row in df_terceros.iterrows():
        val = _nvl(row.get(col_valor, 0))
        if val == 0:
            continue
        if not resumen["nit_detectado"] and len(row) >= 4:
            pos_doc = str(row.iloc[2] or "").strip().split(".")[0]
            pos_nom = str(row.iloc[3] or "").strip()
            if pos_doc.isdigit() and len(pos_doc) >= 6:
                resumen["nit_detectado"] = pos_doc
            if len(pos_nom) >= 5 and "nan" not in pos_nom.lower():
                resumen["nombre_detectado"] = pos_nom

        nit_r = str(row.get(col_nit_rep, "") or "").strip().split(".")[0]
        det_n = str(row.get(col_detalle, "") or "").strip().lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
        uso_n = (str(row.get(col_uso, "") or "").strip().lower() if col_uso else "").replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")

        if "susceptible de beneficio" in det_n:
            resumen["factura_electronica_1pct"] += val
            continue
        if "tope 5" in uso_n or "facturas tras ajustes" in det_n:
            resumen["compras_totales"] += val
            continue
        if "tope 3" in uso_n or "consumos tarjeta" in det_n:
            resumen["tarjetas_credito"] += val
            continue
        if "tope 4" in uso_n or "movimientos en cuentas" in det_n or "inversion efectuada" in det_n:
            resumen["consignaciones_tope4"] += val
            continue
        if "tope 2" in uso_n or "r29" in uso_n or "saldo cuentas" in det_n or "saldo cdt" in det_n:
            if val < 0 or ("r30" in uso_n and "r29" not in uso_n) or "pasivo" in det_n or "deuda" in det_n:
                resumen["deudas"] += abs(val)
            elif any(k in det_n for k in ["cdt", "inversion", "accion", "fiducia", "fondo"]):
                resumen["inversiones_cdt_saldo"] += val
            elif any(k in det_n for k in ["inmueble", "predial", "avaluo", "vehiculo", "automotor"]):
                resumen["inmuebles_vehiculos"] += val
            else:
                resumen["efectivo_bancos"] += val
            continue
        if "r132" in uso_n or "retencion" in det_n or "valor retenido" in det_n:
            resumen["retenciones"] += val
            continue
        if "r58" in uso_n or "cdt rendimientos" in det_n or "rendimientos financieros" in det_n:
            resumen["ing_capital"] += val
            continue
        if "documentos soporte" in det_n or "facturas electronicas de venta" in det_n:
            ing_por_nit.setdefault(nit_r, {"trab": 0.0, "hon": 0.0, "nolab": 0.0, "doc_sop": 0.0})
            if det_n.startswith("ajuste") or "descuenta automaticamente" in uso_n:
                ing_por_nit[nit_r]["doc_sop"] -= val
            else:
                ing_por_nit[nit_r]["doc_sop"] += val
            continue
        if "r33" in uso_n or "r44" in uso_n or "salud obligatoria" in det_n or "pension obligatoria" in det_n:
            resumen["salud_pension"] += val
            continue
        if "r38" in uso_n or "intereses de vivienda" in det_n or "credito hipotecario" in det_n:
            resumen["intereses_vivienda"] += val
            continue
        if "r39" in uso_n or "prepagada" in det_n or "poliza de salud" in det_n:
            resumen["medicina_prepagada"] += val
            continue
        if "gmf" in det_n or "4x1000" in det_n:
            resumen["gmf_100"] += val
            continue
        if "r32" in uso_n or "salarios" in det_n or "pagos laborales" in det_n:
            ing_por_nit.setdefault(nit_r, {"trab": 0.0, "hon": 0.0, "nolab": 0.0, "doc_sop": 0.0})
            ing_por_nit[nit_r]["trab"] += val
        elif "r43" in uso_n or "honorarios" in det_n or "servicios" in det_n:
            ing_por_nit.setdefault(nit_r, {"trab": 0.0, "hon": 0.0, "nolab": 0.0, "doc_sop": 0.0})
            ing_por_nit[nit_r]["hon"] += val
        elif "r74" in uso_n or "tope 1" in uso_n:
            ing_por_nit.setdefault(nit_r, {"trab": 0.0, "hon": 0.0, "nolab": 0.0, "doc_sop": 0.0})
            ing_por_nit[nit_r]["nolab"] += val

    for _, d_nit in ing_por_nit.items():
        exo_sum = d_nit["trab"] + d_nit["hon"] + d_nit["nolab"]
        doc_neto = max(0.0, d_nit["doc_sop"])
        if exo_sum >= doc_neto:
            resumen["ing_trabajo"] += d_nit["trab"]
            resumen["ing_honorarios"] += d_nit["hon"]
            resumen["ing_no_laborales"] += d_nit["nolab"]
        else:
            dif = doc_neto - exo_sum
            resumen["ing_trabajo"] += d_nit["trab"]
            resumen["ing_honorarios"] += d_nit["hon"] + dif
            resumen["ing_no_laborales"] += d_nit["nolab"]

    if "tope_4_movimientos" in topes_oficiales and topes_oficiales["tope_4_movimientos"] > 0:
        resumen["consignaciones_tope4"] = topes_oficiales["tope_4_movimientos"]
    if "tope_5_compras" in topes_oficiales and topes_oficiales["tope_5_compras"] > 0:
        resumen["compras_totales"] = topes_oficiales["tope_5_compras"]
    if "tope_3_tc" in topes_oficiales:
        resumen["tarjetas_credito"] = topes_oficiales["tope_3_tc"]

    anio_ref = anio_meta or anio_gravable
    if resumen["ing_capital"] > 0:
        pct_inf = COMP_INFLACIONARIO_HISTORICO.get(anio_ref, 0.5544)
        resumen["incrngo_capital_sugerido"] = float(redondear_dian(resumen["ing_capital"] * pct_inf))

    df_vista = df_terceros.copy().reset_index(drop=True)
    for c in df_vista.columns:
        if c == col_valor:
            df_vista[c] = df_vista[c].apply(_nvl)
        else:
            df_vista[c] = df_vista[c].apply(lambda x: "" if pd.isna(x) else str(x).strip())

    return df_vista, resumen


def extraer_con_gemini_hibrido(api_key: str, archivo_bytes: bytes, mime_type: str) -> dict:
    contenidos = [types.Part.from_bytes(data=archivo_bytes, mime_type=mime_type)]
    if mime_type == "application/pdf" and OCR_DISPONIBLE:
        try:
            paginas = convert_from_bytes(archivo_bytes, first_page=1, last_page=3)
            texto_ocr = "\n".join(pytesseract.image_to_string(p, lang="spa") for p in paginas)
            if texto_ocr.strip():
                contenidos.append(f"Texto auxiliar OCR:\n{texto_ocr[:8000]}")
        except Exception:
            pass

    contenidos.append(
        "Eres un auditor tributario experto en Colombia (DIAN). Extrae en pesos (COP) los valores "
        "exactos de este soporte tributario (Certificado 220, CDT, Extracto, Predial, Prepagada, Hipotecario o PILA escaneada). "
        "Si un concepto no aparece, devuelve 0. Suma salud obligatoria + pensión obligatoria en 'salud_pension_obligatoria'."
    )
    cfg = types.GenerateContentConfig(
        response_mime_type="application/json",
        response_schema=CertificadoExtraido,
        temperature=0.0,
    )
    return json.loads(ejecutar_gemini_seguro(api_key, contenidos, cfg))


def auditar_exogena_ia(api_key: str, df_exo: pd.DataFrame) -> str:
    csv_muestra = df_exo.head(140).to_csv(index=False)
    prompt = (
        "Actúa como Auditor Tributario Senior DIAN Colombia. Revisa este reporte limpio de terceros y explica en lenguaje sencillo:\n"
        "1. Por qué la persona supera o no los topes para declarar.\n"
        "2. Qué movimientos bancarios o renovaciones de CDT NO son ingreso gravado.\n"
        "3. Qué soportes adicionales debe tener listos.\n\n"
        f"DATOS EXÓGENA:\n{csv_muestra}"
    )
    return ejecutar_gemini_seguro(api_key, prompt, types.GenerateContentConfig(temperature=0.1))


def consultar_estratega_rag(api_key: str, pregunta: str, res_liq: dict, uvt: float, anio: int) -> str:
    instruccion = (
        f"Eres un Contador Público Tributarista Senior en Colombia experto en Formulario 210 y Ley 2277 de 2022. "
        f"Año Gravable: {anio} | UVT aplicada: ${uvt:,.0f} COP. "
        f"Liquidación oficial calculada por el motor determinista: {json.dumps(res_liq, ensure_ascii=False)}. "
        "Responde de forma clara, amable y citando casillas del Formulario 210 y artículos del Estatuto Tributario."
    )
    return ejecutar_gemini_seguro(api_key, pregunta, types.GenerateContentConfig(system_instruction=instruccion, temperature=0.2))


# ==============================================================================
# 5. EXPORTADOR PROFESIONAL ESTILIZADO (.XLSX MULTI-HOJA + .DOCX DICTAMEN)
# ==============================================================================
def construir_filas_210(res: dict) -> List[Tuple[str, str, int]]:
    return [
        ("28", "Uno por ciento (1%) de compras con factura electrónica", res["c28_factura_electronica"]),
        ("29", "Total patrimonio bruto", res["c29_patrimonio_bruto"]),
        ("30", "Deudas", res["c30_deudas"]),
        ("31", "Total patrimonio líquido (29 - 30)", res["c31_patrimonio_liquido"]),
        ("32", "Ingresos brutos Rentas de Trabajo (Salarios / Servicios con 25% exento)", res["c32_ing_trabajo"]),
        ("33", "Ingresos no constitutivos de renta Trabajo (Salud + Pensión)", res["c33_inc_trabajo"]),
        ("34", "Renta líquida Trabajo (32 - 33)", res["c34_renta_liq_trabajo"]),
        ("35", "Aportes voluntarios AFC, FVP y/o AVC", res["c35_afc_fpv"]),
        ("36", "Otras rentas exentas Trabajo (25% laboral + Cesantías)", res["c36_exenta_25"]),
        ("38", "Intereses de vivienda / leasing habitacional", res["c38_vivienda"]),
        ("39", "Otras deducciones imputables (Prepagada, 10% dep., ICETEX, 50% GMF)", res["c39_prepagada"]),
        ("43", "Ingresos brutos Honorarios sujetos a costos y gastos", res["c43_ing_honorarios"]),
        ("44", "Ingresos no constitutivos de renta Honorarios (Salud + Pensión)", res["c44_inc_honorarios"]),
        ("45", "Costos y deducciones procedentes Honorarios (Incluye ARL)", res["c45_costos_honorarios"]),
        ("46", "Renta líquida Honorarios (43 - 44 - 45)", res["c46_renta_liq_honorarios"]),
        ("58", "Ingresos brutos Rentas de Capital (CDT / Intereses / Arriendos)", res["c58_ing_capital"]),
        ("59", "Ingresos no constitutivos Capital (Componente Inflacionario CDT)", res["c59_inc_capital"]),
        ("60", "Costos y deducciones procedentes Rentas de Capital", res["c60_costos_capital"]),
        ("61", "Renta líquida Capital (58 - 59 - 60)", res["c61_renta_liq_capital"]),
        ("74", "Ingresos brutos Rentas No Laborales (Comercio / Ventas)", res["c74_ing_nolab"]),
        ("76", "Ingresos no constitutivos Rentas No Laborales", res["c76_inc_nolab"]),
        ("77", "Costos y gastos procedentes Rentas No Laborales", res["c77_costos_nolab"]),
        ("78", "Renta líquida No Laborales", res["c78_renta_liq_nolab"]),
        ("91", "Renta líquida Cédula General (34 + 46 + 61 + 78)", res["c91_renta_liq_cg"]),
        ("92", "Rentas exentas y deducciones imputables limitadas", res["c92_exentas_ded_limitadas"]),
        ("97", "Renta líquida gravable Cédula General (91 - 92)", res["c97_rlg_cedula_general"]),
        ("99", "Ingresos brutos por rentas de pensiones", res["c99_pen_brutos"]),
        ("100", "Ingresos no constitutivos de renta pensiones (Salud)", res["c100_pen_salud"]),
        ("102", "Rentas exentas de pensiones", res["c102_pen_exenta"]),
        ("103", "Renta líquida gravable Cédula de Pensiones", res["c103_rlg_pensiones"]),
        ("107", "Dividendos y participaciones gravables", res["c107_dividendos"]),
        ("111", "TOTAL RENTA LÍQUIDA GRAVABLE (Base Art. 241 E.T.)", res["c111_rlg_total"]),
        ("112", "Ingresos por ganancias ocasionales", res["c112_go_ingresos"]),
        ("113", "Costos por ganancias ocasionales", res["c113_go_costos"]),
        ("114", "Ganancias ocasionales no gravadas y exentas", res["c114_go_exentas"]),
        ("115", "Ganancias ocasionales gravables", res["c115_go_gravable"]),
        ("116", "Impuesto sobre la renta líquida gravable", res["c116_impuesto_rlg"]),
        ("125", "Total descuentos tributarios (Donaciones / Dividendos)", res["c125_descuentos_trib"]),
        ("126", "Impuesto neto de renta (116 - 125)", res["c126_impuesto_neto_renta"]),
        ("127", "Impuesto de ganancias ocasionales", res["c127_impuesto_go"]),
        ("129", "Total impuesto a cargo (126 + 127)", res["c129_total_impuesto_cargo"]),
        ("130", "Anticipo renta liquidado año gravable anterior", res["c130_anticipo_anterior"]),
        ("131", "Saldo a favor del año gravable anterior sin solicitud", res["c131_saldo_favor_anterior"]),
        ("132", "Retenciones año gravable a declarar", res["c132_retenciones_anio"]),
        ("133", "Anticipo renta para el año gravable siguiente (Art. 807 E.T.)", res["c133_anticipo_siguiente"]),
        ("134", "Saldo a pagar por impuesto", res["c134_saldo_pagar_impuesto"]),
        ("135", "Sanciones (Extemporaneidad Art. 641 / 640 E.T.)", res["c135_sanciones"]),
        ("136", "TOTAL SALDO A PAGAR", res["c136_total_saldo_pagar"]),
        ("137", "TOTAL SALDO A FAVOR", res["c137_total_saldo_favor"]),
        ("138", "Número de dependientes económicos", res["c138_num_dependientes"]),
        ("139", "Adición por dependientes a la casilla 92 (72 UVT c/u)", res["c139_adicion_dependientes"]),
    ]


def exportar_excel_dian(res: dict, d: dict, anio: int, df_pila: pd.DataFrame = None) -> bytes:
    df_210 = pd.DataFrame(
        construir_filas_210(res),
        columns=["Casilla 210", "Concepto Oficial Formulario 210 DIAN", "Valor Oficial (COP)"],
    )

    dv_calc = calcular_dv_dian(d.get("nit_declarante", ""))
    nit_fmt = f"{d.get('nit_declarante', '')}-{dv_calc}" if d.get("nit_declarante") else "Sin especificar"
    balance_neto = res["c136_total_saldo_pagar"] if res["c136_total_saldo_pagar"] > 0 else -res["c137_total_saldo_favor"]

    filas_dash = [
        ("Año Gravable Declarado (Casilla 1)", anio),
        ("Declarante (Casillas 7 a 10)", d.get("nombre_declarante") or "Contribuyente Persona Natural"),
        ("NIT / Cédula + DV (Casillas 5 y 6)", nit_fmt),
        ("Dirección Seccional DIAN (Casilla 12)", d.get("seccional_dian", "32 - Bogotá D.C.")),
        ("Actividad Económica CIIU (Casilla 24)", d.get("cod_actividad") or "8211"),
        ("Estado Obligación DIAN", res["estado_obligacion"]),
        ("Estrategia Tributaria Aplicada", res["estrategia"]),
        ("Patrimonio Bruto (Casilla 29)", res["c29_patrimonio_bruto"]),
        ("Deudas (Casilla 30)", res["c30_deudas"]),
        ("Patrimonio Líquido (Casilla 31)", res["c31_patrimonio_liquido"]),
        ("Ingresos Trabajo + Honorarios (Casillas 32 + 43)", res["c32_ing_trabajo"] + res["c43_ing_honorarios"]),
        ("Ingresos de Capital CDT / Arriendos (Casilla 58)", res["c58_ing_capital"]),
        ("Exenciones y Deducciones Aplicadas (Casilla 92)", res["c92_exentas_ded_limitadas"]),
        ("Renta Líquida Gravable Total (Casilla 111)", res["c111_rlg_total"]),
        ("Total Impuesto a Cargo (Casilla 129)", res["c129_total_impuesto_cargo"]),
        ("Retenciones + Saldos a Favor (Casillas 130 a 132)", res["alivios_totales"]),
        ("Balance Consolidado (+ Pagar / - Saldo a Favor)", balance_neto),
    ]
    df_dash = pd.DataFrame(filas_dash, columns=["Indicador Resumen Ejecutivo", "Valor / Resultado"])

    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    zebra_fill = PatternFill(start_color="F4F7FA", end_color="F4F7FA", fill_type="solid")
    highlight_fill = PatternFill(start_color="E8F4EA", end_color="E8F4EA", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="D9E1E8"),
        right=Side(style="thin", color="D9E1E8"),
        top=Side(style="thin", color="D9E1E8"),
        bottom=Side(style="thin", color="D9E1E8"),
    )

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        nombre_hoja_210 = f"Form_210_AG{anio}"
        df_210.to_excel(writer, index=False, sheet_name=nombre_hoja_210)
        df_dash.to_excel(writer, index=False, sheet_name="Resumen_Ejecutivo")
        if df_pila is not None and not df_pila.empty:
            df_pila.to_excel(writer, index=False, sheet_name="Detalle_PILA")

        ws1 = writer.sheets[nombre_hoja_210]
        ws1.column_dimensions["A"].width = 14
        ws1.column_dimensions["B"].width = 68
        ws1.column_dimensions["C"].width = 24
        ws1.freeze_panes = "A2"
        for cell in ws1[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for r_idx, row in enumerate(ws1.iter_rows(min_row=2, max_row=ws1.max_row), start=2):
            cas_num = str(row[0].value)
            es_clave = cas_num in ("29", "31", "97", "111", "129", "132", "136", "137")
            for c_idx, cell in enumerate(row, start=1):
                cell.border = thin_border
                if es_clave:
                    cell.fill = highlight_fill
                    cell.font = Font(bold=True)
                elif r_idx % 2 == 1:
                    cell.fill = zebra_fill
                if c_idx == 1:
                    cell.alignment = Alignment(horizontal="center")
                elif c_idx == 3:
                    cell.number_format = "#,##0" if cas_num == "138" else "$#,##0"

        ws2 = writer.sheets["Resumen_Ejecutivo"]
        ws2.column_dimensions["A"].width = 52
        ws2.column_dimensions["B"].width = 36
        ws2.freeze_panes = "A2"
        for cell in ws2[1]:
            cell.fill = header_fill
            cell.font = header_font
        for r_idx, row in enumerate(ws2.iter_rows(min_row=2, max_row=ws2.max_row), start=2):
            for c_idx, cell in enumerate(row, start=1):
                cell.border = thin_border
                if r_idx % 2 == 1:
                    cell.fill = zebra_fill
                if c_idx == 2 and isinstance(cell.value, (int, float)) and r_idx > 2:
                    cell.number_format = "$#,##0"

    return buf.getvalue()


def exportar_dictamen_word(res: dict, d: dict, anio: int, uvt: float) -> bytes:
    doc = Document()
    doc.add_heading(f"Dictamen Técnico y Borrador Formulario 210 — AG {anio}", level=1)
    titular = d.get("nombre_declarante") or "Contribuyente"
    nit = d.get("nit_declarante") or "N/A"
    dv = calcular_dv_dian(nit)
    doc.add_paragraph(
        f"Declarante: {titular} | NIT/CC: {nit}{f'-{dv}' if dv else ''} | Seccional: {d.get('seccional_dian', '32')} | CIIU: {d.get('cod_actividad', '8211')}\n"
        f"Año Gravable: {anio} (UVT: ${uvt:,.0f} COP) | Estrategia Óptima: {res['estrategia']}"
    )

    doc.add_heading("1. Diagnóstico de Obligación Tributaria (6 Criterios DIAN)", level=2)
    doc.add_paragraph(f"Estado: {res['estado_obligacion']}")
    for c in res["causales_obligacion"]:
        doc.add_paragraph(c, style="List Bullet")

    doc.add_heading("2. Casillas Oficiales para Transcribir al Portal Muisca DIAN", level=2)
    for cas, nom, val in construir_filas_210(res):
        if val > 0 or cas in ("29", "31", "97", "111", "129", "132", "136", "137"):
            pref = "" if cas == "138" else "$"
            doc.add_paragraph(f"Casilla {cas} — {nom}: {pref}{val:,.0f}", style="List Bullet")

    if res["alertas_auditoria"] or res["sugerencias_optimizacion"]:
        doc.add_heading("3. Auditoría Preventiva y Recomendaciones", level=2)
        for al in res["alertas_auditoria"]:
            doc.add_paragraph(f"[ALERTA] {al}", style="List Bullet")
        for sg in res["sugerencias_optimizacion"]:
            doc.add_paragraph(f"[AVISO] {sg}", style="List Bullet")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ==============================================================================
# 6. INTERFAZ GUIADA ESTILO TRIBUTI + PROGRAMA AYUDA RENTA DIAN
# ==============================================================================
if not verificar_acceso():
    st.stop()

if "datos" not in st.session_state:
    st.session_state.datos = estado_inicial_datos()
if "selector_anio_gravable" not in st.session_state:
    st.session_state["selector_anio_gravable"] = datetime.date.today().year - 1
if "form_version" not in st.session_state:
    st.session_state.form_version = 0
if "lote_certificados" not in st.session_state:
    st.session_state.lote_certificados = []
if "df_pila_cargada" not in st.session_state:
    st.session_state.df_pila_cargada = None
if "auditoria_exogena" not in st.session_state:
    st.session_state.auditoria_exogena = None
if "chat_mensajes" not in st.session_state:
    st.session_state.chat_mensajes = []

v = st.session_state.form_version
d = st.session_state.datos

# --- BARRA LATERAL ---
with st.sidebar:
    st.title("⚙️ Configuración DIAN")
    api_secret = obtener_secreto("GEMINI_API_KEY")
    if api_secret:
        api_key = api_secret
        st.success("🔒 Gemini API Key activa en Secrets")
    else:
        raw_input_key = st.text_input(
            "🔑 Gemini API Key (Solo para lector IA / Chat)",
            type="password",
            help="Exógena y Planillas PILA funcionan sin API Key. Ingresa tu clave AIza... para certificados escaneados o Consultor IA.",
        )
        api_key = limpiar_api_key(raw_input_key)

    anio_gravable = int(
        st.number_input(
            "📅 Año Gravable a Declarar",
            min_value=2020,
            max_value=2100,
            step=1,
            key="selector_anio_gravable",
            help="Cambia el año aquí (ej. 2025, 2026, 2027...) y el título, la UVT y todos los cálculos se actualizan al instante.",
        )
    )

    valor_uvt = st.number_input(
        f"Valor UVT AG {anio_gravable} ($ COP)",
        min_value=10_000,
        value=int(uvt_sugerida_por_anio(anio_gravable)),
        step=100,
        key=f"uvt_{anio_gravable}_{v}",
        help="Para años futuros nuevos, escribe el valor oficial de la UVT que publique la DIAN.",
    )

    with st.expander("Ajustar Topes UVT y Sanción"):
        uvt_pat = st.number_input("Tope Patrimonio (UVT)", value=4_500, step=100)
        uvt_ing = st.number_input("Tope Ingresos / Consumos (UVT)", value=1_400, step=100)
        uvt_sancion = st.number_input(
            f"UVT Año Presentación {anio_gravable + 1} (Sanción Mín.)",
            value=uvt_sugerida_por_anio(anio_gravable + 1),
            step=100,
            key=f"uvtsan_{anio_gravable}_{v}",
        )

    d["anio_declara_num"] = st.selectbox(
        "Historial declarando (Anticipo Casilla 133)",
        options=[0, 1, 2, 3],
        index=int(d["anio_declara_num"]),
        format_func=lambda x: {
            0: "Sin Anticipo ($0 si impuesto es $0)",
            1: "1ra vez que declara (25%)",
            2: "2do año declarando (50%)",
            3: "3er año o más (75% Optimizado)",
        }[x],
    )

    with st.expander("💾 Guardar / Abrir Cliente (.json)"):
        respaldo_bytes = json.dumps(
            {"anio_gravable": anio_gravable, "datos": d}, ensure_ascii=False, indent=2
        ).encode("utf-8")
        st.download_button(
            "⬇️ Guardar Borrador en mi PC",
            data=respaldo_bytes,
            file_name=f"Renta_{d.get('nit_declarante') or 'Cliente'}_AG{anio_gravable}.json",
            mime="application/json",
            use_container_width=True,
        )
        up_json = st.file_uploader("Abrir borrador .json guardado", type=["json"], key=f"up_json_{v}")
        if up_json and st.button("📂 Restaurar Declaración", use_container_width=True):
            cargado = json.loads(up_json.getvalue().decode("utf-8"))
            if "anio_gravable" in cargado:
                st.session_state["selector_anio_gravable"] = int(cargado["anio_gravable"])
            if "datos" in cargado:
                base = estado_inicial_datos()
                base.update(cargado["datos"])
                st.session_state.datos = base
            st.session_state.form_version += 1
            st.rerun()

    clave_m = obtener_secreto("CLAVE_MAESTRA")
    if clave_m and st.session_state.get("es_admin", False):
        with st.expander("🛡️ Generador de Licencias (Admin)"):
            id_cli = st.text_input("Cédula o Correo del Cliente")
            if id_cli:
                st.code(generar_licencia_cliente(id_cli, clave_m))

    st.divider()
    with st.popover("🧹 Nueva Declaración / Limpiar Todo", use_container_width=True):
        st.warning("¿Confirmas borrar todos los datos actuales para iniciar una nueva declaración desde cero?")
        if st.button("✅ Sí, limpiar todo", type="primary", use_container_width=True):
            st.session_state.datos = estado_inicial_datos()
            st.session_state.lote_certificados = []
            st.session_state.df_pila_cargada = None
            st.session_state.auditoria_exogena = None
            st.session_state.chat_mensajes = []
            st.session_state.form_version += 1
            st.rerun()

res = liquidar_declaracion_210(d, valor_uvt, uvt_pat, uvt_ing, uvt_sancion)

# --- ENCABEZADO DINÁMICO MULTIAÑO Y TABLERO KPI ---
st.title(f"📊 RentaIA Pro — Liquidador & Auditor Formulario 210 (AG {anio_gravable})")
st.caption(
    f"Declaración Año Gravable **{anio_gravable}** (Presentación en **{anio_gravable + 1}**) • "
    f"UVT Aplicada: **${valor_uvt:,.0f} COP** • Privacidad total en memoria RAM"
)

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric(
    "Obligación DIAN",
    "SÍ OBLIGADO" if res["obligado_bool"] else "NO OBLIGADO",
    delta=f"{len(res['causales_obligacion'])} causal(es)" if res["obligado_bool"] else "Declaración Voluntaria",
)
m2.metric("Patrimonio Líquido (C.31)", f"${res['c31_patrimonio_liquido']:,.0f}")
m3.metric("Renta Líquida Grav. (C.111)", f"${res['c111_rlg_total']:,.0f}", f"{res['rlg_en_uvt']} UVT")
m4.metric("Impuesto a Cargo (C.129)", f"${res['c129_total_impuesto_cargo']:,.0f}")
if res["c137_total_saldo_favor"] > 0:
    m5.metric("🎉 SALDO A FAVOR (C.137)", f"${res['c137_total_saldo_favor']:,.0f}", delta="Devolución / Compensación")
else:
    m5.metric("💳 SALDO A PAGAR (C.136)", f"${res['c136_total_saldo_pagar']:,.0f}")

# --- 5 PASOS GUIADOS ESTILO TRIBUTI + AYUDA RENTA DIAN ---
t1, t2, t3, t4, t5 = st.tabs([
    "🚀 Paso 1. Carga Inteligente (Exógena y PILA)",
    "👤 Paso 2. Perfil, Patrimonio y Deudas",
    "💼 Paso 3. Ingresos (Todas las Cédulas)",
    "🛡️ Paso 4. Beneficios, Ahorro y Retenciones",
    "📋 Paso 5. Formulario 210, Descargas y Chat IA",
])

# ==============================================================================
# PASO 1: CARGA AUTOMÁTICA (EXÓGENA + MULTI-PDF PILA + CERTIFICADOS IA)
# ==============================================================================
with t1:
    col_izq, col_der = st.columns(2)

    with col_izq:
        st.subheader("📊 A. Reporte de Información Exógena DIAN")
        up_exo = st.file_uploader(
            "Sube el archivo Excel (.xlsx) descargado del portal DIAN (No requiere IA)",
            type=["xlsx", "xls"],
            key=f"up_exo_{v}",
        )
        if up_exo:
            try:
                df_exo, res_exo = limpiar_y_clasificar_exogena_dian(up_exo.getvalue(), anio_gravable)
                st.caption(
                    f"Terceros procesados: **{len(df_exo)}** | "
                    f"Titular: **{res_exo['nombre_detectado'] or 'Detectado'}** (ID: **{res_exo['nit_detectado'] or 'N/A'}**)"
                )
                st.dataframe(df_exo, use_container_width=True, height=210)

                df_res_exo = pd.DataFrame([
                    ("Saldos Cuentas Bancarias a dic 31", f"${res_exo['efectivo_bancos']:,.0f}"),
                    ("Saldos CDT / Inversiones a dic 31", f"${res_exo['inversiones_cdt_saldo']:,.0f}"),
                    ("Movimientos y Consignaciones (Tope 4)", f"${res_exo['consignaciones_tope4']:,.0f}"),
                    ("Compras Totales (Tope 5)", f"${res_exo['compras_totales']:,.0f}"),
                    ("Base Factura Electrónica susceptible 1% (C.28)", f"${res_exo['factura_electronica_1pct']:,.0f}"),
                    ("Ingresos Salarios + Servicios/Honorarios", f"${res_exo['ing_trabajo'] + res_exo['ing_honorarios']:,.0f}"),
                    ("Ingresos Capital — Rendimientos CDT (C.58)", f"${res_exo['ing_capital']:,.0f}"),
                    ("Comp. Inflacionario CDT sugerido (C.59)", f"${res_exo['incrngo_capital_sugerido']:,.0f}"),
                    ("Retenciones Practicadas (C.132)", f"${res_exo['retenciones']:,.0f}"),
                ], columns=["Concepto Tributario DIAN", "Valor Detectado"])
                st.table(df_res_exo)

                if st.button("⚡ Autollenar Declaración con esta Exógena", type="primary", use_container_width=True):
                    if res_exo["anio_detectado"]:
                        st.session_state["selector_anio_gravable"] = int(res_exo["anio_detectado"])
                    if res_exo["nombre_detectado"]:
                        d["nombre_declarante"] = res_exo["nombre_detectado"]
                    if res_exo["nit_detectado"]:
                        d["nit_declarante"] = res_exo["nit_detectado"]
                    d["efectivo"] = res_exo["efectivo_bancos"]
                    d["otros_activos"] = res_exo["inversiones_cdt_saldo"]
                    if res_exo["inmuebles_vehiculos"] > 0:
                        d["inmuebles"] = res_exo["inmuebles_vehiculos"]
                    d["deudas"] = res_exo["deudas"]
                    d["consignaciones"] = res_exo["consignaciones_tope4"]
                    d["tarjetas_credito"] = res_exo["tarjetas_credito"]
                    d["compras_totales"] = res_exo["compras_totales"]
                    d["compras_factura_electronica"] = res_exo["factura_electronica_1pct"]
                    d["ing_brutos_trabajo"] = res_exo["ing_trabajo"]
                    d["ing_brutos_honorarios"] = res_exo["ing_honorarios"]
                    d["ing_brutos_no_laborales"] = res_exo["ing_no_laborales"]
                    if res_exo["salud_pension"] > 0:
                        d["salud_pension_trabajo"] = res_exo["salud_pension"]
                    d["ing_brutos_capital"] = res_exo["ing_capital"]
                    d["incrngo_capital"] = res_exo["incrngo_capital_sugerido"]
                    d["retenciones_anio"] = res_exo["retenciones"]
                    if res_exo["intereses_vivienda"] > 0:
                        d["intereses_vivienda"] = res_exo["intereses_vivienda"]
                    if res_exo["medicina_prepagada"] > 0:
                        d["medicina_prepagada"] = res_exo["medicina_prepagada"]
                    if res_exo["gmf_100"] > 0:
                        d["gmf_certificado_100"] = res_exo["gmf_100"]
                    st.session_state.form_version += 1
                    st.rerun()

                if st.button("🤖 Explicar mi Exógena en Lenguaje Sencillo (IA)", use_container_width=True):
                    if not api_key:
                        st.error("Configura tu Gemini API Key en Secrets o en la barra lateral.")
                    else:
                        try:
                            with st.spinner("Analizando reporte de terceros..."):
                                st.session_state.auditoria_exogena = auditar_exogena_ia(api_key, df_exo)
                        except Exception as e:
                            st.error(f"⚠️ {e}")
            except Exception as e:
                st.error(f"Error al procesar Exógena: {e}")

        if st.session_state.auditoria_exogena:
            st.success(st.session_state.auditoria_exogena)

    with col_der:
        st.subheader("🏥 B. Lector Masivo de Planillas PILA (Multi-PDF sin IA)")
        archivos_pila = st.file_uploader(
            "Arrastra TODAS tus planillas PILA PDF juntas (MiPlanilla, Compensar, SOI, Aportes en Línea, Asopagos)",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"up_pila_{v}",
        )
        if archivos_pila:
            registros_pila = []
            planillas_vistas = set()
            no_reconocidos = []
            for arc in archivos_pila:
                info_p = extraer_pila_pdf_nativo(arc.name, arc.getvalue())
                if info_p.get("es_pila"):
                    id_pl = info_p["planilla"]
                    if id_pl not in planillas_vistas:
                        planillas_vistas.add(id_pl)
                        registros_pila.append({
                            "Planilla": id_pl,
                            "Periodo": info_p["periodo"],
                            "Salud ($)": float(info_p["salud"]),
                            "Pensión ($)": float(info_p["pension"]),
                            "ARL ($)": float(info_p["arl"]),
                        })
                else:
                    no_reconocidos.append(arc.name)

            if registros_pila:
                df_pila_base = pd.DataFrame(registros_pila).sort_values(by="Periodo").reset_index(drop=True)
                st.caption("Puedes verificar o ajustar cualquier celda directamente en la tabla antes de pasar el total:")
                df_pila_edit = st.data_editor(
                    df_pila_base,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="dynamic",
                    key=f"ed_pila_{v}",
                )
                df_pila_edit["Salud + Pensión ($)"] = df_pila_edit["Salud ($)"].fillna(0) + df_pila_edit["Pensión ($)"].fillna(0)
                st.session_state.df_pila_cargada = df_pila_edit

                tot_sp = float(df_pila_edit["Salud + Pensión ($)"].sum())
                tot_arl = float(df_pila_edit["ARL ($)"].sum())

                st.info(
                    f"✅ **{len(df_pila_edit)} planilla(s) sumadas sin errores:**\n"
                    f"- **Salud + Pensión Obligatoria (No gravado C.33 / C.44):** `${tot_sp:,.0f}`\n"
                    f"- **Riesgos Laborales ARL (Costo Honorarios C.45):** `${tot_arl:,.0f}`"
                )
                if st.button("⚡ Pasar Total PILA a mi Declaración", type="primary", use_container_width=True):
                    d["salud_pension_honorarios"] = tot_sp
                    d["arl_honorarios"] = tot_arl
                    st.session_state.form_version += 1
                    st.rerun()

            if no_reconocidos:
                st.warning(
                    "Estos PDFs parecen escaneados como imagen o son certificados distintos a PILA (súbelos abajo en el Lector IA): "
                    + ", ".join(no_reconocidos)
                )

        st.divider()
        st.subheader("📄 C. Lector IA de Otros Certificados (PDF / Foto — Múltiples)")
        archivos_cert = st.file_uploader(
            "Sube uno o varios soportes a la vez (Formulario 220, CDT, Extractos, Predial, Prepagada, Hipotecario)",
            type=["pdf", "png", "jpg", "jpeg"],
            accept_multiple_files=True,
            key=f"up_cert_{v}",
        )
        if archivos_cert and st.button(f"🔍 Leer {len(archivos_cert)} soporte(s) con IA Gemini", use_container_width=True):
            if not api_key:
                st.error("Configura primero tu Gemini API Key en Secrets o en la barra lateral.")
            else:
                extraidos = []
                for arc in archivos_cert:
                    try:
                        with st.spinner(f"Leyendo {arc.name}..."):
                            item = extraer_con_gemini_hibrido(api_key, arc.getvalue(), arc.type)
                            item["archivo"] = arc.name
                            extraidos.append(item)
                    except Exception as e:
                        st.error(f"Error en {arc.name}: {e}")
                st.session_state.lote_certificados = extraidos

        if st.session_state.lote_certificados:
            st.dataframe(pd.DataFrame(st.session_state.lote_certificados), use_container_width=True)
            if st.button("➕ Sumar estos certificados a mi Declaración", type="primary", use_container_width=True):
                for ext in st.session_state.lote_certificados:
                    d["efectivo"] += ext.get("efectivo_bancos_dic31", 0.0)
                    d["otros_activos"] += ext.get("valor_inmueble_o_cdt_dic31", 0.0)
                    d["deudas"] += ext.get("saldo_deuda_dic31", 0.0)
                    d["ing_brutos_trabajo"] += ext.get("ingresos_brutos_trabajo", 0.0)
                    d["cesantias_exentas"] += ext.get("cesantias_exentas", 0.0)
                    d["ing_brutos_honorarios"] += ext.get("ingresos_brutos_honorarios", 0.0)
                    d["salud_pension_trabajo"] += ext.get("salud_pension_obligatoria", 0.0)
                    d["ing_brutos_capital"] += ext.get("ingresos_capital_cdt", 0.0)
                    d["incrngo_capital"] += ext.get("componente_inflacionario_cdt", 0.0)
                    d["intereses_vivienda"] += ext.get("intereses_vivienda", 0.0)
                    d["medicina_prepagada"] += ext.get("medicina_prepagada", 0.0)
                    d["gmf_certificado_100"] += ext.get("gmf_4x1000_certificado", 0.0)
                    d["aportes_afc_fpv"] += ext.get("aportes_afc_fpv", 0.0)
                    d["retenciones_anio"] += ext.get("retenciones_practicadas", 0.0)
                st.session_state.lote_certificados = []
                st.session_state.form_version += 1
                st.rerun()

# ==============================================================================
# PASO 2: PERFIL DIAN, PATRIMONIO Y DEUDAS (COMO TRIBUTI + AYUDA RENTA DIAN)
# ==============================================================================
with t2:
    st.subheader("1. Datos de Identificación del Declarante (Encabezado Formulario 210)")
    id1, id2, id3, id4 = st.columns([2.2, 1.3, 1.3, 1.8])
    d["nombre_declarante"] = id1.text_input("Nombre Completo (Casillas 7 a 10)", value=d["nombre_declarante"], key=f"nom_{v}")
    d["nit_declarante"] = id2.text_input("Cédula / NIT sin DV (Casilla 5)", value=d["nit_declarante"], key=f"nit_{v}")
    dv_auto = calcular_dv_dian(d["nit_declarante"])
    id2.caption(f"Dígito de Verificación (Casilla 6): **{dv_auto or 'Automático'}**")

    idx_sec = SECCIONALES_DIAN.index(d["seccional_dian"]) if d.get("seccional_dian") in SECCIONALES_DIAN else 0
    d["seccional_dian"] = id3.selectbox("Dirección Seccional DIAN (Casilla 12)", options=SECCIONALES_DIAN, index=idx_sec, key=f"sec_{v}")

    opciones_cod = [op.split(" - ")[0] for op in ACTIVIDADES_CIIU_COMUNES]
    idx_ciiu = opciones_cod.index(d["cod_actividad"]) if d.get("cod_actividad") in opciones_cod else len(ACTIVIDADES_CIIU_COMUNES) - 1
    sel_ciiu = id4.selectbox("Actividad Económica Principal (Casilla 24)", options=ACTIVIDADES_CIIU_COMUNES, index=idx_ciiu, key=f"selciiu_{v}")
    if sel_ciiu.startswith("OTRO"):
        d["cod_actividad"] = id4.text_input("Escribe tu código CIIU de 4 dígitos (RUT)", value=d["cod_actividad"], key=f"ciiu_man_{v}")
    else:
        d["cod_actividad"] = sel_ciiu.split(" - ")[0]

    info_venc = orientacion_vencimiento_dian(d["nit_declarante"], anio_gravable + 1)
    if info_venc:
        st.caption(f"🗓️ {info_venc} — Verifica el día exacto en el calendario tributario oficial de la DIAN.")

    st.divider()
    p1, p2 = st.columns(2)
    with p1:
        st.subheader(f"2. Tus Bienes y Deudas a 31 de Diciembre de {anio_gravable}")
        d["efectivo"] = st.number_input(
            "🏦 Saldos en cuentas bancarias (Ahorros, Corriente, Nequi, Daviplata) y efectivo",
            min_value=0.0, value=float(d["efectivo"]), step=100_000.0, key=f"ef_{v}",
            help="Toma el saldo final a 31 de diciembre que aparece en tus certificados tributarios bancarios.",
        )
        d["otros_activos"] = st.number_input(
            "📈 Inversiones vigentes a dic 31 (CDTs abiertos, fiducias, fondos, acciones, cuentas por cobrar)",
            min_value=0.0, value=float(d["otros_activos"]), step=200_000.0, key=f"ot_{v}",
            help="Solo incluye los CDTs o inversiones que seguían vigentes el 31 de diciembre (no sumes renovaciones ya vencidas).",
        )

        with st.expander("🏡 Asistente DIAN para Inmuebles: Comparar Compra vs. Predial (Art. 277 E.T.)"):
            st.caption("La DIAN exige declarar cada inmueble por el **mayor valor** entre el precio de compra (escritura) y el avalúo catastral del recibo Predial.")
            ca1, ca2 = st.columns(2)
            costo_esc = ca1.number_input("Precio de compra / Costo fiscal ($)", min_value=0.0, step=1_000_000.0, key=f"esc_{v}")
            avaluo_pred = ca2.number_input("Avalúo Catastral recibo Predial ($)", min_value=0.0, step=1_000_000.0, key=f"pred_{v}")
            mayor_inm = max(costo_esc, avaluo_pred)
            if mayor_inm > 0 and st.button(f"➕ Sumar mayor valor oficial (${mayor_inm:,.0f}) a Bienes Raíces", use_container_width=True):
                d["inmuebles"] += float(mayor_inm)
                st.session_state.form_version += 1
                st.rerun()

        d["inmuebles"] = st.number_input(
            "🏠 Total Bienes Raíces: Casas, apartamentos, lotes, fincas (Mayor entre compra y predial)",
            min_value=0.0, value=float(d["inmuebles"]), step=1_000_000.0, key=f"inm_{v}",
        )
        d["vehiculos"] = st.number_input(
            "🚗 Vehículos, motocicletas y otros bienes (Precio de compra / factura)",
            min_value=0.0, value=float(d["vehiculos"]), step=500_000.0, key=f"ve_{v}",
        )
        d["deudas"] = st.number_input(
            "💳 Total Deudas a 31 de diciembre (Hipotecas, créditos bancarios, tarjetas — Casilla 30)",
            min_value=0.0, value=float(d["deudas"]), step=200_000.0, key=f"deu_{v}",
        )

    with p2:
        st.subheader("3. Semáforo de los 6 Topes DIAN y Control Patrimonial")
        d["consignaciones"] = st.number_input("Movimientos / Consignaciones bancarias del año (Tope 4)", min_value=0.0, value=float(d["consignaciones"]), step=500_000.0, key=f"con_{v}")
        d["tarjetas_credito"] = st.number_input("Consumos con Tarjeta de Crédito en el año (Tope 3)", min_value=0.0, value=float(d["tarjetas_credito"]), step=500_000.0, key=f"tc_{v}")
        d["compras_totales"] = st.number_input("Compras y consumos totales del año (Tope 5)", min_value=0.0, value=float(d["compras_totales"]), step=500_000.0, key=f"com_{v}")
        d["responsable_iva"] = st.toggle("¿Fuiste Responsable de IVA a 31 de diciembre?", value=bool(d["responsable_iva"]), key=f"iva_{v}")

        st.markdown("**Progreso frente a los Topes DIAN para estar obligado a declarar:**")
        for nombre_t, actual_t, limite_t in [
            ("Patrimonio Bruto", res["c29_patrimonio_bruto"], res["tope_pat_pesos"]),
            ("Ingresos Brutos", res["ing_brutos_totales"], res["tope_ing_pesos"]),
            ("Consignaciones / Inversiones", d["consignaciones"], res["tope_ing_pesos"]),
            ("Compras / Tarjetas", max(d["compras_totales"], d["tarjetas_credito"]), res["tope_ing_pesos"]),
        ]:
            pct_t = min(1.0, (actual_t / limite_t) if limite_t > 0 else 0.0)
            icono_t = "🔴 Supera tope" if actual_t >= limite_t else "🟢 Bajo el tope"
            st.progress(pct_t, text=f"{icono_t} — {nombre_t}: ${actual_t:,.0f} / Tope: ${limite_t:,.0f}")

        st.divider()
        d["patrimonio_liquido_anterior"] = st.number_input(
            "Patrimonio Líquido de tu declaración del año pasado (Opcional — revisa riesgo Art. 236)",
            min_value=0.0, value=float(d["patrimonio_liquido_anterior"]), step=500_000.0, key=f"pat_ant_{v}",
        )

# ==============================================================================
# PASO 3: INGRESOS DEL AÑO (EXPLICADOS FÁCIL PARA CUALQUIER PERSONA)
# ==============================================================================
with t3:
    col_i1, col_i2 = st.columns(2)

    with col_i1:
        st.subheader("👔 1. Empleados / Salarios (Certificado Formulario 220)")
        d["ing_brutos_trabajo"] = st.number_input("Total pagos laborales / salarios brutos (Casilla 32)", min_value=0.0, value=float(d["ing_brutos_trabajo"]), step=200_000.0, key=f"itr_{v}")
        d["salud_pension_trabajo"] = st.number_input("Aportes obligatorios a Salud + Pensión descontados por la empresa (Casilla 33)", min_value=0.0, value=float(d["salud_pension_trabajo"]), step=50_000.0, key=f"sp_{v}")
        d["cesantias_exentas"] = st.number_input("Cesantías e intereses de cesantías recibidas/consignadas (Exentas Art. 206)", min_value=0.0, value=float(d["cesantias_exentas"]), step=100_000.0, key=f"ces_{v}")

        st.divider()
        st.subheader("💼 2. Independientes: Servicios y Honorarios")
        d["ing_brutos_honorarios"] = st.number_input("Ingresos por prestación de servicios u honorarios", min_value=0.0, value=float(d["ing_brutos_honorarios"]), step=200_000.0, key=f"ihon_{v}")
        d["salud_pension_honorarios"] = st.number_input("Tus pagos anuales de Salud + Pensión en planillas PILA", min_value=0.0, value=float(d["salud_pension_honorarios"]), step=50_000.0, key=f"sphon_{v}")
        c_arl, c_cost = st.columns(2)
        d["arl_honorarios"] = c_arl.number_input("Pagos ARL Riesgos Laborales", min_value=0.0, value=float(d["arl_honorarios"]), step=10_000.0, key=f"arl_{v}")
        d["costos_honorarios"] = c_cost.number_input("Otros costos/gastos con factura", min_value=0.0, value=float(d["costos_honorarios"]), step=100_000.0, key=f"chon_{v}")

        if d["ing_brutos_honorarios"] > 0:
            st.info(
                f"🤖 **Comparador Automático DIAN:**\n"
                f"- Base gravable usando **25% Renta Exenta (Casilla 32)**: `${res['remanente_ruta_25']:,.0f}`\n"
                f"- Base gravable usando **Costos + ARL (Casilla 45)**: `${res['remanente_ruta_costos']:,.0f}`\n"
                f"👉 El sistema aplicó automáticamente: **{res['estrategia']}**"
            )

    with col_i2:
        st.subheader("🏦 3. Rentas de Capital: CDTs, Intereses y Arriendos")
        d["ing_brutos_capital"] = st.number_input("Intereses/rendimientos CDT o ingresos por arriendos (Casilla 58)", min_value=0.0, value=float(d["ing_brutos_capital"]), step=50_000.0, key=f"icap_{v}")
        d["incrngo_capital"] = st.number_input("Componente Inflacionario no gravado del CDT (Casilla 59)", min_value=0.0, value=float(d["incrngo_capital"]), step=50_000.0, key=f"inccap_{v}")
        if d["ing_brutos_capital"] > 0:
            pct_sug = COMP_INFLACIONARIO_HISTORICO.get(anio_gravable, 0.5544)
            if st.button(f"✨ Calcular Componente Inflacionario ({pct_sug*100:.2f}%)", use_container_width=True):
                d["incrngo_capital"] = float(redondear_dian(d["ing_brutos_capital"] * pct_sug))
                st.session_state.form_version += 1
                st.rerun()
        d["costos_capital"] = st.number_input("Gastos de inmuebles arrendados (Predial, admon., reparaciones — Casilla 60)", min_value=0.0, value=float(d["costos_capital"]), step=50_000.0, key=f"ccap_{v}")

        st.divider()
        st.subheader("🏪 4. Comercio, Pensiones, Dividendos y Venta de Bienes")
        with st.expander("🛒 Ventas comerciales / Negocios / Rentas No Laborales (Casillas 74 a 78)"):
            d["ing_brutos_no_laborales"] = st.number_input("Ingresos brutos ventas o negocios (Casilla 74)", min_value=0.0, value=float(d["ing_brutos_no_laborales"]), step=200_000.0, key=f"inol_{v}")
            d["incrngo_no_laborales"] = st.number_input("Salud y pensión pagada sobre rentas no laborales (Casilla 76)", min_value=0.0, value=float(d["incrngo_no_laborales"]), step=50_000.0, key=f"incnol_{v}")
            d["costos_no_laborales"] = st.number_input("Compras de mercancía y costos soportados (Casilla 77)", min_value=0.0, value=float(d["costos_no_laborales"]), step=100_000.0, key=f"cnol_{v}")

        with st.expander("👵 Mesadas Pensionales y Dividendos (Casillas 99 a 107)"):
            d["ing_brutos_pensiones"] = st.number_input("Total mesadas pensionales del año (Casilla 99)", min_value=0.0, value=float(d["ing_brutos_pensiones"]), step=200_000.0, key=f"ipen_{v}")
            d["salud_pensiones"] = st.number_input("Descuento de salud en pensión (Casilla 100)", min_value=0.0, value=float(d["salud_pensiones"]), step=50_000.0, key=f"spen_{v}")
            d["dividendos_gravables"] = st.number_input("Dividendos recibidos de sociedades nacionales (Casilla 107)", min_value=0.0, value=float(d["dividendos_gravables"]), step=100_000.0, key=f"div_{v}")

        with st.expander("🏡 Ganancias Ocasionales: Venta de casa/lote/carro (>= 2 años), herencias o loterías"):
            d["ing_ganancia_ocasional"] = st.number_input("Precio de venta de inmueble/vehículo (>= 2 años) o herencia (tarifa 15%)", min_value=0.0, value=float(d["ing_ganancia_ocasional"]), step=1_000_000.0, key=f"igo_{v}")
            d["costos_ganancia_ocasional"] = st.number_input("Costo fiscal de compra del bien vendido (Casilla 113)", min_value=0.0, value=float(d["costos_ganancia_ocasional"]), step=1_000_000.0, key=f"cgo_{v}")
            d["exenta_ganancia_ocasional"] = st.number_input("Parte exenta de herencia / seguro de vida / venta vivienda (Casilla 114)", min_value=0.0, value=float(d["exenta_ganancia_ocasional"]), step=500_000.0, key=f"ego_{v}")
            d["ing_loterias_rifas"] = st.number_input("Premios por loterías, rifas o apuestas (tarifa 20%)", min_value=0.0, value=float(d["ing_loterias_rifas"]), step=500_000.0, key=f"lot_{v}")

# ==============================================================================
# PASO 4: BENEFICIOS PARA BAJAR IMPUESTO, RETENCIONES Y SANCIONES
# ==============================================================================
with t4:
    st.subheader("🎯 Termómetro de Beneficios Tributarios (Tope Ley 2277: 40% o 1.340 UVT)")
    pct_bolsa = min(1.0, (res["usado_40_cg"] / res["limite_40_cg"]) if res["limite_40_cg"] > 0 else 0.0)
    st.progress(
        pct_bolsa,
        text=f"Beneficios usados en bolsa 40%: ${res['usado_40_cg']:,.0f} de ${res['limite_40_cg']:,.0f} permitidos (Cupo disponible: ${res['cupo_libre_40']:,.0f})",
    )

    b_col1, b_col2, b_col3 = st.columns(3)
    with b_col1:
        st.markdown("#### ⭐ Alivios Estrella (Fuera del tope 40%)")
        d["num_dependientes"] = st.number_input(
            "Hijos, cónyuge, padres o hermanos dependientes (Casilla 138)",
            min_value=0, max_value=10, value=int(d["num_dependientes"]), step=1, key=f"dep_{v}",
            help="Aplica automáticamente hasta 4 dependientes (72 UVT c/u fuera del límite) + 10% laboral dentro del límite.",
        )
        d["compras_factura_electronica"] = st.number_input(
            "Monto compras con Factura Electrónica (Base 1%)",
            min_value=0.0, value=float(d["compras_factura_electronica"]), step=100_000.0, key=f"fe_{v}",
        )
        d["aplicar_beneficio_fe"] = st.toggle(
            "✅ Aplicar 1% Factura Electrónica en Casilla 28 (Si pagaste con tarjeta/transferencia)",
            value=bool(d["aplicar_beneficio_fe"]), key=f"apfe_{v}",
        )

    with b_col2:
        st.markdown("#### 🏥 Deducciones y Ahorro (Dentro del 40%)")
        d["intereses_vivienda"] = st.number_input("Intereses pagados por crédito hipotecario / leasing (Casilla 38)", min_value=0.0, value=float(d["intereses_vivienda"]), step=100_000.0, key=f"viv_{v}")
        d["medicina_prepagada"] = st.number_input("Medicina prepagada o póliza de salud privada (Casilla 39)", min_value=0.0, value=float(d["medicina_prepagada"]), step=100_000.0, key=f"prep_{v}")
        d["intereses_icetex"] = st.number_input("Intereses pagados por crédito educativo ICETEX (Tope 100 UVT)", min_value=0.0, value=float(d["intereses_icetex"]), step=50_000.0, key=f"ice_{v}")
        d["gmf_certificado_100"] = st.number_input("4x1000 GMF (Escribe el 100% del certificado; restamos el 50%)", min_value=0.0, value=float(d["gmf_certificado_100"]), step=20_000.0, key=f"gmf_{v}")
        d["aportes_afc_fpv"] = st.number_input("Aportes voluntarios a Pensión o Cuentas AFC (Casilla 35)", min_value=0.0, value=float(d["aportes_afc_fpv"]), step=100_000.0, key=f"afc_{v}")
        d["donaciones_entidades"] = st.number_input("Donaciones certificadas a fundaciones (Descuento 25% — Casilla 125)", min_value=0.0, value=float(d["donaciones_entidades"]), step=100_000.0, key=f"don_{v}")

    with b_col3:
        st.markdown("#### 💰 Retenciones, Saldos y Extemporaneidad")
        d["retenciones_anio"] = st.number_input("Retenciones en la fuente que te practicaron en el año (Casilla 132)", min_value=0.0, value=float(d["retenciones_anio"]), step=10_000.0, key=f"ret_{v}")
        d["saldo_favor_anterior"] = st.number_input("Saldo a favor del año pasado sin pedir devolución (Casilla 131)", min_value=0.0, value=float(d["saldo_favor_anterior"]), step=50_000.0, key=f"sfa_{v}")
        d["anticipo_anio_anterior"] = st.number_input("Anticipo pagado en la declaración del año pasado (Casilla 130)", min_value=0.0, value=float(d["anticipo_anio_anterior"]), step=50_000.0, key=f"ana_{v}")
        d["impuesto_neto_anterior"] = st.number_input("Impuesto neto del año pasado (Casilla 126 anterior — optimiza anticipo)", min_value=0.0, value=float(d["impuesto_neto_anterior"]), step=50_000.0, key=f"ima_{v}")
        st.divider()
        d["meses_extemporaneidad"] = st.number_input("Meses o fracción de retraso si presentas vencido (0 = A tiempo)", min_value=0, max_value=60, value=int(d["meses_extemporaneidad"]), step=1, key=f"mora_{v}")
        d["aplica_reduccion_640"] = st.checkbox("Reducir sanción al 50% por no tener sanciones en los últimos 2 años (Art. 640)", value=bool(d["aplica_reduccion_640"]), key=f"r640_{v}")

    if st.button("🔄 Recalcular Liquidación Oficial Ahora", type="primary"):
        st.rerun()

# ==============================================================================
# PASO 5: FORMULARIO 210 OFICIAL, DESCARGAS EJECUTIVAS Y CONSULTOR IA
# ==============================================================================
with t5:
    if res["c137_total_saldo_favor"] > 0:
        st.success(
            f"### 🎉 ¡TOTAL SALDO A FAVOR (Casilla 137): ${res['c137_total_saldo_favor']:,.0f}!\n"
            f"- **Estado DIAN:** {res['estado_obligacion']} | **Estrategia aplicada:** {res['estrategia']}\n"
            f"- **Impuesto a Cargo (C.129):** ${res['c129_total_impuesto_cargo']:,.0f} | **Retenciones y Saldos a Favor:** ${res['alivios_totales']:,.0f}"
        )
    elif res["c136_total_saldo_pagar"] > 0:
        st.warning(
            f"### 💳 TOTAL SALDO A PAGAR (Casilla 136): ${res['c136_total_saldo_pagar']:,.0f}\n"
            f"- **Estado DIAN:** {res['estado_obligacion']} | **Estrategia aplicada:** {res['estrategia']}\n"
            f"- **Impuesto a Cargo (C.129):** ${res['c129_total_impuesto_cargo']:,.0f} | **Retenciones y Saldos a Favor:** ${res['alivios_totales']:,.0f}"
        )
    else:
        st.info(f"### ✅ Total Saldo a Pagar: $0 — Estado: {res['estado_obligacion']}")

    if res["causales_obligacion"]:
        st.caption("**Causales activadas:** " + " • ".join(res["causales_obligacion"]))
    for al in res["alertas_auditoria"]:
        st.error(f"🚨 {al}")
    for sg in res["sugerencias_optimizacion"]:
        st.info(f"💡 **Recomendación Inteligente:** {sg}")

    dv_mostrar = calcular_dv_dian(d.get("nit_declarante", ""))
    st.markdown(
        f"**Encabezado Muisca DIAN:** `Casilla 1 Año:` **{anio_gravable}** | "
        f"`Casilla 5 NIT:` **{d.get('nit_declarante') or 'Pendiente'}** | "
        f"`Casilla 6 DV:` **{dv_mostrar or '-'}** | "
        f"`Casilla 12 Seccional:` **{d.get('seccional_dian', '32 - Bogotá D.C.').split(' - ')[0]}** | "
        f"`Casilla 24 Actividad:` **{d.get('cod_actividad', '8211')}**"
    )

    solo_activas = st.checkbox("Mostrar únicamente las casillas con valor mayor a $0 (Modo lectura rápida para Muisca)", value=False)
    filas_vista = construir_filas_210(res)
    if solo_activas:
        filas_vista = [
            (cas, conc, val)
            for cas, conc, val in filas_vista
            if val > 0 or cas in ("29", "31", "97", "111", "129", "136", "137")
        ]

    df_vista = pd.DataFrame(
        [
            (cas, conc, f"{val:,.0f}" if cas == "138" else f"${val:,.0f}")
            for cas, conc, val in filas_vista
        ],
        columns=["Casilla Muisca", "Concepto Oficial Formulario 210 DIAN", "Valor Redondeado DIAN"],
    )
    st.dataframe(df_vista, use_container_width=True, hide_index=True, height=460)

    xlsx_bytes = exportar_excel_dian(res, d, anio_gravable, st.session_state.df_pila_cargada)
    docx_bytes = exportar_dictamen_word(res, d, anio_gravable, valor_uvt)

    b1, b2 = st.columns(2)
    b1.download_button(
        "📥 Descargar Excel Profesional (Formulario 210 + Resumen + Detalle PILA)",
        data=xlsx_bytes,
        file_name=f"Declaracion_Renta_AG{anio_gravable}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    b2.download_button(
        "📄 Descargar Dictamen Tributario (.docx)",
        data=docx_bytes,
        file_name=f"Dictamen_Renta_AG{anio_gravable}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

    st.divider()
    st.subheader("⚖️ Consultor Tributario IA con Gemini (Conectado a tu Declaración)")
    for msj in st.session_state.chat_mensajes:
        with st.chat_message(msj["role"]):
            st.markdown(msj["content"])

    consulta = st.chat_input("Pregunta cualquier duda sobre tu declaración, casillas del Muisca o cómo pagar menos impuesto...")
    if consulta:
        if not api_key:
            st.error("Configura tu Gemini API Key en Secrets o en la barra lateral para usar el Consultor IA.")
        else:
            st.session_state.chat_mensajes.append({"role": "user", "content": consulta})
            with st.chat_message("user"):
                st.markdown(consulta)
            with st.chat_message("assistant"):
                try:
                    with st.spinner("Analizando Estatuto Tributario sobre tus casillas..."):
                        r_rag = consultar_estratega_rag(api_key, consulta, res, valor_uvt, anio_gravable)
                        st.markdown(r_rag)
                        st.session_state.chat_mensajes.append({"role": "assistant", "content": r_rag})
                except Exception as e:
                    st.error(f"⚠️ {e}")
