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
    page_title="RentaIA Pro | Liquidador & Auditor DIAN 210",
    page_icon="📊",
    layout="wide",
)

# ==============================================================================
# 1. PARÁMETROS OFICIALES DIAN MULTIAÑO (PREPARADO PARA EL FUTURO)
# ==============================================================================
UVT_OFICIAL = {
    2022: 38_004,
    2023: 42_412,  # Resolución DIAN 001264 de 2022
    2024: 47_065,  # Resolución DIAN 000187 de 2023
    2025: 49_799,  # Resolución DIAN 000193 de 2024 (AG 2025 declarado en 2026)
    2026: 52_374,  # Resolución DIAN 000238 de 2025 (AG 2026 declarado en 2027)
}

COMP_INFLACIONARIO_HISTORICO = {
    2023: 0.6671,
    2024: 0.5088,
    2025: 0.5544,  # Ej: $781.461 * 55,44% -> $433.000 (Casilla 59 Form 210)
}

LIMITES_LEY_2277 = {
    "pct_bolsa_global": 0.40,
    "tope_bolsa_uvt": 1_340,
    "exenta_25_uvt": 790,
    "vivienda_uvt": 1_200,
    "prepagada_mes_uvt": 16,
    "dep_10pct_mes_uvt": 32,
    "dep_extra_uvt": 72,
    "max_dep_extra": 4,
    "factura_elec_pct": 0.01,
    "factura_elec_uvt": 240,
    "afc_fpv_uvt": 3_800,
    "pensiones_exentas_uvt": 12_000,
}

# Tabla Progresiva Oficial Art. 241 E.T.: (desde_uvt, hasta_uvt, tarifa, base_uvt)
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
            "La Gemini API Key ingresada no tiene un formato válido (debe iniciar por 'AIza...'). "
            "Verifica que no tenga comillas ni espacios adicionales."
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
                    f"Google rechazó la API Key (Código {codigo}). Revisa en Google AI Studio que la llave esté activa."
                )
            continue
        except Exception as e:
            ultimo_error = e
            continue

    raise RuntimeError(
        f"El servicio de IA alcanzó temporalmente el límite de consultas por minuto. "
        f"Espera unos segundos y vuelve a intentar. ({type(ultimo_error).__name__})"
    )


# ==============================================================================
# 3. MOTOR MATEMÁTICO DETERMINISTA CALCARDO AL MUISCA DIAN (CASILLAS 28 A 139)
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
        "nombre_declarante": "",
        "nit_declarante": "",
        "cod_actividad": "8211",
        # Patrimonio y Topes DIAN
        "efectivo": 0.0,
        "vehiculos": 0.0,
        "inmuebles": 0.0,
        "otros_activos": 0.0,
        "consignaciones": 0.0,
        "tarjetas_credito": 0.0,
        "compras_totales": 0.0,
        "responsable_iva": False,
        "deudas": 0.0,
        "patrimonio_liquido_anterior": 0.0,
        # Subcédulas Cédula General
        "ing_brutos_trabajo": 0.0,
        "salud_pension_trabajo": 0.0,
        "ing_brutos_honorarios": 0.0,
        "salud_pension_honorarios": 0.0,
        "costos_honorarios": 0.0,
        "ing_brutos_capital": 0.0,
        "incrngo_capital": 0.0,
        "costos_capital": 0.0,
        "ing_brutos_no_laborales": 0.0,
        "incrngo_no_laborales": 0.0,
        "costos_no_laborales": 0.0,
        # Pensiones, Dividendos y Ganancias Ocasionales
        "ing_brutos_pensiones": 0.0,
        "salud_pensiones": 0.0,
        "dividendos_gravables": 0.0,
        "ing_ganancia_ocasional": 0.0,
        "costos_ganancia_ocasional": 0.0,
        "exenta_ganancia_ocasional": 0.0,
        # Beneficios Ley 2277
        "intereses_vivienda": 0.0,
        "medicina_prepagada": 0.0,
        "meses_prepagada": 12,
        "num_dependientes": 0,
        "gmf_certificado_100": 0.0,
        "aportes_afc_fpv": 0.0,
        "compras_factura_electronica": 0.0,
        "aplicar_beneficio_fe": False,
        # Liquidación Privada
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
    pat_bruto_raw = d["efectivo"] + d["vehiculos"] + d["inmuebles"] + d["otros_activos"]
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

    # 3. SIMULADOR MUISCA: RUTA A (Honorarios en C.32 con 25% Exento) vs RUTA B (Honorarios en C.43 con Costos C.45)
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
    ded_gmf = redondear_dian(d["gmf_certificado_100"] * 0.50)
    c35_afc = redondear_dian(min(d["aportes_afc_fpv"], ing_brutos_cg * 0.30, L["afc_fpv_uvt"] * uvt))

    deducciones_comunes = c38_vivienda + ded_prep + ded_dep10 + ded_gmf + c35_afc
    limite_40_cg = redondear_dian(min(ing_netos_cg * L["pct_bolsa_global"], L["tope_bolsa_uvt"] * uvt))
    tope_790_pesos = L["exenta_25_uvt"] * uvt

    # RUTA A (Igual al Muisca cuando se elige el 25% exento: Trabajo + Servicios se declaran en Casillas 32-36)
    c32_a = ing_trab_total_bruto
    c33_a = inc_trab_total
    c34_a = max(0, c32_a - c33_a)
    c91_a = c34_a + c61_renta_liq_capital + c78_renta_liq_nolab
    base_25_a = max(0, c34_a - deducciones_comunes)
    exenta_25_a = redondear_dian(min(base_25_a * 0.25, tope_790_pesos))
    solicitados_a = deducciones_comunes + exenta_25_a
    aceptados_40_a = min(solicitados_a, limite_40_cg, c91_a)
    remanente_a = max(0, c91_a - aceptados_40_a)

    # RUTA B (Honorarios van a Casillas 43-46 con costos en Casilla 45, sin 25% exento sobre honorarios)
    c32_b = redondear_dian(d["ing_brutos_trabajo"])
    c33_b = min(c32_b, redondear_dian(d["salud_pension_trabajo"]))
    c34_b = max(0, c32_b - c33_b)
    c43_b = redondear_dian(d["ing_brutos_honorarios"])
    pila_disp_hon = redondear_dian(d["salud_pension_honorarios"]) + max(0, redondear_dian(d["salud_pension_trabajo"]) - c33_b)
    c44_b = min(c43_b, pila_disp_hon)
    neto_hon_b = max(0, c43_b - c44_b)
    c45_b = min(neto_hon_b, redondear_dian(d["costos_honorarios"]))
    c46_b = max(0, neto_hon_b - c45_b)
    c91_b = c34_b + c46_b + c61_renta_liq_capital + c78_renta_liq_nolab
    base_25_b = max(0, c34_b - deducciones_comunes)
    exenta_25_b = redondear_dian(min(base_25_b * 0.25, tope_790_pesos))
    solicitados_b = deducciones_comunes + exenta_25_b
    aceptados_40_b = min(solicitados_b, limite_40_cg, c91_b)
    remanente_b = max(0, c91_b - aceptados_40_b)

    if d["costos_honorarios"] > 0 and c43_b > 0 and remanente_b < remanente_a:
        estrategia = "Honorarios con Costos y Gastos (Casillas 43-46 — menor base gravable)"
        c32_ing_trabajo, c33_inc_trabajo, c34_renta_liq_trabajo = c32_b, c33_b, c34_b
        c43_ing_honorarios, c44_inc_honorarios, c45_costos_hon, c46_renta_liq_hon = c43_b, c44_b, c45_b, c46_b
        c91_renta_liq_cg = c91_b
        c36_exenta_25 = exenta_25_b
        beneficios_40_limitados = aceptados_40_b
        solicitados_bolsa_40 = solicitados_b
        remanente_cg = remanente_b
    else:
        estrategia = "Rentas de Trabajo con 25% Exento (Casillas 32-36 — óptima Ley 2277)"
        c32_ing_trabajo, c33_inc_trabajo, c34_renta_liq_trabajo = c32_a, c33_a, c34_a
        c43_ing_honorarios, c44_inc_honorarios, c45_costos_hon, c46_renta_liq_hon = 0, 0, 0, 0
        c91_renta_liq_cg = c91_a
        c36_exenta_25 = exenta_25_a
        beneficios_40_limitados = aceptados_40_a
        solicitados_bolsa_40 = solicitados_a
        remanente_cg = remanente_a

    incrngo_cg = c33_inc_trabajo + c44_inc_honorarios + c59_inc_capital + c76_inc_nolab
    costos_totales_cg = c45_costos_hon + c60_costos_capital + c77_costos_nolab

    # 4. BENEFICIOS FUERA DEL LÍMITE DEL 40% (Casilla 28 Factura Electrónica + Casilla 139 Dependientes 72 UVT)
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

    c112_go_ingresos = redondear_dian(d["ing_ganancia_ocasional"])
    c113_go_costos = redondear_dian(d["costos_ganancia_ocasional"])
    c114_go_exentas = redondear_dian(d["exenta_ganancia_ocasional"])
    c115_go_gravable = max(0, c112_go_ingresos - c113_go_costos - c114_go_exentas)

    # 6. EVALUACIÓN DE LOS 6 CRITERIOS DIAN
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

    # 7. LIQUIDACIÓN PRIVADA (Casillas 126 a 137)
    c126_impuesto_neto_renta = redondear_dian(calcular_impuesto_art241(c111_rlg_total, uvt))
    c127_impuesto_go = redondear_dian(c115_go_gravable * 0.15)
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
            "Tienes ingresos de trabajo/servicios pero $0 en Salud y Pensión (Casilla 33): recuerda ingresar tus aportes PILA en la Pestaña 3."
        )
    if ded_fe_potencial > 0 and not d.get("aplicar_beneficio_fe", False):
        sugerencias.append(
            f"Detectamos compras con Factura Electrónica susceptibles del 1% (deducción adicional de ${ded_fe_potencial:,.0f} en Casilla 28). "
            "Si las pagaste con tarjeta o medios electrónicos, activa el switch en la Pestaña 3."
        )
    cupo_40 = max(0, limite_40_cg - solicitados_bolsa_40)
    if cupo_40 > 300_000 and c126_impuesto_neto_renta > 0:
        sugerencias.append(
            f"Tienes un cupo disponible de ${cupo_40:,.0f} en la bolsa del 40% / 1.340 UVT "
            f"(puedes restar 50% del GMF 4x1000, intereses de vivienda, medicina prepagada o aportes AFC/FPV)."
        )

    return {
        "obligado_bool": obligado_bool,
        "estado_obligacion": estado_obligacion,
        "causales_obligacion": causales,
        "estrategia": estrategia,
        "c28_factura_electronica": c28_ded_factura_1pct,
        "c29_patrimonio_bruto": c29_patrimonio_bruto,
        "c30_deudas": c30_deudas,
        "c31_patrimonio_liquido": c31_patrimonio_liquido,
        "c32_ing_trabajo": c32_ing_trabajo,
        "c33_inc_trabajo": c33_inc_trabajo,
        "c34_renta_liq_trabajo": c34_renta_liq_trabajo,
        "c36_exenta_25": c36_exenta_25,
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
        "c103_rlg_pensiones": c103_rlg_pensiones,
        "c107_dividendos": c107_dividendos,
        "c111_rlg_total": c111_rlg_total,
        "rlg_en_uvt": round(c111_rlg_total / uvt, 2) if uvt > 0 else 0,
        "c115_go_gravable": c115_go_gravable,
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
# 4. MOTOR EXÓGENA DIAN CALIBRADO AL 100% + MIGRADOR MACRO + GEMINI MULTIMODAL
# ==============================================================================
class CertificadoExtraido(BaseModel):
    tipo_documento: str = Field(description="Ej: Certificado 220, Planilla PILA, Certificado CDT, Predial")
    entidad: str = Field(description="Banco, operador o empresa emisora")
    efectivo_bancos_dic31: float = Field(default=0.0, description="Saldo en cuentas a 31 de diciembre")
    valor_inmueble_o_cdt_dic31: float = Field(default=0.0, description="Avalúo predial o capital CDT vigente a dic 31")
    saldo_deuda_dic31: float = Field(default=0.0, description="Saldo capital de deudas/hipotecas a dic 31")
    ingresos_brutos_trabajo: float = Field(default=0.0, description="Ingresos laborales o salarios (Casilla 32)")
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


def importar_macro_xlsm(archivo_bytes: bytes, anio_gravable: int) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(archivo_bytes), data_only=True)
    if "Ingreso_Datos" not in wb.sheetnames:
        raise ValueError("El archivo Excel no contiene la hoja 'Ingreso_Datos'.")
    ws = wb["Ingreso_Datos"]
    iva_txt = str(ws["C11"].value or "").strip().upper()
    ing_cap = _nvl(ws["C15"].value)
    pct_inf = COMP_INFLACIONARIO_HISTORICO.get(anio_gravable, 0.5544)
    out = {
        "efectivo": _nvl(ws["C4"].value),
        "vehiculos": _nvl(ws["C5"].value),
        "inmuebles": _nvl(ws["C6"].value),
        "otros_activos": _nvl(ws["C7"].value),
        "consignaciones": _nvl(ws["C8"].value),
        "tarjetas_credito": _nvl(ws["C9"].value),
        "compras_totales": _nvl(ws["C10"].value),
        "responsable_iva": iva_txt in ("SI", "SÍ", "YES", "1", "TRUE"),
        "deudas": _nvl(ws["C12"].value),
        "ing_brutos_honorarios": _nvl(ws["C14"].value),
        "ing_brutos_trabajo": 0.0,
        "ing_brutos_capital": ing_cap,
        "incrngo_capital": float(redondear_dian(ing_cap * pct_inf)) if ing_cap > 0 else 0.0,
        "retenciones_anio": _nvl(ws["C16"].value),
        "salud_pension_honorarios": _nvl(ws["C18"].value),
        "salud_pension_trabajo": 0.0,
        "intereses_vivienda": _nvl(ws["C19"].value),
        "saldo_favor_anterior": _nvl(ws["C20"].value),
    }
    if "Parametros_UVT" in wb.sheetnames:
        uvt_leida = _nvl(wb["Parametros_UVT"]["B2"].value)
        if uvt_leida > 0:
            out["uvt_detectada"] = int(uvt_leida)
    return out


def limpiar_y_clasificar_exogena_dian(archivo_bytes: bytes, anio_gravable: int) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """Procesa el formato oficial de Exógena DIAN desduplicando columnas NIT, topes resumen y documentos soporte."""
    xls = pd.ExcelFile(io.BytesIO(archivo_bytes))
    hoja = xls.sheet_names[0]
    for h in xls.sheet_names:
        if any(p in h.lower() for p in ["report", "tercer", "detalle", "informac", "exogen"]):
            hoja = h
            break

    df_raw = pd.read_excel(xls, sheet_name=hoja, header=None)

    # 1. Extraer metadatos del encabezado superior DIAN (Filas 0 a 14)
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

    # 2. Detectar fila exacta de encabezados y eliminar nombres de columna duplicados ('NIT')
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
        nuevas_cols = []
        vistos = {}
        contador_nit = 0
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

    # 3. Separar las 5 filas resumen 'Tope 1..5' de la DIAN frente a los terceros reales
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
        "notas": [],
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
        det = str(row.get(col_detalle, "") or "").strip().lower()
        uso = str(row.get(col_uso, "") or "").strip().lower() if col_uso else ""

        det_n = det.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
        uso_n = uso.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")

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

    # Regla Oficial DIAN por NIT: Mayor valor entre Exógena e Ingresos por Documentos Soporte netos
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

    resumen["notas"].append(
        "Completa manualmente en las Pestañas 2 y 3 lo que nunca viene en Exógena: "
        "1) Inmuebles/lotes y vehículos, 2) Planilla PILA de Salud y Pensión como independiente, y 3) Saldo a favor año anterior."
    )

    # Normalizar tipos para evitar errores de visualización PyArrow en st.dataframe
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
        "exactos de este soporte tributario. Si un concepto no aparece, devuelve 0. Suma salud obligatoria "
        "y pensión obligatoria en 'salud_pension_obligatoria'."
    )

    cfg = types.GenerateContentConfig(
        response_mime_type="application/json",
        response_schema=CertificadoExtraido,
        temperature=0.0,
    )
    texto_json = ejecutar_gemini_seguro(api_key, contenidos, cfg)
    return json.loads(texto_json)


def auditar_exogena_ia(api_key: str, df_exo: pd.DataFrame) -> str:
    csv_muestra = df_exo.head(140).to_csv(index=False)
    prompt = (
        "Actúa como Auditor Tributario Senior DIAN Colombia. Revisa este reporte limpio de terceros y explica:\n"
        "1. Qué entidades explican el cumplimiento de topes (ej. renovaciones de CDT y movimientos bancarios).\n"
        "2. Cruce entre Servicios reportados y Documentos Soporte / Notas de Ajuste.\n"
        "3. Soportes físicos que el contribuyente debe guardar en su carpeta tributaria.\n\n"
        f"DATOS EXÓGENA:\n{csv_muestra}"
    )
    cfg = types.GenerateContentConfig(temperature=0.1)
    return ejecutar_gemini_seguro(api_key, prompt, cfg)


def consultar_estratega_rag(api_key: str, pregunta: str, res_liq: dict, uvt: float, anio: int) -> str:
    instruccion = (
        f"Eres un Contador Público Tributarista Senior en Colombia experto en Formulario 210 y Ley 2277 de 2022. "
        f"Año Gravable: {anio} | UVT aplicada: ${uvt:,.0f} COP. "
        f"Liquidación oficial calculada por el motor determinista: {json.dumps(res_liq, ensure_ascii=False)}. "
        "Responde de forma clara, estratégica y citando casillas del Formulario 210 y artículos del Estatuto Tributario."
    )
    cfg = types.GenerateContentConfig(system_instruction=instruccion, temperature=0.2)
    return ejecutar_gemini_seguro(api_key, pregunta, cfg)


# ==============================================================================
# 5. EXPORTADOR OFICIAL EN MEMORIA RAM (.XLSX CASILLAS MUISCA + .DOCX DICTAMEN)
# ==============================================================================
def construir_filas_210(res: dict) -> List[Tuple[str, str, int]]:
    return [
        ("28", "Uno por ciento (1%) de compras con factura electrónica", res["c28_factura_electronica"]),
        ("29", "Total patrimonio bruto", res["c29_patrimonio_bruto"]),
        ("30", "Deudas", res["c30_deudas"]),
        ("31", "Total patrimonio líquido (29 - 30)", res["c31_patrimonio_liquido"]),
        ("32", "Ingresos brutos Rentas de Trabajo (Salarios / Servicios con 25% exento)", res["c32_ing_trabajo"]),
        ("33", "Ingresos no constitutivos de renta Trabajo (Salud + Pensión PILA)", res["c33_inc_trabajo"]),
        ("34", "Renta líquida Trabajo (32 - 33)", res["c34_renta_liq_trabajo"]),
        ("36", "Renta exenta 25% laboral calculada (Art. 206 num. 10)", res["c36_exenta_25"]),
        ("43", "Ingresos brutos Honorarios sujetos a costos y gastos", res["c43_ing_honorarios"]),
        ("44", "Ingresos no constitutivos de renta Honorarios", res["c44_inc_honorarios"]),
        ("45", "Costos y deducciones procedentes Honorarios", res["c45_costos_honorarios"]),
        ("46", "Renta líquida Honorarios (43 - 44 - 45)", res["c46_renta_liq_honorarios"]),
        ("58", "Ingresos brutos Rentas de Capital (CDT / Intereses / Arriendos)", res["c58_ing_capital"]),
        ("59", "Ingresos no constitutivos Capital (Componente Inflacionario CDT)", res["c59_inc_capital"]),
        ("60", "Costos y deducciones procedentes Rentas de Capital", res["c60_costos_capital"]),
        ("61", "Renta líquida Capital (58 - 59 - 60)", res["c61_renta_liq_capital"]),
        ("74", "Ingresos brutos Rentas No Laborales", res["c74_ing_nolab"]),
        ("78", "Renta líquida No Laborales", res["c78_renta_liq_nolab"]),
        ("91", "Renta líquida Cédula General (34 + 46 + 61 + 78)", res["c91_renta_liq_cg"]),
        ("92", "Rentas exentas y deducciones imputables limitadas", res["c92_exentas_ded_limitadas"]),
        ("97", "Renta líquida gravable Cédula General (91 - 92)", res["c97_rlg_cedula_general"]),
        ("103", "Renta líquida gravable Cédula de Pensiones", res["c103_rlg_pensiones"]),
        ("107", "Dividendos y participaciones gravables", res["c107_dividendos"]),
        ("111", "TOTAL RENTA LÍQUIDA GRAVABLE (Base Art. 241 E.T.)", res["c111_rlg_total"]),
        ("115", "Ganancias ocasionales gravables", res["c115_go_gravable"]),
        ("126", "Impuesto neto de renta", res["c126_impuesto_neto_renta"]),
        ("127", "Impuesto de ganancias ocasionales (15%)", res["c127_impuesto_go"]),
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


def exportar_excel_dian(res: dict, d: dict, anio: int) -> bytes:
    df_210 = pd.DataFrame(
        construir_filas_210(res),
        columns=["Casilla 210", "Concepto Oficial Formulario 210 DIAN", "Valor Oficial (COP)"],
    )

    balance_neto = res["c136_total_saldo_pagar"] if res["c136_total_saldo_pagar"] > 0 else -res["c137_total_saldo_favor"]
    filas_dash = [
        ("Declarante", d.get("nombre_declarante") or "Contribuyente Persona Natural"),
        ("NIT / Cédula", d.get("nit_declarante") or "Sin especificar"),
        ("Actividad Económica CIIU (Casilla 24)", d.get("cod_actividad") or "8211"),
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
        ("Estado Obligación DIAN", res["estado_obligacion"]),
    ]
    df_dash = pd.DataFrame(filas_dash, columns=["Indicador Dashboard", "Resultado"])

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_210.to_excel(writer, index=False, sheet_name=f"Form_210_AG{anio}")
        df_dash.to_excel(writer, index=False, sheet_name="Dashboard")
        ws1 = writer.sheets[f"Form_210_AG{anio}"]
        ws1.column_dimensions["A"].width = 14
        ws1.column_dimensions["B"].width = 66
        ws1.column_dimensions["C"].width = 24
        for idx, c in enumerate(ws1["C"][1:], start=1):
            c.number_format = "#,##0" if df_210.iloc[idx - 1, 0] == "138" else "$#,##0"
        ws2 = writer.sheets["Dashboard"]
        ws2.column_dimensions["A"].width = 52
        ws2.column_dimensions["B"].width = 30
        for c in ws2["B"][4:14]:
            c.number_format = "$#,##0"
    return buf.getvalue()


def exportar_dictamen_word(res: dict, d: dict, anio: int, uvt: float) -> bytes:
    doc = Document()
    doc.add_heading(f"Dictamen Técnico y Borrador Formulario 210 — AG {anio}", level=1)
    titular = d.get("nombre_declarante") or "Contribuyente"
    nit = d.get("nit_declarante") or "N/A"
    doc.add_paragraph(
        f"Declarante: {titular} (ID: {nit}) | CIIU: {d.get('cod_actividad', '8211')}\n"
        f"UVT Aplicada AG {anio}: ${uvt:,.0f} COP | Estrategia Óptima: {res['estrategia']}"
    )

    doc.add_heading("1. Diagnóstico de Obligación Tributaria (6 Criterios DIAN)", level=2)
    doc.add_paragraph(f"Estado: {res['estado_obligacion']}")
    for c in res["causales_obligacion"]:
        doc.add_paragraph(c, style="List Bullet")

    doc.add_heading("2. Resumen Oficial por Casillas — Formulario 210 DIAN", level=2)
    for cas, nom, val in construir_filas_210(res):
        if val > 0 or cas in ("29", "31", "97", "111", "129", "132", "136", "137"):
            pref = "" if cas == "138" else "$"
            doc.add_paragraph(f"Casilla {cas} — {nom}: {pref}{val:,.0f}", style="List Bullet")

    if res["alertas_auditoria"] or res["sugerencias_optimizacion"]:
        doc.add_heading("3. Auditoría Preventiva y Recomendaciones", level=2)
        for al in res["alertas_auditoria"]:
            doc.add_paragraph(f"[ALERTA] {al}", style="List Bullet")
        for sg in res["sugerencias_optimizacion"]:
            doc.add_paragraph(f"[AHORRO] {sg}", style="List Bullet")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ==============================================================================
# 6. EJECUCIÓN DE INTERFAZ PRINCIPAL STREAMLIT
# ==============================================================================
if not verificar_acceso():
    st.stop()

if "datos" not in st.session_state:
    st.session_state.datos = estado_inicial_datos()
if "form_version" not in st.session_state:
    st.session_state.form_version = 0
if "ultimo_certificado" not in st.session_state:
    st.session_state.ultimo_certificado = None
if "auditoria_exogena" not in st.session_state:
    st.session_state.auditoria_exogena = None
if "chat_mensajes" not in st.session_state:
    st.session_state.chat_mensajes = []

v = st.session_state.form_version
d = st.session_state.datos

# --- BARRA LATERAL ---
with st.sidebar:
    st.title("⚙️ Parámetros DIAN")
    api_secret = obtener_secreto("GEMINI_API_KEY")
    if api_secret:
        api_key = api_secret
        st.success("🔒 Gemini API Key activa desde Secrets")
    else:
        raw_input_key = st.text_input(
            "🔑 Gemini API Key (Opcional para IA)",
            type="password",
            help="Pega solo tu clave que inicia por AIza... sin comillas ni espacios.",
        )
        api_key = limpiar_api_key(raw_input_key)

    anio_actual = datetime.date.today().year
    anio_default = st.session_state.get("anio_detectado_exo", anio_actual - 1)
    anio_gravable = int(
        st.number_input(
            "📅 Año Gravable a Declarar (Cualquier año)",
            min_value=2020,
            max_value=2100,
            value=int(anio_default),
            step=1,
            help="Funciona para cualquier año presente o futuro (2025, 2026, 2027, 2030...).",
        )
    )

    uvt_sugerida = st.session_state.get("uvt_migrada", uvt_sugerida_por_anio(anio_gravable))
    valor_uvt = st.number_input(
        f"Valor UVT AG {anio_gravable} ($ COP)",
        min_value=10_000,
        value=int(uvt_sugerida),
        step=100,
        key=f"uvt_{anio_gravable}_{v}",
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
        "Cálculo de Anticipo (Casilla 133 - Art. 807)",
        options=[0, 1, 2, 3],
        index=int(d["anio_declara_num"]),
        format_func=lambda x: {
            0: "Sin Anticipo ($0 si impuesto es $0)",
            1: "1er Año declarando (25%)",
            2: "2do Año declarando (50%)",
            3: "3er Año o más (75% Optimizado)",
        }[x],
    )

    clave_m = obtener_secreto("CLAVE_MAESTRA")
    if clave_m and st.session_state.get("es_admin", False):
        with st.expander("🛡️ Generador de Licencias (Admin)"):
            id_cli = st.text_input("Cédula o Correo del Cliente")
            if id_cli:
                st.code(generar_licencia_cliente(id_cli, clave_m))

    st.divider()
    with st.popover("🧹 Limpiar Formulario y Tablero", use_container_width=True):
        st.warning("¿Confirmas limpiar todos los datos ingresados y reiniciar el Dashboard?")
        if st.button("✅ Sí, limpiar todo", type="primary", use_container_width=True):
            st.session_state.datos = estado_inicial_datos()
            st.session_state.pop("uvt_migrada", None)
            st.session_state.pop("anio_detectado_exo", None)
            st.session_state.ultimo_certificado = None
            st.session_state.auditoria_exogena = None
            st.session_state.chat_mensajes = []
            st.session_state.form_version += 1
            st.rerun()

res = liquidar_declaracion_210(d, valor_uvt, uvt_pat, uvt_ing, uvt_sancion)

# --- ENCABEZADO Y TABLERO KPI ---
st.title(f"📊 RentaIA Pro — Liquidador & Auditor Formulario 210 (AG {anio_gravable})")

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric(
    "Obligación DIAN",
    "SÍ OBLIGADO" if res["obligado_bool"] else "NO OBLIGADO",
    delta=f"{len(res['causales_obligacion'])} causal(es)" if res["obligado_bool"] else "Voluntaria",
)
m2.metric("Patrimonio Líquido (C.31)", f"${res['c31_patrimonio_liquido']:,.0f}")
m3.metric("Renta Líquida Grav. (C.111)", f"${res['c111_rlg_total']:,.0f}", f"{res['rlg_en_uvt']} UVT")
m4.metric("Impuesto a Cargo (C.129)", f"${res['c129_total_impuesto_cargo']:,.0f}")
if res["c137_total_saldo_favor"] > 0:
    m5.metric("🎉 SALDO A FAVOR (C.137)", f"${res['c137_total_saldo_favor']:,.0f}", delta="Devolución / Compensación")
else:
    m5.metric("💳 SALDO A PAGAR (C.136)", f"${res['c136_total_saldo_pagar']:,.0f}")

# --- PESTAÑAS PRINCIPALES ---
t1, t2, t3, t4, t5 = st.tabs([
    "🚀 1. Carga Inteligente (Exógena / PDF / Macro)",
    "🏠 2. Datos, Patrimonio y Topes DIAN",
    "💼 3. Ingresos, CDT y Beneficios",
    "📋 4. Formulario 210 Oficial & Descargas",
    "⚖️ 5. Consultor Tributario IA",
])

# PESTAÑA 1: EXÓGENA NATIVA + LECTOR GEMINI + MIGRADOR MACRO
with t1:
    col_izq, col_der = st.columns(2)

    with col_izq:
        st.subheader("📊 A. Autollenado desde Información Exógena DIAN")
        up_exo = st.file_uploader(
            "Sube el Excel de Información Exógena descargado de la DIAN (Cálculo instantáneo sin IA)",
            type=["xlsx", "xls"],
        )
        if up_exo:
            try:
                df_exo, res_exo = limpiar_y_clasificar_exogena_dian(up_exo.getvalue(), anio_gravable)
                st.caption(
                    f"Terceros procesados: **{len(df_exo)}** | "
                    f"Declarante: **{res_exo['nombre_detectado'] or 'Detectado'}** (ID: **{res_exo['nit_detectado'] or 'N/A'}**)"
                )
                st.dataframe(df_exo, use_container_width=True, height=220)

                st.markdown("**Resumen Oficial Clasificado de tu Exógena:**")
                df_res_exo = pd.DataFrame([
                    ("Saldos Cuentas Bancarias a dic 31", f"${res_exo['efectivo_bancos']:,.0f}"),
                    ("Saldos CDT / Inversiones a dic 31", f"${res_exo['inversiones_cdt_saldo']:,.0f}"),
                    ("Movimientos y Consignaciones (Tope 4)", f"${res_exo['consignaciones_tope4']:,.0f}"),
                    ("Compras Totales (Tope 5)", f"${res_exo['compras_totales']:,.0f}"),
                    ("Base Factura Electrónica susceptible 1% (C.28)", f"${res_exo['factura_electronica_1pct']:,.0f}"),
                    ("Ingresos Salarios (C.32) + Servicios/Honorarios", f"${res_exo['ing_trabajo'] + res_exo['ing_honorarios']:,.0f}"),
                    ("Ingresos Capital — Rendimientos CDT (C.58)", f"${res_exo['ing_capital']:,.0f}"),
                    ("Comp. Inflacionario CDT sugerido (C.59)", f"${res_exo['incrngo_capital_sugerido']:,.0f}"),
                    ("Retenciones Practicadas (C.132)", f"${res_exo['retenciones']:,.0f}"),
                ], columns=["Concepto Tributario DIAN", "Valor Exacto"])
                st.table(df_res_exo)

                for nota in res_exo["notas"]:
                    st.info(f"📌 **Guía Contable:** {nota}")

                if st.button("⚡ Pasar Valores de Exógena al Formulario (Editable)", type="primary", use_container_width=True):
                    if res_exo["anio_detectado"]:
                        st.session_state["anio_detectado_exo"] = res_exo["anio_detectado"]
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

                if st.button("🤖 Explicación Detallada de Terceros con Gemini (Opcional)", use_container_width=True):
                    if not api_key:
                        st.error("Configura una Gemini API Key válida en Secrets o en la barra lateral.")
                    else:
                        try:
                            with st.spinner("Auditando terceros con Gemini..."):
                                st.session_state.auditoria_exogena = auditar_exogena_ia(api_key, df_exo)
                        except Exception as e:
                            st.error(f"⚠️ {e}")
            except Exception as e:
                st.error(f"No se pudo procesar el archivo de Exógena: {e}")

        if st.session_state.auditoria_exogena:
            st.success(st.session_state.auditoria_exogena)

    with col_der:
        st.subheader("📄 B. Lector IA de Certificados (PDF / Imagen)")
        up_doc = st.file_uploader(
            "Sube Certificado 220, Planilla PILA, Certificado CDT, Predial o Prepagada",
            type=["pdf", "png", "jpg", "jpeg"],
        )
        if up_doc and st.button("🔍 Extraer cifras del soporte con Gemini", type="primary", use_container_width=True):
            if not api_key:
                st.error("Configura primero tu Gemini API Key en Secrets o en la barra lateral.")
            else:
                try:
                    with st.spinner("Leyendo soporte en memoria RAM..."):
                        st.session_state.ultimo_certificado = extraer_con_gemini_hibrido(
                            api_key, up_doc.getvalue(), up_doc.type
                        )
                except Exception as e:
                    st.error(f"⚠️ {e}")

        if st.session_state.ultimo_certificado:
            ext = st.session_state.ultimo_certificado
            st.json(ext)
            if st.button("➕ Sumar valores extraídos a mi Declaración", use_container_width=True):
                d["efectivo"] += ext.get("efectivo_bancos_dic31", 0.0)
                d["otros_activos"] += ext.get("valor_inmueble_o_cdt_dic31", 0.0)
                d["deudas"] += ext.get("saldo_deuda_dic31", 0.0)
                d["ing_brutos_trabajo"] += ext.get("ingresos_brutos_trabajo", 0.0)
                d["ing_brutos_honorarios"] += ext.get("ingresos_brutos_honorarios", 0.0)
                d["salud_pension_trabajo"] += ext.get("salud_pension_obligatoria", 0.0)
                d["ing_brutos_capital"] += ext.get("ingresos_capital_cdt", 0.0)
                d["incrngo_capital"] += ext.get("componente_inflacionario_cdt", 0.0)
                d["intereses_vivienda"] += ext.get("intereses_vivienda", 0.0)
                d["medicina_prepagada"] += ext.get("medicina_prepagada", 0.0)
                d["gmf_certificado_100"] += ext.get("gmf_4x1000_certificado", 0.0)
                d["aportes_afc_fpv"] += ext.get("aportes_afc_fpv", 0.0)
                d["retenciones_anio"] += ext.get("retenciones_practicadas", 0.0)
                st.session_state.ultimo_certificado = None
                st.session_state.form_version += 1
                st.rerun()

        st.divider()
        with st.expander("📥 C. Migrar desde tu antiguo archivo Excel `.xlsm` (Opcional)"):
            up_macro = st.file_uploader("Sube 'PROGRAMA PARA DECLARAR RENTA.xlsm'", type=["xlsm", "xlsx"])
            if up_macro and st.button("⚡ Importar Hoja 'Ingreso_Datos' (C4:C20)", use_container_width=True):
                try:
                    migrados = importar_macro_xlsm(up_macro.getvalue(), anio_gravable)
                    if "uvt_detectada" in migrados:
                        st.session_state["uvt_migrada"] = migrados.pop("uvt_detectada")
                    st.session_state.datos.update(migrados)
                    st.session_state.form_version += 1
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al importar el archivo Excel: {e}")

# PESTAÑA 2: IDENTIFICACIÓN, PATRIMONIO Y LOS 6 CRITERIOS DIAN
with t2:
    id1, id2, id3 = st.columns([2, 1, 1])
    d["nombre_declarante"] = id1.text_input("Nombre Completo del Declarante (Casillas 7-10)", value=d["nombre_declarante"], key=f"nom_{v}")
    d["nit_declarante"] = id2.text_input("Cédula / NIT sin DV (Casilla 5)", value=d["nit_declarante"], key=f"nit_{v}")
    d["cod_actividad"] = id3.text_input("Código Actividad CIIU (Casilla 24)", value=d["cod_actividad"], key=f"ciiu_{v}")

    st.divider()
    p1, p2 = st.columns(2)
    with p1:
        st.subheader("Patrimonio a 31 de Diciembre (Casillas 29 a 31)")
        d["efectivo"] = st.number_input("Saldos en Bancos, Ahorros, Nequi, Daviplata, Efectivo", min_value=0.0, value=float(d["efectivo"]), step=100_000.0, key=f"ef_{v}")
        d["otros_activos"] = st.number_input("Saldos CDTs vigentes a dic 31, Fondos, Acciones", min_value=0.0, value=float(d["otros_activos"]), step=200_000.0, key=f"ot_{v}")
        d["inmuebles"] = st.number_input("Bienes Raíces: Lotes, Casas, Apartamentos (Mayor entre costo y predial)", min_value=0.0, value=float(d["inmuebles"]), step=1_000_000.0, key=f"inm_{v}")
        d["vehiculos"] = st.number_input("Vehículos, Motocicletas y Otros Bienes", min_value=0.0, value=float(d["vehiculos"]), step=500_000.0, key=f"ve_{v}")
        d["deudas"] = st.number_input("Deudas y Pasivos Vigentes a dic 31 (Casilla 30)", min_value=0.0, value=float(d["deudas"]), step=200_000.0, key=f"deu_{v}")
    with p2:
        st.subheader("Verificación de los 6 Topes DIAN y Control Patrimonial")
        d["consignaciones"] = st.number_input("Consignaciones e Inversiones Acumuladas (Tope 4)", min_value=0.0, value=float(d["consignaciones"]), step=500_000.0, key=f"con_{v}")
        d["tarjetas_credito"] = st.number_input("Consumos con Tarjeta de Crédito (Tope 3)", min_value=0.0, value=float(d["tarjetas_credito"]), step=500_000.0, key=f"tc_{v}")
        d["compras_totales"] = st.number_input("Compras y Consumos Totales (Tope 5)", min_value=0.0, value=float(d["compras_totales"]), step=500_000.0, key=f"com_{v}")
        d["responsable_iva"] = st.toggle("¿Responsable de IVA a 31 de diciembre?", value=bool(d["responsable_iva"]), key=f"iva_{v}")
        st.divider()
        d["patrimonio_liquido_anterior"] = st.number_input("Patrimonio Líquido Declaración Año Anterior (Auditoría Art. 236 E.T.)", min_value=0.0, value=float(d["patrimonio_liquido_anterior"]), step=500_000.0, key=f"pat_ant_{v}")

# PESTAÑA 3: INGRESOS, CDT, DEDUCCIONES Y LIQUIDACIÓN PRIVADA
with t3:
    c1, c2, c3 = st.columns(3)
    with c1:
        st.subheader("Trabajo, Servicios y Honorarios")
        d["ing_brutos_trabajo"] = st.number_input("Salarios / Ingresos Laborales (Casilla 32)", min_value=0.0, value=float(d["ing_brutos_trabajo"]), step=200_000.0, key=f"itr_{v}")
        d["salud_pension_trabajo"] = st.number_input("Salud + Pensión Obligatoria Salarios", min_value=0.0, value=float(d["salud_pension_trabajo"]), step=50_000.0, key=f"sp_{v}")
        st.divider()
        d["ing_brutos_honorarios"] = st.number_input("Servicios / Honorarios Independientes", min_value=0.0, value=float(d["ing_brutos_honorarios"]), step=200_000.0, help="El motor los ubica automáticamente en Casilla 32 (si conviene el 25% exento) o en Casilla 43 (si tus costos son mayores).", key=f"ihon_{v}")
        d["salud_pension_honorarios"] = st.number_input("Aportes PILA Salud + Pensión Independiente", min_value=0.0, value=float(d["salud_pension_honorarios"]), step=50_000.0, key=f"sphon_{v}")
        d["costos_honorarios"] = st.number_input("Costos/Gastos de Honorarios (Casilla 45 — Simula vs 25%)", min_value=0.0, value=float(d["costos_honorarios"]), step=100_000.0, key=f"chon_{v}")

    with c2:
        st.subheader("Capital (CDT) y Otras Cédulas")
        d["ing_brutos_capital"] = st.number_input("Ingresos de Capital: Intereses CDT / Arriendos (Casilla 58)", min_value=0.0, value=float(d["ing_brutos_capital"]), step=50_000.0, key=f"icap_{v}")
        d["incrngo_capital"] = st.number_input("Componente Inflacionario CDT — No Gravado (Casilla 59)", min_value=0.0, value=float(d["incrngo_capital"]), step=50_000.0, key=f"inccap_{v}")
        if d["ing_brutos_capital"] > 0:
            pct_sug = COMP_INFLACIONARIO_HISTORICO.get(anio_gravable, 0.5544)
            if st.button(f"✨ Recalcular Comp. Inflacionario ({pct_sug*100:.2f}%)", use_container_width=True):
                d["incrngo_capital"] = float(redondear_dian(d["ing_brutos_capital"] * pct_sug))
                st.session_state.form_version += 1
                st.rerun()
        d["costos_capital"] = st.number_input("Costos y Gastos Rentas de Capital (Casilla 60)", min_value=0.0, value=float(d["costos_capital"]), step=50_000.0, key=f"ccap_{v}")

        with st.expander("➕ Rentas No Laborales, Pensiones y Dividendos"):
            d["ing_brutos_no_laborales"] = st.number_input("Ingresos No Laborales (Casilla 74)", min_value=0.0, value=float(d["ing_brutos_no_laborales"]), step=200_000.0, key=f"inol_{v}")
            d["costos_no_laborales"] = st.number_input("Costos No Laborales (Casilla 77)", min_value=0.0, value=float(d["costos_no_laborales"]), step=100_000.0, key=f"cnol_{v}")
            d["ing_brutos_pensiones"] = st.number_input("Ingresos por Pensiones (Casilla 99)", min_value=0.0, value=float(d["ing_brutos_pensiones"]), step=200_000.0, key=f"ipen_{v}")
            d["salud_pensiones"] = st.number_input("Salud en Pensión (Casilla 100)", min_value=0.0, value=float(d["salud_pensiones"]), step=50_000.0, key=f"spen_{v}")
            d["dividendos_gravables"] = st.number_input("Dividendos Gravables (Casilla 107)", min_value=0.0, value=float(d["dividendos_gravables"]), step=100_000.0, key=f"div_{v}")

        with st.expander("🏡 Ganancias Ocasionales (Venta Lotes/Inmuebles >= 2 años)"):
            d["ing_ganancia_ocasional"] = st.number_input("Ingresos por Ganancia Ocasional (Casilla 112)", min_value=0.0, value=float(d["ing_ganancia_ocasional"]), step=1_000_000.0, key=f"igo_{v}")
            d["costos_ganancia_ocasional"] = st.number_input("Costos por Ganancia Ocasional (Casilla 113)", min_value=0.0, value=float(d["costos_ganancia_ocasional"]), step=1_000_000.0, key=f"cgo_{v}")
            d["exenta_ganancia_ocasional"] = st.number_input("Ganancia Ocasional Exenta (Casilla 114)", min_value=0.0, value=float(d["exenta_ganancia_ocasional"]), step=500_000.0, key=f"ego_{v}")

    with c3:
        st.subheader("Beneficios Ley 2277 y Retenciones")
        d["num_dependientes"] = st.number_input("Número de Dependientes Económicos (Casilla 138)", min_value=0, max_value=10, value=int(d["num_dependientes"]), step=1, key=f"dep_{v}")
        d["compras_factura_electronica"] = st.number_input("Base Compras Factura Electrónica (Susceptible 1%)", min_value=0.0, value=float(d["compras_factura_electronica"]), step=100_000.0, key=f"fe_{v}")
        d["aplicar_beneficio_fe"] = st.toggle("✅ Aplicar 1% Factura Electrónica en Casilla 28 (Si pagaste con tarjeta/transferencia)", value=bool(d["aplicar_beneficio_fe"]), key=f"apfe_{v}")
        d["intereses_vivienda"] = st.number_input("Intereses Crédito Vivienda (Casilla 38)", min_value=0.0, value=float(d["intereses_vivienda"]), step=100_000.0, key=f"viv_{v}")
        d["medicina_prepagada"] = st.number_input("Medicina Prepagada / Póliza Salud (Casilla 39)", min_value=0.0, value=float(d["medicina_prepagada"]), step=100_000.0, key=f"prep_{v}")
        d["gmf_certificado_100"] = st.number_input("GMF 4x1000 (Ingresa el 100% certificado)", min_value=0.0, value=float(d["gmf_certificado_100"]), step=20_000.0, key=f"gmf_{v}")
        d["aportes_afc_fpv"] = st.number_input("Aportes Voluntarios AFC / FPV (Casilla 35)", min_value=0.0, value=float(d["aportes_afc_fpv"]), step=100_000.0, key=f"afc_{v}")
        st.divider()
        d["retenciones_anio"] = st.number_input("Retenciones Año Gravable a Declarar (Casilla 132)", min_value=0.0, value=float(d["retenciones_anio"]), step=10_000.0, key=f"ret_{v}")
        d["saldo_favor_anterior"] = st.number_input("Saldo a Favor Año Anterior sin solicitar (Casilla 131)", min_value=0.0, value=float(d["saldo_favor_anterior"]), step=50_000.0, key=f"sfa_{v}")
        with st.expander("⏱️ Extemporaneidad (Casilla 135) y Datos Año Anterior"):
            d["meses_extemporaneidad"] = st.number_input("Meses o fracción de retraso (0 = En plazo)", min_value=0, max_value=60, value=int(d["meses_extemporaneidad"]), step=1, key=f"mora_{v}")
            d["aplica_reduccion_640"] = st.checkbox("Reducir sanción al 50% (Art. 640 E.T.)", value=bool(d["aplica_reduccion_640"]), key=f"r640_{v}")
            d["anticipo_anio_anterior"] = st.number_input("Anticipo Renta Año Anterior (Casilla 130)", min_value=0.0, value=float(d["anticipo_anio_anterior"]), step=50_000.0, key=f"ana_{v}")
            d["impuesto_neto_anterior"] = st.number_input("Impuesto Neto Año Anterior (Promediar Anticipo)", min_value=0.0, value=float(d["impuesto_neto_anterior"]), step=50_000.0, key=f"ima_{v}")

    if st.button("🔄 Actualizar Liquidación y Formulario 210", type="primary"):
        st.rerun()

# PESTAÑA 4: BORRADOR OFICIAL MUISCA CASILLAS 28 A 139 + DESCARGAS
with t4:
    if res["c137_total_saldo_favor"] > 0:
        st.success(
            f"### 🎉 ¡TOTAL SALDO A FAVOR (Casilla 137): ${res['c137_total_saldo_favor']:,.0f}!\n"
            f"- **Estado DIAN:** {res['estado_obligacion']} | **Estrategia:** {res['estrategia']}\n"
            f"- **Impuesto a Cargo (C.129):** ${res['c129_total_impuesto_cargo']:,.0f} | **Retenciones y Alivios:** ${res['alivios_totales']:,.0f}"
        )
    elif res["c136_total_saldo_pagar"] > 0:
        st.warning(
            f"### 💳 TOTAL SALDO A PAGAR (Casilla 136): ${res['c136_total_saldo_pagar']:,.0f}\n"
            f"- **Estado DIAN:** {res['estado_obligacion']} | **Estrategia:** {res['estrategia']}\n"
            f"- **Impuesto a Cargo (C.129):** ${res['c129_total_impuesto_cargo']:,.0f} | **Retenciones y Alivios:** ${res['alivios_totales']:,.0f}"
        )
    else:
        st.info(f"### ✅ Total Saldo a Pagar: $0 — Estado: {res['estado_obligacion']}")

    if res["causales_obligacion"]:
        st.caption("**Causales activadas:** " + " • ".join(res["causales_obligacion"]))
    for al in res["alertas_auditoria"]:
        st.error(f"🚨 {al}")
    for sg in res["sugerencias_optimizacion"]:
        st.info(f"💡 **Aviso Inteligente:** {sg}")

    st.subheader(f"Borrador Oficial Casilla por Casilla — Formulario 210 DIAN (AG {anio_gravable})")
    filas_vista = construir_filas_210(res)
    df_vista = pd.DataFrame(
        [
            (cas, conc, f"{val:,.0f}" if cas == "138" else f"${val:,.0f}")
            for cas, conc, val in filas_vista
        ],
        columns=["Casilla Muisca", "Concepto Oficial Formulario 210", "Valor Redondeado DIAN"],
    )
    st.dataframe(df_vista, use_container_width=True, hide_index=True, height=520)

    xlsx_bytes = exportar_excel_dian(res, d, anio_gravable)
    docx_bytes = exportar_dictamen_word(res, d, anio_gravable, valor_uvt)

    b1, b2 = st.columns(2)
    b1.download_button(
        "📥 Descargar Excel Oficial (Casillas 210 + Hoja Dashboard)",
        data=xlsx_bytes,
        file_name=f"Formulario_210_AG{anio_gravable}.xlsx",
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

# PESTAÑA 5: CONSULTOR TRIBUTARIO IA
with t5:
    st.subheader("⚖️ Estratega Tributario con Gemini (Conectado a tus Casillas 210)")
    for msj in st.session_state.chat_mensajes:
        with st.chat_message(msj["role"]):
            st.markdown(msj["content"])

    consulta = st.chat_input("Ej: ¿Por qué mis consignaciones de $142M me obligan a declarar pero mi impuesto es $0?")
    if consulta:
        if not api_key:
            st.error("Configura una Gemini API Key válida en la barra lateral o en Secrets.")
        else:
            st.session_state.chat_mensajes.append({"role": "user", "content": consulta})
            with st.chat_message("user"):
                st.markdown(consulta)
            with st.chat_message("assistant"):
                try:
                    with st.spinner("Consultando Estatuto Tributario sobre tus casillas..."):
                        r_rag = consultar_estratega_rag(api_key, consulta, res, valor_uvt, anio_gravable)
                        st.markdown(r_rag)
                        st.session_state.chat_mensajes.append({"role": "assistant", "content": r_rag})
                except Exception as e:
                    st.error(f"⚠️ {e}")
