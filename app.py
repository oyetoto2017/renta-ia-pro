import io
import os
import json
import hmac
import hashlib
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, Any, List

import pandas as pd
import openpyxl
import streamlit as st
from docx import Document
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

# Importaciones opcionales para OCR híbrido en PDFs escaneados
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
# 1. PARÁMETROS OFICIALES DIAN (UVT HISTÓRICAS Y TABLA ART. 241 CORREGIDA)
# ==============================================================================
UVT_OFICIAL = {
    2025: 49_799,  # Resolución DIAN 000193 de 2024 (AG 2025 declarado en 2026)
    2026: 52_374,  # Resolución DIAN 000238 de 2025 (Vigente en 2026 / Sanción mínima)
    2024: 47_065,  # Resolución DIAN 000187 de 2023
    2023: 42_412,  # Resolución DIAN 001264 de 2022
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


# ==============================================================================
# 2. BLINDAJE DE SEGURIDAD Y LICENCIAS SHA-256 (STREAMLIT SECRETS)
# ==============================================================================
def obtener_secreto(nombre: str) -> str:
    try:
        return str(st.secrets.get(nombre, "")).strip()
    except Exception:
        return os.getenv(nombre, "").strip()


def generar_licencia_cliente(identificador: str, semilla: str) -> str:
    limpio = identificador.strip().upper()
    firma = hmac.new(semilla.encode("utf-8"), limpio.encode("utf-8"), hashlib.sha256).hexdigest()[:12].upper()
    return f"RENTA-{firma}"


def verificar_acceso() -> bool:
    clave_maestra = obtener_secreto("CLAVE_MAESTRA")
    if not clave_maestra:
        # Si no ha configurado CLAVE_MAESTRA en Secrets, permite entrada directa
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


# ==============================================================================
# 3. MOTOR MATEMÁTICO DETERMINISTA (FORMULARIO 210 + LEY 2277 + SANCIONES)
# ==============================================================================
def redondear_dian(valor: float) -> int:
    """Redondea al múltiplo de 1.000 más cercano según Art. 577 E.T."""
    if valor <= 0:
        return 0
    return int(
        (Decimal(str(valor)) / Decimal("1000")).quantize(Decimal("1"), rounding=ROUND_HALF_UP) * 1000
    )


def estado_inicial_datos() -> Dict[str, Any]:
    return {
        # Celdas C4 a C12 de la macro original + control patrimonial
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
        # Cédula General: Trabajo / Honorarios, Capital (CDT) y No Laborales
        "ing_brutos_trabajo": 0.0,
        "salud_pension_trabajo": 0.0,
        "costos_honorarios": 0.0,
        "ing_brutos_capital": 0.0,
        "incrngo_capital": 0.0,
        "costos_capital": 0.0,
        "ing_brutos_no_laborales": 0.0,
        "incrngo_no_laborales": 0.0,
        "costos_no_laborales": 0.0,
        # Cédula Pensiones y Ganancias Ocasionales (Venta lotes/inmuebles)
        "ing_brutos_pensiones": 0.0,
        "salud_pensiones": 0.0,
        "ing_ganancia_ocasional": 0.0,
        "costos_ganancia_ocasional": 0.0,
        "exenta_ganancia_ocasional": 0.0,
        # Beneficios y Alivios Ley 2277
        "intereses_vivienda": 0.0,
        "medicina_prepagada": 0.0,
        "meses_prepagada": 12,
        "num_dependientes": 0,
        "gmf_certificado_100": 0.0,
        "aportes_afc_fpv": 0.0,
        "compras_factura_electronica": 0.0,
        # Retenciones, Anticipo y Extemporaneidad
        "retenciones_anio": 0.0,
        "saldo_favor_anterior": 0.0,
        "anio_declara_num": 0,  # 0=Sin anticipo, 1=25%, 2=50%, 3=75%
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

    # 1. Patrimonio (Renglones 29 a 31)
    pat_bruto_raw = d["efectivo"] + d["vehiculos"] + d["inmuebles"] + d["otros_activos"]
    patrimonio_bruto = redondear_dian(pat_bruto_raw)
    deudas = redondear_dian(d["deudas"])
    patrimonio_liquido = max(0, patrimonio_bruto - deudas)

    # 2. Ingresos Netos Cédula General
    ing_trab = d["ing_brutos_trabajo"]
    inc_trab = min(ing_trab, d["salud_pension_trabajo"])
    neto_trab = max(0.0, ing_trab - inc_trab)

    ing_cap = d["ing_brutos_capital"]
    inc_cap = min(ing_cap, d["incrngo_capital"])
    neto_cap = max(0.0, ing_cap - inc_cap)
    cost_cap = min(neto_cap, d["costos_capital"])

    ing_nolab = d["ing_brutos_no_laborales"]
    inc_nolab = min(ing_nolab, d["incrngo_no_laborales"])
    neto_nolab = max(0.0, ing_nolab - inc_nolab)
    cost_nolab = min(neto_nolab, d["costos_no_laborales"])

    ing_brutos_cg = ing_trab + ing_cap + ing_nolab
    incrngo_cg = inc_trab + inc_cap + inc_nolab
    ing_netos_cg = neto_trab + neto_cap + neto_nolab

    # 3. Bolsa del 40% / 1.340 UVT y Simulador Dual (25% Exento vs Costos de Honorarios)
    ded_viv = min(d["intereses_vivienda"], L["vivienda_uvt"] * uvt)
    ded_prep = min(d["medicina_prepagada"], d["meses_prepagada"] * L["prepagada_mes_uvt"] * uvt) if ing_trab > 0 else 0.0
    ded_dep10 = min(ing_trab * 0.10, 12 * L["dep_10pct_mes_uvt"] * uvt) if (d["num_dependientes"] > 0 and ing_trab > 0) else 0.0
    ded_gmf = d["gmf_certificado_100"] * 0.50
    ex_afc = min(d["aportes_afc_fpv"], ing_brutos_cg * 0.30, L["afc_fpv_uvt"] * uvt)

    ded_comunes = ded_viv + ded_prep + ded_dep10 + ded_gmf + ex_afc
    limite_40 = min(ing_netos_cg * L["pct_bolsa_global"], L["tope_bolsa_uvt"] * uvt)

    # Escenario A: Renta Exenta 25% (tope 790 UVT)
    costos_a = cost_cap + cost_nolab
    base_25 = max(0.0, neto_trab - ded_comunes)
    exenta_25 = min(base_25 * 0.25, L["exenta_25_uvt"] * uvt)
    renta_antes_a = max(0.0, ing_netos_cg - costos_a)
    aceptados_a = min(ded_comunes + exenta_25, limite_40, renta_antes_a)
    remanente_a = max(0.0, renta_antes_a - aceptados_a)

    # Escenario B: Imputación de Costos en Honorarios (renunciando al 25% exento)
    costos_hon = min(neto_trab, d["costos_honorarios"])
    costos_b = costos_hon + cost_cap + cost_nolab
    renta_antes_b = max(0.0, ing_netos_cg - costos_b)
    aceptados_b = min(ded_comunes, limite_40, renta_antes_b)
    remanente_b = max(0.0, renta_antes_b - aceptados_b)

    if d["costos_honorarios"] > 0 and remanente_b < remanente_a:
        estrategia = "Honorarios con Costos y Gastos (Menor impuesto que el 25% exento)"
        costos_cg = costos_b
        beneficios_40 = aceptados_b
        solicitados_40 = ded_comunes
        remanente_cg = remanente_b
    else:
        estrategia = "Renta Exenta del 25% Laboral (Óptima según Ley 2277)"
        costos_cg = costos_a
        beneficios_40 = aceptados_a
        solicitados_40 = ded_comunes + exenta_25
        remanente_cg = remanente_a

    # 4. Beneficios FUERA del límite del 40% (Ley 2277: 72 UVT Dep. + 1% Factura Electrónica)
    ded_dep72 = min(int(d["num_dependientes"]), L["max_dep_extra"]) * L["dep_extra_uvt"] * uvt if ing_trab > 0 else 0.0
    ded_fe1 = min(d["compras_factura_electronica"] * L["factura_elec_pct"], L["factura_elec_uvt"] * uvt)
    beneficios_fuera_40 = min(remanente_cg, ded_dep72 + ded_fe1)
    rlg_cg = redondear_dian(max(0.0, remanente_cg - beneficios_fuera_40))

    # 5. Pensiones y Ganancias Ocasionales (15%)
    neto_pen = max(0.0, d["ing_brutos_pensiones"] - d["salud_pensiones"])
    ex_pen = min(neto_pen, L["pensiones_exentas_uvt"] * uvt)
    rlg_pensiones = redondear_dian(max(0.0, neto_pen - ex_pen))

    gan_ocasional_gravable = redondear_dian(
        max(0.0, d["ing_ganancia_ocasional"] - d["costos_ganancia_ocasional"] - d["exenta_ganancia_ocasional"])
    )
    impuesto_ganancia_ocasional = redondear_dian(gan_ocasional_gravable * 0.15)

    # 6. Evaluación de los 6 Criterios Oficiales DIAN
    ing_brutos_totales = ing_brutos_cg + d["ing_brutos_pensiones"] + d["ing_ganancia_ocasional"]
    tope_pat_pesos = uvt_pat * uvt
    tope_ing_pesos = uvt_ing * uvt

    causales: List[str] = []
    if pat_bruto_raw >= tope_pat_pesos:
        causales.append(f"Patrimonio Bruto >= {uvt_pat:,.0f} UVT (${tope_pat_pesos:,.0f})")
    if ing_brutos_totales >= tope_ing_pesos:
        causales.append(f"Ingresos Brutos >= {uvt_ing:,.0f} UVT (${tope_ing_pesos:,.0f})")
    if d["consignaciones"] >= tope_ing_pesos:
        causales.append(f"Consignaciones >= {uvt_ing:,.0f} UVT (${tope_ing_pesos:,.0f})")
    if d["tarjetas_credito"] >= tope_ing_pesos:
        causales.append(f"Tarjetas de Crédito >= {uvt_ing:,.0f} UVT")
    if d["compras_totales"] >= tope_ing_pesos:
        causales.append(f"Compras Totales >= {uvt_ing:,.0f} UVT")
    if d["responsable_iva"]:
        causales.append("Responsable de IVA a 31 de diciembre")

    estado_obligacion = "Sí obligado a declarar" if causales else "No obligado a declarar"

    # 7. Impuesto Renta Art. 241, Anticipo Art. 807 y Sanción Extemporaneidad Art. 641
    rlg_total = rlg_cg + rlg_pensiones
    impuesto_renta = redondear_dian(calcular_impuesto_art241(rlg_total, uvt))
    total_impuesto_cargo = impuesto_renta + impuesto_ganancia_ocasional

    retenciones = redondear_dian(d["retenciones_anio"])
    saldo_fav_ant = redondear_dian(d["saldo_favor_anterior"])
    ant_ant = redondear_dian(d["anticipo_anio_anterior"])
    imp_ant = redondear_dian(d["impuesto_neto_anterior"])

    anio_n = int(d["anio_declara_num"])
    if anio_n == 0 or impuesto_renta == 0:
        anticipo_sig = 0
        ahorro_ant = 0
    else:
        pct = {1: 0.25, 2: 0.50}.get(anio_n, 0.75)
        op1 = max(0.0, (impuesto_renta * pct) - retenciones)
        if anio_n > 1:
            op2 = max(0.0, (((impuesto_renta + imp_ant) / 2.0) * pct) - retenciones)
            anticipo_sig = redondear_dian(min(op1, op2))
            ahorro_ant = redondear_dian(abs(op1 - op2))
        else:
            anticipo_sig = redondear_dian(op1)
            ahorro_ant = 0

    sancion_extemp = 0
    meses_mora = int(d["meses_extemporaneidad"])
    if meses_mora > 0:
        sancion_minima = redondear_dian(10 * uvt_sancion)
        if total_impuesto_cargo > 0:
            bruta = min(total_impuesto_cargo, total_impuesto_cargo * 0.05 * meses_mora)
        else:
            bruta = min(ing_brutos_totales * 0.05, ing_brutos_totales * 0.005 * meses_mora)
        if d["aplica_reduccion_640"]:
            bruta *= 0.50
        sancion_extemp = max(sancion_minima, redondear_dian(bruta))

    alivios_totales = retenciones + saldo_fav_ant + ant_ant
    balance = total_impuesto_cargo + anticipo_sig + sancion_extemp - alivios_totales
    saldo_pagar = max(0, balance)
    saldo_favor = abs(min(0, balance))

    # 8. Auditoría Patrimonial y Sugerencias de Ahorro
    alertas: List[str] = []
    pat_ant = redondear_dian(d["patrimonio_liquido_anterior"])
    if pat_ant > 0:
        delta_pat = patrimonio_liquido - pat_ant
        justificable = ing_netos_cg + neto_pen + gan_ocasional_gravable - imp_ant
        if delta_pat > justificable:
            alertas.append(
                f"Riesgo de Comparación Patrimonial (Art. 236 E.T.): Tu patrimonio líquido creció "
                f"${redondear_dian(delta_pat - justificable):,.0f} por encima de tus ingresos netos."
            )

    sugerencias: List[str] = []
    cupo_40 = max(0.0, limite_40 - solicitados_40)
    if cupo_40 > 300_000 and impuesto_renta > 0:
        sugerencias.append(
            f"Tienes un cupo disponible de ${redondear_dian(cupo_40):,.0f} en la bolsa del 40% / 1.340 UVT "
            f"(puedes incluir 50% del 4x1000, intereses de vivienda, medicina prepagada o aportes AFC/FPV)."
        )
    if d["ing_brutos_capital"] > 0 and d["incrngo_capital"] == 0:
        sugerencias.append(
            "Reportaste rendimientos de CDT: revisa en el certificado bancario el 'Componente Inflacionario' para restarlo como ingreso no gravado."
        )

    return {
        "obligado_bool": len(causales) > 0,
        "estado_obligacion": estado_obligacion,
        "causales_obligacion": causales,
        "estrategia": estrategia,
        "patrimonio_bruto": patrimonio_bruto,
        "deudas": deudas,
        "patrimonio_liquido": patrimonio_liquido,
        "ingresos_trabajo": redondear_dian(ing_trab),
        "ingresos_capital": redondear_dian(ing_cap),
        "ingresos_brutos_cg": redondear_dian(ing_brutos_cg),
        "incrngo_cg": redondear_dian(incrngo_cg),
        "ingresos_netos_cg": redondear_dian(ing_netos_cg),
        "costos_gastos_cg": redondear_dian(costos_cg),
        "beneficios_aceptados_40": redondear_dian(beneficios_40),
        "beneficios_fuera_40": redondear_dian(beneficios_fuera_40),
        "rlg_cedula_general": rlg_cg,
        "rlg_pensiones": rlg_pensiones,
        "renta_liquida_gravable_total": rlg_total,
        "rlg_en_uvt": round(rlg_total / uvt, 2) if uvt > 0 else 0,
        "ganancia_ocasional_gravable": gan_ocasional_gravable,
        "impuesto_renta": impuesto_renta,
        "impuesto_ganancia_ocasional": impuesto_ganancia_ocasional,
        "total_impuesto_cargo": total_impuesto_cargo,
        "retenciones": retenciones,
        "alivios_totales": alivios_totales,
        "anticipo_siguiente_anio": anticipo_sig,
        "ahorro_anticipo": ahorro_ant,
        "sancion_extemporaneidad": sancion_extemp,
        "saldo_a_pagar": saldo_pagar,
        "saldo_a_favor": saldo_favor,
        "alertas_auditoria": alertas,
        "sugerencias_optimizacion": sugerencias,
    }


# ==============================================================================
# 4. MIGRADOR MACRO .XLSM + GEMINI MULTIMODAL / OCR + AUDITOR EXÓGENA
# ==============================================================================
class CertificadoExtraido(BaseModel):
    tipo_documento: str = Field(description="Ej: Certificado 220, Extracto Bancario, Certificado CDT, Predial")
    entidad: str = Field(description="Banco o empresa emisora")
    efectivo_bancos_dic31: float = Field(default=0.0, description="Saldo en cuentas a 31 de diciembre")
    valor_inmueble_o_cdt_dic31: float = Field(default=0.0, description="Avalúo predial o capital CDT/inversión a dic 31")
    saldo_deuda_dic31: float = Field(default=0.0, description="Saldo capital de deudas/hipotecas a dic 31")
    ingresos_brutos_trabajo: float = Field(default=0.0, description="Total ingresos brutos laborales, servicios u honorarios")
    salud_pension_obligatoria: float = Field(default=0.0, description="Suma de salud obligatoria + pensión obligatoria")
    ingresos_capital_cdt: float = Field(default=0.0, description="Intereses o rendimientos financieros pagados")
    componente_inflacionario_cdt: float = Field(default=0.0, description="Parte no gravada / componente inflacionario CDT")
    intereses_vivienda: float = Field(default=0.0, description="Intereses crédito hipotecario o leasing")
    medicina_prepagada: float = Field(default=0.0, description="Pagos anuales medicina prepagada o póliza")
    gmf_4x1000_certificado: float = Field(default=0.0, description="100% del GMF 4x1000 certificado")
    aportes_afc_fpv: float = Field(default=0.0, description="Aportes voluntarios AFC o pensión voluntaria")
    retenciones_practicadas: float = Field(default=0.0, description="Retención en la fuente practicada en el año")


def _nvl(val) -> float:
    try:
        if val is None or isinstance(val, bool):
            return 0.0
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def importar_macro_xlsm(archivo_bytes: bytes) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(archivo_bytes), data_only=True)
    if "Ingreso_Datos" not in wb.sheetnames:
        raise ValueError("El archivo no contiene la hoja 'Ingreso_Datos'.")
    ws = wb["Ingreso_Datos"]
    iva_txt = str(ws["C11"].value or "").strip().upper()
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
        "ing_brutos_trabajo": _nvl(ws["C14"].value),
        "ing_brutos_capital": _nvl(ws["C15"].value),
        "retenciones_anio": _nvl(ws["C16"].value),
        "salud_pension_trabajo": _nvl(ws["C18"].value),
        "intereses_vivienda": _nvl(ws["C19"].value),
        "saldo_favor_anterior": _nvl(ws["C20"].value),
    }
    if "Parametros_UVT" in wb.sheetnames:
        uvt_leida = _nvl(wb["Parametros_UVT"]["B2"].value)
        if uvt_leida > 0:
            out["uvt_detectada"] = int(uvt_leida)
    return out


def extraer_con_gemini_hibrido(api_key: str, archivo_bytes: bytes, mime_type: str) -> dict:
    client = genai.Client(api_key=api_key)
    contenidos = [types.Part.from_bytes(data=archivo_bytes, mime_type=mime_type)]

    # Refuerzo OCR híbrido si es un PDF escaneado y pytesseract está activo
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
        "y pensión obligatoria en 'salud_pension_obligatoria'. No inventes cifras."
    )

    resp = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=contenidos,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=CertificadoExtraido,
            temperature=0.0,
        ),
    )
    return json.loads(resp.text)


def leer_exogena_excel(archivo_bytes: bytes) -> pd.DataFrame:
    xls = pd.ExcelFile(io.BytesIO(archivo_bytes))
    hoja = xls.sheet_names[0]
    for h in xls.sheet_names:
        if any(p in h.lower() for p in ["report", "tercer", "detalle", "informac"]):
            hoja = h
            break
    return pd.read_excel(xls, sheet_name=hoja).dropna(how="all")


def auditar_exogena_ia(api_key: str, df_exo: pd.DataFrame) -> str:
    client = genai.Client(api_key=api_key)
    csv_muestra = df_exo.head(180).to_csv(index=False)
    prompt = (
        "Actúa como Auditor Tributario DIAN Colombia. Analiza este reporte de Información Exógena y entrega:\n"
        "1. Totales sugeridos para Patrimonio (Bancos/Nequi/Daviplata, CDTs, Inmuebles, Pasivos).\n"
        "2. Totales sugeridos de Ingresos (Trabajo/Honorarios, Capital) y Retenciones practicadas.\n"
        "3. Revisión de topes de Consignaciones y Consumos con Tarjeta.\n"
        "4. Alertas críticas sobre movimientos que NO son ingreso real (ej. traslados entre cuentas propias o préstamos).\n\n"
        f"DATOS EXÓGENA:\n{csv_muestra}"
    )
    resp = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(temperature=0.1),
    )
    return resp.text


def consultar_estratega_rag(api_key: str, pregunta: str, res_liq: dict, uvt: float) -> str:
    client = genai.Client(api_key=api_key)
    instruccion = (
        f"Eres un Contador Público Tributarista Senior en Colombia experto en Formulario 210 y Ley 2277 de 2022. "
        f"UVT aplicada: ${uvt:,.0f} COP. Datos liquidados oficialmente por el motor determinista: "
        f"{json.dumps(res_liq, ensure_ascii=False)}. "
        "Responde de forma directa, estratégica y citando artículos del Estatuto Tributario. "
        "Respeta siempre las cifras exactas ya calculadas sin alterarlas."
    )
    resp = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=pregunta,
        config=types.GenerateContentConfig(system_instruction=instruccion, temperature=0.2),
    )
    return resp.text


# ==============================================================================
# 5. EXPORTADOR PROFESIONAL EN RAM (.XLSX MULTI-HOJA + .DOCX DICTAMEN)
# ==============================================================================
def exportar_excel_dian(res: dict, anio: int) -> bytes:
    filas_210 = [
        ("29", "Patrimonio bruto", res["patrimonio_bruto"]),
        ("30", "Deudas vigentes a 31 de diciembre", res["deudas"]),
        ("31", "Total patrimonio líquido (29 - 30)", res["patrimonio_liquido"]),
        ("32 / 43", "Ingresos brutos Trabajo / Servicios / Honorarios", res["ingresos_trabajo"]),
        ("58", "Ingresos brutos Rentas de Capital (CDT / Arriendos)", res["ingresos_capital"]),
        ("33 / 59", "Ingresos no constitutivos de renta (Salud, Pensión, Comp. Inflacionario)", res["incrngo_cg"]),
        ("44 / 60", "Costos y gastos procedentes aceptados", res["costos_gastos_cg"]),
        ("90", "Rentas exentas y deducciones aceptadas (Bolsa 40% / 1.340 UVT)", res["beneficios_aceptados_40"]),
        ("91", "Deducciones fuera del 40% Ley 2277 (72 UVT Dep. + 1% Factura Elec.)", res["beneficios_fuera_40"]),
        ("97", "Renta líquida gravable Cédula General", res["rlg_cedula_general"]),
        ("103", "Renta líquida gravable Cédula de Pensiones", res["rlg_pensiones"]),
        ("111", "Total Renta Líquida Gravable (Base Art. 241 E.T.)", res["renta_liquida_gravable_total"]),
        ("116", "Ganancia ocasional gravable (Venta activos >= 2 años)", res["ganancia_ocasional_gravable"]),
        ("127", "Impuesto neto de renta (Art. 241 E.T.)", res["impuesto_renta"]),
        ("128", "Impuesto de ganancias ocasionales (15%)", res["impuesto_ganancia_ocasional"]),
        ("133", "Total retenciones año gravable a declarar", res["retenciones"]),
        ("134", "Anticipo renta año siguiente (Optimizado Art. 807 E.T.)", res["anticipo_siguiente_anio"]),
        ("136", "Sanciones (Extemporaneidad Art. 641 / 640 E.T.)", res["sancion_extemporaneidad"]),
        ("137", "SALDO A PAGAR POR IMPUESTO", res["saldo_a_pagar"]),
        ("138", "TOTAL SALDO A FAVOR", res["saldo_a_favor"]),
    ]
    df_210 = pd.DataFrame(filas_210, columns=["Renglón 210", "Concepto Oficial DIAN", "Valor Redondeado (COP)"])

    balance_neto = res["saldo_a_pagar"] if res["saldo_a_pagar"] > 0 else -res["saldo_a_favor"]
    filas_dash = [
        ("Patrimonio Bruto", res["patrimonio_bruto"]),
        ("Deudas", res["deudas"]),
        ("Patrimonio Líquido", res["patrimonio_liquido"]),
        ("Ingresos de Trabajo / Honorarios", res["ingresos_trabajo"]),
        ("Ingresos de Capital (CDT / Rendimientos)", res["ingresos_capital"]),
        ("Exenciones y Deducciones Totales Aplicadas", res["beneficios_aceptados_40"] + res["beneficios_fuera_40"]),
        ("Renta Líquida Gravable", res["renta_liquida_gravable_total"]),
        ("Impuesto Total a Cargo", res["total_impuesto_cargo"]),
        ("Retenciones + Saldos a Favor Aplicados", res["alivios_totales"]),
        ("Balance Consolidado (+ Pagar / - Favor)", balance_neto),
        ("Estado Obligación DIAN", res["estado_obligacion"]),
    ]
    df_dash = pd.DataFrame(filas_dash, columns=["Indicador Dashboard", "Resultado"])

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_210.to_excel(writer, index=False, sheet_name=f"Form_210_AG{anio}")
        df_dash.to_excel(writer, index=False, sheet_name="Dashboard")
        ws1 = writer.sheets[f"Form_210_AG{anio}"]
        ws1.column_dimensions["A"].width = 15
        ws1.column_dimensions["B"].width = 68
        ws1.column_dimensions["C"].width = 25
        for c in ws1["C"][1:]:
            c.number_format = "$#,##0"
        ws2 = writer.sheets["Dashboard"]
        ws2.column_dimensions["A"].width = 50
        ws2.column_dimensions["B"].width = 28
        for c in ws2["B"][1:11]:
            c.number_format = "$#,##0"
    return buf.getvalue()


def exportar_dictamen_word(res: dict, anio: int, uvt: float) -> bytes:
    doc = Document()
    doc.add_heading(f"Dictamen Técnico de Declaración de Renta — AG {anio}", level=1)
    doc.add_paragraph(f"Parámetro UVT aplicado: ${uvt:,.0f} COP | Estrategia óptima: {res['estrategia']}")

    doc.add_heading("1. Verificación de Obligación Tributaria (6 Criterios DIAN)", level=2)
    doc.add_paragraph(f"Resultado: {res['estado_obligacion']}")
    for c in res["causales_obligacion"]:
        doc.add_paragraph(c, style="List Bullet")

    doc.add_heading("2. Liquidación Oficial Formulario 210", level=2)
    resumen = [
        f"Patrimonio Líquido: ${res['patrimonio_liquido']:,.0f} (Bruto: ${res['patrimonio_bruto']:,.0f} - Deudas: ${res['deudas']:,.0f})",
        f"Renta Líquida Gravable Total: ${res['renta_liquida_gravable_total']:,.0f} ({res['rlg_en_uvt']} UVT)",
        f"Impuesto Renta Art. 241 E.T.: ${res['impuesto_renta']:,.0f} | Ganancias Ocasionales: ${res['impuesto_ganancia_ocasional']:,.0f}",
        f"Retenciones y Saldos a Favor: ${res['alivios_totales']:,.0f}",
        f"Anticipo Año Siguiente (Art. 807): ${res['anticipo_siguiente_anio']:,.0f} (Ahorro logrado: ${res['ahorro_anticipo']:,.0f})",
        f"Sanción Extemporaneidad (si aplica): ${res['sancion_extemporaneidad']:,.0f}",
        f"SALDO A PAGAR: ${res['saldo_a_pagar']:,.0f} | TOTAL SALDO A FAVOR: ${res['saldo_a_favor']:,.0f}",
    ]
    for r in resumen:
        doc.add_paragraph(r, style="List Bullet")

    if res["alertas_auditoria"] or res["sugerencias_optimizacion"]:
        doc.add_heading("3. Auditoría Preventiva y Recomendaciones de Ahorro", level=2)
        for al in res["alertas_auditoria"]:
            doc.add_paragraph(f"[ALERTA] {al}", style="List Bullet")
        for sg in res["sugerencias_optimizacion"]:
            doc.add_paragraph(f"[OPORTUNIDAD] {sg}", style="List Bullet")

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
        api_key = st.text_input("🔑 Gemini API Key", type="password")

    anio_gravable = st.selectbox("📅 Año Gravable a Declarar", options=[2025, 2024, 2023, 2026], index=0)
    uvt_base = st.session_state.get("uvt_migrada", UVT_OFICIAL.get(anio_gravable, 49_799))
    valor_uvt = st.number_input("Valor UVT AG ($ COP)", min_value=10_000, value=int(uvt_base), step=100)

    with st.expander("Ajustar Topes UVT y Sanción"):
        uvt_pat = st.number_input("Tope Patrimonio (UVT)", value=4_500, step=100)
        uvt_ing = st.number_input("Tope Ingresos / Consumos (UVT)", value=1_400, step=100)
        uvt_sancion = st.number_input("UVT Año Presentación (Sanción Mínima)", value=UVT_OFICIAL[2026], step=100)

    d["anio_declara_num"] = st.selectbox(
        "Cálculo de Anticipo (Art. 807 E.T.)",
        options=[0, 1, 2, 3],
        index=int(d["anio_declara_num"]),
        format_func=lambda x: {
            0: "Sin Anticipo (Modo directo Macro)",
            1: "1er Año declarando (25%)",
            2: "2do Año declarando (50%)",
            3: "3er Año o más (75% Optimizado)",
        }[x],
    )

    # Panel Generador de Licencias (Solo visible si el usuario es Admin y configuró CLAVE_MAESTRA)
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
            st.session_state.ultimo_certificado = None
            st.session_state.auditoria_exogena = None
            st.session_state.chat_mensajes = []
            st.session_state.form_version += 1
            st.rerun()

res = liquidar_declaracion_210(d, valor_uvt, uvt_pat, uvt_ing, uvt_sancion)

# --- ENCABEZADO Y TABLERO KPI ---
st.title(f"📊 RentaIA Pro — Liquidador & Auditor Formulario 210 (AG {anio_gravable})")

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Estado DIAN", res["estado_obligacion"])
m2.metric("Patrimonio Líquido", f"${res['patrimonio_liquido']:,.0f}")
m3.metric("Renta Líquida Gravable", f"${res['renta_liquida_gravable_total']:,.0f}", f"{res['rlg_en_uvt']} UVT")
m4.metric("Impuesto a Cargo", f"${res['total_impuesto_cargo']:,.0f}")
if res["saldo_a_favor"] > 0:
    m5.metric("🎉 SALDO A FAVOR", f"${res['saldo_a_favor']:,.0f}", delta="Devolución / Compensación")
else:
    m5.metric("💳 NETO A PAGAR", f"${res['saldo_a_pagar']:,.0f}")

# --- PESTAÑAS PRINCIPALES ---
t1, t2, t3, t4, t5 = st.tabs([
    "🚀 1. Importar Macro / Soportes IA / Exógena",
    "🏠 2. Patrimonio y 6 Topes DIAN",
    "💼 3. Ingresos, Honorarios y Beneficios",
    "📋 4. Formulario 210 & Descargas",
    "⚖️ 5. Consultor Tributario IA",
])

# PESTAÑA 1: MIGRADOR + LECTOR GEMINI + EXÓGENA
with t1:
    col_izq, col_der = st.columns(2)
    with col_izq:
        st.subheader("📥 A. Migrar tu Excel `.xlsm` en 1 Clic")
        up_macro = st.file_uploader("Sube tu 'PROGRAMA PARA DECLARAR RENTA.xlsm'", type=["xlsm", "xlsx"])
        if up_macro and st.button("⚡ Importar Hoja 'Ingreso_Datos' (C4:C20)", use_container_width=True):
            try:
                migrados = importar_macro_xlsm(up_macro.getvalue())
                if "uvt_detectada" in migrados:
                    st.session_state["uvt_migrada"] = migrados.pop("uvt_detectada")
                st.session_state.datos.update(migrados)
                st.session_state.form_version += 1
                st.rerun()
            except Exception as e:
                st.error(f"Error al importar el archivo Excel: {e}")

        st.divider()
        st.subheader("📄 B. Lector Multimodal + OCR Híbrido (PDF / Foto)")
        up_doc = st.file_uploader("Sube Certificado 220, CDT, Extracto Bancario o Prepagada", type=["pdf", "png", "jpg", "jpeg"])
        if up_doc and st.button("🔍 Extraer cifras con Gemini", type="primary", use_container_width=True):
            if not api_key:
                st.error("Configura tu Gemini API Key en Secrets o en la barra lateral.")
            else:
                with st.spinner("Leyendo soporte en memoria RAM..."):
                    st.session_state.ultimo_certificado = extraer_con_gemini_hibrido(
                        api_key, up_doc.getvalue(), up_doc.type
                    )

        if st.session_state.ultimo_certificado:
            ext = st.session_state.ultimo_certificado
            st.json(ext)
            if st.button("➕ Sumar valores extraídos a mi Declaración", use_container_width=True):
                d["efectivo"] += ext.get("efectivo_bancos_dic31", 0.0)
                d["otros_activos"] += ext.get("valor_inmueble_o_cdt_dic31", 0.0)
                d["deudas"] += ext.get("saldo_deuda_dic31", 0.0)
                d["ing_brutos_trabajo"] += ext.get("ingresos_brutos_trabajo", 0.0)
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

    with col_der:
        st.subheader("📊 C. Auditor de Información Exógena DIAN")
        up_exo = st.file_uploader("Sube el Excel de Información Exógena descargado de la DIAN", type=["xlsx", "xls"])
        if up_exo:
            df_exo = leer_exogena_excel(up_exo.getvalue())
            st.caption(f"Registros reportados por terceros: **{len(df_exo)}**")
            st.dataframe(df_exo, use_container_width=True, height=250)
            if st.button("🤖 Auditar y Clasificar Exógena con Gemini", use_container_width=True):
                if not api_key:
                    st.error("Configura tu Gemini API Key primero.")
                else:
                    with st.spinner("Auditando terceros, topes y posibles traslados propios..."):
                        st.session_state.auditoria_exogena = auditar_exogena_ia(api_key, df_exo)
        if st.session_state.auditoria_exogena:
            st.info(st.session_state.auditoria_exogena)

# PESTAÑA 2: PATRIMONIO Y LOS 6 CRITERIOS DIAN
with t2:
    p1, p2 = st.columns(2)
    with p1:
        st.subheader("Activos y Pasivos a 31 de Diciembre")
        d["efectivo"] = st.number_input("Efectivo, Bancos, Nequi, Daviplata (C4)", min_value=0.0, value=float(d["efectivo"]), step=100_000.0, key=f"ef_{v}")
        d["vehiculos"] = st.number_input("Vehículos y Motocicletas (C5)", min_value=0.0, value=float(d["vehiculos"]), step=500_000.0, key=f"ve_{v}")
        d["inmuebles"] = st.number_input("Bienes Raíces: Lotes, Casas, Apartamentos (C6)", min_value=0.0, value=float(d["inmuebles"]), step=1_000_000.0, key=f"inm_{v}")
        d["otros_activos"] = st.number_input("Otros Activos, CDTs, Acciones (C7)", min_value=0.0, value=float(d["otros_activos"]), step=200_000.0, key=f"ot_{v}")
        d["deudas"] = st.number_input("Deudas y Obligaciones Vigentes (C12)", min_value=0.0, value=float(d["deudas"]), step=200_000.0, key=f"deu_{v}")
    with p2:
        st.subheader("Topes DIAN (C8 a C11) y Comparación Patrimonial")
        d["consignaciones"] = st.number_input("Consignaciones Bancarias Acumuladas (C8)", min_value=0.0, value=float(d["consignaciones"]), step=500_000.0, key=f"con_{v}")
        d["tarjetas_credito"] = st.number_input("Consumos con Tarjeta de Crédito (C9)", min_value=0.0, value=float(d["tarjetas_credito"]), step=500_000.0, key=f"tc_{v}")
        d["compras_totales"] = st.number_input("Compras y Consumos Totales (C10)", min_value=0.0, value=float(d["compras_totales"]), step=500_000.0, key=f"com_{v}")
        d["responsable_iva"] = st.toggle("¿Responsable de IVA a 31 de diciembre? (C11)", value=bool(d["responsable_iva"]), key=f"iva_{v}")
        st.divider()
        d["patrimonio_liquido_anterior"] = st.number_input("Patrimonio Líquido Año Anterior (Auditoría Art. 236 E.T.)", min_value=0.0, value=float(d["patrimonio_liquido_anterior"]), step=500_000.0, key=f"pat_ant_{v}")

# PESTAÑA 3: INGRESOS, DEDUCCIONES, GANANCIAS OCASIONALES Y SANCIONES
with t3:
    c1, c2, c3 = st.columns(3)
    with c1:
        st.subheader("Trabajo, Honorarios y Pensiones")
        d["ing_brutos_trabajo"] = st.number_input("Ingresos Trabajo / Servicios / Honorarios (C14)", min_value=0.0, value=float(d["ing_brutos_trabajo"]), step=200_000.0, key=f"itr_{v}")
        d["salud_pension_trabajo"] = st.number_input("Aporte Obligatorio Salud + Pensión (C18)", min_value=0.0, value=float(d["salud_pension_trabajo"]), step=50_000.0, key=f"sp_{v}")
        d["costos_honorarios"] = st.number_input("Costos/Gastos de Honorarios (Simula vs 25% exento)", min_value=0.0, value=float(d["costos_honorarios"]), step=100_000.0, key=f"chon_{v}")
        st.divider()
        d["ing_brutos_pensiones"] = st.number_input("Ingresos Brutos por Pensiones", min_value=0.0, value=float(d["ing_brutos_pensiones"]), step=200_000.0, key=f"ipen_{v}")
        d["salud_pensiones"] = st.number_input("Salud Obligatoria en Pensión", min_value=0.0, value=float(d["salud_pensiones"]), step=50_000.0, key=f"spen_{v}")

    with c2:
        st.subheader("Capital, No Laborales y Venta Activos")
        d["ing_brutos_capital"] = st.number_input("Ingresos Capital: CDT / Arriendos (C15)", min_value=0.0, value=float(d["ing_brutos_capital"]), step=100_000.0, key=f"icap_{v}")
        d["incrngo_capital"] = st.number_input("Componente Inflacionario CDT (No Gravado)", min_value=0.0, value=float(d["incrngo_capital"]), step=50_000.0, key=f"inccap_{v}")
        d["costos_capital"] = st.number_input("Costos y Gastos Rentas de Capital", min_value=0.0, value=float(d["costos_capital"]), step=50_000.0, key=f"ccap_{v}")
        d["ing_brutos_no_laborales"] = st.number_input("Ingresos Brutos No Laborales", min_value=0.0, value=float(d["ing_brutos_no_laborales"]), step=200_000.0, key=f"inol_{v}")
        d["costos_no_laborales"] = st.number_input("Costos y Gastos No Laborales", min_value=0.0, value=float(d["costos_no_laborales"]), step=100_000.0, key=f"cnol_{v}")
        st.divider()
        with st.expander("🏡 Ganancias Ocasionales (Venta Lotes/Inmuebles >= 2 años)"):
            d["ing_ganancia_ocasional"] = st.number_input("Precio de venta activo fijo / herencia", min_value=0.0, value=float(d["ing_ganancia_ocasional"]), step=1_000_000.0, key=f"igo_{v}")
            d["costos_ganancia_ocasional"] = st.number_input("Costo fiscal del activo vendido", min_value=0.0, value=float(d["costos_ganancia_ocasional"]), step=1_000_000.0, key=f"cgo_{v}")
            d["exenta_ganancia_ocasional"] = st.number_input("Ganancia ocasional exenta", min_value=0.0, value=float(d["exenta_ganancia_ocasional"]), step=500_000.0, key=f"ego_{v}")

    with c3:
        st.subheader("Deducciones Ley 2277 y Retenciones")
        d["intereses_vivienda"] = st.number_input("Intereses Crédito Vivienda (C19)", min_value=0.0, value=float(d["intereses_vivienda"]), step=100_000.0, key=f"viv_{v}")
        d["num_dependientes"] = st.number_input("Número de Dependientes Económicos", min_value=0, max_value=10, value=int(d["num_dependientes"]), step=1, key=f"dep_{v}")
        d["compras_factura_electronica"] = st.number_input("Compras con Factura Electrónica + Pago Digital", min_value=0.0, value=float(d["compras_factura_electronica"]), step=200_000.0, key=f"fe_{v}")
        d["medicina_prepagada"] = st.number_input("Medicina Prepagada / Seguros de Salud", min_value=0.0, value=float(d["medicina_prepagada"]), step=100_000.0, key=f"prep_{v}")
        d["gmf_certificado_100"] = st.number_input("GMF 4x1000 (Ingresa el 100% certificado)", min_value=0.0, value=float(d["gmf_certificado_100"]), step=20_000.0, key=f"gmf_{v}")
        d["aportes_afc_fpv"] = st.number_input("Aportes Voluntarios AFC / FPV", min_value=0.0, value=float(d["aportes_afc_fpv"]), step=100_000.0, key=f"afc_{v}")
        st.divider()
        d["retenciones_anio"] = st.number_input("Retenciones en la Fuente del Año (C16)", min_value=0.0, value=float(d["retenciones_anio"]), step=50_000.0, key=f"ret_{v}")
        d["saldo_favor_anterior"] = st.number_input("Saldo a Favor Año Anterior (C20)", min_value=0.0, value=float(d["saldo_favor_anterior"]), step=50_000.0, key=f"sfa_{v}")
        with st.expander("⏱️ Extemporaneidad (Art. 641) y Datos Año Anterior"):
            d["meses_extemporaneidad"] = st.number_input("Meses o fracción de retraso (0 = En plazo)", min_value=0, max_value=60, value=int(d["meses_extemporaneidad"]), step=1, key=f"mora_{v}")
            d["aplica_reduccion_640"] = st.checkbox("Aplicar reducción al 50% (Art. 640 E.T.)", value=bool(d["aplica_reduccion_640"]), key=f"r640_{v}")
            d["impuesto_neto_anterior"] = st.number_input("Impuesto Neto Año Anterior", min_value=0.0, value=float(d["impuesto_neto_anterior"]), step=50_000.0, key=f"ima_{v}")
            d["anticipo_anio_anterior"] = st.number_input("Anticipo Liquidado Año Anterior", min_value=0.0, value=float(d["anticipo_anio_anterior"]), step=50_000.0, key=f"ana_{v}")

    if st.button("🔄 Actualizar Liquidación y Tablero", type="primary"):
        st.rerun()

# PESTAÑA 4: DASHBOARD CONSOLIDADO + RENGLONES 210 + EXPORTACIÓN
with t4:
    if res["saldo_a_favor"] > 0:
        st.success(
            f"### 🎉 ¡TIENES UN SALDO A FAVOR DE: ${res['saldo_a_favor']:,.0f}!\n"
            f"- **Estado DIAN:** {res['estado_obligacion']} | **Ruta Elegida:** {res['estrategia']}\n"
            f"- **Impuesto Total:** ${res['total_impuesto_cargo']:,.0f} | **Retenciones + Alivios:** ${res['alivios_totales']:,.0f}"
        )
    elif res["saldo_a_pagar"] > 0:
        st.warning(
            f"### 💳 VALOR NETO A PAGAR: ${res['saldo_a_pagar']:,.0f}\n"
            f"- **Estado DIAN:** {res['estado_obligacion']} | **Ruta Elegida:** {res['estrategia']}\n"
            f"- **Impuesto Total:** ${res['total_impuesto_cargo']:,.0f} | **Retenciones / Alivios:** ${res['alivios_totales']:,.0f}"
        )
    else:
        st.info(f"### ✅ El valor neto a pagar es $0 — Estado: {res['estado_obligacion']}")

    if res["causales_obligacion"]:
        st.caption("**Causales activadas:** " + " • ".join(res["causales_obligacion"]))
    for al in res["alertas_auditoria"]:
        st.error(f"🚨 {al}")
    for sg in res["sugerencias_optimizacion"]:
        st.info(f"💡 **Oportunidad de Ahorro:** {sg}")

    st.subheader("Borrador Casilla por Casilla — Formulario 210 DIAN")
    st.table({
        "Renglón Oficial Formulario 210": [
            "29. Patrimonio Bruto",
            "30. Deudas",
            "31. Total Patrimonio Líquido",
            "32/43. Ingresos Brutos Trabajo / Servicios / Honorarios",
            "58. Ingresos Brutos Capital (CDT / Rendimientos)",
            "33/59. Ingresos No Constitutivos de Renta (Salud, Pensión, Comp. Inflacionario)",
            "44/60. Costos y Gastos Procedentes Aplicados",
            "90. Exenciones y Deducciones Aceptadas (Tope 40% / 1.340 UVT)",
            "91. Deducciones Fuera del 40% Ley 2277 (72 UVT Dep. + 1% Factura Elec.)",
            "111. Total Renta Líquida Gravable",
            "127. Impuesto Neto de Renta (Art. 241 E.T.)",
            "128. Impuesto de Ganancias Ocasionales (15%)",
            "133. Retenciones + Saldos a Favor Aplicados",
            "134. Anticipo Año Siguiente (Optimizado Art. 807 E.T.)",
            "136. Sanción por Extemporaneidad (Art. 641 / 640 E.T.)",
            "137. SALDO A PAGAR",
            "138. TOTAL SALDO A FAVOR",
        ],
        "Valor Redondeado DIAN (COP)": [
            f"${res['patrimonio_bruto']:,.0f}",
            f"${res['deudas']:,.0f}",
            f"${res['patrimonio_liquido']:,.0f}",
            f"${res['ingresos_trabajo']:,.0f}",
            f"${res['ingresos_capital']:,.0f}",
            f"${res['incrngo_cg']:,.0f}",
            f"${res['costos_gastos_cg']:,.0f}",
            f"${res['beneficios_aceptados_40']:,.0f}",
            f"${res['beneficios_fuera_40']:,.0f}",
            f"${res['renta_liquida_gravable_total']:,.0f}",
            f"${res['impuesto_renta']:,.0f}",
            f"${res['impuesto_ganancia_ocasional']:,.0f}",
            f"${res['alivios_totales']:,.0f}",
            f"${res['anticipo_siguiente_anio']:,.0f}",
            f"${res['sancion_extemporaneidad']:,.0f}",
            f"${res['saldo_a_pagar']:,.0f}",
            f"${res['saldo_a_favor']:,.0f}",
        ],
    })

    xlsx_bytes = exportar_excel_dian(res, anio_gravable)
    docx_bytes = exportar_dictamen_word(res, anio_gravable, valor_uvt)

    b1, b2 = st.columns(2)
    b1.download_button(
        "📥 Descargar Excel Oficial (Formulario 210 + Hoja Dashboard)",
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

# PESTAÑA 5: CONSULTOR TRIBUTARIO IA (RAG)
with t5:
    st.subheader("⚖️ Estratega Tributario con Gemini (Conectado a tus Cifras)")
    for msj in st.session_state.chat_mensajes:
        with st.chat_message(msj["role"]):
            st.markdown(msj["content"])

    consulta = st.chat_input("Ej: ¿Me conviene más declarar costos en honorarios o el 25% exento? ¿Cómo justifico mi patrimonio?")
    if consulta:
        if not api_key:
            st.error("Configura tu Gemini API Key primero.")
        else:
            st.session_state.chat_mensajes.append({"role": "user", "content": consulta})
            with st.chat_message("user"):
                st.markdown(consulta)
            with st.chat_message("assistant"):
                with st.spinner("Analizando Estatuto Tributario sobre tu declaración..."):
                    r_rag = consultar_estratega_rag(api_key, consulta, res, valor_uvt)
                    st.markdown(r_rag)
                    st.session_state.chat_mensajes.append({"role": "assistant", "content": r_rag})