import io
import os
import re
import time
import json
import hmac
import hashlib
import datetime
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, Any, List, Tuple

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill
import pypdf
import streamlit as st
from docx import Document
from google import genai
from google.genai import types
from google.genai.errors import APIError

try:
    from PIL import Image
    PIL_DISPONIBLE = True
except Exception:
    PIL_DISPONIBLE = False

try:
    from pdf2image import convert_from_bytes
    PDF2IMG_DISPONIBLE = True
except Exception:
    PDF2IMG_DISPONIBLE = False

try:
    import pytesseract
    TESSERACT_DISPONIBLE = True
except Exception:
    TESSERACT_DISPONIBLE = False

OCR_DISPONIBLE = PIL_DISPONIBLE and TESSERACT_DISPONIBLE

st.set_page_config(
    page_title="RentaIA Pro | Formulario 210 DIAN",
    page_icon="📊",
    layout="wide",
)

# ==============================================================================
# 1. PARÁMETROS NORMATIVOS DIAN MULTIAÑO (FUNCIONA PARA CUALQUIER AÑO FUTURO)
# ==============================================================================
UVT_OFICIAL = {2022: 38_004, 2023: 42_412, 2024: 47_065, 2025: 49_799, 2026: 52_374}
COMP_INFLACIONARIO = {2023: 0.6671, 2024: 0.5088, 2025: 0.5544, 2026: 0.5544}

SECCIONALES_DIAN = [
    "32 - Bogotá D.C.", "11 - Medellín", "05 - Cali", "02 - Barranquilla", "04 - Bucaramanga",
    "06 - Cartagena", "09 - Ibagué", "10 - Manizales", "14 - Pasto", "16 - Pereira",
    "20 - Tunja", "26 - Villavicencio", "43 - Yopal", "13 - Neiva", "07 - Cúcuta",
    "19 - Santa Marta", "01 - Armenia", "12 - Montería", "15 - Popayán", "21 - Valledupar", "00 - Otra",
]

ACTIVIDADES_CIIU = [
    "0010 - Asalariado / Empleado",
    "8211 - Servicios administrativos y de apoyo de oficina",
    "0090 - Rentista de capital (CDTs, intereses, arriendos)",
    "0020 - Pensionado",
    "7020 - Consultoría de gestión / administración",
    "6910 - Actividades jurídicas",
    "6920 - Contabilidad, auditoría y asesoría tributaria",
    "6201 - Desarrollo de sistemas informáticos",
    "7112 - Actividades de ingeniería",
    "4711 - Comercio al por menor",
    "6810 - Actividades inmobiliarias / arrendamientos",
    "OTRO - Digitar otro código CIIU",
]

TABLA_ART_241 = [
    (0, 1_090, 0.00, 0),
    (1_090, 1_700, 0.19, 0),
    (1_700, 4_100, 0.28, 116),
    (4_100, 8_670, 0.33, 788),
    (8_670, 18_970, 0.35, 2_296),
    (18_970, 31_000, 0.37, 5_901),
    (31_000, float("inf"), 0.39, 10_352),
]

MODELOS_PREFERIDOS_GEMINI = [
    "gemini-3.5-flash",
    "gemini-3.1-flash-lite",
    "gemini-3.1-flash-lite-preview",
    "gemini-3.8-flash",
    "gemini-3.6-flash",
    "gemini-flash-latest",
]


def uvt_por_anio(anio: int) -> int:
    return UVT_OFICIAL.get(anio, UVT_OFICIAL[max(UVT_OFICIAL.keys())])


def redondear_dian(valor: float) -> int:
    if valor <= 0:
        return 0
    return int((Decimal(str(valor)) / Decimal("1000")).quantize(Decimal("1"), rounding=ROUND_HALF_UP) * 1000)


def calcular_dv(nit_str: str) -> str:
    limpio = "".join(ch for ch in str(nit_str) if ch.isdigit())
    if not limpio:
        return ""
    primos = [3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71]
    suma = sum(int(d) * primos[i] for i, d in enumerate(reversed(limpio)) if i < len(primos))
    res = suma % 11
    return str(res if res in (0, 1) else 11 - res)


def normalizar_entidad(texto: str) -> str:
    t = (texto or "").lower()
    t = t.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    if any(k in t for k in ("bbva", "bilbao", "860003020")):
        return "BBVA Colombia"
    if any(k in t for k in ("bogota", "860002964")):
        return "Banco de Bogotá"
    if any(k in t for k in ("bancolombia", "890903938")):
        return "Bancolombia"
    if any(k in t for k in ("davivienda", "860034313")):
        return "Davivienda"
    if any(k in t for k in ("colpatria", "scotiabank", "860034594")):
        return "Scotiabank Colpatria"
    if any(k in t for k in ("occidente", "890300279")):
        return "Banco de Occidente"
    if any(k in t for k in ("popular", "860007738")):
        return "Banco Popular"
    if any(k in t for k in ("av villas", "860035827")):
        return "Banco AV Villas"
    if "nequi" in t:
        return "Nequi"
    if "daviplata" in t:
        return "Daviplata"
    limpio = re.sub(r"\s+", " ", (texto or "Otra Entidad")).strip()
    return limpio or "Otra Entidad"


# ==============================================================================
# 2. SEGURIDAD Y CONEXIÓN GEMINI INTELIGENTE ('AIza...' + 'AQ....' + CACHÉ MODELO)
# ==============================================================================
def limpiar_key(raw: str) -> str:
    if not raw:
        return ""
    txt = str(raw).strip()
    m_aiza = re.search(r"(AIza[0-9A-Za-z_\-]{25,})", txt)
    if m_aiza:
        return m_aiza.group(1)
    m_aq = re.search(r"(AQ\.[0-9A-Za-z_\-\.]{20,})", txt)
    if m_aq:
        return m_aq.group(1)
    return ""


def obtener_secreto(nombre: str) -> str:
    try:
        val = str(st.secrets.get(nombre, "")).strip()
    except Exception:
        val = os.getenv(nombre, "").strip()
    return limpiar_key(val) if nombre == "GEMINI_API_KEY" else val


def generar_licencia(identificador: str, semilla: str) -> str:
    h = hmac.new(
        semilla.encode("utf-8"),
        identificador.strip().upper().encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()
    return f"RENTA-{h[:12].upper()}"


def verificar_acceso() -> bool:
    clave_m = obtener_secreto("CLAVE_MAESTRA")
    if not clave_m:
        st.session_state["es_admin"] = True
        return True
    if st.session_state.get("autenticado", False):
        return True

    st.title("🔐 RentaIA Pro — Acceso Protegido")
    c1, c2 = st.columns(2)
    doc_u = c1.text_input("Documento / Correo del Cliente")
    cod_u = c2.text_input("Clave Maestra o Licencia", type="password")
    if st.button("Ingresar", type="primary"):
        if hmac.compare_digest(cod_u.strip(), clave_m):
            st.session_state["autenticado"] = True
            st.session_state["es_admin"] = True
            st.rerun()
        elif doc_u.strip() and hmac.compare_digest(cod_u.strip().upper(), generar_licencia(doc_u, clave_m)):
            st.session_state["autenticado"] = True
            st.session_state["es_admin"] = False
            st.rerun()
        else:
            st.error("Credenciales inválidas.")
    return False


def crear_clientes_gemini(llave: str) -> List[Tuple[str, Any]]:
    clientes: List[Tuple[str, Any]] = []
    try:
        clientes.append(("AIStudio", genai.Client(api_key=llave)))
    except Exception:
        pass
    if llave.startswith("AQ."):
        try:
            clientes.append(("VertexAI", genai.Client(vertexai=True, api_key=llave)))
        except Exception:
            pass
    return clientes


def _descubrir_modelos_activos(cli: Any, ya_probados: List[str]) -> List[str]:
    excluidos = ("tts", "live", "image", "embed", "veo", "lyria", "robotics", "aqa", "imagen", "audio")
    hallados: List[str] = []
    try:
        for m in cli.models.list():
            nom = str(getattr(m, "name", "") or "").replace("models/", "").strip()
            acciones = getattr(m, "supported_actions", None) or []
            if not nom or nom in ya_probados:
                continue
            if "gemini" not in nom.lower() or any(x in nom.lower() for x in excluidos):
                continue
            if acciones and "generateContent" not in acciones:
                continue
            hallados.append(nom)
    except Exception:
        return []
    hallados.sort(key=lambda x: (0 if "flash" in x.lower() else 1, x))
    return hallados[:4]


def _construir_config_gemini(modelo: str, system_instruction: str = None, temperatura: float = 0.0) -> types.GenerateContentConfig:
    cfg_kwargs: Dict[str, Any] = {}
    if system_instruction:
        cfg_kwargs["system_instruction"] = system_instruction
    es_serie_3 = modelo.startswith("gemini-3") or modelo.endswith("-latest")
    if not es_serie_3 and temperatura is not None:
        cfg_kwargs["temperature"] = temperatura
    return types.GenerateContentConfig(**cfg_kwargs)


def llamar_gemini(
    api_key: str,
    contents: Any,
    system_instruction: str = None,
    temperatura: float = 0.0,
) -> str:
    llave = limpiar_key(api_key)
    if not llave:
        raise ValueError(
            "Falta una clave válida de Gemini ('AIza...' o 'AQ....') en Configuración -> Secrets o en la barra lateral."
        )

    clientes = crear_clientes_gemini(llave)
    if not clientes:
        raise RuntimeError("No se pudo inicializar el cliente de Gemini. Revisa la clave configurada.")

    ultimo_err = None
    errores_modelos = []
    modelo_cache = st.session_state.get("modelo_gemini_ok")

    for nombre_cli, cli in clientes:
        abortar_este_cliente = False
        modelos_cola = []
        if modelo_cache:
            modelos_cola.append(modelo_cache)
        for m_pref in MODELOS_PREFERIDOS_GEMINI:
            if m_pref not in modelos_cola:
                modelos_cola.append(m_pref)

        descubrimiento_hecho = False
        idx_m = 0

        while idx_m < len(modelos_cola):
            if abortar_este_cliente:
                break
            mod = modelos_cola[idx_m]
            idx_m += 1
            config = _construir_config_gemini(mod, system_instruction=system_instruction, temperatura=temperatura)

            for intento in range(3):
                try:
                    respuesta = cli.models.generate_content(model=mod, contents=contents, config=config)
                    texto = getattr(respuesta, "text", None)
                    if not texto or not str(texto).strip():
                        raise RuntimeError(f"Gemini devolvió respuesta vacía en {mod}.")
                    st.session_state["modelo_gemini_ok"] = mod
                    return str(texto)
                except APIError as e:
                    ultimo_err = e
                    codigo = getattr(e, "code", None) or getattr(e, "status_code", None)
                    detalle = str(e).lower()

                    if codigo == 429:
                        indicadores_cuota = (
                            "quota_exceeded", "daily limit", "limit: 0", "billing", "generaterequestsperday"
                        )
                        if any(p in detalle for p in indicadores_cuota):
                            errores_modelos.append(f"{nombre_cli}/{mod}: cuota agotada")
                            break
                        if intento < 2:
                            time.sleep((2, 5)[intento])
                            continue
                        errores_modelos.append(f"{nombre_cli}/{mod}: HTTP 429 (límite por minuto)")
                        break

                    if codigo in (401, 403):
                        errores_modelos.append(f"{nombre_cli}/{mod}: HTTP {codigo} (permisos/superficie)")
                        abortar_este_cliente = True
                        break

                    if codigo == 404:
                        errores_modelos.append(f"{nombre_cli}/{mod}: HTTP 404 (no disponible)")
                        if idx_m == len(modelos_cola) and not descubrimiento_hecho:
                            descubrimiento_hecho = True
                            nuevos = _descubrir_modelos_activos(cli, modelos_cola)
                            modelos_cola.extend(nuevos)
                        break

                    if codigo == 400:
                        errores_modelos.append(f"{nombre_cli}/{mod}: HTTP 400 ({str(e)[:120]})")
                        break

                    if codigo is not None and codigo >= 500 and intento < 2:
                        time.sleep(2 * (intento + 1))
                        continue

                    errores_modelos.append(f"{nombre_cli}/{mod}: HTTP {codigo} ({str(e)[:120]})")
                    break
                except (TimeoutError, ConnectionError) as e:
                    ultimo_err = e
                    if intento < 2:
                        time.sleep(2 * (intento + 1))
                        continue
                    errores_modelos.append(f"{nombre_cli}/{mod}: error de red ({e})")
                    break
                except Exception as e:
                    ultimo_err = e
                    errores_modelos.append(f"{nombre_cli}/{mod}: {type(e).__name__}: {str(e)[:120]}")
                    break

    resumen_err = " | ".join(errores_modelos[-6:]) or str(ultimo_err)
    raise RuntimeError(f"No se pudo completar la consulta a Gemini. Detalle técnico: {resumen_err}") from ultimo_err


# ==============================================================================
# 3. MOTOR MATEMÁTICO OFICIAL FORMULARIO 210 (LEY 2277 + MUISCA DIAN)
# ==============================================================================
def datos_vacios() -> Dict[str, Any]:
    return {
        "nombre_declarante": "", "nit_declarante": "", "seccional_dian": "32 - Bogotá D.C.", "cod_actividad": "8211",
        "efectivo": 0.0, "otros_activos": 0.0, "inmuebles": 0.0, "vehiculos": 0.0, "deudas": 0.0,
        "consignaciones": 0.0, "tarjetas_credito": 0.0, "compras_totales": 0.0, "responsable_iva": False,
        "patrimonio_liquido_anterior": 0.0,
        "ing_brutos_trabajo": 0.0, "salud_pension_trabajo": 0.0, "cesantias_exentas": 0.0,
        "ing_brutos_honorarios": 0.0, "salud_pension_honorarios": 0.0, "arl_honorarios": 0.0, "costos_honorarios": 0.0,
        "ing_brutos_capital": 0.0, "incrngo_capital": 0.0, "auto_comp_inflacionario": True, "costos_capital": 0.0,
        "ing_brutos_no_laborales": 0.0, "incrngo_no_laborales": 0.0, "costos_no_laborales": 0.0,
        "ing_brutos_pensiones": 0.0, "salud_pensiones": 0.0, "dividendos_gravables": 0.0,
        "ing_ganancia_ocasional": 0.0, "ing_loterias_rifas": 0.0, "costos_ganancia_ocasional": 0.0, "exenta_ganancia_ocasional": 0.0,
        "num_dependientes": 0, "compras_factura_electronica": 0.0, "aplicar_beneficio_fe": False,
        "intereses_vivienda": 0.0, "medicina_prepagada": 0.0, "intereses_icetex": 0.0,
        "gmf_certificado_100": 0.0, "aportes_afc_fpv": 0.0, "donaciones_entidades": 0.0,
        "retenciones_anio": 0.0, "saldo_favor_anterior": 0.0, "anio_declara_num": 0,
        "impuesto_neto_anterior": 0.0, "anticipo_anio_anterior": 0.0, "meses_extemporaneidad": 0, "aplica_reduccion_640": False,
    }


def impuesto_art241(rlg: float, uvt: float) -> float:
    if rlg <= 0 or uvt <= 0:
        return 0.0
    u = rlg / uvt
    for d_u, h_u, tar, base in TABLA_ART_241:
        if d_u < u <= h_u:
            return (base + ((u - d_u) * tar)) * uvt
    return 0.0


def liquidar_210(
    d: Dict[str, Any], uvt: float, uvt_pat: float, uvt_ing: float, uvt_san: float, anio_grav: int = 2025
) -> Dict[str, Any]:
    pat_raw = d["efectivo"] + d["otros_activos"] + d["inmuebles"] + d["vehiculos"]
    c29 = redondear_dian(pat_raw)
    c30 = redondear_dian(d["deudas"])
    c31 = max(0, c29 - c30)

    pct_inf = COMP_INFLACIONARIO.get(anio_grav, 0.5544)
    c58 = redondear_dian(d["ing_brutos_capital"])
    if d.get("auto_comp_inflacionario", True) and d["ing_brutos_capital"] > 0 and d["costos_capital"] == 0:
        inc_cap_base = redondear_dian(d["ing_brutos_capital"] * pct_inf)
    else:
        inc_cap_base = redondear_dian(d["incrngo_capital"])
    c59 = min(c58, inc_cap_base)
    neto_cap = max(0, c58 - c59)
    c60 = min(neto_cap, redondear_dian(d["costos_capital"]))
    c61 = max(0, neto_cap - c60)

    c74 = redondear_dian(d["ing_brutos_no_laborales"])
    c76 = min(c74, redondear_dian(d["incrngo_no_laborales"]))
    neto_nl = max(0, c74 - c76)
    c77 = min(neto_nl, redondear_dian(d["costos_no_laborales"]))
    c78 = max(0, neto_nl - c77)

    trab_bruto_tot = redondear_dian(d["ing_brutos_trabajo"] + d["ing_brutos_honorarios"])
    inc_trab_tot = min(
        trab_bruto_tot,
        redondear_dian(d["salud_pension_trabajo"] + d["salud_pension_honorarios"] + d["arl_honorarios"]),
    )
    ing_brutos_cg = trab_bruto_tot + c58 + c74
    ing_netos_cg = max(0, trab_bruto_tot - inc_trab_tot) + neto_cap + neto_nl
    hay_trab = trab_bruto_tot > 0

    c38 = redondear_dian(min(d["intereses_vivienda"], 1_200 * uvt))
    ded_prep = redondear_dian(min(d["medicina_prepagada"], 192 * uvt)) if hay_trab else 0
    ded_dep10 = redondear_dian(min(trab_bruto_tot * 0.10, 384 * uvt)) if (d["num_dependientes"] > 0 and hay_trab) else 0
    ded_ice = redondear_dian(min(d["intereses_icetex"], 100 * uvt))
    ded_gmf = redondear_dian(d["gmf_certificado_100"] * 0.50)
    c39 = ded_prep + ded_dep10 + ded_ice + ded_gmf
    c40 = c38 + c39
    c35 = redondear_dian(min(d["aportes_afc_fpv"], ing_brutos_cg * 0.30, 3_800 * uvt))
    ces_ex = redondear_dian(d["cesantias_exentas"])

    ded_comunes = c40 + c35 + ces_ex
    lim_40 = redondear_dian(min(ing_netos_cg * 0.40, 1_340 * uvt))
    tope_790 = 790 * uvt

    c34_a = max(0, trab_bruto_tot - inc_trab_tot)
    c91_a = c34_a + c61 + c78
    ex25_a = redondear_dian(min(max(0, c34_a - ded_comunes) * 0.25, tope_790))
    acep_a = min(ded_comunes + ex25_a, lim_40, c91_a)
    rem_a = max(0, c91_a - acep_a)

    cost_hon = redondear_dian(d["costos_honorarios"])
    c32_b = redondear_dian(d["ing_brutos_trabajo"])
    c33_b = min(c32_b, redondear_dian(d["salud_pension_trabajo"]))
    c34_b = max(0, c32_b - c33_b)
    c43_b = redondear_dian(d["ing_brutos_honorarios"])
    pila_hon = redondear_dian(d["salud_pension_honorarios"] + d["arl_honorarios"]) + max(
        0, redondear_dian(d["salud_pension_trabajo"]) - c33_b
    )
    c44_b = min(c43_b, pila_hon)
    neto_h_b = max(0, c43_b - c44_b)
    c45_b = min(neto_h_b, cost_hon)
    c46_b = max(0, neto_h_b - c45_b)
    c91_b = c34_b + c46_b + c61 + c78
    ex25_b = redondear_dian(min(max(0, c34_b - ded_comunes) * 0.25, tope_790))
    acep_b = min(ded_comunes + ex25_b, lim_40, c91_b)
    rem_b = max(0, c91_b - acep_b)

    if cost_hon > 0 and c43_b > 0 and rem_b < rem_a:
        estrategia = "Honorarios con Costos y Gastos (Casillas 43-46)"
        c32, c33, c34 = c32_b, c33_b, c34_b
        c43, c44, c45, c46 = c43_b, c44_b, c45_b, c46_b
        c91, c36, ben_40, sol_40, rem_cg = c91_b, ex25_b + ces_ex, acep_b, ded_comunes + ex25_b, rem_b
    else:
        estrategia = "Rentas de Trabajo con 25% Exento (Casillas 32-36)"
        c32, c33, c34 = trab_bruto_tot, inc_trab_tot, c34_a
        c43, c44, c45, c46 = 0, 0, 0, 0
        c91, c36, ben_40, sol_40, rem_cg = c91_a, ex25_a + ces_ex, acep_a, ded_comunes + ex25_a, rem_a

    c37 = c35 + c36
    c41 = min(c34, ben_40)
    c42 = max(0, c34 - c41)

    fe_pot = redondear_dian(min(d["compras_factura_electronica"] * 0.01, 240 * uvt))
    c28 = fe_pot if d["aplicar_beneficio_fe"] else 0
    c139 = redondear_dian(min(int(d["num_dependientes"]), 4) * 72 * uvt) if hay_trab else 0
    fuera_40 = min(rem_cg, c28 + c139)

    c92 = ben_40 + fuera_40
    c93 = max(0, c91 - c92)
    c97 = c93

    c99 = redondear_dian(d["ing_brutos_pensiones"])
    c100 = min(c99, redondear_dian(d["salud_pensiones"]))
    pen_liq = max(0, c99 - c100)
    c102 = min(pen_liq, redondear_dian(12_000 * uvt))
    c103 = max(0, pen_liq - c102)

    c107 = redondear_dian(d["dividendos_gravables"])
    c111 = c97 + c103 + c107

    go_act = redondear_dian(d["ing_ganancia_ocasional"])
    go_lot = redondear_dian(d["ing_loterias_rifas"])
    c112 = go_act + go_lot
    c113 = redondear_dian(d["costos_ganancia_ocasional"])
    c114 = redondear_dian(d["exenta_ganancia_ocasional"])
    go_act_grav = max(0, go_act - c113 - c114)
    c115 = go_act_grav + go_lot

    ing_tot = ing_brutos_cg + c99 + c107 + c112
    t_pat, t_ing = uvt_pat * uvt, uvt_ing * uvt
    causales = []
    if pat_raw >= t_pat:
        causales.append(f"Patrimonio >= {uvt_pat:,.0f} UVT")
    if ing_tot >= t_ing:
        causales.append(f"Ingresos >= {uvt_ing:,.0f} UVT")
    if d["consignaciones"] >= t_ing:
        causales.append(f"Consignaciones >= {uvt_ing:,.0f} UVT")
    if d["tarjetas_credito"] >= t_ing:
        causales.append(f"Tarjetas >= {uvt_ing:,.0f} UVT")
    if d["compras_totales"] >= t_ing:
        causales.append(f"Compras >= {uvt_ing:,.0f} UVT")
    if d["responsable_iva"]:
        causales.append("Responsable de IVA")

    c116 = redondear_dian(impuesto_art241(c111, uvt))
    desc_don = redondear_dian(d["donaciones_entidades"] * 0.25)
    desc_div = redondear_dian(max(0.0, (c107 - (1_090 * uvt)) * 0.19)) if c107 > (1_090 * uvt) else 0
    c125 = min(c116, desc_don + desc_div)
    c126 = max(0, c116 - c125)
    c127 = redondear_dian((go_act_grav * 0.15) + (go_lot * 0.20))
    c129 = c126 + c127

    c130 = redondear_dian(d["anticipo_anio_anterior"])
    c131 = redondear_dian(d["saldo_favor_anterior"])
    c132 = redondear_dian(d["retenciones_anio"])
    imp_ant = redondear_dian(d["impuesto_neto_anterior"])

    an = int(d["anio_declara_num"])
    if an == 0 or c126 == 0:
        c133 = 0
    else:
        pct = {1: 0.25, 2: 0.50}.get(an, 0.75)
        op1 = max(0.0, (c126 * pct) - c132)
        op2 = max(0.0, (((c126 + imp_ant) / 2.0) * pct) - c132) if an > 1 else op1
        c133 = redondear_dian(min(op1, op2))

    c135 = 0
    mora = int(d["meses_extemporaneidad"])
    if mora > 0:
        s_min = redondear_dian(10 * uvt_san)
        bruta = min(c129, c129 * 0.05 * mora) if c129 > 0 else min(ing_tot * 0.05, ing_tot * 0.005 * mora)
        if d["aplica_reduccion_640"]:
            bruta *= 0.50
        c135 = max(s_min, redondear_dian(bruta))

    alivios = c130 + c131 + c132
    c134 = max(0, c129 + c133 - alivios)
    bal = c129 + c133 + c135 - alivios
    c136, c137 = max(0, bal), abs(min(0, bal))

    alertas, sugerencias = [], []
    pat_ant = redondear_dian(d["patrimonio_liquido_anterior"])
    if pat_ant > 0 and (c31 - pat_ant) > (ing_netos_cg + pen_liq + c115 - imp_ant):
        dif_p = redondear_dian((c31 - pat_ant) - (ing_netos_cg + pen_liq + c115 - imp_ant))
        alertas.append(f"Comparación Patrimonial (Art. 236 E.T.): Patrimonio creció ${dif_p:,.0f} por encima de ingresos netos.")
    if c58 > 0 and c60 == 0:
        c59_esp = redondear_dian(d["ing_brutos_capital"] * pct_inf)
        if abs(c59 - c59_esp) >= 2_000:
            sugerencias.append(
                f"Casilla 59 (${c59:,.0f}) difiere del {pct_inf*100:.2f}% oficial para rendimientos financieros (${c59_esp:,.0f}). "
                "Verifica en el Paso 3 si Casilla 58 incluye arriendos o activa el cálculo automático."
            )
    if hay_trab and inc_trab_tot == 0:
        sugerencias.append("Recuerda cargar tus planillas PILA en el Paso 1 o digitar EPS, Pensión y ARL en el Paso 3.")
    if fe_pot > 0 and not d["aplicar_beneficio_fe"]:
        sugerencias.append(f"Puedes restar ${fe_pot:,.0f} adicionales en Casilla 28 (1% Factura Electrónica) activándolo en el Paso 4.")

    return {
        "obligado": len(causales) > 0,
        "estado": "Sí obligado a declarar" if causales else "No obligado a declarar",
        "causales": causales, "estrategia": estrategia, "rem_25": rem_a, "rem_costos": rem_b,
        "lim_40": lim_40, "usado_40": ben_40, "libre_40": max(0, lim_40 - sol_40),
        "ing_tot": ing_tot, "t_pat": t_pat, "t_ing": t_ing, "alivios": alivios,
        "rlg_uvt": round(c111 / uvt, 2) if uvt > 0 else 0, "alertas": alertas, "sugerencias": sugerencias,
        "c28": c28, "c29": c29, "c30": c30, "c31": c31, "c32": c32, "c33": c33, "c34": c34, "c35": c35,
        "c36": c36, "c37": c37, "c38": c38, "c39": c39, "c40": c40, "c41": c41, "c42": c42,
        "c43": c43, "c44": c44, "c45": c45, "c46": c46, "c58": c58,
        "c59": c59, "c60": c60, "c61": c61, "c74": c74, "c76": c76, "c77": c77, "c78": c78, "c91": c91,
        "c92": c92, "c93": c93, "c97": c97, "c99": c99, "c100": c100, "c102": c102, "c103": c103, "c107": c107,
        "c111": c111, "c112": c112, "c113": c113, "c114": c114, "c115": c115, "c116": c116, "c125": c125,
        "c126": c126, "c127": c127, "c129": c129, "c130": c130, "c131": c131, "c132": c132, "c133": c133,
        "c134": c134, "c135": c135, "c136": c136, "c137": c137, "c138": int(d["num_dependientes"]), "c139": c139,
    }


# ==============================================================================
# 4. LECTORES CON DESENCRIPTADO PDF BANCARIO + CONVERSIÓN JPEG + ANTI-DUPLICADOS
# ==============================================================================
CAMPOS_SOPORTE = [
    "efectivo_bancos_dic31", "valor_inmueble_o_cdt_dic31", "saldo_deuda_dic31",
    "ingresos_brutos_trabajo", "cesantias_exentas", "ingresos_brutos_honorarios",
    "salud_pension_obligatoria", "ingresos_capital_cdt", "componente_inflacionario_cdt",
    "intereses_vivienda", "medicina_prepagada", "gmf_4x1000_certificado",
    "aportes_afc_fpv", "retenciones_practicadas",
]


def _nvl(val) -> float:
    try:
        if val is None or isinstance(val, bool) or pd.isna(val):
            return 0.0
        if isinstance(val, str):
            s = re.sub(r"[^\d.\-]", "", val.replace(",", ""))
            return float(s) if s else 0.0
        return float(val)
    except Exception:
        return 0.0


def _cop_txt(s: str) -> float:
    s = str(s).strip().replace("$", "").replace("COP", "").replace(" ", "")
    if not s:
        return 0.0
    if re.match(r"^-?\d{1,3}(\.\d{3})+(,\d+)?$", s):
        s = s.replace(".", "").replace(",", ".")
    elif re.match(r"^-?\d{1,3}(,\d{3})+(\.\d+)?$", s):
        s = s.replace(",", "")
    elif re.match(r"^-?\d+,\d{1,2}$", s):
        s = s.replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except Exception:
        return 0.0


def validar_archivo_soporte(nombre: str, b: bytes, mime: str) -> None:
    if not b:
        raise ValueError(f"El archivo '{nombre}' está vacío.")
    mime_permitidos = {"application/pdf", "image/png", "image/jpeg", "image/jpg"}
    if mime not in mime_permitidos:
        raise ValueError(f"Tipo de archivo no admitido en '{nombre}': {mime or 'desconocido'}.")
    if len(b) > 20 * 1024 * 1024:
        raise ValueError(f"El archivo '{nombre}' supera los 20 MB permitidos.")


def extraer_texto_hibrido(b: bytes, mime: str) -> str:
    """Desencripta PDFs bancarios con clave vacía y aplica OCR con fallback de idioma."""
    if mime == "application/pdf":
        try:
            lector = pypdf.PdfReader(io.BytesIO(b))
            if getattr(lector, "is_encrypted", False):
                try:
                    lector.decrypt("")
                except Exception:
                    pass
            txt = "\n".join((p.extract_text() or "") for p in lector.pages[:4]).strip()
            if len(txt) >= 35:
                return txt
        except Exception:
            pass
        if PDF2IMG_DISPONIBLE and OCR_DISPONIBLE:
            try:
                imgs = convert_from_bytes(b, first_page=1, last_page=2, dpi=150)
                textos_ocr = []
                for im in imgs:
                    try:
                        textos_ocr.append(pytesseract.image_to_string(im, lang="spa"))
                    except Exception:
                        textos_ocr.append(pytesseract.image_to_string(im))
                return "\n".join(textos_ocr).strip()
            except Exception:
                pass
    elif OCR_DISPONIBLE:
        try:
            im = Image.open(io.BytesIO(b))
            try:
                return pytesseract.image_to_string(im, lang="spa").strip()
            except Exception:
                return pytesseract.image_to_string(im).strip()
        except Exception:
            pass
    return ""


def convertir_soporte_a_partes_seguras(nombre: str, b: bytes, mime: str) -> List[Any]:
    """
    Convierte PDFs bancarios protegidos (/Encrypt) o escaneados en imágenes JPEG ligeras (~120 KB)
    para evitar que la API de Gemini devuelva 'HTTP 400 Request contains an invalid argument'.
    """
    partes: List[Any] = [types.Part.from_text(text=f"=== ARCHIVO_ID: {nombre} ===")]

    if mime in ("image/png", "image/jpeg", "image/jpg"):
        if PIL_DISPONIBLE:
            try:
                im = Image.open(io.BytesIO(b)).convert("RGB")
                im.thumbnail((1600, 1600))
                buf = io.BytesIO()
                im.save(buf, format="JPEG", quality=82)
                partes.append(types.Part.from_bytes(data=buf.getvalue(), mime_type="image/jpeg"))
                return partes
            except Exception:
                pass
        partes.append(types.Part.from_bytes(data=b, mime_type=mime))
        return partes

    if mime == "application/pdf":
        if PDF2IMG_DISPONIBLE and PIL_DISPONIBLE:
            try:
                imgs = convert_from_bytes(b, first_page=1, last_page=2, dpi=140)
                for im in imgs:
                    im_rgb = im.convert("RGB")
                    im_rgb.thumbnail((1600, 1600))
                    buf = io.BytesIO()
                    im_rgb.save(buf, format="JPEG", quality=82)
                    partes.append(types.Part.from_bytes(data=buf.getvalue(), mime_type="image/jpeg"))
                if len(partes) > 1:
                    return partes
            except Exception:
                pass

        try:
            lector = pypdf.PdfReader(io.BytesIO(b))
            if getattr(lector, "is_encrypted", False):
                lector.decrypt("")
            escritor = pypdf.PdfWriter()
            for p in lector.pages[:3]:
                escritor.add_page(p)
            buf_pdf = io.BytesIO()
            escritor.write(buf_pdf)
            partes.append(types.Part.from_bytes(data=buf_pdf.getvalue(), mime_type="application/pdf"))
            return partes
        except Exception:
            pass

    return []


def parsear_json_soportes(texto_raw: str) -> List[dict]:
    limpio = (texto_raw or "").strip()
    limpio = re.sub(r"^```(?:json)?\s*", "", limpio, flags=re.IGNORECASE)
    limpio = re.sub(r"\s*```$", "", limpio)
    try:
        data = json.loads(limpio)
    except Exception:
        m = re.search(r"(\{.*\}|\[.*\])", limpio, flags=re.DOTALL)
        if not m:
            raise ValueError(f"Respuesta IA sin JSON válido: {limpio[:140]}")
        data = json.loads(m.group(1))

    lista = data.get("documentos", [data]) if isinstance(data, dict) else (data if isinstance(data, list) else [])
    salida = []
    for item in lista:
        if not isinstance(item, dict):
            continue
        reg = {
            "incluir": True,
            "archivo": str(item.get("archivo") or "Soporte"),
            "tipo_documento": str(item.get("tipo_documento") or "Certificado"),
            "entidad": normalizar_entidad(str(item.get("entidad") or "Entidad")),
        }
        for c in CAMPOS_SOPORTE:
            reg[c] = _nvl(item.get(c, 0.0))
        salida.append(reg)
    return salida


def deduplicar_certificados_lote(registros: List[dict]) -> Tuple[List[dict], List[str]]:
    avisos: List[str] = []
    limpios = [dict(r) for r in registros]
    por_entidad: Dict[str, List[dict]] = {}

    for r in limpios:
        r["incluir"] = bool(r.get("incluir", True))
        r["entidad"] = normalizar_entidad(r.get("entidad", ""))
        por_entidad.setdefault(r["entidad"], []).append(r)

    campos_base = [
        "efectivo_bancos_dic31", "valor_inmueble_o_cdt_dic31", "saldo_deuda_dic31",
        "ingresos_brutos_trabajo", "ingresos_brutos_honorarios", "ingresos_capital_cdt",
        "componente_inflacionario_cdt",
    ]

    def _es_doc_retefuente(doc: dict) -> bool:
        etiqueta = f"{doc.get('tipo_documento', '')} {doc.get('archivo', '')}".lower()
        return any(k in etiqueta for k in ("retefuente", "retencion", "retención")) and "220" not in etiqueta

    for ent, docs in por_entidad.items():
        if len(docs) < 2:
            continue
        principales = [x for x in docs if not _es_doc_retefuente(x)]
        retes = [x for x in docs if _es_doc_retefuente(x)]

        if principales and retes:
            for c in campos_base:
                suma_prin = sum(_nvl(p.get(c, 0)) for p in principales)
                if suma_prin > 0:
                    for rt in retes:
                        val_rt = _nvl(rt.get(c, 0))
                        if val_rt > 0:
                            avisos.append(
                                f"Anti-duplicado ({ent}): se ignoró base repetida de ${val_rt:,.0f} en '{rt.get('archivo')}' porque ya viene en el certificado principal."
                            )
                            rt[c] = 0.0

        for c in ("retenciones_practicadas", "efectivo_bancos_dic31"):
            vistos = set()
            for doc in docs:
                v = round(_nvl(doc.get(c, 0)), 2)
                if v > 0:
                    if v in vistos:
                        avisos.append(
                            f"Anti-duplicado ({ent}): valor idéntico repetido de ${v:,.0f} en '{c}' ({doc.get('archivo')}) ajustado a $0."
                        )
                        doc[c] = 0.0
                    else:
                        vistos.add(v)

    return limpios, avisos


def extraer_soporte_nativo_sin_ia(nombre: str, b: bytes, mime: str) -> dict:
    texto = extraer_texto_hibrido(b, mime)
    t_clean = re.sub(r"[ \t]+", " ", texto or "")
    lineas = [l.strip() for l in t_clean.splitlines() if l.strip()]
    tl = t_clean.lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")

    entidad = normalizar_entidad(tl) if any(
        k in tl for k in ("bbva", "bilbao", "bancolombia", "davivienda", "bogota", "nequi", "daviplata", "colpatria", "occidente", "popular", "av villas")
    ) else "Entidad Detectada"

    nom_low = nombre.lower()
    es_220 = "220" in tl and ("ingresos y retenciones" in tl or "pagos laborales" in tl)
    es_solo_rete = (
        any(k in nom_low for k in ("retefuente", "retencion", "retención"))
        or ("certificado de retencion" in tl and "saldo a 31" not in tl and not es_220)
    )
    es_cdt = ("cdt" in tl or "deposito a termino" in tl) and not es_solo_rete
    es_predial = "predial" in tl or "avaluo catastral" in tl

    res = {
        "incluir": True,
        "archivo": nombre,
        "tipo_documento": (
            "Formulario 220" if es_220 else (
                "Certificado Retefuente" if es_solo_rete else (
                    "Certificado CDT" if es_cdt else ("Predial" if es_predial else "Certificado Tributario")
                )
            )
        ),
        "entidad": entidad,
    }
    for c in CAMPOS_SOPORTE:
        res[c] = 0.0

    pat_m = r"\$\s*-?\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{1,2})?|\b\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?\b|\b\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?\b"
    for l in lineas:
        ll = l.lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
        montos = [m for m in (_cop_txt(x) for x in re.findall(pat_m, l)) if 0 <= abs(m) <= 50_000_000_000]
        if not montos:
            continue
        if any(k in ll for k in ["retencion en la fuente", "retencion practicada", "retefuente", "valor retenido", "total retenciones"]):
            res["retenciones_practicadas"] = max(res["retenciones_practicadas"], montos[-1])
        elif any(k in ll for k in ["gravamen a los movimientos", "gmf", "4 x 1000", "4x1000", "cuatro por mil"]):
            res["gmf_4x1000_certificado"] = max(res["gmf_4x1000_certificado"], montos[-1])
        elif not es_solo_rete and any(k in ll for k in ["rendimientos financieros", "intereses pagados", "rendimientos pagados", "intereses abonados", "total rendimientos"]):
            res["ingresos_capital_cdt"] = max(res["ingresos_capital_cdt"], montos[-1])
        elif ("componente inflacionario" in ll or "parte no gravada" in ll) and ("$" in l or "%" not in l):
            res["componente_inflacionario_cdt"] = max(res["componente_inflacionario_cdt"], montos[-1])
        elif any(k in ll for k in ["saldo a 31 de diciembre", "saldo al 31 de diciembre", "saldo dic 31", "saldo a dic", "saldo capital a 31"]):
            if es_cdt or "cdt" in ll:
                res["valor_inmueble_o_cdt_dic31"] = max(res["valor_inmueble_o_cdt_dic31"], montos[-1])
            elif any(w in ll for w in ["deuda", "hipotec", "obligacion", "tarjeta"]):
                res["saldo_deuda_dic31"] = max(res["saldo_deuda_dic31"], abs(montos[-1]))
            else:
                res["efectivo_bancos_dic31"] = max(res["efectivo_bancos_dic31"], montos[-1])
        elif "avaluo catastral" in ll or "avaluo vigencia" in ll:
            res["valor_inmueble_o_cdt_dic31"] = max(res["valor_inmueble_o_cdt_dic31"], montos[-1])
        elif not es_solo_rete and ("honorarios" in ll or "servicios" in ll) and "retencion" not in ll:
            res["ingresos_brutos_honorarios"] = max(res["ingresos_brutos_honorarios"], montos[-1])

    return res


def _ejecutar_lote_ia(api_key: str, lote: List[Tuple[Any, bytes, str]]) -> List[dict]:
    bloques_texto = []
    partes_visuales: List[Any] = []
    respaldo_local_directo: List[dict] = []

    for fc, b, _ in lote:
        validar_archivo_soporte(fc.name, b, fc.type)
        txt = extraer_texto_hibrido(b, fc.type)
        if len(txt) >= 35:
            bloques_texto.append(f"=== ARCHIVO_ID: {fc.name} ===\n{txt[:9500]}")
        else:
            p_seguras = convertir_soporte_a_partes_seguras(fc.name, b, fc.type)
            if p_seguras:
                partes_visuales.extend(p_seguras)
            else:
                respaldo_local_directo.append(extraer_soporte_nativo_sin_ia(fc.name, b, fc.type))

    esperados = len(lote) - len(respaldo_local_directo)
    if esperados <= 0:
        return respaldo_local_directo

    instruccion = (
        "Actúa como extractor contable DIAN Colombia (no inventes valores). "
        f"Devuelve exactamente {esperados} elemento(s) en 'documentos', uno por cada ARCHIVO_ID, "
        "usando el nombre exacto en 'archivo'. En 'tipo_documento' usa exactamente uno de: "
        "'Certificado CDT', 'Certificado Saldo e Intereses', 'Certificado Retefuente', 'Formulario 220', 'Predial', 'Otro'.\n"
        "REGLA CRÍTICA ANTI-DUPLICIDAD: Si un documento es únicamente un 'Certificado Retefuente' (certificado anual de retención en la fuente), "
        "extrae SOLO 'retenciones_practicadas' y 'gmf_4x1000_certificado', y pon 0 en 'ingresos_capital_cdt' y 'ingresos_brutos_honorarios' "
        "si en el lote ya vienen certificados de CDT/Saldo de esa misma entidad.\n"
        "Responde SOLO con JSON válido:\n"
        '{"documentos": [{"archivo": "nombre.pdf", "tipo_documento": "Certificado CDT", '
        '"entidad": "BBVA Colombia", "efectivo_bancos_dic31": 0, "valor_inmueble_o_cdt_dic31": 0, '
        '"saldo_deuda_dic31": 0, "ingresos_brutos_trabajo": 0, "cesantias_exentas": 0, '
        '"ingresos_brutos_honorarios": 0, "salud_pension_obligatoria": 0, "ingresos_capital_cdt": 0, '
        '"componente_inflacionario_cdt": 0, "intereses_vivienda": 0, "medicina_prepagada": 0, '
        '"gmf_4x1000_certificado": 0, "aportes_afc_fpv": 0, "retenciones_practicadas": 0}]}\n'
        "Si el componente inflacionario aparece solo como porcentaje (ej. 55.44%) y no en pesos, pon 0 en 'componente_inflacionario_cdt'."
    )

    if not partes_visuales:
        prompt_final = "\n\n".join(bloques_texto) + "\n\n" + instruccion
        resp_txt = llamar_gemini(api_key, prompt_final, temperatura=0.0)
    else:
        todas_partes: List[Any] = list(partes_visuales)
        if bloques_texto:
            todas_partes.append(types.Part.from_text(text="\n\n".join(bloques_texto)))
        todas_partes.append(types.Part.from_text(text=instruccion))
        contenido_estructurado = [types.Content(role="user", parts=todas_partes)]
        resp_txt = llamar_gemini(api_key, contenido_estructurado, temperatura=0.0)

    return parsear_json_soportes(resp_txt) + respaldo_local_directo


def extraer_pendientes_con_ia(api_key: str, pendientes: List[Tuple[Any, bytes, str]]) -> List[dict]:
    """Procesa en sublotes seguros de hasta 4 archivos y aísla cualquier PDF corrupto sin tumbar los demás."""
    salida_total: List[dict] = []
    tam_lote = 4
    for i in range(0, len(pendientes), tam_lote):
        sublote = pendientes[i : i + tam_lote]
        try:
            salida_total.extend(_ejecutar_lote_ia(api_key, sublote))
        except Exception:
            if len(sublote) == 1:
                arc, cont, _ = sublote[0]
                salida_total.append(extraer_soporte_nativo_sin_ia(arc.name, cont, arc.type))
            else:
                for item_ind in sublote:
                    try:
                        salida_total.extend(_ejecutar_lote_ia(api_key, [item_ind]))
                    except Exception:
                        arc, cont, _ = item_ind
                        salida_total.append(extraer_soporte_nativo_sin_ia(arc.name, cont, arc.type))
    return salida_total


def leer_pila_pdf(nombre: str, b: bytes) -> dict:
    try:
        lector = pypdf.PdfReader(io.BytesIO(b))
        if getattr(lector, "is_encrypted", False):
            try:
                lector.decrypt("")
            except Exception:
                pass
        txt = "\n".join((p.extract_text() or "") for p in lector.pages)
    except Exception:
        return {"es_pila": False}

    tl = txt.lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    if not any(k in tl for k in ["planilla integrada", "autoliquidacion", "miplanilla", "soi", "asopagos", "subsistema"]):
        return {"es_pila": False}

    lineas = [re.sub(r"\s+", " ", l).strip() for l in txt.splitlines() if l.strip()]
    m_per = re.search(r"\b(20\d{2}-\d{2})\b", txt)
    periodo = m_per.group(1) if m_per else "Detectado"
    m_nom = re.search(r"(\d{7,12})", nombre)
    planilla = m_nom.group(1) if m_nom else nombre

    for idx, l in enumerate(lineas):
        if re.match(r"^20\d{2}-\d{2}$", l):
            for k in range(idx + 1, min(idx + 5, len(lineas))):
                if re.match(r"^\d{7,12}$", lineas[k]):
                    planilla = lineas[k]
                    break
            break

    salud, pension, arl = 0.0, 0.0, 0.0
    for idx, l in enumerate(lineas):
        ll = l.lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
        nums = []
        for v in lineas[idx + 1 : min(idx + 5, len(lineas))]:
            if re.match(r"^\$?\s*\d{1,3}(?:[.,]\d{3})+$", v) or re.match(r"^\$?\s*\d{4,8}$", v):
                vn = _cop_txt(v)
                if 1_000 <= vn <= 35_000_000:
                    nums.append(vn)
            elif any(st_w in v.lower() for st_w in ["salud", "pension", "pensión", "riesgos", "ccf", "totales"]):
                break
        if ll == "salud" and nums and salud == 0:
            salud = nums[0]
        elif ll == "pension" and nums and pension == 0:
            pension = nums[0]
        elif ll in ("riesgos laborales", "riesgos profesionales", "arl") and nums and arl == 0:
            arl = nums[0]

    return {"es_pila": (salud + pension) > 0, "planilla": planilla, "periodo": periodo, "salud": salud, "pension": pension, "arl": arl}


def procesar_exogena(b: bytes, anio_ref: int) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    xls = pd.ExcelFile(io.BytesIO(b))
    hoja = next((h for h in xls.sheet_names if any(p in h.lower() for p in ["report", "tercer", "detalle", "exogen"])), xls.sheet_names[0])
    df_raw = pd.read_excel(xls, sheet_name=hoja, header=None)

    nit_m, nom_m, anio_m = "", "", None
    for i in range(min(15, len(df_raw))):
        fv = [str(x).strip() if pd.notna(x) else "" for x in df_raw.iloc[i].values]
        et = fv[0].lower().replace("ó", "o") if fv else ""
        vc = fv[2] if len(fv) > 2 else ""
        if "identificacion:" in et and vc:
            nit_m = vc.split(".")[0].strip()
        elif ("nombres" in et or "razon social" in et) and vc:
            nom_m = vc.strip()
        elif "año al que se refiere" in et and vc.isdigit():
            anio_m = int(vc)

    df_s = df_raw.dropna(how="all").reset_index(drop=True)
    f_head = next(
        (i for i in range(min(30, len(df_s))) if sum(1 for c in df_s.iloc[i].values if pd.notna(c) and any(k in str(c).lower() for k in ["nit", "detalle", "valor", "uso"])) >= 3),
        None,
    )

    if f_head is not None:
        cols, vistos, c_nit = [], {}, 0
        for i, c in enumerate(df_s.iloc[f_head].values):
            nm = re.sub(r"\s+", " ", str(c).strip()) if pd.notna(c) and str(c).strip() else f"Col_{i+1}"
            if nm.upper() == "NIT":
                c_nit += 1
                nm = "NIT Reportante" if c_nit == 1 else ("NIT Declarante" if c_nit == 2 else f"NIT_{c_nit}")
            if nm in vistos:
                vistos[nm] += 1
                nm = f"{nm}_{vistos[nm]}"
            else:
                vistos[nm] = 1
            cols.append(nm)
        df = df_s.iloc[f_head + 1 :].copy()
        df.columns = cols
    else:
        df = df_s.copy()
        df.columns = [f"Col_{i+1}" for i in range(df.shape[1])]

    df = df.dropna(how="all").reset_index(drop=True)
    c_nit_r = df.columns[0]
    c_nom_r = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    c_det = next((c for c in df.columns if "detalle" in c.lower()), df.columns[min(4, len(df.columns) - 1)])
    c_val = next((c for c in df.columns if "valor" in c.lower()), df.columns[min(5, len(df.columns) - 1)])
    c_uso = next((c for c in df.columns if "uso" in c.lower()), None)

    m_tope = df[c_nit_r].isna() & df[c_det].astype(str).str.strip().str.lower().str.startswith("tope")
    df_topes, df_ter = df[m_tope], df[~m_tope].copy()

    topes_ofi = {}
    for _, rt in df_topes.iterrows():
        kt, vt = str(rt[c_det]).lower(), _nvl(rt[c_val])
        for n_t in ["1", "2", "3", "4", "5"]:
            if f"tope {n_t}" in kt:
                topes_ofi[n_t] = vt

    res: Dict[str, Any] = {
        "nombre": nom_m, "nit": nit_m, "anio": anio_m, "efectivo": 0.0, "cdt_saldo": 0.0, "inmuebles": 0.0,
        "deudas": 0.0, "consignaciones": 0.0, "tc": 0.0, "compras": 0.0, "fe_1pct": 0.0,
        "ing_trab": 0.0, "ing_hon": 0.0, "ing_nolab": 0.0, "salud_pen": 0.0, "ing_cap": 0.0,
        "inc_cap": 0.0, "retenciones": 0.0, "vivienda": 0.0, "prepagada": 0.0, "gmf": 0.0,
        "por_entidad": {},
    }

    def _sum_ent(ent_k: str, campo_k: str, v_k: float) -> None:
        res["por_entidad"].setdefault(ent_k, {})
        res["por_entidad"][ent_k][campo_k] = res["por_entidad"][ent_k].get(campo_k, 0.0) + v_k

    por_nit = {}
    for _, row in df_ter.iterrows():
        val = _nvl(row.get(c_val, 0))
        if val == 0:
            continue
        nit_r = str(row.get(c_nit_r, "") or "").strip().split(".")[0]
        nom_r = str(row.get(c_nom_r, "") or "").strip()
        ent_norm = normalizar_entidad(f"{nit_r} {nom_r}")
        dn = str(row.get(c_det, "") or "").lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
        un = (str(row.get(c_uso, "") or "").lower() if c_uso else "").replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")

        if "susceptible de beneficio" in dn:
            res["fe_1pct"] += val
        elif "tope 5" in un or "facturas tras ajustes" in dn:
            res["compras"] += val
        elif "tope 3" in un or "consumos tarjeta" in dn:
            res["tc"] += val
        elif "tope 4" in un or "movimientos en cuentas" in dn or "inversion efectuada" in dn:
            res["consignaciones"] += val
        elif "tope 2" in un or "r29" in un or "saldo cuentas" in dn or "saldo cdt" in dn:
            if val < 0 or ("r30" in un and "r29" not in un) or "deuda" in dn or "pasivo" in dn:
                res["deudas"] += abs(val)
                _sum_ent(ent_norm, "deudas", abs(val))
            elif any(k in dn for k in ["cdt", "inversion", "accion", "fiducia", "fondo"]):
                res["cdt_saldo"] += val
                _sum_ent(ent_norm, "otros_activos", val)
            elif any(k in dn for k in ["inmueble", "predial", "avaluo", "vehiculo"]):
                res["inmuebles"] += val
            else:
                res["efectivo"] += val
                _sum_ent(ent_norm, "efectivo", val)
        elif "r132" in un or "retencion" in dn:
            res["retenciones"] += val
            _sum_ent(ent_norm, "retenciones_anio", val)
        elif ("r58" in un or "cdt rendimientos" in dn or "rendimientos financieros" in dn or "intereses pagados" in dn) and "retencion" not in dn:
            res["ing_cap"] += val
            _sum_ent(ent_norm, "ing_brutos_capital", val)
        elif "documentos soporte" in dn or "facturas electronicas de venta" in dn:
            por_nit.setdefault(nit_r, {"ent": ent_norm, "t": 0.0, "h": 0.0, "n": 0.0, "ds": 0.0})
            por_nit[nit_r]["ds"] += (-val if (dn.startswith("ajuste") or "descuenta automaticamente" in un) else val)
        elif "r33" in un or "r44" in un or "salud obligatoria" in dn or "pension obligatoria" in dn:
            res["salud_pen"] += val
        elif "r38" in un or "vivienda" in dn:
            res["vivienda"] += val
        elif "r39" in un or "prepagada" in dn:
            res["prepagada"] += val
        elif "gmf" in dn or "4x1000" in dn:
            res["gmf"] += val
        elif "r32" in un or "salarios" in dn or "pagos laborales" in dn:
            por_nit.setdefault(nit_r, {"ent": ent_norm, "t": 0.0, "h": 0.0, "n": 0.0, "ds": 0.0})
            por_nit[nit_r]["t"] += val
        elif "r43" in un or "honorarios" in dn or "servicios" in dn:
            por_nit.setdefault(nit_r, {"ent": ent_norm, "t": 0.0, "h": 0.0, "n": 0.0, "ds": 0.0})
            por_nit[nit_r]["h"] += val
        elif "r74" in un or "tope 1" in un:
            por_nit.setdefault(nit_r, {"ent": ent_norm, "t": 0.0, "h": 0.0, "n": 0.0, "ds": 0.0})
            por_nit[nit_r]["n"] += val

    for dn_v in por_nit.values():
        exo_s = dn_v["t"] + dn_v["h"] + dn_v["n"]
        ds_n = max(0.0, dn_v["ds"])
        hon_fin = dn_v["h"] + (max(0.0, ds_n - exo_s) if ds_n > exo_s else 0.0)
        res["ing_trab"] += dn_v["t"]
        res["ing_hon"] += hon_fin
        res["ing_nolab"] += dn_v["n"]
        if dn_v["t"] > 0:
            _sum_ent(dn_v["ent"], "ing_brutos_trabajo", dn_v["t"])
        if hon_fin > 0:
            _sum_ent(dn_v["ent"], "ing_brutos_honorarios", hon_fin)

    if topes_ofi.get("4", 0) > 0:
        res["consignaciones"] = topes_ofi["4"]
    if topes_ofi.get("5", 0) > 0:
        res["compras"] = topes_ofi["5"]
    if "3" in topes_ofi:
        res["tc"] = topes_ofi["3"]
    if res["ing_cap"] > 0:
        res["inc_cap"] = float(redondear_dian(res["ing_cap"] * COMP_INFLACIONARIO.get(anio_m or anio_ref, 0.5544)))

    for c in df_ter.columns:
        df_ter[c] = df_ter[c].apply(_nvl) if c == c_val else df_ter[c].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    return df_ter.reset_index(drop=True), res


# ==============================================================================
# 5. EXPORTADORES EJECUTIVOS (.XLSX ESTILIZADO + .DOCX DICTAMEN)
# ==============================================================================
def filas_formulario_210(r: dict) -> List[Tuple[str, str, int]]:
    return [
        ("28", "1% de compras con factura electrónica", r["c28"]),
        ("29", "Total patrimonio bruto", r["c29"]),
        ("30", "Deudas", r["c30"]),
        ("31", "Total patrimonio líquido (29 - 30)", r["c31"]),
        ("32", "Ingresos brutos Rentas de Trabajo (Salarios / Servicios con 25% exento)", r["c32"]),
        ("33", "Ingresos no constitutivos Trabajo (EPS Salud + Pensión + ARL Independiente)", r["c33"]),
        ("34", "Renta líquida Trabajo (32 - 33)", r["c34"]),
        ("35", "Aportes voluntarios AFC, FVP y/o AVC", r["c35"]),
        ("36", "Otras rentas exentas Trabajo (25% laboral + Cesantías)", r["c36"]),
        ("37", "Total rentas exentas Trabajo (35 + 36)", r["c37"]),
        ("38", "Intereses de vivienda", r["c38"]),
        ("39", "Otras deducciones Trabajo (Prepagada, 10% dep., ICETEX, 50% GMF)", r["c39"]),
        ("40", "Total deducciones imputables Trabajo (38 + 39)", r["c40"]),
        ("41", "Rentas exentas y/o deducciones imputables limitadas Trabajo", r["c41"]),
        ("42", "Renta líquida ordinaria Trabajo (34 - 41)", r["c42"]),
        ("43", "Ingresos brutos Honorarios sujetos a costos", r["c43"]),
        ("44", "Ingresos no constitutivos Honorarios (EPS Salud + Pensión + ARL)", r["c44"]),
        ("45", "Costos y deducciones procedentes Honorarios", r["c45"]),
        ("46", "Renta líquida Honorarios (43 - 44 - 45)", r["c46"]),
        ("58", "Ingresos brutos Rentas de Capital (CDT / Intereses / Arriendos)", r["c58"]),
        ("59", "Ingresos no constitutivos Capital (Componente Inflacionario CDT)", r["c59"]),
        ("60", "Costos y deducciones Rentas de Capital", r["c60"]),
        ("61", "Renta líquida Capital (58 - 59 - 60)", r["c61"]),
        ("74", "Ingresos brutos Rentas No Laborales", r["c74"]),
        ("76", "Ingresos no constitutivos No Laborales", r["c76"]),
        ("77", "Costos y gastos No Laborales", r["c77"]),
        ("78", "Renta líquida No Laborales", r["c78"]),
        ("91", "Renta líquida Cédula General (34 + 46 + 61 + 78)", r["c91"]),
        ("92", "Rentas exentas y deducciones imputables limitadas", r["c92"]),
        ("93", "Renta líquida ordinaria Cédula General (91 - 92)", r["c93"]),
        ("97", "Renta líquida gravable Cédula General", r["c97"]),
        ("99", "Ingresos brutos por pensiones", r["c99"]),
        ("100", "Ingresos no constitutivos pensiones (Salud)", r["c100"]),
        ("102", "Rentas exentas de pensiones", r["c102"]),
        ("103", "Renta líquida gravable Cédula de Pensiones", r["c103"]),
        ("107", "Dividendos y participaciones gravables", r["c107"]),
        ("111", "TOTAL RENTA LÍQUIDA GRAVABLE (Base Art. 241 E.T.)", r["c111"]),
        ("112", "Ingresos por ganancias ocasionales", r["c112"]),
        ("113", "Costos por ganancias ocasionales", r["c113"]),
        ("114", "Ganancias ocasionales no gravadas y exentas", r["c114"]),
        ("115", "Ganancias ocasionales gravables", r["c115"]),
        ("116", "Impuesto sobre la renta líquida gravable", r["c116"]),
        ("125", "Total descuentos tributarios", r["c125"]),
        ("126", "Impuesto neto de renta (116 - 125)", r["c126"]),
        ("127", "Impuesto de ganancias ocasionales", r["c127"]),
        ("129", "Total impuesto a cargo (126 + 127)", r["c129"]),
        ("130", "Anticipo renta año anterior", r["c130"]),
        ("131", "Saldo a favor año anterior sin solicitud", r["c131"]),
        ("132", "Retenciones año gravable a declarar", r["c132"]),
        ("133", "Anticipo renta año siguiente (Art. 807 E.T.)", r["c133"]),
        ("134", "Saldo a pagar por impuesto", r["c134"]),
        ("135", "Sanciones (Extemporaneidad Art. 641 / 640 E.T.)", r["c135"]),
        ("136", "TOTAL SALDO A PAGAR", r["c136"]),
        ("137", "TOTAL SALDO A FAVOR", r["c137"]),
        ("138", "Número de dependientes económicos", r["c138"]),
        ("139", "Adición por dependientes a casilla 92 (72 UVT c/u)", r["c139"]),
    ]


def generar_excel(r: dict, d: dict, anio: int, df_pila: pd.DataFrame = None) -> bytes:
    df_210 = pd.DataFrame(filas_formulario_210(r), columns=["Casilla", "Concepto Oficial Formulario 210", "Valor (COP)"])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        hoja = f"Form_210_AG{anio}"
        df_210.to_excel(writer, index=False, sheet_name=hoja)
        if df_pila is not None and not df_pila.empty:
            df_pila.to_excel(writer, index=False, sheet_name="Detalle_PILA")
        ws = writer.sheets[hoja]
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 68
        ws.column_dimensions["C"].width = 22
        h_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
        k_fill = PatternFill(start_color="E8F4EA", end_color="E8F4EA", fill_type="solid")
        for cell in ws[1]:
            cell.fill = h_fill
            cell.font = Font(color="FFFFFF", bold=True)
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            cas = str(row[0].value)
            if cas in ("29", "31", "97", "111", "129", "132", "136", "137"):
                for c in row:
                    c.fill = k_fill
                    c.font = Font(bold=True)
            row[2].number_format = "#,##0" if cas == "138" else "$#,##0"
    return buf.getvalue()


def generar_word(r: dict, d: dict, anio: int, uvt: float) -> bytes:
    doc = Document()
    doc.add_heading(f"Dictamen y Borrador Formulario 210 — AG {anio}", level=1)
    dv = calcular_dv(d.get("nit_declarante", ""))
    doc.add_paragraph(
        f"Declarante: {d.get('nombre_declarante') or 'Contribuyente'} | NIT/CC: {d.get('nit_declarante') or 'N/A'}{f'-{dv}' if dv else ''}\n"
        f"Año Gravable: {anio} | UVT: ${uvt:,.0f} | Estado: {r['estado']} | Estrategia: {r['estrategia']}"
    )
    doc.add_heading("Casillas Principales Formulario 210 DIAN", level=2)
    for cas, nom, val in filas_formulario_210(r):
        if val > 0 or cas in ("29", "31", "97", "111", "129", "132", "136", "137"):
            doc.add_paragraph(
                f"Casilla {cas} — {nom}: {'' if cas == '138' else '$'}{val:,.0f}",
                style="List Bullet",
            )
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ==============================================================================
# 6. INTERFAZ PRINCIPAL BLINDADA (v5.7 FUSIÓN PRO — VISIÓN JPEG + MUISCA EXACTO)
# ==============================================================================
if not verificar_acceso():
    st.stop()

if "datos" not in st.session_state:
    st.session_state.datos = datos_vacios()
elif "auto_comp_inflacionario" not in st.session_state.datos:
    st.session_state.datos["auto_comp_inflacionario"] = True
if "anio_activo" not in st.session_state:
    st.session_state.anio_activo = datetime.date.today().year - 1
if "uvt_map" not in st.session_state:
    st.session_state.uvt_map = {}
if "form_id" not in st.session_state:
    st.session_state.form_id = 0
if "reset_id" not in st.session_state:
    st.session_state.reset_id = 0
if "mem_exo" not in st.session_state:
    st.session_state.mem_exo = None
if "mem_pila" not in st.session_state:
    st.session_state.mem_pila = {}
if "mem_cert" not in st.session_state:
    st.session_state.mem_cert = []
if "avisos_cert" not in st.session_state:
    st.session_state.avisos_cert = []
if "auditoria_ia" not in st.session_state:
    st.session_state.auditoria_ia = ""
if "chat" not in st.session_state:
    st.session_state.chat = []
if "toast_ok" not in st.session_state:
    st.session_state.toast_ok = ""

fid = st.session_state.form_id
rid = st.session_state.reset_id
d = st.session_state.datos

with st.sidebar:
    st.title("⚙️ Configuración DIAN")
    st.caption("Versión **v5.7 Fusión Pro** • PDF Bancario Desencriptado + Visión JPEG")

    sec_key = obtener_secreto("GEMINI_API_KEY")
    if sec_key:
        api_key = sec_key
        modo_txt = "Auth Key (AQ.)" if sec_key.startswith("AQ.") else "API Key (AIza)"
        st.success(f"🔒 Gemini conectada — {modo_txt} ({sec_key[:3]}...{sec_key[-4:]})")
    else:
        st.warning("⚠️ Sin clave en Secrets. Pégala abajo si vas a usar IA:")
        man_key = st.text_input("Pegar clave Gemini (AIza... o AQ....)", value="", placeholder="AIzaSy... o AQ....")
        api_key = limpiar_key(man_key)
        if api_key:
            st.success(f"✅ Clave detectada ({api_key[:3]}...{api_key[-4:]})")

    anio_gravable = int(
        st.number_input(
            "📅 Año Gravable a Declarar",
            min_value=2020, max_value=2100,
            value=int(st.session_state.anio_activo),
            step=1, key=f"w_anio_{fid}",
        )
    )
    st.session_state.anio_activo = anio_gravable

    uvt_def = st.session_state.uvt_map.get(anio_gravable, uvt_por_anio(anio_gravable))
    valor_uvt = int(
        st.number_input(
            f"Valor UVT AG {anio_gravable} ($ COP)",
            min_value=10_000, value=int(uvt_def), step=100, key=f"w_uvt_{anio_gravable}_{fid}",
        )
    )
    st.session_state.uvt_map[anio_gravable] = valor_uvt

    with st.expander("Ajustar Topes UVT y Sanción"):
        uvt_pat = st.number_input("Tope Patrimonio (UVT)", value=4_500, step=100)
        uvt_ing = st.number_input("Tope Ingresos / Consumos (UVT)", value=1_400, step=100)
        uvt_san = st.number_input(f"UVT Presentación {anio_gravable + 1}", value=uvt_por_anio(anio_gravable + 1), step=100)

    d["anio_declara_num"] = st.selectbox(
        "Anticipo Año Siguiente (Casilla 133)",
        options=[0, 1, 2, 3], index=int(d["anio_declara_num"]),
        format_func=lambda x: {
            0: "Sin Anticipo ($0)",
            1: "1er Año (25%)",
            2: "2do Año (50%)",
            3: "3er Año+ (75% Optimizado)",
        }[x],
    )

    with st.expander("💾 Guardar / Abrir Cliente (.json)"):
        st.download_button(
            "⬇️ Descargar Respaldo (.json)",
            data=json.dumps({"anio": anio_gravable, "datos": d}, ensure_ascii=False, indent=2).encode("utf-8"),
            file_name=f"Renta_{d.get('nit_declarante') or 'Cliente'}_AG{anio_gravable}.json",
            mime="application/json", use_container_width=True,
        )
        up_j = st.file_uploader("Abrir respaldo .json", type=["json"], key=f"up_j_{rid}")
        if up_j and st.button("📂 Cargar Cliente", use_container_width=True):
            obj = json.loads(up_j.getvalue().decode("utf-8"))
            st.session_state.anio_activo = int(obj.get("anio", anio_gravable))
            nv = datos_vacios()
            nv.update(obj.get("datos", {}))
            st.session_state.datos = nv
            st.session_state.form_id += 1
            st.rerun()

    clave_m = obtener_secreto("CLAVE_MAESTRA")
    if clave_m and st.session_state.get("es_admin", False):
        with st.expander("🛡️ Licencias Clientes"):
            cli = st.text_input("Cédula o Correo")
            if cli:
                st.code(generar_licencia(cli, clave_m))

    st.divider()
    with st.popover("🧹 Nueva Declaración / Limpiar Todo", use_container_width=True):
        st.warning("¿Borrar todos los datos y archivos actuales?")
        if st.button("✅ Sí, reiniciar de cero", type="primary", use_container_width=True):
            st.session_state.datos = datos_vacios()
            st.session_state.mem_exo = None
            st.session_state.mem_pila = {}
            st.session_state.mem_cert = []
            st.session_state.avisos_cert = []
            st.session_state.auditoria_ia = ""
            st.session_state.chat = []
            st.session_state.form_id += 1
            st.session_state.reset_id += 1
            st.rerun()

st.title(f"📊 RentaIA Pro — Liquidador & Auditor Formulario 210 (AG {anio_gravable})")
st.caption(
    f"Año Gravable **{anio_gravable}** (Presentación en **{anio_gravable + 1}**) • "
    f"UVT: **${valor_uvt:,.0f}** • Memoria Privada RAM"
)

if st.session_state.toast_ok:
    st.success(st.session_state.toast_ok)
    st.session_state.toast_ok = ""

slot_kpis = st.container()

t1, t2, t3, t4, t5 = st.tabs([
    "🚀 Paso 1. Exógena, PILA y Soportes",
    "👤 Paso 2. Perfil y Patrimonio",
    "💼 Paso 3. Ingresos del Año",
    "🛡️ Paso 4. Beneficios y Retenciones",
    "📋 Paso 5. Formulario 210, Excel y Chat IA",
])

# ==============================================================================
# PASO 1: EXÓGENA PERSISTENTE + PILA PERSISTENTE + SOPORTES CON CONCILIACIÓN
# ==============================================================================
with t1:
    c_izq, c_der = st.columns(2)

    with c_izq:
        st.subheader("📊 A. Información Exógena DIAN (.xlsx)")
        f_exo = st.file_uploader(
            "Sube el Excel descargado de la DIAN (No requiere IA)",
            type=["xlsx", "xls"], key=f"exo_{rid}",
        )
        if f_exo is not None:
            try:
                df_e, rs_e = procesar_exogena(f_exo.getvalue(), anio_gravable)
                st.session_state.mem_exo = {"nombre": f_exo.name, "df": df_e, "res": rs_e}
            except Exception as e:
                st.error(f"Error leyendo Excel de Exógena: {e}")

        if st.session_state.mem_exo is not None:
            df_e = st.session_state.mem_exo["df"]
            rs_e = st.session_state.mem_exo["res"]
            st.caption(f"Archivo fijo en sesión: **{st.session_state.mem_exo['nombre']}** • Titular: **{rs_e['nombre']}** ({rs_e['nit']})")
            st.dataframe(df_e, use_container_width=True, height=190)

            st.table(pd.DataFrame([
                ("Saldos Cuentas Bancarias dic 31", f"${rs_e['efectivo']:,.0f}"),
                ("Saldo CDT / Inversiones dic 31", f"${rs_e['cdt_saldo']:,.0f}"),
                ("Movimientos / Consignaciones (Tope 4)", f"${rs_e['consignaciones']:,.0f}"),
                ("Compras Totales (Tope 5) / Base FE 1%", f"${rs_e['compras']:,.0f} / ${rs_e['fe_1pct']:,.0f}"),
                ("Ingresos Trabajo + Honorarios/Servicios", f"${rs_e['ing_trab'] + rs_e['ing_hon']:,.0f}"),
                ("Rendimientos CDT (C.58) / No Gravado (C.59)", f"${rs_e['ing_cap']:,.0f} / ${rs_e['inc_cap']:,.0f}"),
                ("Retenciones en la Fuente (C.132)", f"${rs_e['retenciones']:,.0f}"),
            ], columns=["Concepto Detectado", "Valor"]))

            if st.button("⚡ Autollenar Declaración con esta Exógena", type="primary", use_container_width=True):
                if rs_e["anio"]:
                    st.session_state.anio_activo = int(rs_e["anio"])
                if rs_e["nombre"]:
                    d["nombre_declarante"] = rs_e["nombre"]
                if rs_e["nit"]:
                    d["nit_declarante"] = rs_e["nit"]
                d["efectivo"] = rs_e["efectivo"]
                d["otros_activos"] = rs_e["cdt_saldo"]
                if rs_e["inmuebles"] > 0:
                    d["inmuebles"] = rs_e["inmuebles"]
                d["deudas"] = rs_e["deudas"]
                d["consignaciones"] = rs_e["consignaciones"]
                d["tarjetas_credito"] = rs_e["tc"]
                d["compras_totales"] = rs_e["compras"]
                d["compras_factura_electronica"] = rs_e["fe_1pct"]
                d["ing_brutos_trabajo"] = rs_e["ing_trab"]
                d["ing_brutos_honorarios"] = rs_e["ing_hon"]
                d["ing_brutos_no_laborales"] = rs_e["ing_nolab"]
                if rs_e["salud_pen"] > 0:
                    d["salud_pension_trabajo"] = rs_e["salud_pen"]
                d["ing_brutos_capital"] = rs_e["ing_cap"]
                d["incrngo_capital"] = rs_e["inc_cap"]
                d["auto_comp_inflacionario"] = True
                d["retenciones_anio"] = rs_e["retenciones"]
                if rs_e["vivienda"] > 0:
                    d["intereses_vivienda"] = rs_e["vivienda"]
                if rs_e["prepagada"] > 0:
                    d["medicina_prepagada"] = rs_e["prepagada"]
                if rs_e["gmf"] > 0:
                    d["gmf_certificado_100"] = rs_e["gmf"]
                st.session_state.toast_ok = "✅ ¡Exógena cargada al tablero! Tus archivos siguen fijos en pantalla."
                st.session_state.form_id += 1
                st.rerun()

            if st.button("🤖 Explicar mi Exógena con IA Gemini", use_container_width=True):
                try:
                    with st.spinner("Analizando terceros..."):
                        pr = f"Explica brevemente en español sencillo estos terceros de Exógena DIAN:\n{df_e.head(120).to_csv(index=False)}"
                        st.session_state.auditoria_ia = llamar_gemini(api_key, pr, temperatura=0.1)
                except Exception as e:
                    st.error(f"⚠️ {e}")

        if st.session_state.auditoria_ia:
            st.info(st.session_state.auditoria_ia)

    with c_der:
        st.subheader("🏥 B. Lector Masivo de Planillas PILA (Multi-PDF sin IA)")
        f_pilas = st.file_uploader(
            "Arrastra tus planillas PILA PDF juntas (quedan guardadas en memoria sin borrarse)",
            type=["pdf"], accept_multiple_files=True, key=f"pila_{rid}",
        )
        if f_pilas:
            for fp in f_pilas:
                inf = leer_pila_pdf(fp.name, fp.getvalue())
                if inf.get("es_pila"):
                    st.session_state.mem_pila[inf["planilla"]] = {
                        "Planilla": inf["planilla"],
                        "Periodo": inf["periodo"],
                        "Salud ($)": float(inf["salud"]),
                        "Pensión ($)": float(inf["pension"]),
                        "ARL ($)": float(inf["arl"]),
                    }

        df_pila_exp = None
        if st.session_state.mem_pila:
            df_p = pd.DataFrame(list(st.session_state.mem_pila.values())).sort_values(by="Periodo").reset_index(drop=True)
            df_p_ed = st.data_editor(
                df_p, use_container_width=True, hide_index=True,
                num_rows="dynamic", key=f"ed_p_{rid}_{len(df_p)}",
            )
            df_p_ed["Salud + Pensión ($)"] = df_p_ed["Salud ($)"].fillna(0) + df_p_ed["Pensión ($)"].fillna(0)
            df_p_ed["Total INCRNGO PILA ($)"] = df_p_ed["Salud + Pensión ($)"] + df_p_ed["ARL ($)"].fillna(0)
            df_pila_exp = df_p_ed
            tot_salud = float(df_p_ed["Salud ($)"].fillna(0).sum())
            tot_pen = float(df_p_ed["Pensión ($)"].fillna(0).sum())
            tot_sp = tot_salud + tot_pen
            tot_arl = float(df_p_ed["ARL ($)"].fillna(0).sum())
            tot_inc = redondear_dian(tot_sp + tot_arl)

            st.info(
                f"✅ **{len(df_p_ed)} planilla(s) fijas en memoria:** "
                f"EPS: **`${tot_salud:,.0f}`** + Pensión: **`${tot_pen:,.0f}`** + ARL: **`${tot_arl:,.0f}`** "
                f"= **`${tot_sp + tot_arl:,.0f}`** (Casilla 33 DIAN: **`${tot_inc:,.0f}`**)"
            )
            cp1, cp2 = st.columns([3, 1])
            if cp1.button("⚡ Pasar Total PILA a mi Declaración", type="primary", use_container_width=True):
                d["salud_pension_honorarios"] = tot_sp
                d["arl_honorarios"] = tot_arl
                st.session_state.toast_ok = (
                    f"✅ ¡PILA aplicada! EPS + Pensión: ${tot_sp:,.0f} + ARL: ${tot_arl:,.0f} "
                    f"→ Casilla 33 INCRNGO: ${tot_inc:,.0f}."
                )
                st.session_state.form_id += 1
                st.rerun()
            if cp2.button("🗑️ Quitar PILA", use_container_width=True):
                st.session_state.mem_pila = {}
                st.session_state.reset_id += 1
                st.rerun()

        st.divider()
        st.subheader("📄 C. Lector de Otros Certificados (220, CDT, Predial, Salud)")
        f_certs = st.file_uploader(
            "Sube PDFs o fotos de certificados (desencripta bancos y filtra duplicados automáticamente)",
            type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key=f"cert_{rid}",
        )
        if f_certs:
            bt_ia, bt_nat = st.columns(2)

            if bt_ia.button(f"🤖 Extraer {len(f_certs)} con IA (Caché SHA-256)", type="primary", use_container_width=True):
                if not api_key:
                    st.error("🔑 Configura tu clave en Configuración -> Secrets o usa el botón de al lado 'Extraer SIN IA'.")
                else:
                    resultados = {it["_sha256"]: it for it in st.session_state.mem_cert if it.get("_sha256")}
                    sin_huella = [it for it in st.session_state.mem_cert if not it.get("_sha256")]
                    pendientes = []
                    for arc in f_certs:
                        cont = arc.getvalue()
                        h_sha = hashlib.sha256(cont).hexdigest()
                        if h_sha not in resultados:
                            pendientes.append((arc, cont, h_sha))

                    if pendientes:
                        with st.spinner(f"Procesando {len(pendientes)} soporte(s) con Visión IA..."):
                            lista_nuevos = extraer_pendientes_con_ia(api_key, pendientes)
                            mapa_nombre = {it.get("archivo", ""): it for it in lista_nuevos}
                            for idx_p, (arc, cont, h_sha) in enumerate(pendientes):
                                item_ok = mapa_nombre.get(arc.name) or (
                                    lista_nuevos[idx_p] if idx_p < len(lista_nuevos) else extraer_soporte_nativo_sin_ia(arc.name, cont, arc.type)
                                )
                                item_ok["archivo"] = arc.name
                                item_ok["_sha256"] = h_sha
                                resultados[h_sha] = item_ok

                    lote_bruto = sin_huella + list(resultados.values())
                    limpios, avisos = deduplicar_certificados_lote(lote_bruto)
                    st.session_state.mem_cert = limpios
                    st.session_state.avisos_cert = avisos

            if bt_nat.button(f"⚡ Extraer {len(f_certs)} SIN IA (Local)", use_container_width=True):
                with st.spinner("Extrayendo cifras localmente (PDF + OCR)..."):
                    nuevos_loc = []
                    for arc in f_certs:
                        cont = arc.getvalue()
                        it_l = extraer_soporte_nativo_sin_ia(arc.name, cont, arc.type)
                        it_l["_sha256"] = hashlib.sha256(cont).hexdigest()
                        nuevos_loc.append(it_l)
                    limpios, avisos = deduplicar_certificados_lote(nuevos_loc)
                    st.session_state.mem_cert = limpios
                    st.session_state.avisos_cert = avisos

        for av in st.session_state.avisos_cert:
            st.info(f"🛡️ {av}")

        if st.session_state.mem_cert:
            st.caption("Verifica las cifras extraídas (desmarca 'incluir' si algún soporte ya está contado en otro documento):")
            df_base = pd.DataFrame(st.session_state.mem_cert).drop(columns=["_sha256"], errors="ignore")
            if "incluir" not in df_base.columns:
                df_base.insert(0, "incluir", True)
            df_c = st.data_editor(df_base, use_container_width=True, hide_index=True, num_rows="dynamic")
            df_act = df_c[df_c["incluir"].fillna(True).astype(bool)].copy()

            mapeo_campos = [
                ("efectivo_bancos_dic31", "efectivo"),
                ("valor_inmueble_o_cdt_dic31", "otros_activos"),
                ("saldo_deuda_dic31", "deudas"),
                ("ingresos_brutos_trabajo", "ing_brutos_trabajo"),
                ("cesantias_exentas", "cesantias_exentas"),
                ("ingresos_brutos_honorarios", "ing_brutos_honorarios"),
                ("salud_pension_obligatoria", "salud_pension_trabajo"),
                ("ingresos_capital_cdt", "ing_brutos_capital"),
                ("componente_inflacionario_cdt", "incrngo_capital"),
                ("intereses_vivienda", "intereses_vivienda"),
                ("medicina_prepagada", "medicina_prepagada"),
                ("gmf_4x1000_certificado", "gmf_certificado_100"),
                ("aportes_afc_fpv", "aportes_afc_fpv"),
                ("retenciones_practicadas", "retenciones_anio"),
            ]

            pct_inf_act = COMP_INFLACIONARIO.get(anio_gravable, 0.5544)
            cc0, cc1, cc2 = st.columns(3)

            if cc0.button("🧠 Conciliar con Exógena (Sin Duplicar)", type="primary", use_container_width=True):
                exo_ent = (st.session_state.mem_exo or {}).get("res", {}).get("por_entidad", {})
                campos_por_entidad = {
                    "efectivo_bancos_dic31": "efectivo",
                    "valor_inmueble_o_cdt_dic31": "otros_activos",
                    "saldo_deuda_dic31": "deudas",
                    "ingresos_brutos_trabajo": "ing_brutos_trabajo",
                    "ingresos_brutos_honorarios": "ing_brutos_honorarios",
                    "ingresos_capital_cdt": "ing_brutos_capital",
                    "retenciones_practicadas": "retenciones_anio",
                }
                if exo_ent:
                    cert_por_ent: Dict[str, Dict[str, float]] = {}
                    for _, fila_c in df_act.iterrows():
                        ent_c = normalizar_entidad(str(fila_c.get("entidad", "")))
                        cert_por_ent.setdefault(ent_c, {})
                        for col_df, col_d in campos_por_entidad.items():
                            if col_df in df_act.columns:
                                cert_por_ent[ent_c][col_d] = cert_por_ent[ent_c].get(col_d, 0.0) + _nvl(fila_c.get(col_df, 0))

                    todas_ents = set(exo_ent.keys()) | set(cert_por_ent.keys())
                    for col_d in campos_por_entidad.values():
                        tot_conciliado = sum(
                            max(exo_ent.get(e_k, {}).get(col_d, 0.0), cert_por_ent.get(e_k, {}).get(col_d, 0.0))
                            for e_k in todas_ents
                        )
                        if tot_conciliado > 0:
                            d[col_d] = tot_conciliado

                    for col_df, col_d in mapeo_campos:
                        if col_df not in campos_por_entidad and col_df in df_act.columns:
                            v_extra = float(df_act[col_df].apply(_nvl).sum())
                            if v_extra > 0:
                                d[col_d] = max(d[col_d], v_extra)
                else:
                    for campo_df, campo_d in mapeo_campos:
                        if campo_df in df_act.columns:
                            suma_segura = float(df_act[campo_df].apply(_nvl).sum())
                            if suma_segura > 0:
                                d[campo_d] = suma_segura

                if d["ing_brutos_capital"] > 0 and d.get("auto_comp_inflacionario", True):
                    d["incrngo_capital"] = float(redondear_dian(d["ing_brutos_capital"] * pct_inf_act))

                st.session_state.mem_cert = []
                st.session_state.avisos_cert = []
                st.session_state.toast_ok = "✅ Soportes conciliados por entidad sin duplicar valores de Exógena."
                st.session_state.form_id += 1
                st.rerun()

            if cc1.button("🔄 Solo Reemplazar con Certificados", use_container_width=True):
                for campo_df, campo_d in mapeo_campos:
                    if campo_df in df_act.columns:
                        suma_segura = float(df_act[campo_df].apply(_nvl).sum())
                        if suma_segura > 0:
                            d[campo_d] = suma_segura
                if d["ing_brutos_capital"] > 0 and (d.get("auto_comp_inflacionario", True) or _nvl(df_act.get("componente_inflacionario_cdt", pd.Series([0])).sum()) == 0):
                    d["incrngo_capital"] = float(redondear_dian(d["ing_brutos_capital"] * pct_inf_act))
                st.session_state.mem_cert = []
                st.session_state.avisos_cert = []
                st.session_state.toast_ok = "✅ Casillas actualizadas únicamente con los certificados seleccionados."
                st.session_state.form_id += 1
                st.rerun()

            if cc2.button("➕ Sumar Manualmente", use_container_width=True):
                for campo_df, campo_d in mapeo_campos:
                    if campo_df in df_act.columns:
                        d[campo_d] += float(df_act[campo_df].apply(_nvl).sum())
                if d["ing_brutos_capital"] > 0 and d.get("auto_comp_inflacionario", True):
                    d["incrngo_capital"] = float(redondear_dian(d["ing_brutos_capital"] * pct_inf_act))
                st.session_state.mem_cert = []
                st.session_state.avisos_cert = []
                st.session_state.toast_ok = "✅ Certificados sumados y componente inflacionario sincronizado."
                st.session_state.form_id += 1
                st.rerun()

# ==============================================================================
# PASO 2: PERFIL DIAN, PATRIMONIO Y DEUDAS
# ==============================================================================
with t2:
    i1, i2, i3, i4 = st.columns([2.2, 1.2, 1.3, 1.8])
    d["nombre_declarante"] = i1.text_input("Nombre Completo (Casillas 7-10)", value=d["nombre_declarante"], key=f"nm_{fid}")
    d["nit_declarante"] = i2.text_input("Cédula / NIT sin DV (Casilla 5)", value=d["nit_declarante"], key=f"nt_{fid}")
    i2.caption(f"DV automático (Casilla 6): **{calcular_dv(d['nit_declarante']) or '-'}**")
    d["seccional_dian"] = i3.selectbox(
        "Seccional DIAN (Casilla 12)", SECCIONALES_DIAN,
        index=SECCIONALES_DIAN.index(d["seccional_dian"]) if d["seccional_dian"] in SECCIONALES_DIAN else 0,
        key=f"sc_{fid}",
    )
    cods_c = [x.split(" - ")[0] for x in ACTIVIDADES_CIIU]
    sel_c = i4.selectbox(
        "Actividad Económica CIIU (Casilla 24)", ACTIVIDADES_CIIU,
        index=cods_c.index(d["cod_actividad"]) if d["cod_actividad"] in cods_c else len(ACTIVIDADES_CIIU) - 1,
        key=f"ci_{fid}",
    )
    d["cod_actividad"] = (
        i4.text_input("Código CIIU manual", value=d["cod_actividad"], key=f"cm_{fid}")
        if sel_c.startswith("OTRO") else sel_c.split(" - ")[0]
    )

    st.divider()
    p1, p2 = st.columns(2)
    with p1:
        st.subheader(f"Patrimonio y Deudas a 31 de Diciembre de {anio_gravable}")
        d["efectivo"] = st.number_input("🏦 Saldos en cuentas bancarias, Nequi, Daviplata y efectivo", min_value=0.0, value=float(d["efectivo"]), step=100_000.0, key=f"ef_{fid}")
        d["otros_activos"] = st.number_input("📈 CDTs vigentes a dic 31, fiducias, acciones y cuentas por cobrar", min_value=0.0, value=float(d["otros_activos"]), step=200_000.0, key=f"oa_{fid}")
        d["inmuebles"] = st.number_input("🏠 Bienes Raíces: Casas, lotes, apartamentos (Mayor entre escritura y predial)", min_value=0.0, value=float(d["inmuebles"]), step=1_000_000.0, key=f"in_{fid}")
        d["vehiculos"] = st.number_input("🚗 Vehículos, motos y otros bienes (Precio de compra)", min_value=0.0, value=float(d["vehiculos"]), step=500_000.0, key=f"vh_{fid}")
        d["deudas"] = st.number_input("💳 Total Deudas vigentes a dic 31 (Hipotecas, bancos, tarjetas — Casilla 30)", min_value=0.0, value=float(d["deudas"]), step=200_000.0, key=f"de_{fid}")
    with p2:
        st.subheader("Topes para Declarar Renta y Control Patrimonial")
        d["consignaciones"] = st.number_input("Movimientos / Consignaciones bancarias del año (Tope 4)", min_value=0.0, value=float(d["consignaciones"]), step=500_000.0, key=f"cn_{fid}")
        d["tarjetas_credito"] = st.number_input("Consumos con Tarjeta de Crédito (Tope 3)", min_value=0.0, value=float(d["tarjetas_credito"]), step=500_000.0, key=f"tc_{fid}")
        d["compras_totales"] = st.number_input("Compras y consumos totales del año (Tope 5)", min_value=0.0, value=float(d["compras_totales"]), step=500_000.0, key=f"cp_{fid}")
        d["responsable_iva"] = st.toggle("¿Responsable de IVA a 31 de diciembre?", value=bool(d["responsable_iva"]), key=f"iv_{fid}")
        d["patrimonio_liquido_anterior"] = st.number_input("Patrimonio Líquido declaración año anterior (Opcional — Art. 236)", min_value=0.0, value=float(d["patrimonio_liquido_anterior"]), step=500_000.0, key=f"pa_{fid}")

# ==============================================================================
# PASO 3: INGRESOS DEL AÑO (TODAS LAS CÉDULAS DIAN)
# ==============================================================================
with t3:
    g1, g2 = st.columns(2)
    with g1:
        st.subheader("👔 1. Salarios / Empleados (Formulario 220)")
        d["ing_brutos_trabajo"] = st.number_input("Salarios y pagos laborales brutos (Casilla 32)", min_value=0.0, value=float(d["ing_brutos_trabajo"]), step=200_000.0, key=f"it_{fid}")
        d["salud_pension_trabajo"] = st.number_input("Salud + Pensión descontada por empleador (Casilla 33)", min_value=0.0, value=float(d["salud_pension_trabajo"]), step=50_000.0, key=f"spt_{fid}")
        d["cesantias_exentas"] = st.number_input("Cesantías e intereses de cesantías exentas", min_value=0.0, value=float(d["cesantias_exentas"]), step=100_000.0, key=f"ce_{fid}")

        st.divider()
        st.subheader("💼 2. Independientes: Servicios y Honorarios")
        d["ing_brutos_honorarios"] = st.number_input("Ingresos por servicios u honorarios", min_value=0.0, value=float(d["ing_brutos_honorarios"]), step=200_000.0, key=f"ih_{fid}")
        d["salud_pension_honorarios"] = st.number_input("Aportes anuales PILA EPS (Salud) + Pensión", min_value=0.0, value=float(d["salud_pension_honorarios"]), step=50_000.0, key=f"sph_{fid}")
        ca1, ca2 = st.columns(2)
        d["arl_honorarios"] = ca1.number_input("ARL Riesgos Laborales (Suma a INCRNGO C.33/C.44)", min_value=0.0, value=float(d["arl_honorarios"]), step=10_000.0, key=f"arl_{fid}")
        d["costos_honorarios"] = ca2.number_input("Costos y gastos con factura (C.45)", min_value=0.0, value=float(d["costos_honorarios"]), step=100_000.0, key=f"ch_{fid}")

    with g2:
        st.subheader("🏦 3. Rentas de Capital: CDTs, Intereses y Arriendos")
        d["ing_brutos_capital"] = st.number_input("Intereses CDT o ingresos por arriendos (Casilla 58)", min_value=0.0, value=float(d["ing_brutos_capital"]), step=50_000.0, key=f"ic_{fid}")
        pct_i = COMP_INFLACIONARIO.get(anio_gravable, 0.5544)
        d["auto_comp_inflacionario"] = st.toggle(
            f"🔄 Sincronizar automáticamente Componente Inflacionario ({pct_i*100:.2f}%) sobre CDT/Intereses",
            value=bool(d.get("auto_comp_inflacionario", True)),
            help="Desactívalo únicamente si en Casilla 58 incluiste ingresos por arrendamientos o certificados mixtos.",
            key=f"aci_{fid}",
        )
        if d["auto_comp_inflacionario"] and d["ing_brutos_capital"] > 0 and d["costos_capital"] == 0:
            d["incrngo_capital"] = float(redondear_dian(d["ing_brutos_capital"] * pct_i))
            st.caption(f"✅ Casilla 59 sincronizada automáticamente al {pct_i*100:.2f}%: **${d['incrngo_capital']:,.0f}**")
        else:
            d["incrngo_capital"] = st.number_input("Componente Inflacionario no gravado CDT manual (Casilla 59)", min_value=0.0, value=float(d["incrngo_capital"]), step=50_000.0, key=f"inc_{fid}")
        d["costos_capital"] = st.number_input("Gastos procedentes de arriendos (Casilla 60)", min_value=0.0, value=float(d["costos_capital"]), step=50_000.0, key=f"cc_{fid}")

        st.divider()
        with st.expander("🛒 Comercio / No Laborales, Pensiones, Dividendos y Venta de Activos"):
            d["ing_brutos_no_laborales"] = st.number_input("Ventas / Ingresos No Laborales (Casilla 74)", min_value=0.0, value=float(d["ing_brutos_no_laborales"]), step=200_000.0, key=f"inl_{fid}")
            d["costos_no_laborales"] = st.number_input("Costos No Laborales (Casilla 77)", min_value=0.0, value=float(d["costos_no_laborales"]), step=100_000.0, key=f"cnl_{fid}")
            d["ing_brutos_pensiones"] = st.number_input("Mesadas Pensionales (Casilla 99)", min_value=0.0, value=float(d["ing_brutos_pensiones"]), step=200_000.0, key=f"ip_{fid}")
            d["salud_pensiones"] = st.number_input("Salud en Pensión (Casilla 100)", min_value=0.0, value=float(d["salud_pensiones"]), step=50_000.0, key=f"spp_{fid}")
            d["dividendos_gravables"] = st.number_input("Dividendos Gravables (Casilla 107)", min_value=0.0, value=float(d["dividendos_gravables"]), step=100_000.0, key=f"dv_{fid}")
            d["ing_ganancia_ocasional"] = st.number_input("Venta casa/lote/vehículo (>=2 años) o herencia (Casilla 112)", min_value=0.0, value=float(d["ing_ganancia_ocasional"]), step=1_000_000.0, key=f"goi_{fid}")
            d["costos_ganancia_ocasional"] = st.number_input("Costo fiscal del activo vendido (Casilla 113)", min_value=0.0, value=float(d["costos_ganancia_ocasional"]), step=1_000_000.0, key=f"goc_{fid}")
            d["exenta_ganancia_ocasional"] = st.number_input("Ganancia ocasional exenta (Casilla 114)", min_value=0.0, value=float(d["exenta_ganancia_ocasional"]), step=500_000.0, key=f"goe_{fid}")
            d["ing_loterias_rifas"] = st.number_input("Loterías, rifas o apuestas (20%)", min_value=0.0, value=float(d["ing_loterias_rifas"]), step=500_000.0, key=f"lot_{fid}")

# ==============================================================================
# PASO 4: BENEFICIOS LEY 2277, RETENCIONES Y SANCIONES
# ==============================================================================
with t4:
    slot_bolsa = st.container()
    b1, b2, b3 = st.columns(3)
    with b1:
        st.markdown("#### ⭐ Alivios Fuera del Tope 40%")
        d["num_dependientes"] = st.number_input("Número de Dependientes (Casilla 138)", min_value=0, max_value=10, value=int(d["num_dependientes"]), step=1, key=f"dp_{fid}")
        d["compras_factura_electronica"] = st.number_input("Base Compras Factura Electrónica", min_value=0.0, value=float(d["compras_factura_electronica"]), step=100_000.0, key=f"fe_{fid}")
        d["aplicar_beneficio_fe"] = st.toggle("✅ Aplicar 1% en Casilla 28 (Pago electrónico)", value=bool(d["aplicar_beneficio_fe"]), key=f"afe_{fid}")
    with b2:
        st.markdown("#### 🏥 Deducciones Dentro del 40%")
        d["intereses_vivienda"] = st.number_input("Intereses Crédito Vivienda (Casilla 38)", min_value=0.0, value=float(d["intereses_vivienda"]), step=100_000.0, key=f"vv_{fid}")
        d["medicina_prepagada"] = st.number_input("Medicina Prepagada / Póliza Salud", min_value=0.0, value=float(d["medicina_prepagada"]), step=100_000.0, key=f"mp_{fid}")
        d["intereses_icetex"] = st.number_input("Intereses ICETEX (Tope 100 UVT)", min_value=0.0, value=float(d["intereses_icetex"]), step=50_000.0, key=f"icx_{fid}")
        d["gmf_certificado_100"] = st.number_input("4x1000 GMF (100% certificado)", min_value=0.0, value=float(d["gmf_certificado_100"]), step=20_000.0, key=f"gm_{fid}")
        d["aportes_afc_fpv"] = st.number_input("Aportes Voluntarios AFC / FPV (Casilla 35)", min_value=0.0, value=float(d["aportes_afc_fpv"]), step=100_000.0, key=f"af_{fid}")
        d["donaciones_entidades"] = st.number_input("Donaciones certificadas (Descuento 25%)", min_value=0.0, value=float(d["donaciones_entidades"]), step=100_000.0, key=f"dn_{fid}")
    with b3:
        st.markdown("#### 💰 Retenciones y Saldos")
        d["retenciones_anio"] = st.number_input("Retenciones del Año (Casilla 132)", min_value=0.0, value=float(d["retenciones_anio"]), step=10_000.0, key=f"rt_{fid}")
        d["saldo_favor_anterior"] = st.number_input("Saldo a Favor Año Anterior (Casilla 131)", min_value=0.0, value=float(d["saldo_favor_anterior"]), step=50_000.0, key=f"sf_{fid}")
        d["anticipo_anio_anterior"] = st.number_input("Anticipo Pagado Año Anterior (Casilla 130)", min_value=0.0, value=float(d["anticipo_anio_anterior"]), step=50_000.0, key=f"aa_{fid}")
        d["impuesto_neto_anterior"] = st.number_input("Impuesto Neto Año Anterior (Promediar anticipo)", min_value=0.0, value=float(d["impuesto_neto_anterior"]), step=50_000.0, key=f"ia_{fid}")
        d["meses_extemporaneidad"] = st.number_input("Meses de retraso si es extemporánea (Casilla 135)", min_value=0, max_value=60, value=int(d["meses_extemporaneidad"]), step=1, key=f"mr_{fid}")
        d["aplica_reduccion_640"] = st.checkbox("Reducir sanción al 50% (Art. 640 E.T.)", value=bool(d["aplica_reduccion_640"]), key=f"r6_{fid}")

# CÁLCULO FINAL EN VIVO PARA TODA LA APP
r = liquidar_210(d, valor_uvt, uvt_pat, uvt_ing, uvt_san, anio_grav=anio_gravable)

with slot_bolsa:
    pct_b = min(1.0, (r["usado_40"] / r["lim_40"]) if r["lim_40"] > 0 else 0.0)
    st.progress(pct_b, text=f"Bolsa 40% usada: ${r['usado_40']:,.0f} de ${r['lim_40']:,.0f} (Cupo libre: ${r['libre_40']:,.0f})")

with slot_kpis:
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Obligación DIAN", "SÍ OBLIGADO" if r["obligado"] else "NO OBLIGADO", delta=f"{len(r['causales'])} causal(es)" if r["obligado"] else "Voluntaria")
    k2.metric("Patrimonio Líquido (C.31)", f"${r['c31']:,.0f}")
    k3.metric("Renta Líquida Grav. (C.111)", f"${r['c111']:,.0f}", f"{r['rlg_uvt']} UVT")
    k4.metric("Impuesto a Cargo (C.129)", f"${r['c129']:,.0f}")
    if r["c137"] > 0:
        k5.metric("🎉 SALDO A FAVOR (C.137)", f"${r['c137']:,.0f}", delta="Devolución / Compensación")
    else:
        k5.metric("💳 SALDO A PAGAR (C.136)", f"${r['c136']:,.0f}")

# ==============================================================================
# PASO 5: FORMULARIO 210 OFICIAL, DESCARGAS Y CONSULTOR IA
# ==============================================================================
with t5:
    if r["c137"] > 0:
        st.success(f"### 🎉 TOTAL SALDO A FAVOR (Casilla 137): ${r['c137']:,.0f}\n**Estado:** {r['estado']} • **Estrategia:** {r['estrategia']}")
    elif r["c136"] > 0:
        st.warning(f"### 💳 TOTAL SALDO A PAGAR (Casilla 136): ${r['c136']:,.0f}\n**Estado:** {r['estado']} • **Estrategia:** {r['estrategia']}")
    else:
        st.info(f"### ✅ Total Saldo a Pagar: $0 — {r['estado']}")

    for al in r["alertas"]:
        st.error(f"🚨 {al}")
    for sg in r["sugerencias"]:
        st.info(f"💡 {sg}")

    solo_con_valor = st.checkbox("Ver solo casillas con valores mayores a $0 (Lectura rápida Muisca)", value=False)
    filas_v = [
        x for x in filas_formulario_210(r)
        if (not solo_con_valor or x[2] > 0 or x[0] in ("29", "31", "97", "111", "129", "136", "137"))
    ]
    st.dataframe(
        pd.DataFrame(
            [(c, n, f"{v_c:,.0f}" if c == "138" else f"${v_c:,.0f}") for c, n, v_c in filas_v],
            columns=["Casilla Muisca", "Concepto Oficial Formulario 210", "Valor Oficial DIAN"],
        ),
        use_container_width=True, hide_index=True, height=460,
    )

    ex_b, wd_b = st.columns(2)
    ex_b.download_button(
        "📥 Descargar Excel Oficial (.xlsx)",
        data=generar_excel(r, d, anio_gravable, df_pila_exp),
        file_name=f"Formulario_210_AG{anio_gravable}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    wd_b.download_button(
        "📄 Descargar Dictamen Word (.docx)",
        data=generar_word(r, d, anio_gravable, valor_uvt),
        file_name=f"Dictamen_Renta_AG{anio_gravable}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

    st.divider()
    st.subheader("⚖️ Consultor Tributario IA con Gemini")
    for m in st.session_state.chat:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])

    preg = st.chat_input("Haz cualquier pregunta sobre tu declaración o casillas del Formulario 210...")
    if preg:
        if not api_key:
            st.error("Configura tu clave en Secrets o pégala en la barra lateral.")
        else:
            st.session_state.chat.append({"role": "user", "content": preg})
            with st.chat_message("user"):
                st.markdown(preg)
            with st.chat_message("assistant"):
                try:
                    with st.spinner("Analizando Estatuto Tributario..."):
                        sis = (
                            f"Eres Contador Tributarista Senior experto en DIAN Colombia. "
                            f"AG {anio_gravable}, UVT ${valor_uvt:,.0f}. Recuerda que según Concepto Unificado DIAN "
                            f"0912 de 2018 (Descriptor 1.12) y Arts. 55-56 E.T., EPS + Pensión + ARL pagados por el "
                            f"independiente van como Ingreso No Constitutivo de Renta (Casilla 33 o 44), y los "
                            f"rendimientos financieros tienen componente inflacionario no gravado del {COMP_INFLACIONARIO.get(anio_gravable, 0.5544)*100:.2f}% "
                            f"(Casilla 59). Liquidación actual: {json.dumps(r, ensure_ascii=False)}."
                        )
                        historial = [
                            types.Content(
                                role="user" if msg["role"] == "user" else "model",
                                parts=[types.Part.from_text(text=msg["content"])],
                            )
                            for msg in st.session_state.chat[-8:]
                        ]
                        ans = llamar_gemini(api_key, historial, system_instruction=sis, temperatura=0.2)
                        st.markdown(ans)
                        st.session_state.chat.append({"role": "assistant", "content": ans})
                except Exception as e:
                    st.error(f"⚠️ {e}")
