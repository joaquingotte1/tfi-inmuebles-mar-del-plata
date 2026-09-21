"""Scraper de avisos de venta publicados en Mitula para Mar del Plata."""
from __future__ import annotations

import argparse
import asyncio
import base64
import logging
import re
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.async_api import async_playwright

BASE = "https://casas.mitula.com.ar/casas/st-partido-de-general-pueyrredon-buenos-aires"
SELECTOR_TARJETA = "article.listing, article[data-listingid], [data-listingid][data-price]"


class VerificacionMitulaBloqueada(RuntimeError):
    """Mitula mostró una verificación que no puede completarse en ese momento."""

NUMEROS_PUBLICADOS = {
    "cero": 0.0, "un": 1.0, "uno": 1.0, "una": 1.0, "dos": 2.0,
    "tres": 3.0, "cuatro": 4.0, "cinco": 5.0, "seis": 6.0,
    "siete": 7.0, "ocho": 8.0, "nueve": 9.0, "diez": 10.0,
    "once": 11.0, "doce": 12.0, "trece": 13.0, "catorce": 14.0,
    "quince": 15.0,
}
TOKEN_CANTIDAD = rf"(?:\d+(?:[.,]\d+)?|{'|'.join(NUMEROS_PUBLICADOS)})"
SINONIMOS_AMENITIES = {
    "tiene_cochera": ("garage", "garaje"),
    "tiene_calefaccion": ("calefactor", "calefactores", "losa radiante", "radiadores"),
    "tiene_piscina": ("pileta",),
    "tiene_baulera": ("baulera incorporada",),
    "tiene_cloacas": ("cloaca", "servicio de cloacas"),
}
AMENITIES = {
    "agua": "tiene_agua", "aire acondicionado": "tiene_aire_acondicionado",
    "balcón": "tiene_balcon", "calefacción": "tiene_calefaccion",
    "chimenea": "tiene_chimenea", "cochera": "tiene_cochera",
    "electricidad": "tiene_electricidad", "gas natural": "tiene_gas_natural",
    "internet": "tiene_internet", "parrilla": "tiene_parrilla",
    "patio": "tiene_patio", "piscina": "tiene_piscina", "terraza": "tiene_terraza",
    "jardín": "tiene_jardin", "ascensor": "tiene_ascensor", "seguridad": "tiene_seguridad",
    "wifi": "tiene_wifi", "lavadero": "tiene_lavadero",
    "cocina equipada": "tiene_cocina_equipada", "cocina integral": "tiene_cocina_integral",
    "cuarto de servicio": "tiene_cuarto_servicio", "vista panorámica": "tiene_vista_panoramica",
    "bodega": "tiene_bodega", "baulera": "tiene_baulera", "cloacas": "tiene_cloacas",
}
BARRIOS = {
    r"bosque peralta ramos": "Bosque Peralta Ramos", r"colinas de peralta ramos": "Colinas de Peralta Ramos",
    r"parque luro": "Parque Luro", r"plaza col.n": "Plaza Colón", r"punta mogotes": "Punta Mogotes",
    r"playa grande": "Playa Grande", r"los troncos": "Los Troncos", r"punta iglesia": "Punta Iglesia",
    r"faro norte": "Faro Norte", r"nueva pompeya": "Nueva Pompeya", r"santa m.nica": "Santa Mónica",
    r"constituci.n": "Constitución", r"la perla": "La Perla", r"acantilados": "Acantilados",
    r"stella maris": "Stella Maris", r"parque camet": "Parque Camet", r"los pinares": "Los Pinares",
    r"san jos.": "San José", r"san juan": "San Juan", r"don bosco": "Don Bosco",
    r"chauv.n": "Chauvín", r"g.emes": "Güemes", r"varese": "Varese", r"alem": "Alem",
    r"caisamar": "Caisamar", r"las avenidas": "Las Avenidas", r"puerto": "Puerto",
    r"terminal": "Terminal", r"zacagnini": "Zacagnini", r"macrocentro": "Macrocentro",
    r"centro": "Centro", r"alfar": "Alfar",
}
COLUMNAS = [
    "id_aviso", "tipo_propiedad", "operacion", "precio", "moneda",
    "precio_publicado", "superficie_publicada",
    "superficie_total_m2", "superficie_cubierta_m2", "superficie_descubierta_m2",
    "superficie_semicubierta_m2", "superficie_terreno_m2",
    "ambientes", "dormitorios", "banos", "toilettes", "cocheras", "antiguedad_anios",
    "anio_construccion", "piso", "unidad", "cantidad_plantas",
    "frente_terreno_m", "fondo_terreno_m", "expensas", "apto_credito", "apto_profesional", "escritura",
    "estado", "disposicion", "orientacion", "ubicacion", "ciudad", "partido", "provincia",
    "latitud", "longitud", "tipo_gas", "portal_origen", "id_aviso_origen",
    "barrio", "barrio_fuente", "direccion", "circunscripcion", "seccion_catastral", "manzana", "parcela", "chacra",
    "titulo", "descripcion", "fecha_publicacion", "actualizacion_relativa", "anunciante",
    "publicacion_destacada", "amoblamiento",
    *dict.fromkeys(AMENITIES.values()),
    "url_aviso", "url_listado", "pagina_origen", "posicion_listado",
    "fecha_scraping_utc", "detalle_extraido", "detalle_http_status", "detalle_error",
    "texto_listado_crudo", "texto_detalle_crudo",
]


def numero(texto: str | None) -> float | None:
    if not texto:
        return None
    encontrado = re.search(r"[-+]?\d[\d.,]*", texto.strip())
    if not encontrado:
        return None
    valor = encontrado.group()
    signo = "-" if valor.startswith("-") else ""
    valor = valor.lstrip("+-")
    if "." in valor and "," in valor:
        decimal = "." if valor.rfind(".") > valor.rfind(",") else ","
        miles = "," if decimal == "." else "."
        valor = valor.replace(miles, "").replace(decimal, ".")
    elif "." in valor or "," in valor:
        separador = "." if "." in valor else ","
        partes = valor.split(separador)
        if len(partes) > 2 or (len(partes) == 2 and len(partes[1]) == 3):
            valor = "".join(partes)
        else:
            valor = ".".join(partes)
    try:
        return float(signo + valor)
    except ValueError:
        return None


def precio_atributo(texto: str | None) -> float | None:
    if not texto:
        return None
    limpio = texto.strip()
    if re.fullmatch(r"\d{1,3}(?:,\d{3})+", limpio):
        return float(limpio.replace(",", ""))
    return numero(limpio)


def limpiar(texto: str | None) -> str | None:
    if not texto:
        return texto
    reemplazos = {
        "Pueyrred�n": "Pueyrredón", "m�": "m²", "Ba�o": "Baño", "ba�o": "baño",
        "Balc�n": "Balcón", "Calefacci�n": "Calefacción", "Jard�n": "Jardín",
        "Habitaci�n": "Habitación", "Descripci�n": "Descripción",
        "Direcci�n": "Dirección", "A�o": "Año", "construcci�n": "construcción",
        "Operaci�n": "Operación", "cr�dito": "crédito", "Caracter�sticas": "Características",
        "Ubicaci�n": "Ubicación", "Secci�n": "Sección", "Circunscripci�n": "Circunscripción",
        "Col�n": "Colón", "Chauv�n": "Chauvín", "G�emes": "Güemes", "Constituci�n": "Constitución",
    }
    for original, corregido in reemplazos.items():
        texto = texto.replace(original, corregido)
    return texto


def extraer(texto: str, patron: str) -> float | None:
    encontrado = re.search(patron, texto, re.I)
    return numero(encontrado.group(1)) if encontrado else None


def numero_publicado(texto: str | None) -> float | None:
    if not texto:
        return None
    limpio = texto.strip().lower()
    return NUMEROS_PUBLICADOS.get(limpio, numero(limpio))


def cantidad_desde_texto(
    texto: str, sustantivo: str, *, singular_implicito: bool = False
) -> float | None:
    """Extrae cantidades expresas en dígitos o palabras; no estima cantidades."""
    patrones = (
        rf"\b({TOKEN_CANTIDAD})\s+(?:{sustantivo})\b",
        rf"(?:^|\n)\s*[-•]?\s*(?:{sustantivo})\s*:\s*({TOKEN_CANTIDAD})\b",
    )
    for patron in patrones:
        encontrado = re.search(patron, texto, re.I)
        if encontrado:
            return numero_publicado(encontrado.group(1))
    if singular_implicito and re.search(rf"\b(?:{sustantivo})\b", texto, re.I):
        return 1.0
    return None


def cantidad_cocheras(texto: str) -> float | None:
    encontrado = re.search(
        rf"\b(?:cochera|garage|garaje)\b[^.\n]{{0,50}}\bpara\s+({TOKEN_CANTIDAD})\s+autos?\b",
        texto,
        re.I,
    )
    if encontrado:
        return numero_publicado(encontrado.group(1))
    return cantidad_desde_texto(
        texto, r"cocheras?|garages?|garajes?", singular_implicito=True
    )


def piso_publicado(texto: str) -> float | None:
    encontrado = re.search(rf"\bpiso\s*:?[ ]*({TOKEN_CANTIDAD})\b", texto, re.I)
    return numero_publicado(encontrado.group(1)) if encontrado else None


def antiguedad_publicada(texto: str) -> float | None:
    encontrado = re.search(
        rf"\bAntig[uü]edad\s*:\s*({TOKEN_CANTIDAD})\b", texto, re.I
    )
    if encontrado:
        return numero_publicado(encontrado.group(1))
    campo = campo_etiquetado(texto, "Antigüedad", "Antiguedad")
    if campo:
        return numero_publicado(campo)
    patrones = (
        rf"\b({TOKEN_CANTIDAD})\s+a[nñ]os?\s+de\s+antig[uü]edad\b",
        rf"\bantig[uü]edad\s+de\s+({TOKEN_CANTIDAD})\s+a[nñ]os?\b",
        rf"\bedificio\s+de\s+({TOKEN_CANTIDAD})\s+a[nñ]os?\b",
    )
    for patron in patrones:
        encontrado = re.search(patron, texto, re.I)
        if encontrado:
            return numero_publicado(encontrado.group(1))
    return None


def anio_construccion_publicado(texto: str) -> float | None:
    encontrado = re.search(
        r"\bA[nñ]o(?: de)? construcci[oó]n\s*:\s*([12]\d{3})\b", texto, re.I
    )
    if encontrado:
        return float(encontrado.group(1))
    campo = campo_etiquetado(texto, "Año de construcción", "Año construcción")
    if campo:
        encontrado = re.search(r"\b([12]\d{3})\b", campo)
        return float(encontrado.group(1)) if encontrado else None
    encontrado = re.search(
        r"\b(?:construid[oa]|construcci[oó]n)\s+(?:en|del\s+a[nñ]o)?\s*([12]\d{3})\b",
        texto,
        re.I,
    )
    return float(encontrado.group(1)) if encontrado else None


def extraer_texto(texto: str, patron: str) -> str | None:
    encontrado = re.search(patron, texto, re.I)
    return encontrado.group(1).strip(" .") if encontrado else None


def campo_etiquetado(texto: str, *etiquetas: str) -> str | None:
    """Obtiene el valor publicado después de una etiqueta, sin inferirlo."""
    if not texto:
        return None
    union = "|".join(re.escape(etiqueta) for etiqueta in etiquetas)
    siguiente_etiqueta = (
        r"Barrio|Direcci[oó]n|Escritura|Estado|Disposici[oó]n|Orientaci[oó]n|"
        r"Antig[uü]edad|Dormitorios|Ba[nñ]os|Sup(?:erficie)?\.?|Circunscripci[oó]n|"
        r"Secci[oó]n|Manzana|Parcela|Chacra"
    )
    encontrado = re.search(
        rf"(?:^|\n)\s*[-•]?\s*(?:{union})\b\s*:?\s*(.+?)(?=\s+(?:{siguiente_etiqueta})\s*:|\n|$)",
        texto,
        re.I,
    )
    return encontrado.group(1).strip(" .") if encontrado else None


def extraer_descripcion(texto: str) -> str | None:
    if not texto:
        return None
    patrones = (
        r"(?:^|\n)\s*Descripci[oó]n(?: de la propiedad)?\s*:?\s*([\s\S]+?)(?=\n\s*(?:Caracter[ií]sticas|Ubicaci[oó]n|Servicios|Comodidades|Anunciante)\b|$)",
        r"(?:^|\n)\s*Descripci[oó]n\s*:?\s*([\s\S]+)$",
    )
    for patron in patrones:
        encontrado = re.search(patron, texto, re.I)
        if encontrado:
            valor = encontrado.group(1).strip()
            return valor or None
    return None


def unidad_publicada(texto: str) -> str | None:
    encontrado = re.search(
        r"(?:^|\n)\s*[-•]?\s*(?:Unidad(?: funcional)?|UF)\s*:\s*([^\n]+)", texto, re.I
    )
    if not encontrado:
        encontrado = re.search(r"\bunidad funcional\s+(?:n[°º]?\s*)?([\w-]+)", texto, re.I)
    return encontrado.group(1).strip(" .") if encontrado else None


def plantas_publicadas(texto: str) -> float | None:
    campo = campo_etiquetado(texto, "Cantidad de plantas", "Plantas")
    if campo:
        return numero(campo)
    patrones = (
        rf"\b(?:casa|chalet|propiedad)\b[^.\n]{{0,50}}\b(?:de|con)\s+({TOKEN_CANTIDAD})\s+plantas?\b",
        rf"\bdesarrollad[oa]\s+en\s+({TOKEN_CANTIDAD})\s+plantas?\b",
    )
    for patron in patrones:
        encontrado = re.search(patron, texto, re.I)
        if encontrado:
            return numero_publicado(encontrado.group(1))
    return None


def medida_terreno(texto: str, lado: str) -> float | None:
    patrones = (
        rf"\b{lado}\s*:?[ ]*([\d.,]+)\s*(?:m|metros?)\b",
        rf"\b([\d.,]+)\s*(?:m|metros?)\s+de\s+{lado}\b",
    )
    for patron in patrones:
        valor = extraer(texto, patron)
        if valor is not None:
            return valor
    return None


def disposicion_publicada(texto: str) -> str | None:
    campo = campo_etiquetado(texto, "Disposición", "Disposicion")
    if campo:
        return campo
    encontrado = re.search(
        r"\b(?:departamento|unidad|propiedad)\b[^.\n]{0,60}\bdispuest[oa]\s+(?:al\s+)?(frente|contrafrente|intern[oa]|lateral)\b",
        texto,
        re.I,
    )
    return encontrado.group(1).capitalize() if encontrado else None


def orientacion_publicada(texto: str) -> str | None:
    campo = campo_etiquetado(texto, "Orientación", "Orientacion")
    if campo:
        return campo
    encontrado = re.search(
        r"\borientaci[oó]n\s+(norte|sur|este|oeste|noreste|noroeste|sudeste|sudoeste)\b",
        texto,
        re.I,
    )
    return encontrado.group(1).capitalize() if encontrado else None


def estado_publicado(texto: str) -> str | None:
    campo = campo_etiquetado(texto, "Estado de la propiedad", "Estado")
    if campo:
        return campo
    patrones = (
        r"\ba estrenar\b", r"\ben construcci[oó]n\b", r"\ba refaccionar\b",
        r"\brefaccionad[oa]\b", r"\breciclad[oa]\b", r"\bimpecable estado\b",
        r"\bexcelente estado\b", r"\bmuy buen estado\b", r"\bbuen estado\b",
        r"\bestado original\b",
    )
    valores = []
    for patron in patrones:
        encontrado = re.search(patron, texto, re.I)
        if encontrado:
            nuevo = encontrado.group(0)
            if not any(nuevo.lower() in existente.lower() for existente in valores):
                valores.append(nuevo)
    return " | ".join(valores) if valores else None


def tipo_gas_publicado(texto: str) -> str | None:
    if re.search(r"\b(?:sin gas|no (?:tiene|posee|cuenta con) gas)\b", texto, re.I):
        return "Sin gas"
    if re.search(r"\bgas (?:natural|de red)|\bred de gas\b|\bgas,? red\b", texto, re.I):
        return "Gas natural / red"
    if re.search(r"\bgas envasado|\bgarrafa", texto, re.I):
        return "Gas envasado"
    return campo_etiquetado(texto, "Tipo de gas", "Gas")


def detectar_barrio(texto: str) -> str | None:
    return next((nombre for patron, nombre in BARRIOS.items() if re.search(rf"\b{patron}\b", texto, re.I)), None)


def normalizar_barrio(campo: str | None, texto: str) -> tuple[str | None, str | None]:
    desde_campo = detectar_barrio(campo or "")
    if desde_campo:
        return desde_campo, "campo_ficha"
    if campo and re.search(r"zona c.ntrica", campo, re.I):
        return "Centro", "campo_ficha"
    desde_texto = detectar_barrio(texto)
    if desde_texto:
        return desde_texto, "titulo_descripcion"
    if campo and len(campo) <= 35 and not re.search(r"excelente|residencial|cercan|unidad|metros|cuadras", campo, re.I):
        return campo, "campo_ficha"
    return None, None


def distancia_mar(texto: str) -> float | None:
    patrones = (
        r"(?:a|a solo|a sólo)\s+([\d.,]+)\s+cuadras?\s+(?:del?|de la)\s+(?:mar|costa|playa)",
        r"([\d.,]+)\s+cuadras?[^\n.]{0,35}(?:mar|costa|playa)",
    )
    for patron in patrones:
        valor = extraer(texto, patron)
        if valor is not None:
            return valor
    return None


def presencia_publicada(
    texto: str, etiqueta: str, variantes_extra: tuple[str, ...] = ()
) -> str | None:
    variantes = [etiqueta, *variantes_extra]
    if etiqueta == "amueblado":
        variantes.append("amueblar")
    if etiqueta == "piscina":
        variantes.append("pileta")
    union = "|".join(re.escape(x) for x in variantes)
    negativo = rf"(?:\bsin\s+(?:{union})|\bno\s+(?:tiene|posee|cuenta con|incluye)?\s*(?:{union})|\bcarece de\s+(?:{union}))"
    if re.search(negativo, texto, re.I):
        return "No"
    if re.search(rf"\b(?:{union})\b", texto, re.I):
        return "Sí"
    return None


def aptitud_credito(texto: str) -> str | None:
    if re.search(r"(?:no|sin)\s+(?:es\s+)?apto\s+(?:a\s+)?cr[eé]dito", texto, re.I):
        return "No"
    return "Sí" if re.search(r"apto\s+(?:a\s+)?cr[eé]dito", texto, re.I) else None


def estado_amoblamiento(texto: str) -> str | None:
    if re.search(r"parcialmente\s+amueblad", texto, re.I):
        return "Parcialmente amueblado"
    if re.search(r"sin\s+amueblar|no\s+amueblad", texto, re.I):
        return "Sin amueblar"
    if re.search(r"\bamueblad", texto, re.I):
        return "Amueblado"
    return None


def pagina_url(pagina: int) -> str:
    return BASE if pagina == 1 else f"{BASE}?page={pagina}"


async def navegar_con_reintentos(page, url: str, intentos: int, espera_ms: int):
    ultimo_error = None
    for intento in range(1, intentos + 1):
        try:
            respuesta = await page.goto(
                url, wait_until="domcontentloaded", timeout=60000
            )
            estado = respuesta.status if respuesta else None
            if estado not in {429, 500, 502, 503, 504}:
                return respuesta, None
            ultimo_error = f"HTTP {estado}"
        except Exception as error:
            ultimo_error = f"{type(error).__name__}: {str(error)[:300]}"
            if "TargetClosedError" in ultimo_error:
                return None, ultimo_error
        if intento < intentos:
            logging.warning(
                "Reintento %s/%s para %s por %s", intento + 1, intentos, url, ultimo_error
            )
            try:
                await page.wait_for_timeout(espera_ms * intento)
            except Exception as error:
                return None, f"{type(error).__name__}: {str(error)[:300]}"
    return None, ultimo_error


async def hay_verificacion_manual(page) -> bool:
    try:
        titulo = (await page.title()).lower()
        cuerpo = (await page.locator("body").inner_text(timeout=5000)).lower()
        if any(
            señal in f"{titulo}\n{cuerpo}"
            for señal in (
                "captcha", "no soy un robot", "verifica que eres humano",
                "verificar que eres humano", "security check", "human verification",
                "comprobando tu navegador", "checking your browser",
            )
        ):
            return True
        return await page.locator(
            'iframe[src*="captcha" i], iframe[src*="recaptcha" i], iframe[src*="hcaptcha" i], [class*="captcha" i]'
        ).count() > 0
    except Exception:
        return False


async def esperar_verificacion_manual(page, minutos: int) -> bool:
    if not await hay_verificacion_manual(page):
        return False
    logging.warning(
        "CAPTCHA detectado. Complete la verificación en la ventana de Edge; se esperará hasta %s minutos.",
        minutos,
    )
    intentos = max(1, minutos * 12)
    for intento in range(intentos):
        if await page.locator(SELECTOR_TARJETA).count() > 0:
            logging.info("Verificación completada; se reanuda el scraping")
            return True
        if intento and intento % 6 == 0:
            logging.info("Esperando que se complete el CAPTCHA en Edge...")
        await page.wait_for_timeout(5000)
    return False


def ruta_detalle(tarjeta: dict) -> str | None:
    if tarjeta.get("destino"):
        ruta = base64.b64decode(tarjeta["destino"]).decode("utf-8")
        return ruta if ruta.startswith("http") else f"https://casas.mitula.com.ar{ruta}"
    return tarjeta.get("href")


def dataframe_unico(filas: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(filas, columns=COLUMNAS)
    if df.empty:
        return df
    clave = df["id_aviso"].fillna(df["url_aviso"])
    sin_clave = clave.isna()
    clave.loc[sin_clave] = [f"sin_clave_{i}" for i in df.index[sin_clave]]
    return df.loc[~clave.duplicated()].reset_index(drop=True)


async def ejecutar(args: argparse.Namespace) -> tuple[pd.DataFrame, dict]:
    filas: list[dict] = []
    if args.reanudar and args.parcial.exists():
        try:
            previo = pd.read_excel(args.parcial, sheet_name="Propiedades")
            previo = previo.reindex(columns=COLUMNAS).astype(object)
            registros_previos = previo.where(pd.notna(previo), None).to_dict("records")
            pendientes = [
                fila for fila in registros_previos
                if fila.get("detalle_extraido") != "Sí" or fila.get("detalle_error")
            ]
            filas = [fila for fila in registros_previos if fila not in pendientes]
            logging.info(
                "Reanudación: %s registros completos cargados y %s pendientes de reintento desde %s",
                len(filas), len(pendientes), args.parcial,
            )
        except Exception as error:
            logging.warning("No se pudo reanudar %s: %s", args.parcial, type(error).__name__)
    vistos_id = {str(f["id_aviso"]) for f in filas if f.get("id_aviso")}
    vistos_url = {str(f["url_aviso"]) for f in filas if f.get("url_aviso")}
    resumen = {
        "estado": "en_progreso", "inicio_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "fin_utc": None, "paginas_procesadas": 0, "tarjetas_encontradas": 0,
        "avisos_aceptados_nuevos": 0, "duplicados_omitidos": 0,
        "fuera_de_alcance": 0, "detalles_exitosos": 0, "detalles_fallidos": 0,
        "errores_listado": 0, "registros_totales": len(filas),
    }
    if args.max_avisos and len(filas) >= args.max_avisos:
        resumen["estado"] = "completa"
        resumen["fin_utc"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
        return dataframe_unico(filas), resumen
    fecha_scraping_utc = datetime.now(timezone.utc).isoformat(timespec="seconds")
    firmas_paginas: set[tuple[str, ...]] = set()
    async with async_playwright() as p:
        args.perfil.mkdir(parents=True, exist_ok=True)
        contexto = await p.chromium.launch_persistent_context(
            user_data_dir=str(args.perfil.resolve()), channel="msedge", headless=args.headless,
            locale="es-AR", timezone_id="America/Argentina/Buenos_Aires",
            viewport={"width": 1366, "height": 900},
        )
        page = contexto.pages[0] if contexto.pages else await contexto.new_page()
        detalle = await contexto.new_page()
        for pagina in range(args.pagina_inicial, args.max_paginas + 1):
            url = pagina_url(pagina)
            logging.info("Listado: %s", url)
            respuesta, error_listado = await navegar_con_reintentos(
                page, url, args.reintentos, args.espera_reintento_ms
            )
            if error_listado:
                resumen["errores_listado"] += 1
                resumen["estado"] = "con_errores"
                logging.error("Listado omitido tras reintentos: %s", error_listado)
                break
            if respuesta and respuesta.status >= 400:
                logging.error("HTTP %s", respuesta.status)
                resumen["errores_listado"] += 1
                resumen["estado"] = "con_errores"
                break
            try:
                await page.locator(SELECTOR_TARJETA).first.wait_for(state="visible", timeout=30000)
            except Exception:
                if await page.locator(SELECTOR_TARJETA).count() == 0:
                    desbloqueada = False
                    if not args.headless:
                        desbloqueada = await esperar_verificacion_manual(
                            page, args.esperar_captcha_minutos
                        )
                    if not desbloqueada:
                        if await hay_verificacion_manual(page):
                            logging.error("CAPTCHA no resuelto dentro del tiempo de espera")
                            resumen["errores_listado"] += 1
                            resumen["estado"] = "con_errores"
                        else:
                            logging.info("Fin detectado: la página %s no contiene tarjetas", pagina)
                            resumen["estado"] = "completa"
                        break
            cantidad_anterior = 0
            for _ in range(8):
                await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                await page.wait_for_timeout(1200)
                cantidad_actual = await page.locator(SELECTOR_TARJETA).count()
                if cantidad_actual == cantidad_anterior and cantidad_actual >= 30:
                    break
                cantidad_anterior = cantidad_actual
            tarjetas = await page.locator(SELECTOR_TARJETA).evaluate_all("""els => els.map(e => ({
                id: e.dataset.listingid, tipo: e.querySelector('.listing-card__property-type')?.innerText || e.innerText.split('\\n')[0],
                operacion: e.dataset.operationType, precio: e.dataset.price, moneda: e.dataset.currency,
                dormitorios: e.dataset.rooms, superficie: e.dataset.floorarea, ubicacion: e.dataset.location,
                ciudad: e.dataset.city, partido: e.dataset.district, provincia: e.dataset.region,
                barrio: e.dataset.neighbourhood, altaDemanda: e.dataset.ishighdemand,
                destino: e.dataset.clickdestination, href: e.querySelector('a[href]')?.href || null,
                texto: e.innerText,
                titulo: e.querySelector('img[alt]')?.alt || e.querySelector('h2,h3,[class*="title"]')?.innerText || null
            }))""")
            logging.info("%s tarjetas encontradas", len(tarjetas))
            firma_pagina = tuple(
                str(tarjeta.get("id") or tarjeta.get("destino") or tarjeta.get("href"))
                for tarjeta in tarjetas
            )
            if firma_pagina in firmas_paginas:
                logging.info("Fin detectado: la página %s repite una página anterior", pagina)
                resumen["estado"] = "completa"
                break
            firmas_paginas.add(firma_pagina)
            resumen["paginas_procesadas"] += 1
            resumen["tarjetas_encontradas"] += len(tarjetas)
            for posicion, tarjeta in enumerate(tarjetas, 1):
                if (tarjeta.get("ciudad") or "").strip().lower() != "mar del plata":
                    resumen["fuera_de_alcance"] += 1
                    continue
                try:
                    url_previa = ruta_detalle(tarjeta)
                except Exception:
                    url_previa = tarjeta.get("href")
                if (tarjeta.get("id") and str(tarjeta["id"]) in vistos_id) or (
                    url_previa and str(url_previa) in vistos_url
                ):
                    resumen["duplicados_omitidos"] += 1
                    continue
                texto = limpiar("\n".join(x.strip() for x in tarjeta["texto"].splitlines() if x.strip()))
                lineas = texto.splitlines()
                banos = next((numero(x) for x in lineas if "baño" in x.lower()), None)
                superficie = numero(tarjeta.get("superficie"))
                precio = precio_atributo(tarjeta.get("precio"))
                titulo = limpiar(tarjeta.get("titulo"))
                tipo_titulo = re.search(r"\b(departamento|casa|ph|chale(?:t)?)\b", f"{titulo or ''}\n{texto}", re.I)
                tipo = tipo_titulo.group(1).title() if tipo_titulo else tarjeta.get("tipo")
                if tipo and tipo.lower() == "chale":
                    tipo = "Chalet"
                if tipo not in {"Casa", "Departamento", "Chalet"}:
                    continue
                texto_detalle = ""
                url_detalle = None
                detalle_http_status = None
                detalle_error = None
                latitud = None
                longitud = None
                if tarjeta.get("destino") or tarjeta.get("href"):
                    try:
                        if detalle.is_closed():
                            logging.warning("La pestaña de detalle se cerró; se crea una nueva")
                            detalle = await contexto.new_page()
                        url_detalle = ruta_detalle(tarjeta)
                        respuesta_detalle, detalle_error = await navegar_con_reintentos(
                            detalle, url_detalle, args.reintentos, args.espera_reintento_ms
                        )
                        if detalle_error and "TargetClosedError" in detalle_error:
                            detalle = await contexto.new_page()
                            respuesta_detalle, detalle_error = await navegar_con_reintentos(
                                detalle, url_detalle, args.reintentos, args.espera_reintento_ms
                            )
                        detalle_http_status = respuesta_detalle.status if respuesta_detalle else None
                        if respuesta_detalle and respuesta_detalle.status < 400:
                            if await hay_verificacion_manual(detalle):
                                raise VerificacionMitulaBloqueada(
                                    f"Security verification bloqueada en {url_detalle}"
                                )
                            await detalle.locator("body").wait_for(state="visible", timeout=30000)
                            texto_detalle = limpiar("\n".join(
                                x.strip() for x in (await detalle.locator("body").inner_text()).splitlines() if x.strip()
                            ))
                            coordenadas = await detalle.evaluate(r"""() => {
                                const html = document.documentElement.innerHTML;
                                const lat = html.match(/latitude\s*:\s*["'](-?\d+(?:\.\d+)?)["']/i);
                                const lon = html.match(/longitude\s*:\s*["'](-?\d+(?:\.\d+)?)["']/i);
                                return {latitud: lat?.[1] || null, longitud: lon?.[1] || null};
                            }""")
                            latitud = numero(coordenadas.get("latitud"))
                            longitud = numero(coordenadas.get("longitud"))
                        elif respuesta_detalle:
                            detalle_error = f"HTTP {respuesta_detalle.status}"
                    except VerificacionMitulaBloqueada:
                        raise
                    except Exception as error:
                        detalle_error = f"{type(error).__name__}: {str(error)[:300]}"
                        logging.warning("Detalle omitido %s: %s", tarjeta.get("id"), type(error).__name__)
                else:
                    detalle_error = "La tarjeta no publicó un destino de detalle"
                texto_completo = f"{texto}\n{texto_detalle}" if texto_detalle else texto
                superficie_total_detalle = extraer(texto_detalle, r"(?:sup(?:erficie)?\.?)[ ]*total\D{0,12}([\d.,]+)")
                superficie_cubierta = extraer(texto_detalle, r"(?:sup(?:erficie)?\.?)[ ]*cubierta\D{0,12}([\d.,]+)")
                superficie_descubierta = extraer(texto_detalle, r"(?:sup(?:erficie)?\.?)[ ]*descubierta\D{0,12}([\d.,]+)")
                superficie_semicubierta = extraer(texto_detalle, r"(?:sup(?:erficie)?\.?)[ ]*semicubierta\D{0,12}([\d.,]+)")
                superficie_terreno = extraer(texto_detalle, r"(?:sup(?:erficie)?\.?)[ ]*(?:de )?terreno\D{0,12}([\d.,]+)")
                anio_construccion = anio_construccion_publicado(texto_detalle)
                antiguedad = antiguedad_publicada(texto_detalle)
                superficie_total_final = superficie_total_detalle or superficie
                barrio_ficha = extraer_texto(texto_detalle, r"Barrio:\s*([^\n.]+)") or limpiar(tarjeta.get("barrio")) or None
                barrio, barrio_fuente = normalizar_barrio(barrio_ficha, f"{titulo or ''}\n{texto_detalle}")
                fila = {
                    "id_aviso": tarjeta.get("id"), "tipo_propiedad": tipo,
                    "operacion": "venta" if tarjeta.get("operacion") == "sell" else tarjeta.get("operacion"),
                    "precio": precio, "moneda": tarjeta.get("moneda"),
                    "precio_publicado": tarjeta.get("precio"), "superficie_publicada": tarjeta.get("superficie"),
                    "superficie_total_m2": superficie_total_final,
                    "superficie_cubierta_m2": superficie_cubierta,
                    "superficie_descubierta_m2": superficie_descubierta,
                    "superficie_semicubierta_m2": superficie_semicubierta,
                    "superficie_terreno_m2": superficie_terreno,
                    "ambientes": cantidad_desde_texto(
                        f"{titulo or ''}\n{texto_completo}", r"ambientes?"
                    ),
                    "dormitorios": numero(tarjeta.get("dormitorios"))
                    or cantidad_desde_texto(texto_completo, r"dormitorios?|habitaciones?"),
                    "banos": cantidad_desde_texto(
                        texto_detalle, r"ba[nñ]os?", singular_implicito=True
                    ) or banos,
                    "toilettes": cantidad_desde_texto(
                        texto_detalle, r"toilettes?|medio ba[nñ]o", singular_implicito=True
                    ),
                    "cocheras": cantidad_cocheras(texto_completo),
                    "antiguedad_anios": antiguedad,
                    "anio_construccion": anio_construccion,
                    "piso": piso_publicado(texto_detalle),
                    "unidad": unidad_publicada(texto_detalle),
                    "cantidad_plantas": plantas_publicadas(texto_detalle),
                    "frente_terreno_m": medida_terreno(texto_detalle, "frente"),
                    "fondo_terreno_m": medida_terreno(texto_detalle, "fondo"),
                    "expensas": extraer(texto_completo, r"expensas?\D{0,12}([\d.,]+)"),
                    "apto_credito": aptitud_credito(texto_detalle),
                    "apto_profesional": presencia_publicada(texto_detalle, "apto profesional"),
                    "escritura": extraer_texto(texto_detalle, r"Escritura:\s*([^\n]+)"),
                    "estado": estado_publicado(texto_detalle),
                    "disposicion": disposicion_publicada(texto_detalle),
                    "orientacion": orientacion_publicada(texto_detalle),
                    "ubicacion": limpiar(tarjeta.get("ubicacion")), "ciudad": limpiar(tarjeta.get("ciudad")),
                    "partido": limpiar(tarjeta.get("partido")), "provincia": limpiar(tarjeta.get("provincia")),
                    "latitud": latitud, "longitud": longitud,
                    "tipo_gas": tipo_gas_publicado(texto_detalle),
                    "portal_origen": campo_etiquetado(texto_detalle, "Portal de origen", "Sitio de origen"),
                    "id_aviso_origen": campo_etiquetado(texto_detalle, "ID del aviso original", "ID de origen"),
                    "barrio": barrio, "barrio_fuente": barrio_fuente,
                    "direccion": campo_etiquetado(texto_detalle, "Dirección", "Direccion"),
                    "circunscripcion": extraer_texto(texto_detalle, r"Circunscripción:\s*([^\s]+)"),
                    "seccion_catastral": extraer_texto(texto_detalle, r"Sección:\s*([^\s]+)"),
                    "manzana": extraer_texto(texto_detalle, r"Manzana:\s*([^\s]+)"),
                    "parcela": extraer_texto(texto_detalle, r"Parcela:\s*([^\s]+)"),
                    "chacra": extraer_texto(texto_detalle, r"Chacra:\s*([^\s]+)"),
                    "titulo": titulo,
                    "descripcion": extraer_descripcion(texto_detalle),
                    "fecha_publicacion": (re.search(r"\b\d{1,2}\s+[a-záéíóú]+\s+20\d{2}\b", texto, re.I).group(0)
                                          if re.search(r"\b\d{1,2}\s+[a-záéíóú]+\s+20\d{2}\b", texto, re.I) else None),
                    "actualizacion_relativa": (re.search(r"Hace [^\n]+", texto, re.I).group(0)
                                               if re.search(r"Hace [^\n]+", texto, re.I) else None),
                    "anunciante": texto.rsplit(" en - ", 1)[1].strip() if " en - " in texto else None,
                    "publicacion_destacada": "Sí" if "destacado" in texto.lower() else None,
                    "amoblamiento": estado_amoblamiento(texto_completo),
                    "url_aviso": url_detalle, "url_listado": url,
                    "pagina_origen": pagina, "posicion_listado": posicion,
                    "fecha_scraping_utc": fecha_scraping_utc,
                    "detalle_extraido": "Sí" if texto_detalle else "No",
                    "detalle_http_status": detalle_http_status, "detalle_error": detalle_error,
                    "texto_listado_crudo": tarjeta.get("texto"), "texto_detalle_crudo": texto_detalle or None,
                }
                reconocidas = []
                for etiqueta, columna in AMENITIES.items():
                    presencia = presencia_publicada(
                        texto_completo, etiqueta, SINONIMOS_AMENITIES.get(columna, ())
                    )
                    fila[columna] = presencia
                    if presencia is not None:
                        reconocidas.append(etiqueta)
                filas.append(fila)
                resumen["avisos_aceptados_nuevos"] += 1
                if texto_detalle:
                    resumen["detalles_exitosos"] += 1
                else:
                    resumen["detalles_fallidos"] += 1
                if tarjeta.get("id"):
                    vistos_id.add(str(tarjeta["id"]))
                if url_detalle:
                    vistos_url.add(str(url_detalle))
                await asyncio.sleep(args.espera_detalle_ms / 1000)
                if args.max_avisos and len(filas) >= args.max_avisos:
                    break
            resumen["registros_totales"] = len(filas)
            guardar(dataframe_unico(filas), args.parcial, resumen)
            logging.info("Parcial: %s (%s registros)", args.parcial, len(filas))
            if args.max_avisos and len(filas) >= args.max_avisos:
                resumen["estado"] = "completa"
                break
            await asyncio.sleep(args.espera_ms / 1000)
        await contexto.close()
    df = dataframe_unico(filas)
    resumen["registros_totales"] = len(df)
    resumen["fin_utc"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
    if resumen["detalles_fallidos"] or resumen["errores_listado"]:
        resumen["estado"] = "con_errores"
    elif resumen["estado"] == "en_progreso":
        resumen["estado"] = "completa"
    return df, resumen


def guardar(df: pd.DataFrame, salida: Path, resumen: dict | None = None) -> None:
    salida.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Propiedades", index=False)
        if resumen:
            pd.DataFrame(
                [{"metrica": clave, "valor": valor} for clave, valor in resumen.items()]
            ).to_excel(writer, sheet_name="Ejecucion", index=False)
        cobertura = pd.DataFrame({
            "variable": df.columns,
            "completos": [int(df[columna].notna().sum()) for columna in df.columns],
            "faltantes": [int(df[columna].isna().sum()) for columna in df.columns],
            "cobertura_pct": [round(float(df[columna].notna().mean() * 100), 1) if len(df) else 0.0 for columna in df.columns],
        })
        cobertura.to_excel(writer, sheet_name="Cobertura", index=False)
        hoja = writer.sheets["Propiedades"]
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions
        hoja.sheet_view.showGridLines = False
        for celda in hoja[1]:
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = PatternFill(fill_type="solid", fgColor="1F4E78")
        columnas_largas = {"titulo", "descripcion", "url_aviso", "url_listado", "detalle_error", "texto_listado_crudo", "texto_detalle_crudo"}
        for i, nombre in enumerate(COLUMNAS, 1):
            hoja.column_dimensions[get_column_letter(i)].width = 55 if nombre in columnas_largas else 22
        for fila in range(2, hoja.max_row + 1):
            hoja.row_dimensions[fila].height = 18
            for celda in hoja[fila]:
                celda.alignment = Alignment(vertical="center", wrap_text=False)
        for nombre_hoja in ("Ejecucion", "Cobertura"):
            if nombre_hoja not in writer.sheets:
                continue
            auxiliar = writer.sheets[nombre_hoja]
            auxiliar.freeze_panes = "A2"
            auxiliar.auto_filter.ref = auxiliar.dimensions
            auxiliar.sheet_view.showGridLines = False
            for celda in auxiliar[1]:
                celda.font = Font(bold=True, color="FFFFFF")
                celda.fill = PatternFill(fill_type="solid", fgColor="1F4E78")
            for columna in range(1, auxiliar.max_column + 1):
                auxiliar.column_dimensions[get_column_letter(columna)].width = 24


def argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scraper Mitula: venta en Mar del Plata")
    parser.add_argument(
        "--pagina-inicial", type=int, default=1,
        help="Primera página del listado que se recorrerá",
    )
    parser.add_argument("--max-paginas", type=int, default=1)
    parser.add_argument("--max-avisos", type=int, default=30)
    parser.add_argument("--espera-ms", type=int, default=2000)
    parser.add_argument("--espera-detalle-ms", type=int, default=700)
    parser.add_argument("--espera-reintento-ms", type=int, default=2500)
    parser.add_argument("--reintentos", type=int, default=3)
    parser.add_argument(
        "--esperar-captcha-minutos", type=int, default=15,
        help="Tiempo para completar manualmente un CAPTCHA en Edge visible",
    )
    parser.add_argument("--salida", type=Path, default=Path("data/mitula_venta_mar_del_plata.xlsx"))
    parser.add_argument("--parcial", type=Path)
    parser.add_argument(
        "--reanudar", action=argparse.BooleanOptionalAction, default=True,
        help="Reanuda desde el XLSX parcial si existe (predeterminado: sí)",
    )
    parser.add_argument("--perfil", type=Path, default=Path("data/.mitula_edge_profile"))
    parser.add_argument("--headless", action="store_true")
    args = parser.parse_args()
    if args.pagina_inicial < 1:
        parser.error("--pagina-inicial debe ser 1 o mayor")
    if args.max_paginas < args.pagina_inicial:
        parser.error("--max-paginas no puede ser menor que --pagina-inicial")
    if args.parcial is None:
        args.parcial = args.salida.with_name(f"{args.salida.stem}_parcial.xlsx")
    return args


def main() -> None:
    args = argumentos()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
    try:
        df, resumen = asyncio.run(ejecutar(args))
    except VerificacionMitulaBloqueada as error:
        logging.error("%s", error)
        logging.error(
            "Ejecución detenida antes del checkpoint de la página actual. Reanude más tarde con --pagina-inicial."
        )
        return
    except KeyboardInterrupt:
        logging.warning("Ejecución interrumpida; se conserva el último archivo parcial")
        if args.parcial.exists():
            df = pd.read_excel(args.parcial, sheet_name="Propiedades")
            resumen = {
                "estado": "interrumpida",
                "fin_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                "registros_totales": len(df),
            }
            guardar(df, args.parcial, resumen)
        return
    guardar(df, args.parcial, resumen)
    guardar(df, args.salida, resumen)
    logging.info(
        "Excel: %s (%s filas, %s columnas) | estado=%s | nuevos=%s | duplicados=%s | detalles_ok=%s | detalles_fallidos=%s",
        args.salida.resolve(), *df.shape, resumen["estado"],
        resumen["avisos_aceptados_nuevos"], resumen["duplicados_omitidos"],
        resumen["detalles_exitosos"], resumen["detalles_fallidos"],
    )


if __name__ == "__main__":
    main()
