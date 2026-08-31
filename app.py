import io
import re
from copy import copy
from pathlib import Path

import openpyxl
import pdfplumber
import streamlit as st

st.set_page_config(page_title="Orden de Producción → Excel", page_icon="📦", layout="centered")

PLANTILLA_PATH = Path(__file__).parent / "ETIQUETAC_APACKCH.xlsx"


# ---------------------------------------------------------------------------
# 1. EXTRACCION DE DATOS DEL PDF
# ---------------------------------------------------------------------------

def extraer_datos_pdf(file_bytes: bytes):
    texto_completo = []
    tabla_sku = None
    tabla_pack = None

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            texto_completo.append(page.extract_text() or "")
            for tabla in page.extract_tables():
                if not tabla or not tabla[0]:
                    continue
                encabezado = [c.strip() if c else "" for c in tabla[0]]
                if encabezado[:3] == ["SKU", "Talla", "Piezas"]:
                    tabla_sku = tabla
                elif encabezado[:2] == ["Pack", "Pack ID"]:
                    tabla_pack = tabla

    texto = "\n".join(texto_completo)

    def buscar(patron, default=None):
        m = re.search(patron, texto)
        return m.group(1).strip() if m else default

    orden = buscar(r"Numero de Orden:\s*(\d+)")
    modelo = buscar(r"Modelo ID:\s*(\d+)")
    descripcion = buscar(r"Descripcion del Articulo:\s*(.+)")
    color = buscar(r"Color Generico:\s*([^\n]+?)\s*PANTONE")
    piezas_totales = buscar(r"Piezas Totales:\s*(\d+)")

    if not all([orden, modelo, descripcion, color]):
        raise ValueError(
            "No se pudieron extraer todos los campos base del PDF "
            "(orden/modelo/descripcion/color). Revisa el formato del documento."
        )

    tallas = []
    piezas_por_talla = {}
    if tabla_sku:
        for fila in tabla_sku[1:]:
            sku, talla, piezas = (fila + [None, None, None])[:3]
            if not talla or talla.strip().lower() == "total":
                continue
            talla = talla.strip()
            tallas.append(talla)
            piezas_por_talla[talla] = int(str(piezas).strip())

    packs = []
    if tabla_pack:
        pack_actual = None
        for fila in tabla_pack[1:]:
            fila = (fila + [None] * 8)[:8]
            letra, pack_id, tipo, upp, total_packs, total_unid, talla, cant = fila

            if letra:
                pack_actual = {
                    "letra": letra.strip(),
                    "pza_pack": int(str(upp).strip()),
                    "total_packs": int(str(total_packs).strip()),
                    "total_unidades": int(str(total_unid).strip()),
                    "ratio": {},
                }
                packs.append(pack_actual)

            if pack_actual is not None and talla:
                pack_actual["ratio"][talla.strip()] = int(str(cant).strip())

    if not packs:
        raise ValueError("No se pudo extraer la tabla 'Detalles PACK / SKU' del PDF.")

    for p in packs:
        ratio_por_pack = {}
        for talla, total_talla in p["ratio"].items():
            if p["total_packs"] == 0:
                ratio_por_pack[talla] = 0
                continue
            cociente, resto = divmod(total_talla, p["total_packs"])
            if resto != 0:
                raise ValueError(
                    f"Pack {p['letra']}: la cantidad de talla '{talla}' "
                    f"({total_talla}) no es divisible entre el total de "
                    f"packs ({p['total_packs']}). Revisa el PDF."
                )
            ratio_por_pack[talla] = cociente
        p["ratio"] = ratio_por_pack

    return {
        "orden": int(orden),
        "modelo": int(modelo),
        "descripcion": descripcion,
        "color": color.strip(),
        "piezas_totales": int(piezas_totales) if piezas_totales else sum(piezas_por_talla.values()),
        "tallas": tallas if tallas else list(packs[0]["ratio"].keys()),
        "packs": packs,
    }


def validar_totales(datos):
    suma = 0
    for p in datos["packs"]:
        suma += sum(p["ratio"].values()) * p["total_packs"]
    if suma != datos["piezas_totales"]:
        raise ValueError(
            f"Los totales no cuadran: suma calculada de packs = {suma}, "
            f"'Piezas Totales' del PDF = {datos['piezas_totales']}."
        )
    return suma


# ---------------------------------------------------------------------------
# 2. LLENADO DE LA PLANTILLA
# ---------------------------------------------------------------------------

def _anchos_columnas(tallas, ratios_de_todos_los_packs):
    """Un ancho por talla: el maximo entre el largo del nombre de la talla
    y el largo de la cantidad mas grande que se va a imprimir en esa
    columna para toda la orden. Asi el encabezado y las cantidades quedan
    siempre alineados, sin importar si el numero es de 1 o 2 digitos."""
    anchos = {}
    for t in tallas:
        max_cant = max((len(str(r.get(t, 0))) for r in ratios_de_todos_los_packs), default=1)
        anchos[t] = max(len(t), max_cant)
    return anchos


def _formatear_encabezado(tallas, anchos):
    return "  ".join(f"{t:>{anchos[t]}}" for t in tallas)


def _formatear_cantidad(ratio, tallas, anchos):
    return "  ".join(f"{ratio.get(t, 0):>{anchos[t]}}" for t in tallas)


def llenar_plantilla(datos, plantilla_path):
    wb = openpyxl.load_workbook(plantilla_path)
    ws = wb["Hoja1"] if "Hoja1" in wb.sheetnames else wb.worksheets[0]

    tallas = datos["tallas"]
    anchos = _anchos_columnas(tallas, [p["ratio"] for p in datos["packs"]])
    encabezado_tallas = _formatear_encabezado(tallas, anchos)

    columnas = list("ABCDEFGHI")
    estilos_ref = {col: copy(ws.cell(row=2, column=i + 1)._style) for i, col in enumerate(columnas)}

    fila = 2
    for pack in datos["packs"]:
        cadena_cantidad = _formatear_cantidad(pack["ratio"], tallas, anchos)
        for n in range(1, pack["total_packs"] + 1):
            valores = [
                datos["orden"], datos["modelo"], pack["letra"], pack["pza_pack"],
                datos["descripcion"], n, datos["color"], encabezado_tallas, cadena_cantidad,
            ]
            for idx, val in enumerate(valores, start=1):
                celda = ws.cell(row=fila, column=idx, value=val)
                celda._style = copy(estilos_ref[columnas[idx - 1]])
            fila += 1

    ultima_fila_datos = fila - 1
    if ws.max_row > ultima_fila_datos:
        ws.delete_rows(ultima_fila_datos + 1, ws.max_row - ultima_fila_datos)

    out_buffer = io.BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer, ultima_fila_datos - 1


# ---------------------------------------------------------------------------
# Interfaz
# ---------------------------------------------------------------------------

st.title("📦 Orden de Producción → Excel de Packs")
st.caption("Sube el PDF de la orden de producción y descarga el Excel de packs/tallas ya llenado")

if not PLANTILLA_PATH.exists():
    st.error("Falta el archivo de plantilla ETIQUETAC_APACKCH.xlsx junto a esta app.")
    st.stop()

uploaded = st.file_uploader("PDF de la orden de producción", type=["pdf"])

if uploaded is not None:
    file_bytes = uploaded.read()

    if st.button("Generar Excel", type="primary"):
        with st.spinner("Leyendo PDF..."):
            try:
                datos = extraer_datos_pdf(file_bytes)
                total = validar_totales(datos)
                out_xlsx, num_renglones = llenar_plantilla(datos, PLANTILLA_PATH)
            except Exception as e:
                st.error(f"No se pudo procesar el archivo: {e}")
            else:
                st.success(f"Listo: {num_renglones} renglones generados, {total} piezas (cuadra con el PDF)")

                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Orden", datos["orden"])
                    st.metric("Modelo", datos["modelo"])
                with col2:
                    st.metric("Color", datos["color"])
                    st.metric("Piezas totales", datos["piezas_totales"])

                st.write(f"**Descripción:** {datos['descripcion']}")
                for p in datos["packs"]:
                    st.write(
                        f"- Pack **{p['letra']}**: {p['pza_pack']} pza/pack × "
                        f"{p['total_packs']} packs = {p['total_unidades']} unidades "
                        f"— ratio {p['ratio']}"
                    )

                out_name = f"ETIQUETAC_APACKCH_{datos['orden']}.xlsx"
                st.download_button(
                    label="⬇️ Descargar Excel",
                    data=out_xlsx,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

st.divider()
st.caption("Sin base de datos · los archivos se procesan en memoria y no se guardan.")
