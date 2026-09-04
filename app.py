import io
import re
from copy import copy
from pathlib import Path

import openpyxl
import pdfplumber
import streamlit as st
from openpyxl.styles import Font

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
    return " ".join(f"{t:>{anchos[t]}}" for t in tallas)


def _formatear_cantidad(ratio, tallas, anchos):
    return " ".join(f"{ratio.get(t, 0):>{anchos[t]}}" for t in tallas)


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
                if idx in (8, 9):
                    # TALLAS y CANTIDAD POR TALLA: fuente monoespaciada para
                    # que los espacios de relleno alineen de verdad en pantalla
                    # (Calibri es de ancho variable, no alinea aunque el
                    # numero de caracteres coincida).
                    f = celda.font
                    celda.font = Font(
                        name="Consolas", size=f.size, bold=f.bold,
                        italic=f.italic, color=f.color,
                    )
            fila += 1

    ultima_fila_datos = fila - 1
    if ws.max_row > ultima_fila_datos:
        ws.delete_rows(ultima_fila_datos + 1, ws.max_row - ultima_fila_datos)

    out_buffer = io.BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer, ultima_fila_datos - 1


# ---------------------------------------------------------------------------
# 3. ETIQUETAS DE PACK (PDF Y ZPL) -- 10 cm x 3.5 cm, diseño Label Matrix
# ---------------------------------------------------------------------------
LABEL_W_CM = 10.0
LABEL_H_CM = 3.5

ROW1_H = 0.55   # Orden de Compra | Modelo
ROW2_H = 0.68   # Destinatario | "Descripcion:"
ROW3_H = 0.65   # Tipo de Pack + Pza. por Pack | (texto de descripción)
ROW4_H = LABEL_H_CM - ROW1_H - ROW2_H - ROW3_H  # fila grande: # Pack | Curva | Color

COL_SPLIT = 6.0   # límite entre columna izq/der en filas 1-3
COL_A_W = 2.0     # fila 4: ancho "# de Pack"
COL_B_W = 5.5     # fila 4: ancho "Cantidad y curva de tallas"
COL_C_W = LABEL_W_CM - COL_A_W - COL_B_W  # fila 4: ancho "Color"

TALLAS_ORDEN_ETIQUETA = ["ECH", "CH", "M", "G", "EG", "XG", "XXG"]

DESTINATARIO_DEFAULT = "C&A México S. de R.L."


def _tallas_presentes_etiqueta(ratio: dict) -> list:
    presentes = [t for t in TALLAS_ORDEN_ETIQUETA if t in ratio]
    extra = [t for t in ratio if t not in TALLAS_ORDEN_ETIQUETA]
    return presentes + extra


def expand_packs_to_labels(datos, destinatario: str = DESTINATARIO_DEFAULT) -> list:
    """Convierte datos['packs'] (tipos A, B, C...) en una lista *plana* de
    etiquetas individuales, una por cada pack físico, numeradas 1..N dentro
    de su tipo -- el campo '# de Pack' del diseño original.

    La curva de tallas de cada etiqueta siempre incluye TODAS las tallas
    de la orden (datos['tallas']), con 0 en las que ese pack no lleva --
    así la etiqueta muestra la curva completa aunque el pack en particular
    no traiga esa talla."""
    tallas_orden = datos.get("tallas") or []
    etiquetas = []
    for p in datos["packs"]:
        tallas_pack = tallas_orden or list(p["ratio"].keys())
        ratio_completo = {t: p["ratio"].get(t, 0) for t in tallas_pack}
        for n in range(1, p["total_packs"] + 1):
            etiquetas.append({
                "orden": datos["orden"], "modelo": datos["modelo"], "color": datos["color"],
                "descripcion": datos["descripcion"], "destinatario": destinatario,
                "letra": p["letra"], "pza_pack": p["pza_pack"],
                "num_pack": n, "total_packs": p["total_packs"],
                "ratio": ratio_completo,
            })
    return etiquetas


def _wrap_two_lines(text, font, size, max_width, string_width_fn):
    words = text.split()
    linea1, i = "", 0
    while i < len(words):
        candidate = (linea1 + " " + words[i]).strip()
        if string_width_fn(candidate, font, size) <= max_width:
            linea1 = candidate
            i += 1
        else:
            break
    resto = words[i:]
    linea2 = ""
    for w in resto:
        candidate = (linea2 + " " + w).strip()
        if string_width_fn(candidate, font, size) <= max_width:
            linea2 = candidate
        else:
            while string_width_fn(linea2 + "...", font, size) > max_width and linea2:
                linea2 = linea2[:-1]
            linea2 += "..."
            break
    return linea1, linea2


def build_labels_pdf(etiquetas: list) -> bytes:
    from reportlab.lib.units import cm
    from reportlab.pdfgen import canvas

    buf = io.BytesIO()
    page_w, page_h = LABEL_W_CM * cm, LABEL_H_CM * cm
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))
    for et in etiquetas:
        _draw_label_pdf(c, et, page_w, page_h)
        c.showPage()
    c.save()
    return buf.getvalue()


def _draw_label_pdf(c, et, page_w, page_h):
    from reportlab.lib.units import cm
    from reportlab.pdfbase.pdfmetrics import stringWidth

    def y_top(cm_from_top):
        return page_h - cm_from_top * cm

    c.setLineWidth(0.6)
    y1 = y_top(ROW1_H)
    y2 = y_top(ROW1_H + ROW2_H)
    y3 = y_top(ROW1_H + ROW2_H + ROW3_H)
    c.line(0, page_h, page_w, page_h)
    c.line(0, 0, page_w, 0)
    c.line(0, 0, 0, page_h)
    c.line(page_w, 0, page_w, page_h)
    c.line(0, y1, page_w, y1)
    c.line(0, y2, page_w, y2)
    c.line(0, y3, page_w, y3)
    c.line(COL_SPLIT * cm, y3, COL_SPLIT * cm, page_h)
    x_b1 = COL_A_W * cm
    x_b2 = (COL_A_W + COL_B_W) * cm
    c.line(x_b1, 0, x_b1, y3)
    c.line(x_b2, 0, x_b2, y3)

    pad = 0.1 * cm

    # Fila 1: Orden de Compra | Modelo
    c.setFont("Helvetica", 6)
    c.drawString(pad, page_h - 0.22 * cm, "Orden de Compra:")
    c.setFont("Helvetica-Bold", 8)
    c.drawString(pad + 1.9 * cm, page_h - 0.23 * cm, str(et["orden"]))
    c.setFont("Helvetica", 6)
    c.drawString(COL_SPLIT * cm + pad, page_h - 0.22 * cm, "Modelo:")
    c.setFont("Helvetica-Bold", 8)
    c.drawString(COL_SPLIT * cm + pad + 1.0 * cm, page_h - 0.23 * cm, str(et["modelo"]))

    # Fila 2: Destinatario | "Descripcion:"
    c.setFont("Helvetica", 5.5)
    c.drawString(pad, y1 - 0.20 * cm, "Destinatario:")
    c.setFont("Helvetica", 6)
    c.drawString(pad, y1 - 0.42 * cm, et["destinatario"][:42])
    c.setFont("Helvetica", 5.5)
    c.drawString(COL_SPLIT * cm + pad, y1 - 0.16 * cm, "Descripcion:")

    # Fila 3: Tipo de Pack + Pza. por Pack | texto de descripción
    c.setFont("Helvetica", 5.5)
    c.drawString(pad, y2 - 0.20 * cm, "Tipo de Pack")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(pad + 1.7 * cm, y2 - 0.30 * cm, str(et["letra"]))
    c.setFont("Helvetica", 5.5)
    c.drawString(pad + 2.6 * cm, y2 - 0.20 * cm, "Pza. por Pack:")
    c.setFont("Helvetica-Bold", 8)
    c.drawString(pad + 4.6 * cm, y2 - 0.22 * cm, str(et["pza_pack"]))

    desc_font, desc_size = "Helvetica", 6.5
    max_w = page_w - COL_SPLIT * cm - 2 * pad
    linea1, linea2 = _wrap_two_lines(et["descripcion"], desc_font, desc_size, max_w, stringWidth)
    c.setFont(desc_font, desc_size)
    c.drawString(COL_SPLIT * cm + pad, y1 - 0.40 * cm, linea1)
    if linea2:
        c.drawString(COL_SPLIT * cm + pad, y1 - 0.60 * cm, linea2)

    # Fila 4: # de Pack | Cantidad y curva de tallas | Color
    c.setFont("Helvetica", 6)
    c.drawCentredString(COL_A_W / 2 * cm, y3 - 0.30 * cm, "# de Pack")
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(COL_A_W / 2 * cm, y3 - 0.75 * cm, f"{et['num_pack']}/{et['total_packs']}")

    c.setFont("Helvetica", 6)
    c.drawCentredString((COL_A_W + COL_B_W / 2) * cm, y3 - 0.28 * cm, "Cantidad y curva de tallas")
    tallas = _tallas_presentes_etiqueta(et["ratio"])
    n = len(tallas) or 1
    col_w = COL_B_W / n
    for i, t in enumerate(tallas):
        cx = (COL_A_W + col_w * (i + 0.5)) * cm
        c.setFont("Helvetica", 6.5)
        c.drawCentredString(cx, y3 - 0.55 * cm, t)
        c.setFont("Helvetica-Bold", 11)
        c.drawCentredString(cx, y3 - 0.95 * cm, str(et["ratio"][t]))

    cx_color = (COL_A_W + COL_B_W + COL_C_W / 2) * cm
    c.setFont("Helvetica", 6)
    c.drawCentredString(cx_color, y3 - 0.30 * cm, "Color:")
    c.setFont("Helvetica-Bold", 9)
    c.drawCentredString(cx_color, y3 - 0.75 * cm, str(et["color"]))


DPI = 203
_DOTS_PER_CM = DPI / 2.54


def _d(cm_val: float) -> int:
    return round(cm_val * _DOTS_PER_CM)


def build_labels_zpl(etiquetas: list) -> str:
    return "\n".join(_zpl_one_label(et) for et in etiquetas)


def _zpl_one_label(et: dict) -> str:
    W, H = _d(LABEL_W_CM), _d(LABEL_H_CM)
    row1, row2, row3 = _d(ROW1_H), _d(ROW1_H + ROW2_H), _d(ROW1_H + ROW2_H + ROW3_H)
    col_split = _d(COL_SPLIT)
    x_b1 = _d(COL_A_W)
    x_b2 = _d(COL_A_W + COL_B_W)

    lines = [
        "^XA",
        f"^PW{W}",
        f"^LL{H}",
        "^CI28",
        "^LH0,0",
        f"^FO0,{row1}^GB{W},1,2^FS",
        f"^FO0,{row2}^GB{W},1,2^FS",
        f"^FO0,{row3}^GB{W},1,2^FS",
        f"^FO{col_split},0^GB1,{row3},2^FS",
        f"^FO{x_b1},{row3}^GB1,{H - row3},2^FS",
        f"^FO{x_b2},{row3}^GB1,{H - row3},2^FS",
        f"^FO0,0^GB{W},{H},2^FS",
        f"^FO10,10^A0N,16,16^FDOrden de Compra:^FS",
        f"^FO190,6^A0N,22,20^FB{col_split - 200},1,0,L,0^FD{et['orden']}^FS",
        f"^FO{col_split + 10},10^A0N,16,16^FDModelo:^FS",
        f"^FO{col_split + 100},6^A0N,22,20^FD{et['modelo']}^FS",
        f"^FO10,{row1 + 5}^A0N,14,14^FDDestinatario:^FS",
        f"^FO10,{row1 + 25}^A0N,16,16^FB{col_split - 20},1,0,L,0^FD{et['destinatario']}^FS",
        f"^FO{col_split + 10},{row1 + 5}^A0N,14,14^FDDescripcion:^FS",
        f"^FO{col_split + 10},{row1 + 25}^A0N,15,15^FB{W - col_split - 20},2,2,L,0^FD{et['descripcion']}^FS",
        f"^FO10,{row2 + 5}^A0N,14,14^FDTipo de Pack^FS",
        f"^FO135,{row2 + 2}^A0N,26,24^FD{et['letra']}^FS",
        f"^FO195,{row2 + 5}^A0N,14,14^FDPza. por Pack:^FS",
        f"^FO370,{row2 + 2}^A0N,22,20^FD{et['pza_pack']}^FS",
        f"^FO0,{row3 + 8}^A0N,16,16^FB{x_b1},1,0,C,0^FD# de Pack^FS",
        f"^FO0,{row3 + 30}^A0N,34,32^FB{x_b1},1,0,C,0^FD{et['num_pack']}/{et['total_packs']}^FS",
        f"^FO{x_b2},{row3 + 8}^A0N,16,16^FB{W - x_b2},1,0,C,0^FDColor:^FS",
        f"^FO{x_b2},{row3 + 40}^A0N,24,22^FB{W - x_b2},1,0,C,0^FD{et['color']}^FS",
    ]

    tallas = _tallas_presentes_etiqueta(et["ratio"])
    n = max(len(tallas), 1)
    col_w = (x_b2 - x_b1) // n
    lines.append(f"^FO{x_b1},{row3 + 8}^A0N,16,16^FB{x_b2 - x_b1},1,0,C,0^FDCantidad y curva de tallas^FS")
    for i, t in enumerate(tallas):
        cx = x_b1 + col_w * i
        lines.append(f"^FO{cx},{row3 + 32}^A0N,15,15^FB{col_w},1,0,C,0^FD{t}^FS")
        lines.append(f"^FO{cx},{row3 + 55}^A0N,30,28^FB{col_w},1,0,C,0^FD{et['ratio'][t]}^FS")

    lines.append("^XZ")
    return "\n".join(lines)


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

                etiquetas = expand_packs_to_labels(datos)
                col_pdf, col_zpl = st.columns(2)
                with col_pdf:
                    st.download_button(
                        label=f"🏷️ Etiquetas PDF ({len(etiquetas)})",
                        data=build_labels_pdf(etiquetas),
                        file_name=f"Etiquetas_{datos['orden']}.pdf",
                        mime="application/pdf",
                    )
                with col_zpl:
                    st.download_button(
                        label=f"🖨️ Etiquetas ZPL ({len(etiquetas)})",
                        data=build_labels_zpl(etiquetas),
                        file_name=f"Etiquetas_{datos['orden']}.zpl",
                        mime="text/plain",
                    )

st.divider()
st.caption("Sin base de datos · los archivos se procesan en memoria y no se guardan.")
