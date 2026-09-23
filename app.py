import io
import os
import textwrap
from collections import Counter, defaultdict
from datetime import datetime

import pandas as pd
import openpyxl
import barcode as barcode_lib
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "equivalencia-app-secret")

DATA_DIR = os.environ.get("DATA_DIR", "/app/data")
os.makedirs(DATA_DIR, exist_ok=True)

GRANDE_CACHE_PATH_XLS = os.path.join(DATA_DIR, "planilla_grande.xls")
GRANDE_CACHE_PATH_XLSX = os.path.join(DATA_DIR, "planilla_grande.xlsx")
GRANDE_CACHE_META = os.path.join(DATA_DIR, "planilla_grande.meta")
STOCK_CACHE_PATH = os.path.join(DATA_DIR, "stock_precios.xlsx")
STOCK_CACHE_META = os.path.join(DATA_DIR, "stock_precios.meta")

ALLOWED_EXT = {".xls", ".xlsx"}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_BOLD = os.path.join(BASE_DIR, "fonts", "DejaVuSans-Bold.ttf")
FONT_REGULAR = os.path.join(BASE_DIR, "fonts", "DejaVuSans.ttf")


def _ext(filename):
    return os.path.splitext(filename)[1].lower()


def _texto(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    texto = str(valor).strip()
    return texto[:-2] if texto.endswith(".0") else texto


def _clave_producto(sku, color):
    return f"{_texto(sku)}{_texto(color)}".upper().replace(" ", "")


def _clave_barra(codigo):
    texto = _texto(codigo)
    return str(int(texto)) if texto.isdigit() else texto.upper()


def _precio(valor):
    texto = _texto(valor)
    if not texto or texto == "-":
        return None
    if isinstance(valor, (int, float)):
        return round(valor)
    return round(float(texto.replace("$", "").replace(".", "").replace(",", ".")))


def _read_grande(path):
    """Lee la planilla de equivalencias y arma un índice SKU + COLOR + TALLE.

    Acepta tanto el export anterior (Código, Artículo, Descripción, ..., Talle)
    como el nuevo, cuya columna de barras se llama CODIGOS.  El dato ``Artículo``
    del export anterior contiene ``SKU COLOR`` en una sola celda.
    """
    ext = _ext(path)
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    xl = pd.ExcelFile(path, engine=engine)

    lookup = defaultdict(list)
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, header=None, dtype=str)
        if df.empty:
            continue
        encabezados = {
            str(valor).strip().upper().replace("Ó", "O"): columna
            for columna, valor in enumerate(df.iloc[0])
            if not pd.isna(valor) and str(valor).strip()
        }
        col_codigo = _buscar_columna(encabezados, "CODIGO", "CODIGOS", "CODIGO DE BARRA")
        col_articulo = _buscar_columna(encabezados, "ARTICULO", "SKU COLOR", "CONCAT")
        col_sku = _buscar_columna(encabezados, "SKU")
        col_color = _buscar_columna(encabezados, "COLOR")
        col_talle = _buscar_columna(encabezados, "TALLE")
        col_descripcion = _buscar_columna(encabezados, "DESCRIPCION", "NOMBRE", "PRODUCTO")

        # Compatibilidad con el export histórico: no siempre trae los mismos títulos.
        tiene_encabezados = col_codigo is not None and (col_articulo is not None or (col_sku is not None and col_color is not None))
        if tiene_encabezados:
            df = df.iloc[1:]
        else:
            col_codigo, col_articulo, col_sku, col_color = 0, 1, None, None
            col_descripcion, col_talle = 2, 5
        for _, row in df.iterrows():
            codigo = row.get(col_codigo)
            descripcion = row.get(col_descripcion) if col_descripcion is not None else None
            talle = row.get(col_talle) if col_talle is not None else None
            if col_articulo is not None:
                articulo = row.get(col_articulo)
            else:
                sku = row.get(col_sku)
                color = row.get(col_color)
                articulo = None if pd.isna(sku) or pd.isna(color) else f"{sku} {color}"
            if pd.isna(codigo) or pd.isna(articulo):
                continue
            codigo = str(codigo).strip()
            articulo = str(articulo).strip()
            descripcion = "" if pd.isna(descripcion) else str(descripcion).strip()
            talle = "" if pd.isna(talle) else str(talle).strip()
            if not codigo or not articulo:
                continue
            lookup[articulo].append((codigo, talle, descripcion))
    return lookup


_LOOKUP_CACHE = {"path": None, "mtime": None, "lookup": None}


def _get_lookup(path):
    """Cachea en memoria el resultado de _read_grande mientras no cambie el archivo,
    para no re-parsear ~300k filas en cada pedido o cada etiqueta."""
    global _LOOKUP_CACHE
    mtime = os.path.getmtime(path)
    if _LOOKUP_CACHE["path"] == path and _LOOKUP_CACHE["mtime"] == mtime:
        return _LOOKUP_CACHE["lookup"]
    lookup = _read_grande(path)
    _LOOKUP_CACHE = {"path": path, "mtime": mtime, "lookup": lookup}
    return lookup


_BARCODE_CACHE = {"path": None, "mtime": None, "index": None}


def _get_barcode_index(path):
    """Crea un índice código de barras -> SKU, color, talle y descripción."""
    global _BARCODE_CACHE
    mtime = os.path.getmtime(path)
    if _BARCODE_CACHE["path"] == path and _BARCODE_CACHE["mtime"] == mtime:
        return _BARCODE_CACHE["index"]

    index = {}
    for articulo, candidatos in _get_lookup(path).items():
        partes = articulo.split(maxsplit=1)
        if len(partes) != 2:
            continue
        sku, color = partes
        for codigo, talle, descripcion in candidatos:
            index.setdefault(
                _clave_barra(codigo),
                {"codigo": codigo, "sku": sku, "color": color, "talle": talle, "descripcion": descripcion},
            )
    _BARCODE_CACHE = {"path": path, "mtime": mtime, "index": index}
    return index


_PRICE_CACHE = {"path": None, "mtime": None, "prices": None}


def _get_price_lookup(path):
    """Lee RETAIL (columna K) del STOCK y lo agrupa por SKU + color."""
    global _PRICE_CACHE
    mtime = os.path.getmtime(path)
    if _PRICE_CACHE["path"] == path and _PRICE_CACHE["mtime"] == mtime:
        return _PRICE_CACHE["prices"]

    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if "STOCK" not in workbook.sheetnames:
        raise ValueError("La planilla de precios debe tener una solapa llamada STOCK.")
    values_by_product = defaultdict(list)
    for row in workbook["STOCK"].iter_rows(min_row=2, max_col=11, values_only=True):
        key = _clave_producto(row[1], row[2])
        amount = _precio(row[10])
        if key and amount is not None:
            values_by_product[key].append(amount)

    prices = {key: Counter(values).most_common(1)[0][0] for key, values in values_by_product.items()}
    _PRICE_CACHE = {"path": path, "mtime": mtime, "prices": prices}
    return prices


def _resolver_grande_cache():
    """Devuelve el path de la planilla grande cacheada en disco, o None si no hay ninguna."""
    if os.path.exists(GRANDE_CACHE_PATH_XLS):
        return GRANDE_CACHE_PATH_XLS
    if os.path.exists(GRANDE_CACHE_PATH_XLSX):
        return GRANDE_CACHE_PATH_XLSX
    return None


def _resolver_stock_cache():
    return STOCK_CACHE_PATH if os.path.exists(STOCK_CACHE_PATH) else None


def _dedupe_codigos(codigos):
    """A veces la planilla grande tiene el mismo código de barra duplicado con
    distinto padding de ceros a la izquierda (ej. '00191743848643' y '191743848643').
    Los trata como el mismo código y se queda con la representación más corta."""
    vistos = {}
    for c in codigos:
        try:
            clave = int(c)
        except (TypeError, ValueError):
            clave = c
        if clave not in vistos or len(c) < len(vistos[clave]):
            vistos[clave] = c
    return list(vistos.values())


def _buscar_codigo(lookup, sku, color, talle):
    """Busca el código de barra exacto para SKU + Color + Talle."""
    if sku is None or color is None or str(sku).strip() == "" or str(color).strip() == "":
        return ""
    sku_str = str(sku).strip()
    if sku_str.endswith(".0"):
        sku_str = sku_str[:-2]
    color_str = str(color).strip()
    talle_str = "" if talle is None else str(talle).strip()

    key = f"{sku_str} {color_str}"
    candidatos = lookup.get(key, [])
    coincidencias = [
        codigo for codigo, t, _desc in candidatos if t.strip().upper() == talle_str.upper()
    ]
    coincidencias = _dedupe_codigos(coincidencias)

    if not coincidencias:
        return "SIN COINCIDENCIA"
    if len(coincidencias) == 1:
        return coincidencias[0]
    # Conflicto real de datos en la planilla grande: mismo SKU+Color+Talle con más
    # de un código de barra distinto. Se muestran todos para que se revise a mano.
    return " / ".join(coincidencias)


def _buscar_producto(lookup, sku, color, talle):
    """Busca código de barra + descripción exactos para SKU + Color + Talle.
    Devuelve (codigo, descripcion) o (None, None) si no hay coincidencia."""
    if not sku or not color:
        return None, None
    sku_str = str(sku).strip()
    if sku_str.endswith(".0"):
        sku_str = sku_str[:-2]
    color_str = str(color).strip()
    talle_str = "" if talle is None else str(talle).strip()

    key = f"{sku_str} {color_str}"
    candidatos = lookup.get(key, [])
    coincidencias = [
        (codigo, desc) for codigo, t, desc in candidatos if t.strip().upper() == talle_str.upper()
    ]
    if not coincidencias:
        return None, None

    codigos_dedup = _dedupe_codigos([c for c, _d in coincidencias])
    descripcion = coincidencias[0][1]
    # Para la etiqueta necesitamos un único código escaneable: si hay un conflicto real
    # de datos (mismo SKU+Color+Talle con más de un código), usamos el primero.
    codigo = codigos_dedup[0]
    return codigo, descripcion


def _buscar_columna(headers, *nombres_posibles):
    for nombre in nombres_posibles:
        if nombre in headers:
            return headers[nombre]
    return None


def _procesar(grande_path, pedido_path):
    """Completa EQUIVALENCIA usando SKU, COLOR y TALLE de la planilla de stock.

    Se preservan formato, fórmulas y todas las demás columnas. Si el archivo tiene
    varias hojas con esos encabezados, se completan todas.
    """
    lookup = _get_lookup(grande_path)

    wb = openpyxl.load_workbook(pedido_path)
    hojas_procesadas = 0
    for ws in wb.worksheets:
        headers = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=1, column=c).value
            if v is not None:
                headers[str(v).strip().upper()] = c

        col_equiv = _buscar_columna(headers, "EQUIVALENCIA")
        col_sku = _buscar_columna(headers, "SKU")
        col_color = _buscar_columna(headers, "COLOR")
        col_talle = _buscar_columna(headers, "TALLE")
        if not all((col_equiv, col_sku, col_color, col_talle)):
            continue

        hojas_procesadas += 1
        for row in range(2, ws.max_row + 1):
            sku = ws.cell(row=row, column=col_sku).value
            color = ws.cell(row=row, column=col_color).value
            talle = ws.cell(row=row, column=col_talle).value
            if sku is None and color is None:
                continue
            resultado = _buscar_codigo(lookup, sku, color, talle)
            cell = ws.cell(row=row, column=col_equiv, value=resultado)
            cell.number_format = "@"  # texto plano, para no perder ceros a la izquierda

    if not hojas_procesadas:
        raise ValueError(
            "No encontré una hoja con las columnas EQUIVALENCIA, SKU, COLOR y TALLE en la fila 1."
        )

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def _generar_imagen_codigo_barra(codigo, module_height_mm=9.0, module_width_mm=0.28):
    """Genera la imagen del código de barra (EAN13 / UPC-A / Code128 según el largo)
    sin el texto legible debajo (eso lo dibujamos aparte con nuestra propia fuente)."""
    codigo = str(codigo).strip()
    solo_digitos = codigo.isdigit()
    if solo_digitos and len(codigo) == 13:
        clase = "ean13"
    elif solo_digitos and len(codigo) == 12:
        clase = "upca"
    elif solo_digitos and len(codigo) == 8:
        clase = "ean8"
    else:
        clase = "code128"

    BarcodeClass = barcode_lib.get_barcode_class(clase)
    bc = BarcodeClass(codigo, writer=ImageWriter())
    buf = io.BytesIO()
    bc.write(
        buf,
        options={
            "write_text": False,
            "quiet_zone": 1.0,
            "module_height": module_height_mm,
            "module_width": module_width_mm,
        },
    )
    buf.seek(0)
    return Image.open(buf).convert("L")


def _texto_ajustado(draw, texto, font, max_width_px, max_lineas=2):
    """Envuelve el texto para que entre en max_width_px, cortando con '…' si no entra."""
    palabras = texto.split()
    lineas = []
    actual = ""
    for palabra in palabras:
        prueba = (actual + " " + palabra).strip()
        ancho = draw.textbbox((0, 0), prueba, font=font)[2]
        if ancho <= max_width_px or not actual:
            actual = prueba
        else:
            lineas.append(actual)
            actual = palabra
            if len(lineas) == max_lineas - 1:
                break
    if actual:
        lineas.append(actual)
    lineas = lineas[:max_lineas]

    # Si sobró texto sin usar, agregamos "…" a la última línea que entre
    usado = " ".join(lineas)
    if len(usado) < len(texto):
        ultima = lineas[-1]
        while draw.textbbox((0, 0), ultima + "…", font=font)[2] > max_width_px and len(ultima) > 1:
            ultima = ultima[:-1]
        lineas[-1] = ultima.rstrip() + "…"
    return lineas


def _money(amount):
    return "$ " + f"{round(amount):,}".replace(",", ".")


def generar_etiqueta(descripcion, color, talle, codigo, precio_sin_iva, sku="", dpi=300):
    """Genera una etiqueta de 30 x 15 mm, PNG a 300 DPI."""
    px_mm = dpi / 25.4
    ancho_mm, alto_mm = 30.0, 15.0
    W, H = round(ancho_mm * px_mm), round(alto_mm * px_mm)
    img = Image.new("L", (W, H), color=255)
    draw = ImageDraw.Draw(img)

    margen = round(0.8 * px_mm)
    col_izq_ancho = round(15.0 * px_mm)
    font_desc = ImageFont.truetype(FONT_REGULAR, size=round(1.55 * px_mm))
    font_meta = ImageFont.truetype(FONT_REGULAR, size=round(1.55 * px_mm))
    font_price_label = ImageFont.truetype(FONT_REGULAR, size=round(1.35 * px_mm))
    font_price = ImageFont.truetype(FONT_BOLD, size=round(2.05 * px_mm))
    font_digitos = ImageFont.truetype(FONT_REGULAR, size=round(1.45 * px_mm))

    y = margen
    max_w = col_izq_ancho - margen
    for linea in _texto_ajustado(draw, str(descripcion).upper(), font_desc, max_w, max_lineas=2):
        draw.text((margen, y), linea, font=font_desc, fill=0)
        y += draw.textbbox((0, 0), linea, font=font_desc)[3] + round(0.25 * px_mm)

    draw.text((margen, y + round(0.25 * px_mm)), f"{str(color).upper()}  {str(talle).upper()}", font=font_meta, fill=0)
    y += round(2.35 * px_mm)
    if sku:
        draw.text((margen, y), str(sku).strip(), font=font_meta, fill=0)
        y += round(2.0 * px_mm)

    draw.text((margen, y), "SIN IVA", font=font_price_label, fill=0)
    y += round(1.55 * px_mm)
    draw.text((margen, y), _money(precio_sin_iva), font=font_price, fill=0)
    y += round(2.75 * px_mm)
    draw.text((margen, y), "CON IVA", font=font_price_label, fill=0)
    y += round(1.55 * px_mm)
    draw.text((margen, y), _money(round(precio_sin_iva * 1.21)), font=font_price, fill=0)

    codigo_str = str(codigo).strip()
    col_der_x = margen + col_izq_ancho + round(0.5 * px_mm)
    col_der_ancho = W - col_der_x - margen
    bc_img = _generar_imagen_codigo_barra(codigo_str, module_height_mm=7.0, module_width_mm=0.22)
    escala = col_der_ancho / bc_img.width
    bc_alto = min(round(bc_img.height * escala), round(9.3 * px_mm))
    bc_img = bc_img.resize((col_der_ancho, bc_alto))
    bc_y = margen
    img.paste(bc_img, (col_der_x, bc_y))
    bbox = draw.textbbox((0, 0), codigo_str, font=font_digitos)
    digitos_x = col_der_x + max(0, (col_der_ancho - (bbox[2] - bbox[0])) // 2)
    draw.text((digitos_x, bc_y + bc_alto + round(0.25 * px_mm)), codigo_str, font=font_digitos, fill=0)

    salida = io.BytesIO()
    img.save(salida, format="PNG", dpi=(dpi, dpi))
    salida.seek(0)
    return salida


@app.route("/etiqueta", methods=["GET"])
def etiqueta():
    grande_cached = _resolver_grande_cache() is not None
    stock_cached = _resolver_stock_cache() is not None
    return render_template("etiqueta.html", grande_cached=grande_cached, stock_cached=stock_cached)


@app.route("/etiqueta/generar", methods=["GET", "POST"])
def etiqueta_generar():
    datos = request.values
    barcode_scanned = (datos.get("barcode") or "").strip()
    sku = (datos.get("sku") or "").strip()
    color = (datos.get("color") or "").strip()
    talle = (datos.get("talle") or "").strip()

    if not barcode_scanned:
        flash("Completá SKU, Color y Talle.")
        return redirect(url_for("etiqueta"))

    grande_path = _resolver_grande_cache()
    if not grande_path:
        flash("No hay planilla grande cargada todavía. Subila primero desde la página principal.")
        return redirect(url_for("etiqueta"))

    try:
        producto = _get_barcode_index(grande_path).get(_clave_barra(barcode_scanned))
    except Exception as e:
        flash(f"Error buscando el producto: {e}")
        return redirect(url_for("etiqueta"))

    if not producto:
        flash(f"No encontré ningún producto con SKU {sku}, Color {color} y Talle {talle} en la planilla grande.")
        return redirect(url_for("etiqueta"))

    stock_path = _resolver_stock_cache()
    if not stock_path:
        flash("No hay una planilla de STOCK con precios cargada. Subila desde la pantalla principal.")
        return redirect(url_for("etiqueta"))

    try:
        precio_sin_iva = _get_price_lookup(stock_path).get(_clave_producto(producto["sku"], producto["color"]))
    except Exception as e:
        flash(f"Error leyendo los precios: {e}")
        return redirect(url_for("etiqueta"))
    if precio_sin_iva is None:
        flash(f"No encontré precio para {producto['sku']} {producto['color']} en la planilla STOCK.")
        return redirect(url_for("etiqueta"))

    try:
        imagen = generar_etiqueta(
            producto["descripcion"], producto["color"], producto["talle"], producto["codigo"], precio_sin_iva, sku=producto["sku"]
        )
    except Exception as e:
        flash(f"No pude generar la etiqueta: {e}")
        return redirect(url_for("etiqueta"))

    descargar = datos.get("descargar") == "1"
    nombre = f"etiqueta_{secure_filename(producto['codigo'])}.png"
    return send_file(
        imagen,
        mimetype="image/png",
        as_attachment=descargar,
        download_name=nombre,
    )


@app.route("/", methods=["GET"])
def index():
    grande_cached = os.path.exists(GRANDE_CACHE_META)
    grande_info = None
    if grande_cached:
        with open(GRANDE_CACHE_META) as f:
            grande_info = f.read().strip()
    stock_cached = os.path.exists(STOCK_CACHE_META)
    stock_info = None
    if stock_cached:
        with open(STOCK_CACHE_META) as f:
            stock_info = f.read().strip()
    return render_template(
        "index.html", grande_cached=grande_cached, grande_info=grande_info,
        stock_cached=stock_cached, stock_info=stock_info,
    )


@app.route("/subir-stock", methods=["POST"])
def subir_stock():
    stock_file = request.files.get("stock")
    if not stock_file or stock_file.filename == "":
        flash("Subí la planilla STOCK con los precios.")
        return redirect(url_for("index"))
    if _ext(stock_file.filename) != ".xlsx":
        flash("La planilla STOCK debe ser un archivo .xlsx.")
        return redirect(url_for("index"))

    stock_file.save(STOCK_CACHE_PATH)
    global _PRICE_CACHE
    _PRICE_CACHE = {"path": None, "mtime": None, "prices": None}
    with open(STOCK_CACHE_META, "w") as f:
        f.write(f"{secure_filename(stock_file.filename)} - cargada {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    flash("Planilla STOCK con precios cargada correctamente.")
    return redirect(url_for("index"))


@app.route("/procesar", methods=["POST"])
def procesar():
    pedido_file = request.files.get("pedido")
    grande_file = request.files.get("grande")

    if not pedido_file or pedido_file.filename == "":
        flash("Subí la planilla de pedido (la que hay que resolver).")
        return redirect(url_for("index"))

    if _ext(pedido_file.filename) != ".xlsx":
        flash("La planilla de pedido debe ser .xlsx (para poder mantener el mismo formato/estilo al insertar la columna).")
        return redirect(url_for("index"))

    # Resolver planilla grande: la subida ahora, o la que quedó cacheada
    if grande_file and grande_file.filename != "":
        if _ext(grande_file.filename) not in ALLOWED_EXT:
            flash("La planilla grande debe ser .xls o .xlsx")
            return redirect(url_for("index"))
        grande_ext = _ext(grande_file.filename)
        grande_path = GRANDE_CACHE_PATH_XLS if grande_ext == ".xls" else GRANDE_CACHE_PATH_XLSX
        # Limpiar el otro formato cacheado para no usar una versión vieja por error
        other_path = GRANDE_CACHE_PATH_XLSX if grande_ext == ".xls" else GRANDE_CACHE_PATH_XLS
        if os.path.exists(other_path):
            os.remove(other_path)
        grande_file.save(grande_path)
        with open(GRANDE_CACHE_META, "w") as f:
            f.write(f"{secure_filename(grande_file.filename)} — cargada {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    else:
        if os.path.exists(GRANDE_CACHE_PATH_XLS):
            grande_path = GRANDE_CACHE_PATH_XLS
        elif os.path.exists(GRANDE_CACHE_PATH_XLSX):
            grande_path = GRANDE_CACHE_PATH_XLSX
        else:
            flash("No hay planilla grande cargada todavía. Subí la planilla de Equivalencia primero.")
            return redirect(url_for("index"))

    pedido_bytes = io.BytesIO(pedido_file.read())
    pedido_bytes.name = pedido_file.filename

    try:
        output = _procesar(grande_path, pedido_bytes)
    except Exception as e:
        flash(f"Error procesando los archivos: {e}")
        return redirect(url_for("index"))

    nombre_salida = f"{secure_filename(os.path.splitext(pedido_file.filename)[0])} - con equivalencia.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=nombre_salida,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/health")
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
