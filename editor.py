"""Lee y actualiza fecha, turno, slides y fotos IDS del reporte PPTX."""

from __future__ import annotations

import io
import re
import shutil
from copy import deepcopy
from datetime import date
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

MESES = [
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
]
MESES_RE = (
    r"enero|febrero|marzo|abril|mayo|junio|julio|agosto|"
    r"septiembre|setiembre|octubre|noviembre|diciembre"
)
DATE_RE = re.compile(rf"\b(\d{{1,2}})\s+de\s+({MESES_RE})\b", re.IGNORECASE)
TURNO_FRASE_RE = re.compile(r"\bTURNO\s+(D[IÍ]A|NOCHE)\b", re.IGNORECASE)
TURNO_SOLO_RE = re.compile(r"^(D[IÍ]A|NOCHE)$", re.IGNORECASE)
WS_RE = re.compile(r"\s+")
MIN_SLIDES = 2
MAX_SLIDES = 15


def fecha_tabla(d: date) -> str:
    return f"{d.day} de {MESES[d.month - 1]}"


def fecha_portada(d: date) -> str:
    return f"{d.day}  de {MESES[d.month - 1].upper()}"


def normalizar_turno(valor: str | None) -> str | None:
    if not valor:
        return None
    t = WS_RE.sub(" ", valor).strip().upper().replace("Í", "I")
    if t in {"DIA", "DÍA", "TURNO DIA", "TURNO DÍA"}:
        return "DIA"
    if t in {"NOCHE", "TURNO NOCHE"}:
        return "NOCHE"
    m = TURNO_FRASE_RE.search(valor)
    if m:
        return "NOCHE" if m.group(1).upper().replace("Í", "I") == "NOCHE" else "DIA"
    m = TURNO_SOLO_RE.match(t)
    if m:
        return "NOCHE" if "NOCHE" in t else "DIA"
    return None


def etiqueta_turno(turno: str) -> str:
    return "NOCHE" if normalizar_turno(turno) == "NOCHE" else "DIA"


def parsear_fecha_texto(texto: str) -> date | None:
    m = DATE_RE.search(texto or "")
    if not m:
        return None
    dia = int(m.group(1))
    mes_nombre = m.group(2).lower().replace("setiembre", "septiembre")
    try:
        mes = MESES.index(mes_nombre) + 1
        anio = date.today().year
        return date(anio, mes, dia)
    except (ValueError, IndexError):
        return None


def iter_shapes(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(sh.shapes)


def set_first_run_text(paragraph, texto: str) -> None:
    if paragraph.runs:
        paragraph.runs[0].text = texto
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.text = texto


def set_cell_text(cell, texto: str) -> None:
    tf = cell.text_frame
    first = True
    for paragraph in tf.paragraphs:
        if first:
            set_first_run_text(paragraph, texto)
            first = False
        else:
            set_first_run_text(paragraph, "")


def texto_shape(shape) -> str:
    if not getattr(shape, "has_text_frame", False):
        return ""
    return "\n".join(p.text for p in shape.text_frame.paragraphs)


def inspeccionar(ruta: str | Path) -> dict:
    ruta = Path(ruta)
    prs = Presentation(str(ruta))
    detalle = []
    fechas = []
    turnos = []

    for i, slide in enumerate(prs.slides, start=1):
        fecha = None
        turno = None
        tipo = "contenido"
        titulo = ""

        for shape in iter_shapes(slide.shapes):
            if shape.has_table:
                tabla = shape.table
                headers = [WS_RE.sub(" ", c.text).strip().lower() for c in tabla.rows[0].cells]
                if len(tabla.rows) >= 2:
                    fila = tabla.rows[1]
                    if "fecha" in headers:
                        fecha = WS_RE.sub(" ", fila.cells[headers.index("fecha")].text).strip()
                    if "turno" in headers:
                        turno = WS_RE.sub(" ", fila.cells[headers.index("turno")].text).strip()
                tipo = "tabla"
            elif shape.has_text_frame:
                t = texto_shape(shape).strip()
                if DATE_RE.search(t) and TURNO_FRASE_RE.search(t):
                    tipo = "portada"
                    titulo = WS_RE.sub(" ", t).strip()
                    m_fecha = DATE_RE.search(t)
                    if m_fecha:
                        fecha = WS_RE.sub(" ", m_fecha.group(0)).strip()
                    m_turno = TURNO_FRASE_RE.search(t)
                    if m_turno:
                        turno = m_turno.group(0)

        if fecha:
            fechas.append(fecha)
        if turno:
            turnos.append(turno)
        if es_slide_ids(slide):
            tipo = "ids"
        fotos_n = 0
        for shape in iter_shapes(slide.shapes):
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and not es_logo(shape):
                fotos_n += 1
        detalle.append(
            {
                "n": i,
                "tipo": tipo,
                "fecha": fecha or "",
                "turno": etiqueta_turno(turno) if turno else "",
                "turno_original": turno or "",
                "titulo": titulo,
                "fotos": fotos_n,
            }
        )

    turno_principal = None
    if turnos:
        counts = {}
        for t in turnos:
            key = etiqueta_turno(t)
            counts[key] = counts.get(key, 0) + 1
        turno_principal = max(counts, key=counts.get)

    fecha_principal = fechas[0] if fechas else ""
    fechas_unicas = []
    for f in fechas:
        clave = WS_RE.sub(" ", f).strip().lower()
        if clave not in {WS_RE.sub(" ", x).strip().lower() for x in fechas_unicas}:
            fechas_unicas.append(f)

    ids_idx = next((d["n"] for d in detalle if d["tipo"] == "ids"), 2)
    fotos_ids = next((d["fotos"] for d in detalle if d["tipo"] == "ids"), 0)
    return {
        "ok": True,
        "ruta": str(ruta),
        "nombre": ruta.name,
        "slides": len(prs.slides),
        "min_slides": MIN_SLIDES,
        "max_slides": MAX_SLIDES,
        "fecha": fecha_principal,
        "fechas": fechas_unicas,
        "turno": turno_principal or "",
        "ids_slide": ids_idx,
        "fotos_ids": fotos_ids,
        "detalle": detalle,
    }


def _reemplazar_fecha_en_texto(texto: str, nueva: date) -> str:
    def _sub(m):
        original = m.group(0)
        mes = m.group(2)
        if mes.isupper():
            return fecha_portada(nueva)
        if mes[:1].isupper() and mes[1:].islower():
            return f"{nueva.day} de {MESES[nueva.month - 1].capitalize()}"
        return fecha_tabla(nueva)

    return DATE_RE.sub(_sub, texto)


def _reemplazar_turno_en_texto(texto: str, nuevo: str) -> str:
    etiqueta = etiqueta_turno(nuevo)
    palabra = "NOCHE" if etiqueta == "NOCHE" else "DIA"

    def _frase(m):
        return re.sub(r"(D[IÍ]A|NOCHE)\s*$", palabra, m.group(0), flags=re.IGNORECASE)

    return TURNO_FRASE_RE.sub(_frase, texto)


def _actualizar_parrafo(paragraph, nueva_fecha: date | None, nuevo_turno: str | None) -> bool:
    original = paragraph.text
    if not original:
        return False
    nuevo = original
    if nueva_fecha:
        nuevo = _reemplazar_fecha_en_texto(nuevo, nueva_fecha)
    if nuevo_turno:
        nuevo = _reemplazar_turno_en_texto(nuevo, nuevo_turno)
    if nuevo == original:
        return False
    set_first_run_text(paragraph, nuevo)
    return True


def es_logo(shape) -> bool:
    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        return False
    return shape.top < 600_000 and shape.left > 6_000_000 and shape.height < 800_000


def es_slide_ids(slide) -> bool:
    for shape in iter_shapes(slide.shapes):
        t = texto_shape(shape).strip().upper()
        if t.startswith("IDS"):
            return True
    return False


def borrar_slide(prs: Presentation, index: int) -> None:
    sld_id_lst = prs.slides._sldIdLst
    sld_id = list(sld_id_lst)[index]
    r_id = sld_id.get(qn("r:id"))
    prs.part.drop_rel(r_id)
    sld_id_lst.remove(sld_id)


def duplicar_slide(prs: Presentation, index: int) -> None:
    src = prs.slides[index]
    dest = prs.slides.add_slide(src.slide_layout)
    for shape in list(dest.shapes):
        shape.element.getparent().remove(shape.element)
    for shape in src.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            dest.shapes.add_picture(
                io.BytesIO(shape.image.blob),
                shape.left,
                shape.top,
                shape.width,
                shape.height,
            )
        else:
            dest.shapes._spTree.append(deepcopy(shape.element))


def ajustar_slides(prs: Presentation, n: int) -> int:
    n = max(MIN_SLIDES, min(MAX_SLIDES, int(n)))
    while len(prs.slides) > n:
        borrar_slide(prs, len(prs.slides) - 1)
    plantilla = len(prs.slides) - 1
    while len(prs.slides) < n:
        duplicar_slide(prs, plantilla)
        plantilla = len(prs.slides) - 1
    return n


def _tamano_imagen(data: bytes) -> tuple[int, int]:
    from PIL import Image

    with Image.open(io.BytesIO(data)) as im:
        return im.size


def _encajar(caja_w: int, caja_h: int, img_w: int, img_h: int) -> tuple[int, int, int, int]:
    if img_w <= 0 or img_h <= 0:
        return 0, 0, caja_w, caja_h
    escala = min(caja_w / img_w, caja_h / img_h)
    w = max(1, int(img_w * escala))
    h = max(1, int(img_h * escala))
    x = (caja_w - w) // 2
    y = (caja_h - h) // 2
    return x, y, w, h


def poner_fotos_ids(prs: Presentation, fotos: list[Path | bytes]) -> int:
    if not fotos:
        return 0
    slide = None
    for s in prs.slides:
        if es_slide_ids(s):
            slide = s
            break
    if slide is None:
        if len(prs.slides) < 2:
            return 0
        slide = prs.slides[1]

    for shape in list(slide.shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and not es_logo(shape):
            shape.element.getparent().remove(shape.element)

    blobs: list[bytes] = []
    for item in fotos[:3]:
        if item is None:
            continue
        if isinstance(item, (bytes, bytearray)):
            blobs.append(bytes(item))
        else:
            blobs.append(Path(item).read_bytes())
    if not blobs:
        return 0

    slide_w = int(prs.slide_width)
    slide_h = int(prs.slide_height)
    margen = 152400
    top = 1686651
    gap = 100000
    n = len(blobs)
    avail_w = slide_w - margen * 2
    avail_h = slide_h - top - 150000
    caja_w = (avail_w - gap * (n - 1)) // n
    caja_h = avail_h

    for i, blob in enumerate(blobs):
        caja_x = margen + i * (caja_w + gap)
        try:
            img_w, img_h = _tamano_imagen(blob)
            ox, oy, w, h = _encajar(caja_w, caja_h, img_w, img_h)
        except Exception:
            ox, oy, w, h = 0, 0, caja_w, caja_h
        slide.shapes.add_picture(io.BytesIO(blob), caja_x + ox, top + oy, w, h)
    return len(blobs)


def llenar_textos_reportes(prs: Presentation, por_slide: dict) -> int:
    """Rellena Labor (tabla) y Descripción / Plan de acción en cada diapo 2+."""
    if not por_slide:
        return 0
    cambios = 0
    for i, slide in enumerate(prs.slides):
        n = i + 1
        data = por_slide.get(n) or por_slide.get(str(n)) or {}
        if not isinstance(data, dict):
            continue
        labor = str(data.get("labor") or "").strip()
        descripcion = str(data.get("descripcion") or "").strip()
        plan = str(data.get("plan") or "").strip()
        if labor:
            for shape in iter_shapes(slide.shapes):
                if not shape.has_table:
                    continue
                tabla = shape.table
                headers = [WS_RE.sub(" ", c.text).strip().lower() for c in tabla.rows[0].cells]
                if "labor" not in headers or len(tabla.rows) < 2:
                    continue
                cell = tabla.rows[1].cells[headers.index("labor")]
                if WS_RE.sub(" ", cell.text).strip() != labor:
                    set_cell_text(cell, labor)
                    cambios += 1
        if n >= 3 and (descripcion or plan):
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                plano = shape.text_frame.text.lower()
                if "descripci" not in plano:
                    continue
                plan_label = "Medidas correctivas:" if "medidas correctivas" in plano else "Plan de acción:"
                tf = shape.text_frame
                while len(tf.paragraphs) < 4:
                    tf.add_paragraph()
                wanted = ["Descripción:", descripcion, plan_label, plan]
                paras = list(tf.paragraphs)
                for idx, texto in enumerate(wanted):
                    set_first_run_text(paras[idx], texto)
                    for run in paras[idx].runs:
                        run.font.bold = idx in (0, 2)
                for extra in paras[4:]:
                    set_first_run_text(extra, "")
                cambios += 1
                break
    return cambios


def _poner_borde_foto(picture, rgb: tuple[int, int, int], grosor_pt: float = 4.5) -> None:
    from pptx.dml.color import RGBColor
    from pptx.util import Pt

    picture.line.color.rgb = RGBColor(*rgb)
    picture.line.width = Pt(grosor_pt)


def poner_fotos_reportes(prs: Presentation, por_slide: dict) -> int:
    """por_slide: {3: {'antes': bytes|Path, 'despues': bytes|Path}, ...}"""
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Emu, Pt

    if not por_slide:
        return 0
    puestas = 0
    for i, slide in enumerate(prs.slides):
        n = i + 1
        if n < 3:
            continue
        par = por_slide.get(n) or por_slide.get(str(n))
        if not par:
            continue
        antes = par.get("antes") if isinstance(par, dict) else None
        despues = par.get("despues") if isinstance(par, dict) else None
        if not antes and not despues:
            continue

        for shape in list(slide.shapes):
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and not es_logo(shape):
                shape.element.getparent().remove(shape.element)

        def _blob(item):
            if item is None:
                return None
            if isinstance(item, (bytes, bytearray)):
                return bytes(item)
            return Path(item).read_bytes()

        blobs = [
            ("ANTES", _blob(antes), (0xC6, 0x28, 0x28)),
            ("DESPUÉS", _blob(despues), (0x2E, 0x7D, 0x32)),
        ]
        area_x = 2680000
        area_w = int(prs.slide_width) - area_x - 160000
        gap = 140000
        col_w = (area_w - gap) // 2
        label_h = 260000
        top_label = 1520000
        top_pic = top_label + label_h + 30000
        pic_h = 3050000

        for col, (label, blob, rgb) in enumerate(blobs):
            x = area_x + col * (col_w + gap)
            tb = slide.shapes.add_textbox(Emu(x), Emu(top_label), Emu(col_w), Emu(label_h))
            tf = tb.text_frame
            tf.word_wrap = False
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = label
            run.font.bold = True
            run.font.size = Pt(16)
            run.font.color.rgb = RGBColor(*rgb)
            run.font.name = "Arial"
            if blob:
                try:
                    img_w, img_h = _tamano_imagen(blob)
                    ox, oy, w, h = _encajar(col_w, pic_h, img_w, img_h)
                except Exception:
                    ox, oy, w, h = 0, 0, col_w, pic_h
                pic = slide.shapes.add_picture(
                    io.BytesIO(blob), Emu(x + ox), Emu(top_pic + oy), Emu(w), Emu(h)
                )
                _poner_borde_foto(pic, rgb)
                puestas += 1
    return puestas


def aplicar(
    ruta: str | Path,
    salida: str | Path,
    fecha: date | None = None,
    turno: str | None = None,
    n_slides: int | None = None,
    fotos_ids: list[Path | bytes] | None = None,
    fotos_reportes: dict | None = None,
) -> dict:
    ruta = Path(ruta)
    salida = Path(salida)
    if ruta.resolve() != salida.resolve():
        salida.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ruta, salida)

    prs = Presentation(str(salida))
    cambios = 0
    turno_norm = etiqueta_turno(turno) if turno else None
    if n_slides:
        antes = len(prs.slides)
        ajustar_slides(prs, n_slides)
        if len(prs.slides) != antes:
            cambios += 1
    if fotos_ids:
        n_fotos = poner_fotos_ids(prs, fotos_ids)
        cambios += n_fotos
    if fotos_reportes:
        cambios += llenar_textos_reportes(prs, fotos_reportes)
        cambios += poner_fotos_reportes(prs, fotos_reportes)

    for slide in prs.slides:
        for shape in iter_shapes(slide.shapes):
            if shape.has_table:
                tabla = shape.table
                headers = [WS_RE.sub(" ", c.text).strip().lower() for c in tabla.rows[0].cells]
                if len(tabla.rows) < 2:
                    continue
                fila = tabla.rows[1]
                if fecha and "fecha" in headers:
                    cell = fila.cells[headers.index("fecha")]
                    actual = WS_RE.sub(" ", cell.text).strip()
                    nuevo = fecha_tabla(fecha)
                    if DATE_RE.search(actual) or actual == "":
                        if actual != nuevo:
                            set_cell_text(cell, nuevo)
                            cambios += 1
                if turno_norm and "turno" in headers:
                    cell = fila.cells[headers.index("turno")]
                    actual = WS_RE.sub(" ", cell.text).strip()
                    estilo = actual if actual else "DIA"
                    if estilo.isupper():
                        nuevo = "NOCHE" if turno_norm == "NOCHE" else "DIA"
                    elif estilo[:1].isupper():
                        nuevo = "Noche" if turno_norm == "NOCHE" else "Dia"
                    else:
                        nuevo = "noche" if turno_norm == "NOCHE" else "dia"
                    if actual != nuevo:
                        set_cell_text(cell, nuevo)
                        cambios += 1
            elif shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    if _actualizar_parrafo(paragraph, fecha, turno_norm):
                        cambios += 1

    prs.save(str(salida))
    info = inspeccionar(salida)
    info["cambios"] = cambios
    info["salida"] = str(salida)
    return info


TITULO_REPORTE = "Reporte de seguridad"


def nombre_descarga(fecha: date | None, turno: str | None) -> str:
    partes = [TITULO_REPORTE]
    if fecha:
        partes.append(fecha_tabla(fecha))
    if turno:
        partes.append(f"TURNO {etiqueta_turno(turno)}")
    nombre = " - ".join(partes) + ".pptx"
    return re.sub(r'[\\/:*?"<>|]', "", nombre)


def nombre_salida(ruta: Path, fecha: date | None, turno: str | None) -> Path:
    return ruta.parent / nombre_descarga(fecha, turno)
