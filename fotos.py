"""Fotos recientes de WhatsApp (Descargas) para el IDS."""

from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path

EXTS = {".jpg", ".jpeg", ".png", ".webp", ".jfif"}
CARPETAS = [
    Path(r"C:\Users\Eduardo\Downloads"),
    Path(r"C:\Users\Eduardo\Pictures"),
]
PERMITIDAS = [
    Path(r"C:\Users\Eduardo\Downloads").resolve(),
    Path(r"C:\Users\Eduardo\Pictures").resolve(),
    Path(r"C:\Users\Eduardo\Documents").resolve(),
]


def ruta_permitida(ruta: Path) -> bool:
    try:
        p = ruta.resolve()
    except OSError:
        return False
    if not p.is_file() or p.suffix.lower() not in EXTS:
        return False
    return any(p == base or base in p.parents for base in PERMITIDAS)


def _es_whatsapp(nombre: str) -> bool:
    n = nombre.lower()
    return "whatsapp" in n or n.startswith("wa ")


def listar_whatsapp(limite: int = 24) -> list[dict]:
    vistos: dict[str, Path] = {}
    for carpeta in CARPETAS:
        if not carpeta.exists():
            continue
        for p in carpeta.iterdir():
            if not p.is_file() or p.suffix.lower() not in EXTS:
                continue
            if not _es_whatsapp(p.name):
                continue
            clave = p.name.lower()
            if clave not in vistos or p.stat().st_mtime > vistos[clave].stat().st_mtime:
                vistos[clave] = p
    items = sorted(vistos.values(), key=lambda x: x.stat().st_mtime, reverse=True)[:limite]
    out = []
    for p in items:
        st = p.stat()
        out.append(
            {
                "ruta": str(p),
                "nombre": p.name,
                "bytes": st.st_size,
                "fecha": datetime.fromtimestamp(st.st_mtime).strftime("%d/%m %H:%M"),
            }
        )
    return out


def miniatura(ruta: Path, ancho: int = 240) -> bytes:
    from PIL import Image

    with Image.open(ruta) as im:
        im = im.convert("RGB")
        im.thumbnail((ancho, int(ancho * 0.8)))
        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=78)
        return buf.getvalue()
