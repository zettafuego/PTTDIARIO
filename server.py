"""Servidor local del editor de reporte de seguridad."""

from __future__ import annotations

import json
import os
import shutil
import socket
import sys
import traceback
import uuid
import webbrowser
from datetime import date
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
PLANTILLA = ROOT / "plantilla.pptx"
ORIGEN_PPT = Path(r"C:\Users\Eduardo\Downloads\Reporte de seguridad (1).pptx")
PORT = 8765

sys.path.insert(0, str(ROOT))
from editor import aplicar, inspeccionar, nombre_descarga  # noqa: E402
from fotos import listar_whatsapp, miniatura, ruta_permitida  # noqa: E402

SESIONES = ROOT / ".sesiones"
DESCARGAS = SESIONES / "descargas"
SESIONES.mkdir(exist_ok=True)
DESCARGAS.mkdir(exist_ok=True)
TOKENS: dict[str, tuple[Path, str]] = {}


def json_bytes(data: dict, code: int = 200) -> tuple[int, bytes, str]:
    raw = json.dumps(data, ensure_ascii=False).encode("utf-8")
    return code, raw, "application/json; charset=utf-8"


def parse_fecha(valor: str | None) -> date | None:
    if not valor:
        return None
    valor = valor.strip()
    if not valor:
        return None
    partes = valor.split("-")
    if len(partes) != 3:
        raise ValueError("Usa una fecha con formato AAAA-MM-DD")
    return date(int(partes[0]), int(partes[1]), int(partes[2]))


def asegurar_plantilla() -> Path:
    if PLANTILLA.exists():
        return PLANTILLA
    if ORIGEN_PPT.exists():
        shutil.copy2(ORIGEN_PPT, PLANTILLA)
        return PLANTILLA
    raise FileNotFoundError("No está la plantilla del reporte de seguridad")


def buscar_default() -> Path | None:
    try:
        return asegurar_plantilla()
    except FileNotFoundError:
        return None


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def log_message(self, fmt: str, *args) -> None:
        sys.stderr.write("[app] " + (fmt % args) + "\n")

    def _send(self, code: int, body: bytes, content_type: str) -> None:
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _send_download(self, path: Path, nombre: str) -> None:
        data = path.read_bytes()
        ascii_name = nombre.encode("ascii", "replace").decode("ascii")
        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        self.send_header("Content-Length", str(len(data)))
        self.send_header(
            "Content-Disposition",
            f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(nombre)}",
        )
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return {}
        raw = self.rfile.read(length)
        if not raw:
            return {}
        return json.loads(raw.decode("utf-8"))

    def _guardar_upload(self, payload: dict) -> Path:
        nombre = payload.get("nombre") or "reporte.pptx"
        b64 = payload.get("archivo_b64")
        if not b64:
            raise ValueError("No llegó el archivo")
        import base64

        data = base64.b64decode(b64)
        destino = SESIONES / Path(nombre).name
        destino.write_bytes(data)
        return destino

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/estado":
            ruta = buscar_default()
            data = {
                "ok": True,
                "hoy": date.today().isoformat(),
                "default": True,
                "titulo": "Reporte de seguridad",
                "nombre_descarga": nombre_descarga(date.today(), "DIA"),
            }
            if ruta:
                try:
                    info = inspeccionar(ruta)
                    info["nombre"] = "Reporte de seguridad"
                    info["ruta"] = "plantilla"
                    data.update(info)
                except Exception as e:
                    data["aviso"] = str(e)
            code, body, ctype = json_bytes(data)
            return self._send(code, body, ctype)

        if parsed.path == "/api/descargar":
            qs = parse_qs(parsed.query)
            token = (qs.get("token") or [""])[0]
            item = TOKENS.get(token)
            if not item:
                return self._send(404, b"no", "text/plain")
            path, nombre = item
            if not path.exists():
                return self._send(404, b"no", "text/plain")
            return self._send_download(path, nombre)

        if parsed.path == "/manifest.webmanifest":
            data = (STATIC / "manifest.webmanifest").read_bytes()
            return self._send(200, data, "application/manifest+json")

        if parsed.path == "/api/abrir":
            qs = parse_qs(parsed.query)
            ruta = (qs.get("ruta") or [""])[0]
            if ruta and Path(ruta).exists():
                os.startfile(ruta)  # type: ignore[attr-defined]
            code, body, ctype = json_bytes({"ok": True})
            return self._send(code, body, ctype)

        if parsed.path == "/api/whatsapp":
            fotos = listar_whatsapp()
            code, body, ctype = json_bytes({"ok": True, "fotos": fotos})
            return self._send(code, body, ctype)

        if parsed.path == "/api/foto":
            qs = parse_qs(parsed.query)
            ruta = Path((qs.get("ruta") or [""])[0])
            if not ruta_permitida(ruta):
                return self._send(404, b"no", "text/plain")
            try:
                if (qs.get("thumb") or ["1"])[0] != "0":
                    data = miniatura(ruta)
                    return self._send(200, data, "image/jpeg")
                raw = ruta.read_bytes()
                ctype = "image/jpeg" if ruta.suffix.lower() in {".jpg", ".jpeg", ".jfif"} else "image/png"
                return self._send(200, raw, ctype)
            except OSError:
                return self._send(404, b"no", "text/plain")

        if parsed.path == "/api/abrir-descargas":
            os.startfile(r"C:\Users\Eduardo\Downloads")  # type: ignore[attr-defined]
            code, body, ctype = json_bytes({"ok": True})
            return self._send(code, body, ctype)

        if parsed.path == "/" or parsed.path == "/index.html":
            html = (ROOT / "index.html").read_bytes()
            return self._send(200, html, "text/html; charset=utf-8")

        if parsed.path == "/app.js":
            return self._send(200, (ROOT / "app.js").read_bytes(), "application/javascript; charset=utf-8")

        if parsed.path == "/styles.css":
            return self._send(200, (ROOT / "styles.css").read_bytes(), "text/css; charset=utf-8")

        if parsed.path == "/favicon.ico":
            return self._send(204, b"", "image/x-icon")

        return super().do_GET()

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        try:
            payload = self._read_json()
            if parsed.path == "/api/inspeccionar":
                if payload.get("archivo_b64"):
                    ruta = self._guardar_upload(payload)
                else:
                    ruta = asegurar_plantilla()
                info = inspeccionar(ruta)
                info["nombre"] = "Reporte de seguridad"
                info["ruta"] = "plantilla"
                code, body, ctype = json_bytes(info)
                return self._send(code, body, ctype)

            if parsed.path == "/api/aplicar":
                fecha = parse_fecha(payload.get("fecha"))
                turno = payload.get("turno") or None
                n_slides = payload.get("n_slides")
                fotos_ids = decodificar_fotos(payload.get("fotos_ids") or [])
                fotos_reportes = decodificar_fotos_reportes(payload.get("fotos_reportes") or {})
                if not fecha and not turno and not n_slides and not fotos_ids and not fotos_reportes:
                    raise ValueError("Elige fecha, turno, diapositivas o fotos")

                if payload.get("archivo_b64"):
                    origen = self._guardar_upload(payload)
                else:
                    origen = asegurar_plantilla()

                nombre = nombre_descarga(fecha, turno)
                token = uuid.uuid4().hex
                salida = DESCARGAS / f"{token}.pptx"
                info = aplicar(
                    origen,
                    salida,
                    fecha=fecha,
                    turno=turno,
                    n_slides=int(n_slides) if n_slides else None,
                    fotos_ids=fotos_ids or None,
                    fotos_reportes=fotos_reportes or None,
                )
                TOKENS[token] = (salida, nombre)
                info["nombre_descarga"] = nombre
                info["token"] = token
                info["descarga"] = f"/api/descargar?token={token}"
                info["ruta"] = "plantilla"
                info["nombre"] = "Reporte de seguridad"
                code, body, ctype = json_bytes(info)
                return self._send(code, body, ctype)

            code, body, ctype = json_bytes({"ok": False, "error": "Ruta no encontrada"}, 404)
            return self._send(code, body, ctype)
        except Exception as e:
            traceback.print_exc()
            code, body, ctype = json_bytes({"ok": False, "error": str(e)}, 400)
            return self._send(code, body, ctype)


def decodificar_fotos(items: list) -> list:
    import base64

    out = []
    for item in items:
        if not item:
            continue
        if isinstance(item, str):
            p = Path(item)
            if not ruta_permitida(p):
                raise ValueError(f"Foto no permitida: {p.name}")
            out.append(p)
            continue
        ruta = item.get("ruta")
        if ruta:
            p = Path(ruta)
            if not ruta_permitida(p):
                raise ValueError(f"Foto no permitida: {p.name}")
            out.append(p)
            continue
        b64 = item.get("b64")
        if b64:
            out.append(base64.b64decode(b64))
    return out[:3]


def decodificar_fotos_reportes(mapa: dict) -> dict:
    import base64

    out = {}
    if not isinstance(mapa, dict):
        return out
    for clave, par in mapa.items():
        try:
            n = int(clave)
        except (TypeError, ValueError):
            continue
        if n < 3 or not isinstance(par, dict):
            continue
        item = {}
        for campo in ("labor", "descripcion", "plan"):
            val = par.get(campo)
            if val:
                item[campo] = str(val)
        for kind in ("antes", "despues"):
            raw = par.get(kind)
            if not raw:
                continue
            if isinstance(raw, str):
                p = Path(raw)
                if ruta_permitida(p):
                    item[kind] = p
                continue
            ruta = raw.get("ruta")
            if ruta:
                p = Path(ruta)
                if ruta_permitida(p):
                    item[kind] = p
                continue
            b64 = raw.get("b64")
            if b64:
                item[kind] = base64.b64decode(b64)
        if item:
            out[n] = item
    return out


def puerto_ocupado(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        return sock.connect_ex(("127.0.0.1", port)) == 0


def main() -> None:
    url = f"http://127.0.0.1:{PORT}/"
    print()
    print("  Reporte de seguridad  ·  fecha, turno, slides y fotos IDS")
    print(f"  {url}")
    print("  Ctrl+C para cerrar")
    print()
    webbrowser.open(url)
    if puerto_ocupado(PORT):
        print("  Ya estaba abierta. Se reabrió el navegador.")
        return
    httpd = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nCerrado.")


if __name__ == "__main__":
    main()
