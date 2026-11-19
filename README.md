# Reporte de seguridad

Editor del reporte fotográfico de seguridad: fecha, turno, fotos del IDS y el antes / después de cada labor.

## Abrir en la PC

Doble clic en `start.bat`.

Abre http://127.0.0.1:8765

Hace falta Python 3. El arranque instala lo de `requirements.txt` (`python-pptx` y `Pillow`) y levanta `server.py`. Sin ese servidor no se arma el PowerPoint.

## En GitHub Pages

[La página](https://zettafuego.github.io/PTTDIARIO/) muestra el frente. El `.pptx` se genera solo cuando corre `server.py` en la PC.

## Qué usa

- `index.html` y `app.js` — la pantalla
- `server.py`, `editor.py`, `fotos.py` — arman la presentación
- `plantilla.pptx` — la plantilla del reporte

## Copia

[PPtS-AUTOMATICO](https://github.com/zettafuego/PPtS-AUTOMATICO) es el mismo editor. Sigue este repositorio.
