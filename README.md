# Editor de reporte de seguridad

App **web estática**. Corre en el navegador (GitHub Pages / online) y también en tu PC.

La plantilla por defecto es `plantilla.pptx` (tu reporte original).

- Fecha y turno (Día / Noche)
- Número de diapositivas
- 3 fotos del IDS (súbelas; en GitHub no se puede leer WhatsApp de tu PC)
- Desde la diapositiva 3: labor, descripción, plan de acción y fotos **ANTES** (rojo) / **DESPUÉS** (verde)
- Fotos del IDS sin deformar (se encajan en el recuadro)
- Descarga: `Reporte de seguridad - 19 de agosto - TURNO DIA.pptx`

## En tu PC

Doble clic en `start.bat` → `http://127.0.0.1:8765`

## En GitHub, online

Así **sí** funciona, porque ya no necesita Python ni carpetas de Windows.

1. Crea un repo (por ejemplo `reporte-seguridad`)
2. Sube **estos archivos** (en la raíz):
   - `index.html`
   - `app.js`
   - `styles.css`
   - `manifest.webmanifest`
   - `plantilla.pptx`
   - `.nojekyll`
3. En el repo: **Settings → Pages → Deploy from a branch → main / root**
4. Entra a:

`https://TU-USUARIO.github.io/reporte-seguridad/`

En internet las 3 fotos del IDS se **suben** (bájala de WhatsApp y elígelas). GitHub no puede ver Descargas de tu PC.
