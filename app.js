const $ = (id) => document.getElementById(id);

const MESES = [
  "enero", "febrero", "marzo", "abril", "mayo", "junio",
  "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
];
const MESES_RE = MESES.join("|") + "|setiembre";
const DATE_RE = new RegExp(`(\\d{1,2})\\s+de\\s+(${MESES_RE})`, "gi");
const TITULO = "Reporte de seguridad";
const PLANTILLA_URL = "plantilla.pptx";

let plantillaBuf = null;
const fotosIds = [null, null, null];
const fotosReportes = {};
let whatsapp = [];

function turnoSeleccionado() {
  const el = document.querySelector('input[name="turno"]:checked');
  return el ? el.value : "DIA";
}

function mostrarMsg(texto, ok) {
  const msg = $("msg");
  msg.hidden = !texto;
  msg.textContent = texto || "";
  msg.className = "msg " + (ok ? "ok" : "err");
}

function fechaDeInput() {
  const raw = $("fecha").value;
  if (!raw) return new Date();
  const [y, m, d] = raw.split("-").map(Number);
  return new Date(y, m - 1, d);
}

function fechaTabla(d) {
  return `${d.getDate()} de ${MESES[d.getMonth()]}`;
}

function nombreDescarga() {
  return `${TITULO} - ${fechaTabla(fechaDeInput())} - TURNO ${turnoSeleccionado()}.pptx`;
}

function pintarNombre() {
  $("nombre-dl").textContent = nombreDescarga();
}

function hintSlides(n) {
  const reportes = Math.max(0, n - 2);
  $("slides-hint").textContent = `1 portada + 1 IDS + ${reportes} reporte${reportes === 1 ? "" : "s"}`;
}

function setNSlides(n) {
  n = Math.max(2, Math.min(15, Number(n) || 2));
  $("n-slides-input").value = String(n);
  hintSlides(n);
  pintarSlotsReportes();
  return n;
}

function initReportesUI() {
  const wrap = $("reportes-fotos");
  if (!wrap || wrap.dataset.ready === "1") return;
  wrap.dataset.ready = "1";
  wrap.addEventListener("input", (e) => {
    const el = e.target.closest("[data-field]");
    if (!el || !wrap.contains(el)) return;
    const s = Number(el.dataset.slide);
    if (!fotosReportes[s]) fotosReportes[s] = datosSlideVacio();
    fotosReportes[s][el.dataset.field] = el.value;
  });
  wrap.addEventListener("change", (e) => {
    const input = e.target;
    if (!input || input.type !== "file") return;
    const slot = input.closest(".slot");
    if (!slot) return;
    const file = input.files && input.files[0];
    if (!file) return;
    const slide = Number(slot.dataset.slide);
    const kind = slot.dataset.kind;
    if (!fotosReportes[slide]) fotosReportes[slide] = datosSlideVacio();
    fotosReportes[slide][kind] = {
      blob: file,
      preview: URL.createObjectURL(file),
      nombre: file.name,
    };
    slot.classList.add("lleno");
    slot.style.backgroundImage = `url("${fotosReportes[slide][kind].preview}")`;
  });
}

function textosSlide(xml) {
  return [...xml.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)].map((m) => m[1]);
}

function clasificar(texts) {
  const plano = texts.join(" ");
  let tipo = "tabla";
  if (/TURNO\s+(D[IÍ]A|NOCHE)/i.test(plano)) tipo = "portada";
  if (texts.some((t) => t.trim().toUpperCase() === "IDS")) tipo = "ids";
  const fechaM = plano.match(DATE_RE);
  DATE_RE.lastIndex = 0;
  let turno = "";
  const tm = plano.match(/TURNO\s+(D[IÍ]A|NOCHE)/i);
  if (tm) turno = /NOCHE/i.test(tm[1]) ? "NOCHE" : "DIA";
  else {
    const solo = texts.find((t) => /^(D[IÍ]A|NOCHE)$/i.test(t.trim()));
    if (solo) turno = /NOCHE/i.test(solo) ? "NOCHE" : "DIA";
  }
  return {
    tipo,
    fecha: fechaM ? fechaM[0] : "",
    turno,
  };
}

function parseRels(xml) {
  const map = {};
  for (const m of xml.matchAll(/<Relationship Id="([^"]+)" Type="([^"]+)" Target="([^"]+)"\/>/g)) {
    map[m[1]] = { type: m[2], target: m[3] };
  }
  return map;
}

function slidesOrden(presXml, rels) {
  const out = [];
  for (const m of presXml.matchAll(/<p:sldId id="(\d+)" r:id="([^"]+)"\/>/g)) {
    const rel = rels[m[2]];
    if (rel) out.push({ id: Number(m[1]), rid: m[2], target: rel.target.replace(/^\//, "") });
  }
  return out;
}

async function inspeccionarBuf(buf) {
  const zip = await JSZip.loadAsync(buf);
  const presXml = await zip.file("ppt/presentation.xml").async("string");
  const relsXml = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
  const orden = slidesOrden(presXml, parseRels(relsXml));
  const detalle = [];
  for (let i = 0; i < orden.length; i += 1) {
    const path = "ppt/" + orden[i].target.replace(/^ppt\//, "");
    const xml = await zip.file(path).async("string");
    const meta = clasificar(textosSlide(xml));
    detalle.push({ n: i + 1, ...meta });
  }
  const fechas = detalle.map((d) => d.fecha).filter(Boolean);
  const turnos = detalle.map((d) => d.turno).filter(Boolean);
  return {
    slides: orden.length,
    fecha: fechas[0] || "",
    turno: turnos[0] || "",
    detalle,
    nombre: TITULO,
  };
}

function pintar(info) {
  $("nombre-archivo").textContent = info.nombre || TITULO;
  $("ruta-archivo").textContent = "Plantilla lista · funciona online y en GitHub";
  $("n-slides").textContent = info.slides ?? "—";
  $("fecha-actual").textContent = info.fecha || "Sin fecha";
  const turno = (info.turno || "").toUpperCase();
  $("turno-actual").textContent = turno || "—";
  $("turno-card").classList.toggle("dia", turno === "DIA");
  $("turno-card").classList.toggle("noche", turno === "NOCHE");
  if (turno === "DIA" || turno === "NOCHE") {
    document.querySelector(`input[name="turno"][value="${turno}"]`).checked = true;
  }
  if (info.slides) setNSlides(info.slides);
  const tbody = $("tbody");
  tbody.innerHTML = "";
  for (const fila of info.detalle || []) {
    const tr = document.createElement("tr");
    const clase = (fila.turno || "").toLowerCase();
    tr.innerHTML = `
      <td>${fila.n}</td>
      <td>${fila.fecha || "—"}</td>
      <td>${fila.turno ? `<span class="pill ${clase}">${fila.turno}</span>` : "—"}</td>
      <td>${fila.tipo}</td>`;
    tbody.appendChild(tr);
  }
  pintarNombre();
}

function nextNum(zip, re) {
  let max = 0;
  zip.forEach((path) => {
    const m = path.match(re);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return max + 1;
}

function nextRid(relsXml) {
  let max = 0;
  for (const m of relsXml.matchAll(/Id="rId(\d+)"/g)) max = Math.max(max, Number(m[1]));
  return `rId${max + 1}`;
}

function nextSldId(presXml) {
  let max = 255;
  for (const m of presXml.matchAll(/<p:sldId id="(\d+)"/g)) max = Math.max(max, Number(m[1]));
  return max + 1;
}

async function ajustarSlides(zip, n) {
  n = Math.max(2, Math.min(15, n));
  let presXml = await zip.file("ppt/presentation.xml").async("string");
  let relsXml = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
  let ct = await zip.file("[Content_Types].xml").async("string");
  let orden = slidesOrden(presXml, parseRels(relsXml));

  if (n < orden.length) {
    const keep = orden.slice(0, n);
    const lst = keep.map((s) => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join("");
    presXml = presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${lst}</p:sldIdLst>`);
  }

  while (orden.length < n) {
    const last = orden[orden.length - 1];
    const srcPath = "ppt/" + last.target.replace(/^ppt\//, "");
    const num = nextNum(zip, /^ppt\/slides\/slide(\d+)\.xml$/);
    const destPath = `ppt/slides/slide${num}.xml`;
    zip.file(destPath, await zip.file(srcPath).async("uint8array"));
    const relsSrc = srcPath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
    let slideRels = await zip.file(relsSrc).async("string");
    slideRels = slideRels.replace(/<Relationship[^>]*notesSlide[^>]*\/>/g, "");
    zip.file(`ppt/slides/_rels/slide${num}.xml.rels`, slideRels);
    const rid = nextRid(relsXml);
    const sid = nextSldId(presXml);
    relsXml = relsXml.replace(
      "</Relationships>",
      `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${num}.xml"/></Relationships>`
    );
    presXml = presXml.replace("</p:sldIdLst>", `<p:sldId id="${sid}" r:id="${rid}"/></p:sldIdLst>`);
    if (!ct.includes(`/ppt/slides/slide${num}.xml`)) {
      ct = ct.replace(
        "</Types>",
        `<Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml" PartName="/ppt/slides/slide${num}.xml"/></Types>`
      );
    }
    orden = slidesOrden(presXml, parseRels(relsXml));
  }

  zip.file("ppt/presentation.xml", presXml);
  zip.file("ppt/_rels/presentation.xml.rels", relsXml);
  zip.file("[Content_Types].xml", ct);
}

function aplicarTexto(xml, fecha, turno, esPortada) {
  const day = fecha.getDate();
  const mes = MESES[fecha.getMonth()];
  xml = xml.replace(DATE_RE, (full, _d, m) => {
    if (m === m.toUpperCase()) return `${day}  de ${mes.toUpperCase()}`;
    if (m[0] === m[0].toUpperCase() && m.slice(1) === m.slice(1).toLowerCase()) {
      return `${day} de ${mes[0].toUpperCase()}${mes.slice(1)}`;
    }
    return `${day} de ${mes}`;
  });
  if (esPortada) {
    xml = xml.replace(/<a:t([^>]*)>\d{1,2} <\/a:t>/, `<a:t$1>${day} </a:t>`);
    xml = xml.replace(
      new RegExp(`<a:t([^>]*)>(${MESES_RE})(\\t)?</a:t>`, "i"),
      `<a:t$1>${mes.toUpperCase()}$3</a:t>`
    );
  }
  xml = xml.replace(/TURNO\s+(D[IÍ]A|NOCHE)/gi, `TURNO ${turno}`);
  xml = xml.replace(/<a:t([^>]*)>(D[IÍ]A|Dia|NOCHE|Noche)<\/a:t>/g, (_, attrs, cur) => {
    let out = turno;
    if (cur === "Dia" || cur === "Noche") out = turno === "NOCHE" ? "Noche" : "Dia";
    return `<a:t${attrs}>${out}</a:t>`;
  });
  return xml;
}

function picXml(id, rid, x, y, w, h, borderHex) {
  const ln = borderHex
    ? `<a:ln w="63500" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="${borderHex}"/></a:solidFill><a:prstDash val="solid"/></a:ln>`
    : `<a:ln><a:noFill/></a:ln>`;
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Foto ${id}"/><p:cNvPicPr preferRelativeResize="0"/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"><a:alphaModFix/></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>${ln}</p:spPr></p:pic>`;
}

function captionXml(id, x, y, w, h, text, hex) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${text}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="es-PE" sz="1600" b="1" dirty="0"><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

function nextShapeId(xml) {
  let max = 300;
  for (const m of xml.matchAll(/cNvPr id="(\d+)"/g)) max = Math.max(max, Number(m[1]));
  return max + 1;
}

function quitarFotosContenido(xml) {
  return xml.replace(/<p:pic>[\s\S]*?<\/p:pic>/g, (pic) => {
    const off = pic.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
    const ext = pic.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
    if (!off || !ext) return pic;
    const x = Number(off[1]);
    const y = Number(off[2]);
    const h = Number(ext[2]);
    if (y < 600000 && x > 6000000 && h < 800000) return pic;
    return "";
  });
}

function escapeXml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function datosSlideVacio() {
  return { antes: null, despues: null, labor: "", descripcion: "", plan: "" };
}

function guardarCamposReportesDesdeDom() {
  document.querySelectorAll("#reportes-fotos [data-slide][data-field]").forEach((el) => {
    const s = Number(el.dataset.slide);
    if (!fotosReportes[s]) fotosReportes[s] = datosSlideVacio();
    fotosReportes[s][el.dataset.field] = el.value;
  });
  const laborIds = $("labor-ids");
  if (laborIds) {
    if (!fotosReportes[2]) fotosReportes[2] = datosSlideVacio();
    fotosReportes[2].labor = laborIds.value;
  }
}

function htmlCardReporte(s) {
  return `<article class="reporte-fotos" data-slide="${s}">
      <h3>Diapositiva ${s}</h3>
      <div class="campos">
        <label class="field">
          <span>Labor</span>
          <input type="text" data-slide="${s}" data-field="labor" placeholder="Ej: TJ 311 B" autocomplete="off" />
        </label>
        <label class="field">
          <span>Descripción</span>
          <textarea data-slide="${s}" data-field="descripcion" rows="3" placeholder="Qué se observó"></textarea>
        </label>
        <label class="field">
          <span>Plan de acción</span>
          <textarea data-slide="${s}" data-field="plan" rows="3" placeholder="Qué se va a hacer"></textarea>
        </label>
      </div>
      <div class="par-fotos">
        <label class="slot antes" data-slide="${s}" data-kind="antes">
          <input type="file" accept="image/*" hidden />
          <span class="slot-n">Antes</span>
          <span class="slot-ph">Clic para subir · borde rojo</span>
        </label>
        <label class="slot despues" data-slide="${s}" data-kind="despues">
          <input type="file" accept="image/*" hidden />
          <span class="slot-n">Después</span>
          <span class="slot-ph">Clic para subir · borde verde</span>
        </label>
      </div>
    </article>`;
}

function crearCardReporte(s) {
  if (!fotosReportes[s]) fotosReportes[s] = datosSlideVacio();
  const par = fotosReportes[s];
  const holder = document.createElement("div");
  holder.innerHTML = htmlCardReporte(s).trim();
  const card = holder.firstElementChild;
  const labor = card.querySelector('[data-field="labor"]');
  const desc = card.querySelector('[data-field="descripcion"]');
  const plan = card.querySelector('[data-field="plan"]');
  if (labor) labor.value = par.labor || "";
  if (desc) desc.value = par.descripcion || "";
  if (plan) plan.value = par.plan || "";
  ["antes", "despues"].forEach((kind) => {
    if (!par[kind] || !par[kind].preview) return;
    const slot = card.querySelector(`.slot[data-kind="${kind}"]`);
    if (!slot) return;
    slot.classList.add("lleno");
    slot.style.backgroundImage = `url("${par[kind].preview}")`;
  });
  return card;
}

function pintarSlotsReportes() {
  initReportesUI();
  guardarCamposReportesDesdeDom();
  const wrap = $("reportes-fotos");
  if (!wrap) return;
  const n = Number(($("n-slides-input") || {}).value) || 2;
  const empty = $("empty-reportes");
  if (empty) empty.hidden = n >= 3;
  wrap.hidden = n < 3;
  if (n < 3) return;
  wrap.querySelectorAll(".reporte-fotos").forEach((card) => {
    card.hidden = Number(card.dataset.slide) > n;
  });
  for (let s = 3; s <= n; s += 1) {
    let card = wrap.querySelector(`.reporte-fotos[data-slide="${s}"]`);
    if (!card) {
      card = crearCardReporte(s);
      wrap.appendChild(card);
    }
    card.hidden = false;
  }
}

function fillLaborTable(xml, labor) {
  if (!labor) return xml;
  return xml.replace(/<a:tbl>[\s\S]*?<\/a:tbl>/g, (tbl) => {
    const rows = tbl.match(/<a:tr[\s\S]*?<\/a:tr>/g);
    if (!rows || rows.length < 2) return tbl;
    const headerCells = rows[0].match(/<a:tc[\s\S]*?<\/a:tc>/g) || [];
    const idx = headerCells.findIndex((c) => />Labor<\/a:t>/.test(c));
    if (idx < 0) return tbl;
    let i = 0;
    const newRow = rows[1].replace(/<a:tc[\s\S]*?<\/a:tc>/g, (cell) => {
      if (i++ !== idx) return cell;
      if (/<a:t\s*\/>/.test(cell)) return cell.replace(/<a:t\s*\/>/, `<a:t>${escapeXml(labor)}</a:t>`);
      if (/<a:t><\/a:t>/.test(cell)) return cell.replace(/<a:t><\/a:t>/, `<a:t>${escapeXml(labor)}</a:t>`);
      return cell.replace(/<a:t([^>]*)>[\s\S]*?<\/a:t>/, `<a:t$1>${escapeXml(labor)}</a:t>`);
    });
    return tbl.replace(rows[1], newRow);
  });
}

function paraXml(pPr, rPr, text) {
  const lines = String(text ?? "").split(/\r?\n/);
  const runs = lines.map((ln, i) => {
    const spc = /^\s|\s$/.test(ln) ? ' xml:space="preserve"' : "";
    const r = `<a:r>${rPr}<a:t${spc}>${escapeXml(ln)}</a:t></a:r>`;
    return i ? `<a:br/>${r}` : r;
  }).join("");
  return `<a:p>${pPr}${runs || `<a:endParaRPr/>`}</a:p>`;
}

function fillCajaDescripcion(xml, descripcion, plan) {
  if (!descripcion && !plan) return xml;
  return xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (sp) => {
    if (!/Descripci[oó]n/i.test(sp) || !sp.includes("<p:txBody>")) return sp;
    const bodyMatch = sp.match(/<p:txBody>([\s\S]*)<\/p:txBody>/);
    if (!bodyMatch) return sp;
    const body = bodyMatch[1];
    const bodyPr = (body.match(/<a:bodyPr[\s\S]*?\/>/) || body.match(/<a:bodyPr[\s\S]*?<\/a:bodyPr>/) || ["<a:bodyPr/>"])[0];
    const lst = body.includes("<a:lstStyle/>") ? "<a:lstStyle/>" : "<a:lstStyle/>";
    const paras = body.match(/<a:p>[\s\S]*?<\/a:p>/g) || [];
    const pDesc = paras.find((p) => /Descripci/i.test(p)) || paras[0] || "";
    const pPlan = paras.find((p) => /Plan de acci|Medidas correctivas/i.test(p)) || "";
    const planLabel = /Medidas correctivas/i.test(pPlan) ? "Medidas correctivas:" : "Plan de acción:";
    const pPrOf = (p) => {
      const m = p && p.match(/<a:pPr[\s\S]*?<\/a:pPr>/);
      return m ? m[0] : "<a:pPr/>";
    };
    const rPrOf = (p, bold) => {
      const runs = [...String(p || "").matchAll(/<a:rPr[\s\S]*?<\/a:rPr>|<a:rPr[^/]*\/>/g)].map((x) => x[0]);
      if (bold) {
        const b = runs.find((r) => /b="1"/.test(r));
        if (b) return b;
      }
      return runs.find((r) => !/b="1"/.test(r)) || runs[0] || '<a:rPr lang="es-PE" sz="1100" dirty="0"/>';
    };
    const rPrLabel = rPrOf(pDesc, true);
    const rPrPlan = rPrOf(pPlan || pDesc, true);
    const rPrBody = rPrOf(paras.find((p) => !/Descripci|Plan de acci|Medidas/i.test(p)) || paras[1], false);
    const newBody =
      `${bodyPr}${lst}` +
      paraXml(pPrOf(pDesc), rPrLabel, "Descripción:") +
      paraXml(pPrOf(paras[1] || pDesc), rPrBody, descripcion || "") +
      paraXml(pPrOf(pPlan || pDesc), rPrPlan, planLabel) +
      paraXml(pPrOf(paras[3] || paras[1] || pDesc), rPrBody, plan || "");
    return sp.replace(/<p:txBody>[\s\S]*<\/p:txBody>/, `<p:txBody>${newBody}</p:txBody>`);
  });
}

async function llenarTextosReportes(zip) {
  guardarCamposReportesDesdeDom();
  const presXml = await zip.file("ppt/presentation.xml").async("string");
  const relsXml = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
  const orden = slidesOrden(presXml, parseRels(relsXml));
  for (let i = 1; i < orden.length; i += 1) {
    const n = i + 1;
    const data = fotosReportes[n] || {};
    const path = "ppt/" + orden[i].target.replace(/^ppt\//, "");
    let xml = await zip.file(path).async("string");
    if (data.labor) xml = fillLaborTable(xml, data.labor);
    if (n >= 3) xml = fillCajaDescripcion(xml, data.descripcion || "", data.plan || "");
    zip.file(path, xml);
  }
}

async function ponerFotosReportes(zip) {
  const presXml = await zip.file("ppt/presentation.xml").async("string");
  const relsXml = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
  const orden = slidesOrden(presXml, parseRels(relsXml));
  let ct = await zip.file("[Content_Types].xml").async("string");
  if (!ct.includes('Extension="jpg"')) {
    ct = ct.replace(
      "</Types>",
      `<Default Extension="jpg" ContentType="image/jpeg"/></Types>`
    );
    zip.file("[Content_Types].xml", ct);
  }

  for (let i = 2; i < orden.length; i += 1) {
    const n = i + 1;
    const par = fotosReportes[n];
    if (!par || (!par.antes && !par.despues)) continue;

    const path = "ppt/" + orden[i].target.replace(/^ppt\//, "");
    let xml = await zip.file(path).async("string");
    const relsPath = path.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
    let slideRels = await zip.file(relsPath).async("string");
    xml = quitarFotosContenido(xml);

    const areaX = 2680000;
    const areaW = 9144000 - areaX - 160000;
    const gap = 140000;
    const colW = Math.floor((areaW - gap) / 2);
    const labelH = 260000;
    const topLabel = 1520000;
    const topPic = topLabel + labelH + 30000;
    const picH = 3050000;
    let id = nextShapeId(xml);
    let extras = "";

    const cols = [
      { kind: "antes", label: "ANTES", hex: "C62828", x: areaX },
      { kind: "despues", label: "DESPUÉS", hex: "2E7D32", x: areaX + colW + gap },
    ];
    for (const col of cols) {
      extras += captionXml(id, col.x, topLabel, colW, labelH, col.label, col.hex);
      id += 1;
      const foto = par[col.kind];
      if (!foto) continue;
      const rid = nextRid(slideRels);
      const jpeg = await fileToJpegBytes(foto.blob);
      const mediaName = `rep${n}${col.kind}.jpg`;
      zip.file(`ppt/media/${mediaName}`, jpeg.bytes);
      slideRels = slideRels.replace(
        "</Relationships>",
        `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaName}"/></Relationships>`
      );
      const fit = encajar(colW, picH, jpeg.w, jpeg.h);
      extras += picXml(id, rid, col.x + fit.x, topPic + fit.y, fit.w, fit.h, col.hex);
      id += 1;
    }
    xml = xml.replace("</p:spTree>", extras + "</p:spTree>");
    zip.file(path, xml);
    zip.file(relsPath, slideRels);
  }
}

function encajar(cajaW, cajaH, imgW, imgH) {
  if (!imgW || !imgH) return { x: 0, y: 0, w: cajaW, h: cajaH };
  const escala = Math.min(cajaW / imgW, cajaH / imgH);
  const w = Math.max(1, Math.round(imgW * escala));
  const h = Math.max(1, Math.round(imgH * escala));
  return {
    x: Math.floor((cajaW - w) / 2),
    y: Math.floor((cajaH - h) / 2),
    w,
    h,
  };
}

async function fileToJpegBytes(fileOrBlob) {
  const blob = fileOrBlob instanceof Blob ? fileOrBlob : new Blob([fileOrBlob]);
  const bmp = await createImageBitmap(blob);
  const scale = Math.min(1, 1600 / Math.max(bmp.width, 1));
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(bmp.width * scale));
  canvas.height = Math.max(1, Math.round(bmp.height * scale));
  canvas.getContext("2d").drawImage(bmp, 0, 0, canvas.width, canvas.height);
  const out = await new Promise((resolve) => canvas.toBlob(resolve, "image/jpeg", 0.85));
  return {
    bytes: new Uint8Array(await out.arrayBuffer()),
    w: canvas.width,
    h: canvas.height,
  };
}

async function ponerFotosIds(zip, fotos) {
  if (!fotos.length) return;
  const presXml = await zip.file("ppt/presentation.xml").async("string");
  const relsXml = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
  const orden = slidesOrden(presXml, parseRels(relsXml));
  let idsPath = null;
  for (const s of orden) {
    const path = "ppt/" + s.target.replace(/^ppt\//, "");
    const xml = await zip.file(path).async("string");
    if (textosSlide(xml).some((t) => t.trim().toUpperCase() === "IDS")) {
      idsPath = path;
      break;
    }
  }
  if (!idsPath) idsPath = "ppt/slides/slide2.xml";
  let xml = await zip.file(idsPath).async("string");
  const relsPath = idsPath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
  let slideRels = await zip.file(relsPath).async("string");

  xml = quitarFotosContenido(xml);
  if (!xml.includes("<!--IDS_FOTOS-->")) {
    xml = xml.replace("</p:spTree>", "<!--IDS_FOTOS--></p:spTree>");
  }

  const n = Math.min(3, fotos.length);
  const margen = 152400;
  const top = 1686651;
  const gap = 100000;
  const cajaW = Math.floor((9144000 - margen * 2 - gap * (n - 1)) / n);
  const cajaH = 5143500 - top - 150000;
  let pics = "";
  let nextId = 400;
  for (let i = 0; i < n; i += 1) {
    const rid = nextRid(slideRels);
    const jpeg = await fileToJpegBytes(fotos[i]);
    const media = `ppt/media/ids${i + 1}.jpg`;
    zip.file(media, jpeg.bytes);
    slideRels = slideRels.replace(
      "</Relationships>",
      `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/ids${i + 1}.jpg"/></Relationships>`
    );
    const x = margen + i * (cajaW + gap);
    const fit = encajar(cajaW, cajaH, jpeg.w, jpeg.h);
    pics += picXml(nextId + i, rid, x + fit.x, top + fit.y, fit.w, fit.h);
  }
  xml = xml.replace("<!--IDS_FOTOS-->", pics);
  zip.file(idsPath, xml);
  zip.file(relsPath, slideRels);
}

async function generarPptx() {
  if (!plantillaBuf) throw new Error("La plantilla aún no cargó");
  if (typeof JSZip === "undefined") throw new Error("No cargó JSZip. Revisa internet / CDN.");
  const zip = await JSZip.loadAsync(plantillaBuf);
  const fecha = fechaDeInput();
  const turno = turnoSeleccionado();
  await ajustarSlides(zip, Number($("n-slides-input").value));

  const presXml = await zip.file("ppt/presentation.xml").async("string");
  const relsXml = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
  const orden = slidesOrden(presXml, parseRels(relsXml));
  for (const s of orden) {
    const path = "ppt/" + s.target.replace(/^ppt\//, "");
    let xml = await zip.file(path).async("string");
    const esPortada = /TURNO\s+(D[IÍ]A|NOCHE)/i.test(xml);
    zip.file(path, aplicarTexto(xml, fecha, turno, esPortada));
  }

  await llenarTextosReportes(zip);
  const fotos = fotosIds.filter(Boolean).map((f) => f.blob);
  await ponerFotosIds(zip, fotos);
  await ponerFotosReportes(zip);

  return zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
}

async function cargarPlantilla() {
  const res = await fetch(PLANTILLA_URL);
  if (!res.ok) throw new Error("No encuentro plantilla.pptx junto a la app");
  plantillaBuf = await res.arrayBuffer();
  const info = await inspeccionarBuf(plantillaBuf);
  pintar(info);
}

function hoyISO() {
  const d = new Date();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${d.getFullYear()}-${m}-${day}`;
}

function pintarSlot(i) {
  const slot = document.querySelector(`#slots .slot[data-i="${i}"]`);
  if (!slot) return;
  const foto = fotosIds[i];
  if (!foto) {
    slot.classList.remove("lleno");
    slot.style.backgroundImage = "";
    return;
  }
  slot.classList.add("lleno");
  slot.style.backgroundImage = `url("${foto.preview}")`;
}

async function asignarFile(i, file) {
  fotosIds[i] = { blob: file, preview: URL.createObjectURL(file), nombre: file.name };
  pintarSlot(i);
}

function bind(id, ev, fn) {
  const el = $(id);
  if (!el) return;
  el.addEventListener(ev, fn);
}

function iniciarUI() {
  try {
    pintarSlotsReportes();
  } catch (err) {
    console.error(err);
  }
  const fecha = $("fecha");
  if (fecha && !fecha.value) fecha.value = hoyISO();
  try {
    pintarNombre();
  } catch (_e) {
    /* ignore */
  }

  bind("slides-menos", "click", () => setNSlides(Number($("n-slides-input").value) - 1));
  bind("slides-mas", "click", () => setNSlides(Number($("n-slides-input").value) + 1));
  bind("n-slides-input", "change", () => setNSlides($("n-slides-input").value));
  bind("fecha", "change", pintarNombre);
  document.querySelectorAll('input[name="turno"]').forEach((el) => el.addEventListener("change", pintarNombre));

  bind("btn-hoy", "click", () => {
    $("fecha").value = hoyISO();
    pintarNombre();
  });

  bind("btn-default", "click", async () => {
    try {
      await cargarPlantilla();
      mostrarMsg("Plantilla por defecto lista.", true);
    } catch (e) {
      mostrarMsg(e.message, false);
    }
  });

  bind("file", "change", async (ev) => {
    const file = ev.target.files[0];
    if (!file) return;
    plantillaBuf = await file.arrayBuffer();
    try {
      pintar(await inspeccionarBuf(plantillaBuf));
      $("ruta-archivo").textContent = file.name;
      mostrarMsg("Archivo leído.", true);
    } catch (e) {
      mostrarMsg(e.message, false);
    }
  });

  document.querySelectorAll("#slots .slot").forEach((slot) => {
    const input = slot.querySelector("input[type=file]");
    if (!input) return;
    input.addEventListener("change", async () => {
      const file = input.files[0];
      if (file) await asignarFile(Number(slot.dataset.i), file);
    });
  });

  bind("fotos-multi", "change", async (ev) => {
    const files = [...ev.target.files].slice(0, 3);
    for (let i = 0; i < files.length; i += 1) await asignarFile(i, files[i]);
    mostrarMsg(`${files.length} foto(s) del IDS listas.`, true);
  });

  bind("form", "submit", async (ev) => {
    ev.preventDefault();
    const btn = ev.submitter;
    if (btn) btn.disabled = true;
    try {
      const blob = await generarPptx();
      const nombre = nombreDescarga();
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = nombre;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 2000);
      mostrarMsg(`Descargado: ${nombre}`, true);
    } catch (e) {
      mostrarMsg(e.message, false);
    } finally {
      if (btn) btn.disabled = false;
    }
  });

  cargarPlantilla()
    .then(() => mostrarMsg("Plantilla cargada. Completa IDS y cada reporte desde la diapo 3.", true))
    .catch((e) => mostrarMsg(e.message, false));
  intentarWhatsappLocal();
}

async function intentarWhatsappLocal() {
  try {
    const res = await fetch("/api/whatsapp");
    if (!res.ok) return;
    const data = await res.json();
    whatsapp = data.fotos || [];
    if (!whatsapp.length) return;
    $("wa-wrap").hidden = false;
    const box = $("gallery");
    box.innerHTML = "";
    for (const item of whatsapp) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "gal";
      btn.innerHTML = `<img alt="" src="/api/foto?ruta=${encodeURIComponent(item.ruta)}" /><span>${item.fecha}</span>`;
      btn.addEventListener("click", async () => {
        const img = await fetch("/api/foto?ruta=" + encodeURIComponent(item.ruta) + "&thumb=0");
        const blob = await img.blob();
        const i = fotosIds.findIndex((x) => !x);
        await asignarFile(i === -1 ? 0 : i, new File([blob], item.nombre, { type: blob.type }));
        mostrarMsg(`Foto ${fotosIds.filter(Boolean).length}/3 del IDS lista.`, true);
      });
      box.appendChild(btn);
    }
    $("btn-3-recientes").addEventListener("click", async () => {
      const tres = whatsapp.slice(0, 3);
      if (tres.length < 3) {
        mostrarMsg("No hay 3 fotos de WhatsApp en Descargas.", false);
        return;
      }
      for (let i = 0; i < 3; i += 1) {
        const img = await fetch("/api/foto?ruta=" + encodeURIComponent(tres[i].ruta) + "&thumb=0");
        const blob = await img.blob();
        await asignarFile(i, new File([blob], tres[i].nombre, { type: blob.type }));
      }
      mostrarMsg("Listas las 3 fotos más recientes de WhatsApp.", true);
    });
  } catch (_e) {
    /* online / GitHub: no hay API local */
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", iniciarUI);
} else {
  iniciarUI();
}
