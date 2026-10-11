/**
 * Tomiño F.C. Veteranos - almacén compartido de la web
 * Guarda quién va a cada partido, la convocatoria publicada y la hora del partido.
 *
 * 1. O PIN de abaixo é só o inicial. Despois pódese cambiar desde a web (Axustes),
 *    e o novo gárdase nas propiedades do script.
 * 2. Implementar > Nueva implementación > Aplicación web
 *    Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario
 */
const PIN_ENTRENADOR = "1234"; // cambia isto polo teu PIN antes de implementar

function pinActual() {
  return PropertiesService.getScriptProperties().getProperty("PIN") || PIN_ENTRENADOR;
}

function doGet() {
  return salida(leerTodo());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === "avail") {
      if (!b.fecha || !b.jugador) return salida({ ok: false, error: "faltan datos" });
      guardar("disponibilidad", ["fecha", "jugador", "estado", "actualizado"], [b.fecha, b.jugador], b.estado ? [b.fecha, b.jugador, b.estado, new Date()] : null);
    } else if (["convocatoria", "partido", "borrarConvocatoria", "comprobarPin", "cambiarPin", "gardarSeguimento", "lerSeguimento", "crearVotacion", "pecharVotacion"].includes(b.action)) {
      if (String(b.pin) !== pinActual()) return salida({ ok: false, error: "pin" });
      if (b.action === "cambiarPin") {
        const novo = String(b.nuevo || "").trim();
        if (novo.length < 4) return salida({ ok: false, error: "curto" });
        PropertiesService.getScriptProperties().setProperty("PIN", novo);
      }
      if (b.action === "convocatoria") {
        guardar("convocatorias", ["fecha", "convocados", "descansan", "actualizado"], [b.fecha], [b.fecha, JSON.stringify(b.convocados || []), JSON.stringify(b.descansan || []), new Date()]);
      } else if (b.action === "borrarConvocatoria") {
        guardar("convocatorias", ["fecha", "convocados", "descansan", "actualizado"], [b.fecha], null);
      } else if (b.action === "partido") {
        guardar("partidos", ["fecha", "hora", "nota", "actualizado"], [b.fecha], [b.fecha, b.hora || "", b.nota || "", new Date()]);
      } else if (b.action === "gardarSeguimento") {
        guardar("seguimento", ["fecha", "datos", "actualizado"], [b.fecha], [b.fecha, JSON.stringify(b.datos || {}), new Date()]);
      } else if (b.action === "crearVotacion") {
        crearVotacion(b);
      } else if (b.action === "pecharVotacion") {
        const v = cfgVotacion(); if (v && !v.pechada) pecharVotacion(v);
      }
      const r = leerTodo(); r.ok = true; r.seguimento = lerSeguimento();
      return salida(r);
    } else {
      return salida({ ok: false, error: "acción desconocida" });
    }
    const r = leerTodo(); r.ok = true;
    return salida(r);
  } catch (err) {
    return salida({ ok: false, error: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

function hoja(nombre, cabecera) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, cabecera.length).setValues([cabecera]).setFontWeight("bold");
    sh.getRange("A:C").setNumberFormat("@");
    sh.setFrozenRows(1);
  }
  return sh;
}

/* A folla pode converter "2026-10-17" ou "17:00" en datas: volvemos sempre a texto */
function txt(v) {
  if (v instanceof Date) {
    return v.getFullYear() < 1901
      ? Utilities.formatDate(v, Session.getScriptTimeZone(), "HH:mm")
      : Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(v);
}

/* Busca as filas cuxas primeiras columnas coinciden con "clave": actualiza a primeira e borra as repetidas; se fila é null, bórraas todas */
function guardar(nombre, cabecera, clave, fila) {
  const sh = hoja(nombre, cabecera);
  const datos = sh.getDataRange().getValues();
  let feito = false;
  for (let i = datos.length - 1; i >= 1; i--) {
    if (clave.every((k, j) => txt(datos[i][j]) === String(k))) {
      if (fila && !feito) { sh.getRange(i + 1, 1, 1, fila.length).setValues([fila]); feito = true; }
      else sh.deleteRow(i + 1);
    }
  }
  if (fila && !feito) sh.appendRow(fila);
}

function filas(nombre, cabecera) {
  const v = hoja(nombre, cabecera).getDataRange().getValues();
  return v.slice(1).filter(r => r[0] !== "");
}

function leerTodo() {
  const avail = {};
  filas("disponibilidad", ["fecha", "jugador", "estado", "actualizado"]).forEach(r => {
    const f = txt(r[0]); (avail[f] = avail[f] || {})[String(r[1])] = String(r[2]);
  });
  const convocatorias = filas("convocatorias", ["fecha", "convocados", "descansan", "actualizado"]).map(r => ({
    fecha: txt(r[0]), convocados: JSON.parse(r[1] || "[]"), descansan: JSON.parse(r[2] || "[]")
  }));
  const partidos = {};
  filas("partidos", ["fecha", "hora", "nota", "actualizado"]).forEach(r => {
    partidos[txt(r[0])] = { hora: txt(r[1]), nota: String(r[2]) };
  });
  return { avail, convocatorias, partidos, votacion: estadoVotacion() };
}

/* Notas do adestrador: só se len co PIN, nunca na lectura pública */
function lerSeguimento() {
  const s = {};
  filas("seguimento", ["fecha", "datos", "actualizado"]).forEach(r => {
    try { s[txt(r[0])] = JSON.parse(r[1] || "{}"); } catch (e) {}
  });
  return s;
}

/* ---------- Votación en Google Forms ----------
   Unha resposta por conta de Google e sen recoller o correo: ninguén vota dúas veces
   e o voto é anónimo, mesmo para quen é dono do formulario. */
function cfgVotacion() {
  return JSON.parse(PropertiesService.getScriptProperties().getProperty("VOTACION") || "null");
}
function gardarCfgVotacion(v) {
  PropertiesService.getScriptProperties().setProperty("VOTACION", JSON.stringify(v));
}
function crearVotacion(b) {
  const vella = cfgVotacion(); if (vella && !vella.pechada) pecharVotacion(vella);
  const opcions = (b.opcions || []).map(String).filter(o => o.trim()).slice(0, 6);
  const ata = Number(b.ata);
  if (!b.pregunta || opcions.length < 2 || !ata) throw new Error("datos da votación incompletos");
  const ataTxt = Utilities.formatDate(new Date(ata), "Europe/Madrid", "dd/MM/yyyy 'ás' HH:mm");
  const form = FormApp.create("Tomiño F.C. Veteranos · Votación");
  form.setDescription("Votación anónima do equipo. Só se pode votar unha vez por persoa. Pecha o " + ataTxt + ".");
  form.setCollectEmail(false);
  form.setLimitOneResponsePerUser(true);
  form.setAllowResponseEdits(true);
  form.setShowLinkToRespondAgain(false);
  form.setPublishingSummary(false);
  form.setProgressBar(false);
  form.setConfirmationMessage("Grazas! O teu voto quedou gardado de forma anónima. O resultado verase na páxina do equipo cando peche a votación.");
  form.setCustomClosedFormMessage("Esta votación xa está pechada. Mira o resultado na páxina do equipo.");
  form.addMultipleChoiceItem().setTitle(String(b.pregunta)).setChoiceValues(opcions).setRequired(true);
  try { if (form.setPublished) form.setPublished(true); } catch (e) {}
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "pecharVotacionAuto").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("pecharVotacionAuto").timeBased().at(new Date(ata)).create();
  gardarCfgVotacion({ id: form.getId(), url: form.getPublishedUrl(), pregunta: String(b.pregunta), opcions, ata, creada: Date.now(), pechada: false });
  CacheService.getScriptCache().remove("votos_total");
}
function contarVotos(v) {
  const c = {}; v.opcions.forEach(o => c[o] = 0); let total = 0;
  try {
    FormApp.openById(v.id).getResponses().forEach(r => {
      const it = r.getItemResponses()[0]; if (!it) return;
      const val = String(it.getResponse()); if (c[val] != null) c[val]++; total++;
    });
  } catch (e) {}
  return { c, total };
}
function pecharVotacion(v) {
  try { FormApp.openById(v.id).setAcceptingResponses(false); } catch (e) {}
  const r = contarVotos(v);
  v.pechada = true; v.pechadaEn = Date.now(); v.resultados = r.c; v.total = r.total;
  gardarCfgVotacion(v);
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "pecharVotacionAuto").forEach(t => ScriptApp.deleteTrigger(t));
  return v;
}
function pecharVotacionAuto() {
  const v = cfgVotacion(); if (v && !v.pechada) pecharVotacion(v);
}
/* O público ve a pregunta, o prazo e cantos votaron; o resultado só cando está pechada */
function estadoVotacion() {
  let v = cfgVotacion(); if (!v) return null;
  if (!v.pechada && v.ata && Date.now() > v.ata) v = pecharVotacion(v);
  const out = { pregunta: v.pregunta, opcions: v.opcions, url: v.url, ata: v.ata, creada: v.creada, pechada: !!v.pechada };
  if (v.pechada) { out.resultados = v.resultados; out.total = v.total; }
  else {
    const cache = CacheService.getScriptCache(); let t = cache.get("votos_total");
    if (t === null) { try { t = String(FormApp.openById(v.id).getResponses().length); } catch (e) { t = "0"; } cache.put("votos_total", t, 60); }
    out.total = Number(t);
  }
  return out;
}
/* Executar unha vez desde o editor para dar permiso a crear formularios e avisos programados */
function autorizar() {
  FormApp.getActiveForm();
  ScriptApp.getProjectTriggers();
  CacheService.getScriptCache();
  Logger.log("Permisos concedidos");
}

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
