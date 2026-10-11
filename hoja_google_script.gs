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
    } else if (["convocatoria", "partido", "borrarConvocatoria", "comprobarPin", "cambiarPin", "gardarSeguimento", "lerSeguimento"].includes(b.action)) {
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
      }
      const r = leerTodo(); r.ok = true; r.seguimento = lerSeguimento();
      return salida(r);
    } else {
      return salida({ ok: false, error: "acción desconocida" });
    }
    const r = leerTodo(); r.ok = true;
    return salida(r);
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
  return { avail, convocatorias, partidos };
}

/* Notas do adestrador: só se len co PIN, nunca na lectura pública */
function lerSeguimento() {
  const s = {};
  filas("seguimento", ["fecha", "datos", "actualizado"]).forEach(r => {
    try { s[txt(r[0])] = JSON.parse(r[1] || "{}"); } catch (e) {}
  });
  return s;
}

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
