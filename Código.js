// ==================================================================
// CONFIGURACIÓN CENTRALIZADA
// ==================================================================
const CONFIG = {
  HOJA_MADRE_ID: "1oCWP8QjUE7WXuusAdH29NYv0-0ZCkGPW__VRUw1qaHM",
  HOJA_MADRE_NAME: "Madre",
  HOJA_ORGANISMOS: "Organismos",
  HOJA_ESTUDIOS: "Estudios",
  HOJA_FECHAS: "Fecha_derivación",
  HOJA_REPORTE: "Reporte",
  HOJA_BITACORA: "Bitacora",
  HABILITAR_BITACORA: false,
  HOJA_VINCULACIONES: "Vinculaciones",
  HOJA_DESVIOS: "Desvios",
  HOJA_RECLAMOS: "Reclamos",
  HOJA_COMUNICACIONES: "Comunicaciones",
  HOJA_EMPLEADOS: "Empleados",
  HOJA_AVANCE: "Reportes Avance",
  HOJA_KPIS: "KPIs",
  HOJA_PRESTADOR_RECO: "prestador_reco",
  HOJA_ESTUDIOS_ADICIONALES: "Estudios_Adicionales",
  HOJA_LINKS: "Links",
  EXTERNA_ID: "1aVbOK010lH4X4QCLAy3J1AoTIu7rE9qZmK07W2EX3ow",
  EXTERNA_NAME: "Gestion_Links",
  EXTERNA_GID: 1054481966,
  CARPETA_PDFS_ID: "12u1_0G8x61yhJjvGumtQmW1lXQT8FpXV",

  // ==============================================================
  // COLUMNAS HOJA MADRE (índice base-0, col_hoja = índice + 1)
  // ==============================================================
  COL_CLAVE:                    0,  // col 1  — Clave
  COL_FECHA_DERIV:              1,  // col 2  — Fecha_derivación
  COL_CUIT_CONTRATO:            2,  // col 3  — Cuit_Contrato
  COL_NOMBRE_CONTRATO:          3,  // col 4  — Nombre_del_contrato
  COL_CUIT_ORGANISMO:           4,  // col 5  — CUIT_Organismo
  COL_NOMBRE_ORGANISMO:         5,  // col 6  — Nombre_Organismo
  COL_NOMBRE_ESTABLECIMIENTO:   6,  // col 7  — Nombre_del_Establecimiento
  COL_SRT:                      7,  // col 8  — Codigo_SRT
  COL_DIRECCION:                8,  // col 9  — Direccion_del_Establecimiento
  COL_LOCALIDAD:                9,  // col 10 — Localidad
  COL_PROVINCIA:               10,  // col 11 — Provincia
  COL_FECHA_RAR:               11,  // col 12 — Fecha_de_Presentacion_RAR
  COL_CUIL:                    12,  // col 13 — CUIL_Empleado
  COL_APELLIDO:                13,  // col 14 — Apellido_y_Nombre
  COL_PRESTADOR:               14,  // col 15 — Prestador
  COL_FECHA_INGRESO:           15,  // col 16 — Fecha_de_Ingreso
  COL_INICIO_EXPOSICION:       16,  // col 17 — Inicio_de_Exposicion
  COL_SECTOR:                  17,  // col 18 — Sector
  COL_PUESTO:                  18,  // col 19 — Puesto_de_Trabajo
  COL_ESTUDIO:                 19,  // col 20 — Estudio_a_Realizar
  COL_CODIGO_RIESGO:           20,  // col 21 — Código_Riesgo_Individual
  COL_ORIENTACIONES:           21,  // col 22 — Orientaciones_Médicas
  COL_RESULTADOS_RAR:          22,  // col 23 — Resultados (campo RAR, no el de auditoría)
  COL_ASISTENCIA:              23,  // col 24 — Asistencia (sync desde hoja externa)
  COL_FECHA_CONTROL:           24,  // col 25 — Fecha_de_control
  COL_COMUNICACION:            25,  // col 26 — Comunicación_de_Resultados
  COL_FECHA_AUDITORIA:         26,  // col 27 — Fecha_auditoria
  COL_FECHA_ESTUDIO:           27,  // col 28 — Fecha (fecha estudio)
  COL_RESULTADO:               28,  // col 29 — Resultado (auditoria)
  COL_OBSERVACION:             29,  // col 30 — Observación
  COL_IMPORTE:                 30,  // col 31 — Importe
  COL_FRECUENCIA:              31,  // col 32 — Frecuencia
  COL_ADJUNTO_1:               32,  // col 33 — Link1
  COL_ADJUNTO_2:               33,  // col 34 — Link2
  COL_FECHA_VIGENCIA:          34,  // col 35 — Fecha_vigencia
  COL_CODIGOS_SEMESTRALES:     35,  // col 36 — Códigos_Semestrales_Identificados
  COL_DESCRIPCION_ESTUDIO:     36,  // col 37 — Descripción_Estudio
  COL_FILA_MADRE:              37,  // col 38 — Fila_Madre (en hojas especiales)
  COL_GUARDADO:                37,  // col 38 (AL) — Guardado (SÍ / Fecha)
  COL_LINK_CARPETA:            38,  // col 39 (AM) — Link a carpeta en Drive
  COL_CLAVE_EMPLEADO:          39,  // col 40 (AN) — Clave_Empleado_SRT (CUIL & "_" & Codigo_SRT)

  // ==============================================================
  // COLUMNAS HOJA VINCULACIONES
  // ==============================================================
  VIN_COL_CONTRATO:       0, VIN_COL_ORGANISMO: 1, VIN_COL_SRT: 2, VIN_COL_ESTABLECIMIENTO: 3,

  // ==============================================================
  // COLUMNAS HOJA EXTERNA (Gestion_Links: A=0, B=1, C=2, D=3, E=4)
  // ==============================================================
  EXT_COL_CLAVE:          0,  // Col A — Clave de cruce
  EXT_COL_ASISTENCIA:     1,  // Col B — Asistencia Calculada → COL_ASISTENCIA en Madre
  EXT_COL_ADJUNTO_1:      2,  // Col C — Link Adjunto 1       → COL_ADJUNTO_1 en Madre
  EXT_COL_ADJUNTO_2:      3,  // Col D — Link Adjunto 2       → COL_ADJUNTO_2 en Madre
  EXT_COL_FECHA_ESTUDIO:  4,  // Col E — Fecha Examen         → COL_FECHA_ESTUDIO en Madre

  // ==============================================================
  // PARÁMETROS DE COMPORTAMIENTO
  // ==============================================================
  SINCRONIZACION_INTERVALO_MIN: 30,
  CACHE_DURACION_MIN: 10,
  TIMEOUT_LOCK: 10000,
  MAX_FILAS_DISPLAY: 150,
  ESTUDIO_EXCLUIR: "No posee estudios asociados al riesgo",
  // [2026-09-06] Incorporación de Desistido y Ausente/Sin realizar manteniendo compatibilidad
  RESULTADOS_PERMITIDOS: [
    "Normal Laboral", "Pendiente de Confirmación", "Desvío",
    "Ausente/Sin realizar", "Ausente", "Enfermedad Profesional", "Práctica defectuosa", "Desistido"
  ]
};


// ==================================================================
// UTILIDADES COMPARTIDAS
// ==================================================================
function formatearFecha(fecha, tz) {
  if (!fecha && fecha !== 0) return "";
  const timeZone = tz || Session.getScriptTimeZone() || "America/Argentina/Buenos_Aires";

  // Si ya es un objeto Date
  if (fecha instanceof Date) {
    if (isNaN(fecha.getTime())) return "";
    return Utilities.formatDate(fecha, timeZone, "dd/MM/yyyy");
  }

  // Si es un número (serial de Sheets/Excel o timestamp ms)
  if (typeof fecha === 'number') {
    if (fecha > 30000 && fecha < 60000) {
      const dt = new Date(Math.round((fecha - 25569) * 86400 * 1000));
      return Utilities.formatDate(dt, timeZone, "dd/MM/yyyy");
    }
    if (fecha > 1000000000000) {
      const dt = new Date(fecha);
      return Utilities.formatDate(dt, timeZone, "dd/MM/yyyy");
    }
  }

  let s = String(fecha).trim();
  if (!s || s === '-' || s.toLowerCase() === 'invalid date' || s.toLowerCase() === 'null' || s.toLowerCase() === 'undefined') {
    return "";
  }

  // 1) Si viene como DD/MM/YYYY o D/M/YYYY
  const mArSlash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (mArSlash) {
    return `${mArSlash[1].padStart(2, '0')}/${mArSlash[2].padStart(2, '0')}/${mArSlash[3]}`;
  }

  // 2) Si viene como DD-MM-YYYY o D-M-YYYY
  const mArDash = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/);
  if (mArDash) {
    return `${mArDash[1].padStart(2, '0')}/${mArDash[2].padStart(2, '0')}/${mArDash[3]}`;
  }

  // 3) Si viene como YYYY-MM-DD o YYYY/MM/DD (incluyendo ISO strings con 'T' o hora)
  const mIso = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (mIso) {
    return `${mIso[3].padStart(2, '0')}/${mIso[2].padStart(2, '0')}/${mIso[1]}`;
  }

  // 4) Si viene con año de 2 dígitos: DD/MM/YY o DD-MM-YY
  const m2Dig = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2})$/);
  if (m2Dig) {
    const anio = parseInt(m2Dig[3], 10) > 50 ? `19${m2Dig[3]}` : `20${m2Dig[3]}`;
    return `${m2Dig[1].padStart(2, '0')}/${m2Dig[2].padStart(2, '0')}/${anio}`;
  }

  // 5) Si es número serial en string (ej. "45424")
  if (/^\d{5}(\.\d+)?$/.test(s)) {
    const num = parseFloat(s);
    if (num > 30000 && num < 60000) {
      const dt = new Date(Math.round((num - 25569) * 86400 * 1000));
      return Utilities.formatDate(dt, timeZone, "dd/MM/yyyy");
    }
  }

  // 6) Fallback para strings de Date generales de JavaScript (ej. "Sun May 12 2026 00:00:00 GMT...")
  const dtGen = new Date(s);
  if (!isNaN(dtGen.getTime()) && dtGen.getFullYear() > 1900 && dtGen.getFullYear() < 2100) {
    return Utilities.formatDate(dtGen, timeZone, "dd/MM/yyyy");
  }

  return s;
}

function obtenerTimestampBitacora() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
}

function registrarEnBitacora(accion, detalles) {
  // Bitácora desactivada permanentemente
  return;
}

// ==================================================================
// ROUTING
// ==================================================================
function doGet(e) {
  if (!e) e = { parameter: {} };
  if (e.parameter.api === 'sync_hojas') {
    const res = sincronizarTodasLasHojasDesdeMadre();
    return ContentService.createTextOutput(JSON.stringify(res, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'sync_empleados') {
    const res = sincronizarHojaEmpleados();
    return ContentService.createTextOutput(JSON.stringify(res, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'limpiar_ausentes') {
    const res = limpiarInconsistenciasAusentesEnMadre();
    return ContentService.createTextOutput(JSON.stringify(res, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'inspect_madre') {
    const ss = SpreadsheetApp.openById(CONFIG.HOJA_MADRE_ID);
    const sheet = ss.getSheetByName(CONFIG.HOJA_MADRE_NAME);
    const ultFila = sheet.getLastRow();
    const cuils = sheet.getRange(2, CONFIG.COL_CUIL + 1, ultFila - 1, 1).getValues();
    const asistencias = sheet.getRange(2, CONFIG.COL_ASISTENCIA + 1, ultFila - 1, 1).getValues();
    const resultados = sheet.getRange(2, CONFIG.COL_RESULTADO + 1, ultFila - 1, 1).getValues();
    const links = sheet.getRange(2, CONFIG.COL_ADJUNTO_1 + 1, ultFila - 1, 1).getValues();
    const colGuardado = (CONFIG.COL_GUARDADO !== undefined ? CONFIG.COL_GUARDADO : 37) + 1;
    const guardados = sheet.getRange(2, colGuardado, ultFila - 1, 1).getValues();
    const colClaveAN = (CONFIG.COL_CLAVE_EMPLEADO !== undefined ? CONFIG.COL_CLAVE_EMPLEADO : 39) + 1;
    const claves = sheet.getRange(2, colClaveAN, ultFila - 1, 1).getValues();

    let total = ultFila - 1;
    let conCuil = 0;
    let conResultado = 0;
    let conLink = 0;
    let conAsistencia = 0;
    let conGuardado = 0;
    const cuilsConAuditoria = new Set();
    const cuilsConLink = new Set();
    const cuilsConAsistencia = new Set();
    const cuilsRelevantes = new Set();
    const cuilsConGuardado = new Set();
    const clavesConGuardado = new Set();

    for (let i = 0; i < total; i++) {
      const c = String(cuils[i][0] || '').trim();
      const as = String(asistencias[i][0] || '').trim();
      const res = String(resultados[i][0] || '').trim();
      const lk = String(links[i][0] || '').trim();
      const gd = String(guardados[i][0] || '').trim();
      const k = String(claves[i][0] || '').trim() || c;

      if (c) {
        conCuil++;
        if (res) { conResultado++; cuilsConAuditoria.add(c); cuilsRelevantes.add(c); }
        if (lk) { conLink++; cuilsConLink.add(c); cuilsRelevantes.add(c); }
        if (as) { conAsistencia++; cuilsConAsistencia.add(c); }
        const esGd = gd.toUpperCase() === 'SÍ' || gd.toUpperCase() === 'SI' || gd.toUpperCase() === 'S' || gd.toUpperCase().startsWith('SÍ') || gd.toUpperCase().startsWith('SI') || gd.toUpperCase().includes('GUARDAD');
        if (esGd) {
          conGuardado++;
          cuilsConGuardado.add(c);
          if (k) clavesConGuardado.add(k);
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      totalFilas: total,
      conCuil,
      conResultado,
      conLink,
      conAsistencia,
      conGuardado,
      empleadosConAuditoria: cuilsConAuditoria.size,
      empleadosConLink: cuilsConLink.size,
      empleadosConAsistencia: cuilsConAsistencia.size,
      empleadosRelevantes: cuilsRelevantes.size,
      empleadosConGuardadoCUIL: cuilsConGuardado.size,
      empleadosConGuardadoAN: clavesConGuardado.size
    }, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'inspect_resultados') {
    const ss = SpreadsheetApp.openById(CONFIG.HOJA_MADRE_ID);
    const sheet = ss.getSheetByName(CONFIG.HOJA_MADRE_NAME);
    const ultFila = sheet.getLastRow();
    const resultados = sheet.getRange(2, CONFIG.COL_RESULTADO + 1, ultFila - 1, 1).getValues();
    const freq = {};
    for (let i = 0; i < resultados.length; i++) {
      const v = String(resultados[i][0] || '').trim();
      if (v) freq[v] = (freq[v] || 0) + 1;
    }
    return ContentService.createTextOutput(JSON.stringify(freq, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'inspect_multi_estab') {
    const ss = SpreadsheetApp.openById(CONFIG.HOJA_MADRE_ID);
    const sheet = ss.getSheetByName(CONFIG.HOJA_MADRE_NAME);
    const ultFila = sheet.getLastRow();
    const cuils = sheet.getRange(2, CONFIG.COL_CUIL + 1, ultFila - 1, 1).getValues();
    const estabs = sheet.getRange(2, CONFIG.COL_NOMBRE_ESTABLECIMIENTO + 1, ultFila - 1, 1).getValues();
    const prests = sheet.getRange(2, CONFIG.COL_PRESTADOR + 1, ultFila - 1, 1).getValues();
    const resuls = sheet.getRange(2, CONFIG.COL_RESULTADO + 1, ultFila - 1, 1).getValues();

    const mapaCuilEstabs = new Map();
    for (let i = 0; i < cuils.length; i++) {
      const c = String(cuils[i][0] || '').trim();
      const est = String(estabs[i][0] || '').trim();
      const pr = String(prests[i][0] || '').trim();
      const res = String(resuls[i][0] || '').trim();
      if (!c) continue;
      if (!mapaCuilEstabs.has(c)) {
        mapaCuilEstabs.set(c, { cuil: c, auditado: false, ests: new Map() });
      }
      const item = mapaCuilEstabs.get(c);
      if (res) item.auditado = true;
      if (!item.ests.has(est)) item.ests.set(est, new Set());
      if (pr) item.ests.get(est).add(pr);
    }

    const multi = [];
    mapaCuilEstabs.forEach((v, c) => {
      if (v.ests.size > 1 && v.auditado) {
        const estList = [];
        v.ests.forEach((prestSet, estName) => {
          estList.push({ estab: estName, prestadores: Array.from(prestSet) });
        });
        multi.push({ cuil: c, establecimientos: estList });
      }
    });

    return ContentService.createTextOutput(JSON.stringify({
      totalMultiAuditados: multi.length,
      ejemplos: multi.slice(0, 5)
    }, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'get_desvios') {
    const res = obtenerRegistrosDesvios();
    return ContentService.createTextOutput(JSON.stringify({ count: res.length, sample: res.slice(0, 3) }, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'get_prestadores_reco') {
    const res = typeof obtenerPrestadoresReconfirmacion === 'function' ? obtenerPrestadoresReconfirmacion() : [];
    return ContentService.createTextOutput(JSON.stringify(res, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'get_estudios_nomenclatura') {
    const res = typeof obtenerNombresEstudios === 'function' ? obtenerNombresEstudios() : [];
    return ContentService.createTextOutput(JSON.stringify(res, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'get_reclamos') {
    const res = obtenerRegistrosReclamos();
    return ContentService.createTextOutput(JSON.stringify({ count: res.length, sample: res.slice(0, 3) }, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'get_avance') {
    const res = typeof obtenerDatosAvance === 'function' ? obtenerDatosAvance() : { error: 'No implementado' };
    return ContentService.createTextOutput(JSON.stringify(res, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'sync_avance') {
    const res = typeof sincronizarHojaAvance === 'function' ? sincronizarHojaAvance() : { error: 'No implementado' };
    return ContentService.createTextOutput(JSON.stringify(res, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'get_kpis') {
    const forzar = e.parameter.forzar === '1' || e.parameter.forzar === 'true';
    const res = typeof obtenerMetricasKPIs === 'function' ? obtenerMetricasKPIs(forzar) : { error: 'No implementado' };
    return ContentService.createTextOutput(JSON.stringify(res, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'sync_kpis') {
    const res = typeof sincronizarHojaKPIs === 'function' ? sincronizarHojaKPIs() : { error: 'No implementado' };
    return ContentService.createTextOutput(JSON.stringify(res, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'inspect_bloques') {
    const t0 = new Date().getTime();
    const sheetMadre = SpreadsheetApp.openById(CONFIG.HOJA_MADRE_ID).getSheetByName(CONFIG.HOJA_MADRE_NAME);
    const ultFila = sheetMadre.getLastRow();
    const t1 = new Date().getTime();
    const cuilVals = sheetMadre.getRange(2, CONFIG.COL_CUIL + 1, ultFila - 1, 1).getValues();
    const resVals = sheetMadre.getRange(2, CONFIG.COL_RESULTADO + 1, ultFila - 1, 1).getValues();
    const colG = (CONFIG.COL_GUARDADO !== undefined ? CONFIG.COL_GUARDADO : 37) + 1;
    const guarVals = sheetMadre.getRange(2, colG, ultFila - 1, 1).getValues();
    const t2 = new Date().getTime();

    const cuilsAuditados = new Set();
    for (let i = 0; i < ultFila - 1; i++) {
      const res = String(resVals[i][0] || '').trim();
      const gd = String(guarVals[i][0] || '').trim();
      if (res !== '' || gd.toUpperCase() === 'SÍ' || gd.toUpperCase() === 'SI') {
        const c = String(cuilVals[i][0] || '').trim();
        if (c) cuilsAuditados.add(c);
      }
    }

    const filasObjetivo = [];
    for (let i = 0; i < ultFila - 1; i++) {
      const c = String(cuilVals[i][0] || '').trim();
      if (c && cuilsAuditados.has(c)) {
        filasObjetivo.push(i + 2);
      }
    }

    const bloques = [];
    let bloque = null;
    for (let j = 0; j < filasObjetivo.length; j++) {
      const fila = filasObjetivo[j];
      if (!bloque) {
        bloque = { start: fila, count: 1 };
      } else if (fila === bloque.start + bloque.count) {
        bloque.count++;
      } else {
        bloques.push(bloque);
        bloque = { start: fila, count: 1 };
      }
    }
    if (bloque) bloques.push(bloque);
    const t3 = new Date().getTime();

    return ContentService.createTextOutput(JSON.stringify({
      ultFila,
      cuilsAuditados: cuilsAuditados.size,
      totalFilasObjetivo: filasObjetivo.length,
      minFila: filasObjetivo.length ? filasObjetivo[0] : 0,
      maxFila: filasObjetivo.length ? filasObjetivo[filasObjetivo.length - 1] : 0,
      totalBloques: bloques.length,
      bloquesSample: bloques.slice(0, 10),
      msOpen: t1 - t0,
      msScan3Cols: t2 - t1,
      msGroup: t3 - t2
    }, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'inspect_bitacora') {
    return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.api === 'test_comunicaciones') {
    try {
      const res = obtenerRegistrosComunicables();
      return ContentService.createTextOutput(JSON.stringify({
        status: "OK",
        count: res.length,
        first: res.slice(0, 1)
      }, null, 2)).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "ERROR",
        error: err.message,
        stack: err.stack
      }, null, 2)).setMimeType(ContentService.MimeType.JSON);
    }
  }
  if (e.parameter.api === 'limpiar_cache' || e.parameter.api === 'reset_cache') {
    try {
      const cache = CacheService.getScriptCache();
      cache.removeAll([
        'datosBase_v5', 'dashData_v4', 'empleadosData_v1', 'comunicaciones_v1'
      ]);
      for (let i = 0; i < 25; i++) {
        cache.remove(`datosBase_v5_${i}`);
        cache.remove(`dashData_v4_${i}`);
      }
      return ContentService.createTextOutput(JSON.stringify({
        status: "OK",
        msg: "Memoria y caché borrados con éxito"
      }, null, 2)).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "ERROR",
        error: err.message
      }, null, 2)).setMimeType(ContentService.MimeType.JSON);
    }
  }
  if (e.parameter.api === 'inspect') {
    const ss = SpreadsheetApp.openById(CONFIG.HOJA_MADRE_ID);
    const sheets = ss.getSheets().map(s => ({
      name: s.getName(),
      rows: s.getLastRow(),
      cols: s.getLastColumn(),
      firstRow: s.getLastRow() > 0 ? s.getRange(1, 1, 1, Math.min(Math.max(s.getLastColumn(), 1), 40)).getValues()[0] : []
    }));
    return ContentService.createTextOutput(JSON.stringify(sheets, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.v === 'reportes') {
    const html = HtmlService.createTemplateFromFile('Reportes');
    html.url = ScriptApp.getService().getUrl();
    return html.evaluate().setTitle('Reportes').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (e.parameter.v === 'cobertura') {
    const html = HtmlService.createTemplateFromFile('Cobertura');
    html.url = ScriptApp.getService().getUrl();
    return html.evaluate().setTitle('Cobertura').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (e.parameter.v === 'comunicacion' || e.parameter.v === 'comunicar') {
    const html = HtmlService.createTemplateFromFile('Comunicacion');
    html.url = ScriptApp.getService().getUrl();
    html.urlMadre = 'https://docs.google.com/spreadsheets/d/' + CONFIG.HOJA_MADRE_ID;
    return html.evaluate().setTitle('Comunicación de Resultados').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // [2026-09-06] Nuevas vistas: Desvíos Médicos y Reclamos a Prestadores
  if (e.parameter.v === 'desvios') {
    const html = HtmlService.createTemplateFromFile('Desvios');
    html.url = ScriptApp.getService().getUrl();
    html.urlMadre = 'https://docs.google.com/spreadsheets/d/' + CONFIG.HOJA_MADRE_ID;
    try {
      html.initialData = JSON.stringify(obtenerRegistrosDesvios()).replace(/<\/script/gi, '<\\/script');
    } catch(err) {
      html.initialData = '[]';
    }
    try {
      html.prestadoresRecoData = JSON.stringify(typeof obtenerPrestadoresReconfirmacion === 'function' ? obtenerPrestadoresReconfirmacion() : []).replace(/<\/script/gi, '<\\/script');
    } catch(err) {
      html.prestadoresRecoData = '[]';
    }
    try {
      html.estudiosAdicionalesData = JSON.stringify(typeof obtenerEstudiosAdicionalesABM === 'function' ? obtenerEstudiosAdicionalesABM() : { status: 'OK', items: [] }).replace(/<\/script/gi, '<\\/script');
    } catch(err) {
      html.estudiosAdicionalesData = '{"status":"OK","items":[]}';
    }
    try {
      html.mapeoEstudiosAdicionalesData = JSON.stringify(typeof obtenerMapeoEstudiosAdicionales === 'function' ? obtenerMapeoEstudiosAdicionales() : {}).replace(/<\/script/gi, '<\\/script');
    } catch(err) {
      html.mapeoEstudiosAdicionalesData = '{}';
    }
    try {
      html.estudiosNomenclaturaData = JSON.stringify(typeof obtenerNombresEstudios === 'function' ? obtenerNombresEstudios() : []).replace(/<\/script/gi, '<\\/script');
    } catch(err) {
      html.estudiosNomenclaturaData = '[]';
    }
    return html.evaluate().setTitle('Desvíos y Confirmaciones').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (e.parameter.v === 'reclamos') {
    const html = HtmlService.createTemplateFromFile('Reclamos');
    html.url = ScriptApp.getService().getUrl();
    try {
      html.initialData = JSON.stringify(obtenerRegistrosReclamos()).replace(/<\/script/gi, '<\\/script');
    } catch(err) {
      html.initialData = '[]';
    }
    return html.evaluate().setTitle('Reclamos a Prestadores').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (e.parameter.v === 'empleados') {
    const html = HtmlService.createTemplateFromFile('Empleados');
    html.url = ScriptApp.getService().getUrl();
    html.urlMadre = 'https://docs.google.com/spreadsheets/d/' + CONFIG.HOJA_MADRE_ID;
    return html.evaluate().setTitle('Empleados Auditados · Auditoría Médica').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (e.parameter.v === 'avance') {
    const html = HtmlService.createTemplateFromFile('Avance');
    html.url = ScriptApp.getService().getUrl();
    const cached = CacheService.getScriptCache().get('datosAvance_cache_v2');
    html.initialData = cached ? cached.replace(/<\/script/gi, '<\\/script') : '{}';
    return html.evaluate().setTitle('Reportes de Avance · Auditoría Médica').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (e.parameter.v === 'kpis') {
    const html = HtmlService.createTemplateFromFile('KPIs');
    html.url = ScriptApp.getService().getUrl();
    html.urlMadre = 'https://docs.google.com/spreadsheets/d/' + CONFIG.HOJA_MADRE_ID;
    try {
      let dataKpis = null;
      const cached = CacheService.getScriptCache().get('metricas_kpis_v2');
      if (cached) {
        dataKpis = cached;
      } else if (typeof obtenerMetricasKPIs === 'function') {
        const res = obtenerMetricasKPIs(false);
        if (res && res.totales && res.totales.estudios > 0) {
          dataKpis = JSON.stringify(res);
        }
      }
      html.initialData = dataKpis ? dataKpis.replace(/<\/script/gi, '<\\/script') : '{}';
    } catch(err) {
      html.initialData = '{}';
    }
    return html.evaluate().setTitle('Panel de KPIs y Métricas · Auditoría Médica').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  const html = HtmlService.createTemplateFromFile('Index');
  html.url = ScriptApp.getService().getUrl();
  html.urlMadre = 'https://docs.google.com/spreadsheets/d/' + CONFIG.HOJA_MADRE_ID;
  return html.evaluate().setTitle('Auditoría Médica').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Función de utilidad para autorizar los permisos de Google Drive directamente
 * desde el editor de Apps Script si la cuenta aún no los aceptó.
 */
function autorizarPermisosDrive() {
  const carpeta = DriveApp.getFolderById(CONFIG.CARPETA_PDFS_ID);
  const nombre = carpeta.getName();
  Logger.log("✅ Acceso a Google Drive autorizado correctamente: " + nombre);
  return "Acceso a Google Drive autorizado correctamente: " + nombre;
}