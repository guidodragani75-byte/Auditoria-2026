// ==================================================================
// CACHÉ — datos base (ScriptCache multiusuario) [2026-09-06]
// ==================================================================
const CACHE_KEY      = 'datosBase_v5';
const DASH_CACHE_KEY = 'dashData_v4';

function _getSpreadsheet_() {
  return SpreadsheetApp.openById(CONFIG.HOJA_MADRE_ID);
}

function _getSheetMadre_() {
  return _getSpreadsheet_().getSheetByName(CONFIG.HOJA_MADRE_NAME);
}

function _setCache_(datos) {
  try {
    const json = JSON.stringify(datos);
    // Si los datos superan los 1.8 MB (más de ~22 chunks), no fragmentar en ScriptCache porque Google desaloja chunks
    if (json.length > 1800000) {
      _invalidarCache_();
      return;
    }
    const cache = CacheService.getScriptCache();
    if (json.length < 90000) {
      cache.put(CACHE_KEY, json, CONFIG.CACHE_DURACION_MIN * 60);
    } else {
      const chunkSize = 80000;
      const chunks = Math.ceil(json.length / chunkSize);
      const entries = {};
      for (let i = 0; i < chunks; i++) {
        entries[`${CACHE_KEY}_${i}`] = json.slice(i * chunkSize, (i + 1) * chunkSize);
      }
      entries[`${CACHE_KEY}_meta`] = String(chunks);
      cache.putAll(entries, CONFIG.CACHE_DURACION_MIN * 60);
      cache.put(CACHE_KEY, `__chunked__${chunks}`, CONFIG.CACHE_DURACION_MIN * 60);
    }
  } catch (e) {
    console.log("No se pudo cachear: " + e.message);
  }
}

function _getCache_() {
  try {
    const cache = CacheService.getScriptCache();
    const main  = cache.get(CACHE_KEY);
    if (!main) return null;
    if (!main.startsWith('__chunked__')) {
      try { return JSON.parse(main); } catch (e) { return null; }
    }
    const chunks = parseInt(main.replace('__chunked__', ''));
    if (isNaN(chunks) || chunks <= 0 || chunks > 25) {
      _invalidarCache_();
      return null;
    }
    const keys   = Array.from({ length: chunks }, (_, i) => `${CACHE_KEY}_${i}`);
    const parts  = cache.getAll(keys);
    // Verificar que absolutamente todos los chunks estén presentes sin faltantes
    for (let i = 0; i < chunks; i++) {
      if (!parts[`${CACHE_KEY}_${i}`]) {
        _invalidarCache_();
        return null;
      }
    }
    const json = keys.map(k => parts[k]).join('');
    return JSON.parse(json);
  } catch (err) {
    console.warn("Error leyendo caché, invalidando: " + err.message);
    _invalidarCache_();
    return null;
  }
}

function _invalidarCache_() {
  try {
    const cache = CacheService.getScriptCache();
    for (const prefix of ['datosBase_v4', 'datosBase_v5']) {
      const main = cache.get(prefix);
      if (main && main.startsWith('__chunked__')) {
        const chunks = parseInt(main.replace('__chunked__', '')) || 100;
        const keys = Array.from({ length: chunks }, (_, i) => `${prefix}_${i}`);
        keys.forEach(k => cache.remove(k));
        cache.remove(`${prefix}_meta`);
      }
      cache.remove(prefix);
    }
    cache.remove('metricas_links_empleados_v3');
    try {
      PropertiesService.getScriptProperties().deleteProperty('METRICAS_LINKS_EMPLEADOS_V3');
    } catch (eProp) {}
  } catch (e) {}
}

function calentarCacheDatos() {
  try {
    _invalidarCache_();
    cargarDatosBase();
  } catch (e) { console.error("Error calentando caché: " + e.message); }
}

function limpiarMemoria() {
  try {
    _invalidarCache_();
    const cache = CacheService.getScriptCache();
    cache.removeAll([
      'datosBase_v5', 'dashData_v4', 'empleadosData_v1', 'comunicaciones_v1',
      'datosBase_v4'
    ]);
    for (let i = 0; i < 30; i++) {
      cache.remove(`datosBase_v5_${i}`);
      cache.remove(`dashData_v4_${i}`);
      cache.remove(`datosBase_v4_${i}`);
    }
    return { status: "OK", msg: "Memoria del servidor y caché limpiados con éxito." };
  } catch (e) {
    return { status: "ERROR", msg: e.message };
  }
}

function limpiarCacheCompleto() {
  return limpiarMemoria();
}

// ==================================================================
// CARGA DE FILTROS — paralela con Promise-style via batch
// ==================================================================
function cargarFiltros() {
  const ss = _getSpreadsheet_();

  const estudios     = _leerColumnaSencilla_(ss, CONFIG.HOJA_ESTUDIOS, "A2:A");
  const fechasDeriv  = _leerFechasDerivacion_(ss);
  const vinculaciones = _leerVinculaciones_(ss);

  registrarEnBitacora('CARGA_FILTROS', { estudios: estudios.length, vinculaciones: vinculaciones.length });

  // Devuelve solo lo necesario para poblar los selects — sin comunicaciones
  // (se eliminó la lectura de comunicaciones del inicio para no releer Madre dos veces)
  return { estudios, fechasDeriv, vinculaciones };
}

function _leerColumnaSencilla_(ss, nombreHoja, rango) {
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) return [];
  return hoja.getRange(rango).getValues().map(r => String(r[0]).trim()).filter(Boolean);
}

function _leerFechasDerivacion_(ss) {
  const hoja = ss.getSheetByName(CONFIG.HOJA_FECHAS);
  if (!hoja) return [];
  return hoja.getRange("A2:A").getValues().map(r => {
    if (r[0] instanceof Date) return Utilities.formatDate(r[0], Session.getScriptTimeZone(), "dd/MM/yyyy");
    return String(r[0]).trim();
  }).filter(Boolean);
}

function _leerVinculaciones_(ss) {
  const hoja = ss.getSheetByName(CONFIG.HOJA_VINCULACIONES);
  if (!hoja) return [];
  const ultFila = hoja.getLastRow();
  if (ultFila < 2) return [];
  return hoja.getRange(2, 1, ultFila - 1, 4).getValues()
    .map(row => ({
      contrato: String(row[CONFIG.VIN_COL_CONTRATO]).trim(),
      organismo: String(row[CONFIG.VIN_COL_ORGANISMO]).trim(),
      srt:      String(row[CONFIG.VIN_COL_SRT]).trim(),
      estab:    String(row[CONFIG.VIN_COL_ESTABLECIMIENTO]).trim()
    }))
    .filter(v => v.contrato && v.organismo && v.srt && v.estab);
}

// ==================================================================
// CARGA DE DATOS BASE — lectura en batch único
// ==================================================================
function cargarDatosBase() {
  const cached = _getCache_();
  if (cached) return cached;

  const sheet    = _getSheetMadre_();
  if (!sheet) return [];
  const ultFila  = sheet.getLastRow();
  if (ultFila < 2) return [];

  // Asegurar columnas y encabezados para AL (Guardado) y AM (Link_carpeta)
  const colsNecesarias = Math.max(CONFIG.COL_ADJUNTO_2 + 1, (CONFIG.COL_LINK_CARPETA || 38) + 1);
  const maxCols = sheet.getMaxColumns();
  if (maxCols < colsNecesarias) {
    sheet.insertColumnsAfter(maxCols, colsNecesarias - maxCols);
  }

  const colAL = (CONFIG.COL_GUARDADO !== undefined ? CONFIG.COL_GUARDADO : 37) + 1;
  if (sheet.getLastRow() >= 1 && !sheet.getRange(1, colAL).getValue()) {
    sheet.getRange(1, colAL, 1, 2).setValues([["Guardado", "Link_carpeta"]]);
  }

  const rows = sheet.getRange(2, 1, ultFila - 1, colsNecesarias).getValues();

  // Pre-calcular timezone una sola vez
  const tz = Session.getScriptTimeZone();

  const datos = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row[CONFIG.COL_CUIL]) continue;
    datos.push(_mapearFila_(row, i, tz));
  }

  _setCache_(datos);
  registrarEnBitacora('CARGA_DATOS_BASE', { registros: datos.length });
  return datos;
}

function _mapearFila_(row, i, tz) {
  return {
    rowIndex:           i + 2,
    Clave:              String(row[CONFIG.COL_CLAVE] || "").trim(),
    Contrato:           String(row[CONFIG.COL_CUIT_CONTRATO] || "").trim(),
    Nombre_Contrato:    String(row[CONFIG.COL_NOMBRE_CONTRATO] || "").trim(),
    Fecha_Derivacion:   _parseFechaConTz_(row[CONFIG.COL_FECHA_DERIV], tz),
    Nombre_Organismo:   String(row[CONFIG.COL_NOMBRE_ORGANISMO] || "").trim(),
    Establecimiento:    String(row[CONFIG.COL_NOMBRE_ESTABLECIMIENTO] || "").trim(),
    Codigo_SRT:         String(row[CONFIG.COL_SRT] || "").trim(),
    CUIL_Empleado:      String(row[CONFIG.COL_CUIL] || "").trim(),
    Apellido_y_Nombre:  String(row[CONFIG.COL_APELLIDO] || "").trim(),
    Estudio_a_Realizar: String(row[CONFIG.COL_ESTUDIO] || "").trim(),
    Asistencia:         String(row[CONFIG.COL_ASISTENCIA] || "").toUpperCase().trim(),
    Fecha_Auditoria:    _parseFechaConTz_(row[CONFIG.COL_FECHA_AUDITORIA], tz),
    Fecha_Estudio:      _parseFechaConTz_(row[CONFIG.COL_FECHA_ESTUDIO], tz),
    Resultado:          String(row[CONFIG.COL_RESULTADO] || "").trim(),
    Observación:        String(row[CONFIG.COL_OBSERVACION] || ""),
    Comunicacion:       _parseFechaConTz_(row[CONFIG.COL_COMUNICACION], tz),
    Adjunto1:           String(row[CONFIG.COL_ADJUNTO_1] || "").trim(),
    Adjunto2:           String(row[CONFIG.COL_ADJUNTO_2] || "").trim(),
    Guardado:           String(row[CONFIG.COL_GUARDADO] || "").trim(),
    Link_Carpeta:       String(row[CONFIG.COL_LINK_CARPETA] || "").trim()
  };
}

function _parseFechaConTz_(val, tz) {
  if (!val && val !== 0) return "";
  if (typeof formatearFecha === 'function') {
    return formatearFecha(val, tz);
  }
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "";
    return Utilities.formatDate(val, tz || Session.getScriptTimeZone() || "America/Argentina/Buenos_Aires", "dd/MM/yyyy");
  }
  const s = String(val).trim();
  const mIso = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (mIso) return `${mIso[3].padStart(2, '0')}/${mIso[2].padStart(2, '0')}/${mIso[1]}`;
  const mAr = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (mAr) return `${mAr[1].padStart(2, '0')}/${mAr[2].padStart(2, '0')}/${mAr[3]}`;
  return s;
}

function _esGuardado_(val) {
  if (!val && val !== true) return false;
  if (typeof val === 'boolean') return val;
  const s = String(val).trim().toUpperCase();
  if (!s || s === 'NO' || s === 'FALSE' || s === 'FALSO' || s === '-' || s === 'PENDIENTE') return false;
  if (s.startsWith('SIN ') || s.startsWith('SIN/') || s.startsWith('ERROR')) return false;
  return s === 'SÍ' || s === 'SI' || s === 'TRUE' || s === 'VERDADERO' ||
         s.startsWith('SÍ') || s.startsWith('SI ') || s.startsWith('SI(') || s.startsWith('SI/') || s.startsWith('SI-') ||
         s.includes('GUARDAD') || s.includes('ARCHIV') || s === 'OK';
}

function _esResultadoApto_(res) {
  if (!res) return false;
  const r = (typeof _normalizarTexto_ === 'function')
    ? _normalizarTexto_(res)
    : String(res).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
  return r.includes('normal laboral') || r.includes('enfermedad profesional') || r.includes('desistido');
}

function _mapearFilaHojaEspecial_(row, i, tz) {
  const mapped = _mapearFila_(row, i, tz);
  const filaMadre = row[CONFIG.COL_FILA_MADRE];
  if (filaMadre && !isNaN(Number(filaMadre))) {
    mapped.rowIndex = Number(filaMadre);
  }
  return mapped;
}

// ==================================================================
// GUARDADO
// ==================================================================
function normalizarFechaBackend_(f) {
  if (!f) return "";
  if (typeof formatearFecha === 'function') return formatearFecha(f);
  f = String(f).trim();
  if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(f)) {
    const p = f.split(/[-/]/);
    return `${p[2].padStart(2, '0')}/${p[1].padStart(2, '0')}/${p[0]}`;
  }
  if (/^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/.test(f)) {
    const p = f.split(/[-/]/);
    return `${p[0].padStart(2, '0')}/${p[1].padStart(2, '0')}/${p[2]}`;
  }
  return f;
}
function validarFecha_(f) { return /^\d{2}[/-]\d{2}[/-]\d{4}$/.test(f) || /^\d{4}[/-]\d{2}[/-]\d{2}$/.test(f) || f === ""; }
function _validarResultado_(r) { return CONFIG.RESULTADOS_PERMITIDOS.includes(r) || r === ""; }

function guardarMultiples(cambios) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.TIMEOUT_LOCK);

    for (const c of cambios) {
      c.fechaEstudio = normalizarFechaBackend_(c.fechaEstudio);
      if (!validarFecha_(c.fechaEstudio)) throw new Error(`Fecha inválida: "${c.fechaEstudio}"`);
      if (!_validarResultado_(c.res))     throw new Error(`Resultado no permitido: "${c.res}"`);
      if (c.res === "Desistido" && (!c.obs || !String(c.obs).trim())) {
        throw new Error(`La observación es obligatoria para estudios con resultado Desistido (fila ${c.idx})`);
      }
    }

    const sheet    = _getSheetMadre_();
    const fechaHoy = new Date();

    // Escritura en batch de las 4 columnas contiguas (cols 27 a 30: Fecha Auditoría, Fecha Estudio, Resultado, Observación)
    const indices = cambios.map(c => c.idx);
    const minIdx  = Math.min(...indices);
    const maxIdx  = Math.max(...indices);
    const rango   = maxIdx - minIdx + 1;

    const colInicio = CONFIG.COL_FECHA_AUDITORIA + 1; // col 27

    if (cambios.length > 3 && rango < 800) {
      // Bloque contiguo único (1 lectura + 1 escritura)
      const bloque = sheet.getRange(minIdx, colInicio, rango, 4).getValues();
      cambios.forEach(c => {
        const off = c.idx - minIdx;
        bloque[off][0] = fechaHoy;
        bloque[off][1] = c.fechaEstudio || "";
        bloque[off][2] = c.res;
        bloque[off][3] = c.obs || "";
      });
      sheet.getRange(minIdx, colInicio, rango, 4).setValues(bloque);
    } else {
      // Rango amplio o pocos registros: escribir 4 celdas juntas por fila (1 sola llamada por fila)
      cambios.forEach(c => {
        sheet.getRange(c.idx, colInicio, 1, 4).setValues([[fechaHoy, c.fechaEstudio || "", c.res, c.obs || ""]]);
      });
    }

    // [2026-09-11] Si el resultado es Ausente / Sin realizar, borrar automáticamente los links de adjuntos y guardado en la Hoja Madre
    const colAdj1 = CONFIG.COL_ADJUNTO_1 + 1; // col 33 (AG)
    const colG = (CONFIG.COL_GUARDADO !== undefined ? CONFIG.COL_GUARDADO : 37) + 1; // col 38 (AL)
    cambios.forEach(c => {
      if (_normalizarTexto_(c.res).includes('ausente')) {
        sheet.getRange(c.idx, colAdj1, 1, 2).setValues([["", ""]]); // Limpia Link1 (AG) y Link2 (AH)
        sheet.getRange(c.idx, colG, 1, 2).setValues([["", ""]]);    // Limpia Guardado (AL) y Link_Carpeta (AM)
      }
    });

    // Sincronizar en tiempo real en las hojas dedicadas (Desvios, Reclamos y Empleados)
    try {
      _sincronizarCambiosHojasEspeciales_(sheet, cambios, fechaHoy, Session.getScriptTimeZone());
      _actualizarAuditoriaEnHojaEmpleados_(sheet, cambios, fechaHoy, Session.getScriptTimeZone());
    } catch (eSync) {
      console.error('Error sincronizando hojas especiales: ' + eSync.message);
    }

    // Obtener lista deduplicada por empleado de PDFs pendientes de archivar
    let pendientesArchivar = [];
    try {
      pendientesArchivar = _obtenerPendientesArchivarDeduplicados_(sheet, cambios);
    } catch (ePend) {
      console.error('Error calculando pendientes de archivar: ' + ePend.message);
    }

    SpreadsheetApp.flush();
    _invalidarCache_();
    registrarEnBitacora('GUARDADO_EXITOSO', { cantidad: cambios.length, pendientesPDF: pendientesArchivar.length });

    return {
      status: "OK",
      fecha: _parseFechaConTz_(fechaHoy, Session.getScriptTimeZone()),
      cantidad: cambios.length,
      pendientesArchivar: pendientesArchivar
    };

  } catch (e) {
    registrarEnBitacora('ERROR_GUARDADO', { error: e.message });
    if (String(e).includes("lock")) return { status: "Error", msg: "Reintentá en unos segundos" };
    return { status: "Error", msg: String(e) };
  } finally {
    lock.releaseLock();
  }
}

function limpiarRegistroIndividual(idx) {
  try {
    const sheet = _getSheetMadre_();
    sheet.getRange(idx, CONFIG.COL_FECHA_AUDITORIA + 1, 1, 4).clearContent();
    SpreadsheetApp.flush();
    _invalidarCache_();
    try {
      sincronizarTodasLasHojasDesdeMadre(true);
    } catch (eSync) {
      console.warn("Aviso al sincronizar hojas tras borrar: " + eSync.message);
    }
    registrarEnBitacora('REGISTRO_LIMPIADO', { rowIndex: idx });
    return "OK";
  } catch (e) {
    registrarEnBitacora('ERROR_LIMPIAR', { error: e.message });
    return "Error: " + String(e);
  }
}

// ==================================================================
// CACHÉ DEL DASHBOARD — independiente de la caché de datos base
// ==================================================================
function _getDashCache_() {
  const cache = CacheService.getScriptCache();
  const main  = cache.get(DASH_CACHE_KEY);
  if (!main) return null;
  if (!main.startsWith('__chunked__')) return JSON.parse(main);
  const chunks = parseInt(main.replace('__chunked__', ''));
  const keys   = Array.from({ length: chunks }, (_, i) => `${DASH_CACHE_KEY}_${i}`);
  const parts  = cache.getAll(keys);
  const json   = keys.map(k => parts[k] || '').join('');
  return json ? JSON.parse(json) : null;
}

function _setDashCache_(datos) {
  try {
    const cache = CacheService.getScriptCache();
    const json  = JSON.stringify(datos);
    const TTL   = 60 * 60; // 1 hora
    if (json.length < 90000) {
      cache.put(DASH_CACHE_KEY, json, TTL);
    } else {
      const chunkSize = 80000;
      const chunks    = Math.ceil(json.length / chunkSize);
      const entries   = {};
      for (let i = 0; i < chunks; i++) {
        entries[`${DASH_CACHE_KEY}_${i}`] = json.slice(i * chunkSize, (i + 1) * chunkSize);
      }
      entries[`${DASH_CACHE_KEY}_meta`] = String(chunks);
      cache.putAll(entries, TTL);
      cache.put(DASH_CACHE_KEY, `__chunked__${chunks}`, TTL);
    }
  } catch (e) { console.log('No se pudo cachear dashboard: ' + e.message); }
}

function _invalidarCacheDashboard_() {
  const cache = CacheService.getScriptCache();
  const main  = cache.get(DASH_CACHE_KEY);
  if (main && main.startsWith('__chunked__')) {
    const chunks = parseInt(main.replace('__chunked__', ''));
    Array.from({ length: chunks }, (_, i) => `${DASH_CACHE_KEY}_${i}`)
         .forEach(k => cache.remove(k));
    cache.remove(`${DASH_CACHE_KEY}_meta`);
  }
  cache.remove(DASH_CACHE_KEY);
}

// Pre-calcula y guarda en caché. Llamado por trigger.
function calentarCacheDashboard() {
  try {
    _invalidarCacheDashboard_();
    obtenerDatosDashboard(); // recalcula y deja en caché
    registrarEnBitacora('DASH_CACHE_CALENTADO', { timestamp: new Date().toISOString() });
  } catch (e) { console.error('Error calentando caché dashboard: ' + e.message); }
}

// ==================================================================
// DASHBOARD — métricas calculadas desde Hoja Madre
// ==================================================================
function obtenerDatosDashboard(contratoFiltro) {
  // No usamos caché de dashboard porque el filtro es variable
  try {
    const datos = cargarDatosBase();

    // Lista de contratos únicos para el dropdown
    const contratosSet = new Set();
    datos.forEach(r => { if (r.Contrato) contratosSet.add(r.Contrato); });
    const contratosDisponibles = Array.from(contratosSet).sort();

    // Filtro por contrato (columna C) + exclusiones estándar
    const rows = datos.filter(r =>
      !String(r.Estudio_a_Realizar).includes('No posee') &&
      (!contratoFiltro || r.Contrato === String(contratoFiltro))
    );

    const cuilesGlobal   = new Set();
    const organismos     = new Map();   // key: nombre lowercase
    const establecimientos = new Map(); // key: org|||estab
    const porAsistencia  = new Map();
    const porResultado   = new Map();
    let totalConComunicacion = 0;

    rows.forEach(r => {
      const orgKey   = r.Nombre_Organismo   || '(sin organismo)';
      const estKey   = orgKey + '|||' + (r.Establecimiento || '(sin establecimiento)');
      const cuil     = r.CUIL_Empleado || r.CUIL_Original;
      const asist    = r.Asistencia || 'SIN DATO';
      const resultado = r.Resultado || '';
      const esPresente = asist.includes('PRESENTE');
      const esAuditado = resultado !== '';

      cuilesGlobal.add(cuil);
      if (r.Comunicacion) totalConComunicacion++;

      // — Asistencia —
      if (!porAsistencia.has(asist)) porAsistencia.set(asist, { cuiles: new Set(), estudios: 0 });
      porAsistencia.get(asist).cuiles.add(cuil);
      porAsistencia.get(asist).estudios++;

      // — Resultado (solo auditados) —
      if (resultado) porResultado.set(resultado, (porResultado.get(resultado) || 0) + 1);

      // — Por organismo —
      if (!organismos.has(orgKey)) {
        organismos.set(orgKey, {
          cuiles: new Set(), estudios: 0,
          presentesCuiles: new Set(), auditados: 0
        });
      }
      const o = organismos.get(orgKey);
      o.cuiles.add(cuil);
      o.estudios++;
      if (esPresente) o.presentesCuiles.add(cuil);
      if (esAuditado) o.auditados++;

      // — Por establecimiento —
      if (!establecimientos.has(estKey)) {
        establecimientos.set(estKey, {
          organismo: orgKey,
          nombre: r.Establecimiento || '(sin establecimiento)',
          cuiles: new Set(),
          presentesCuiles: new Set(), // CUILs con al menos una fila PRESENTE
          estudios: 0, presentes: 0, auditados: 0
        });
      }
      const e = establecimientos.get(estKey);
      e.cuiles.add(cuil);
      e.estudios++;
      if (esPresente) { e.presentes++; e.presentesCuiles.add(cuil); }
      if (esAuditado) e.auditados++;
    });

    // Serializar organismos
    // Primero indexar establecimientos por organismo para calcular métricas
    const estabsPorOrg = new Map();
    establecimientos.forEach((v, k) => {
      if (!estabsPorOrg.has(v.organismo)) estabsPorOrg.set(v.organismo, []);
      estabsPorOrg.get(v.organismo).push(v);
    });

    const porOrganismo = Array.from(organismos.entries()).map(([k, v]) => {
      const orgEstabs        = estabsPorOrg.get(k) || [];
      const estabConPresente = orgEstabs.filter(e => e.presentesCuiles.size > 0).length;
      const estabTodoAusente = orgEstabs.length - estabConPresente; // garantiza suma exacta
      return {
        nombre:    k,
        empleados: v.cuiles.size,
        estudios:  v.estudios,
        presentes: v.presentesCuiles.size,
        auditados: v.auditados,
        pendientes: v.estudios - v.auditados,
        pctPresentismo: v.cuiles.size > 0 ? Math.round(v.presentesCuiles.size / v.cuiles.size * 100) : 0,
        pctAuditado:   v.estudios > 0 ? Math.round(v.auditados / v.estudios * 100) : 0,
        // Métricas de establecimientos
        totalEstab:        orgEstabs.length,
        estabConPresente,
        estabTodoAusente
      };
    }).sort((a, b) => b.estudios - a.estudios);

    // Serializar establecimientos
    const porEstablecimiento = Array.from(establecimientos.values()).map(v => ({
      organismo: v.organismo,
      nombre:    v.nombre,
      empleados: v.cuiles.size,
      estudios:  v.estudios,
      presentes: v.presentes,
      auditados: v.auditados,
      pendientes: v.estudios - v.auditados,
      pctAuditado: v.estudios > 0 ? Math.round(v.auditados / v.estudios * 100) : 0
    })).sort((a, b) => b.estudios - a.estudios);

    // Serializar asistencia
    const asistenciaArr = Array.from(porAsistencia.entries()).map(([k, v]) => ({
      etiqueta: k,
      empleados: v.cuiles.size,
      estudios:  v.estudios
    })).sort((a, b) => b.estudios - a.estudios);

    // Serializar resultados
    const resultadoArr = Array.from(porResultado.entries()).map(([k, v]) => ({
      etiqueta: k,
      cantidad: v
    })).sort((a, b) => b.cantidad - a.cantidad);

    const totalAuditados  = rows.filter(r => r.Resultado !== '').length;
    const totalPendientes = rows.length - totalAuditados;

    registrarEnBitacora('DASHBOARD_CONSULTADO', { filas: rows.length });

    const resultado = {
      kpis: {
        organismos:       organismos.size,
        establecimientos: establecimientos.size,
        empleados:        cuilesGlobal.size,
        estudios:         rows.length,
        auditados:        totalAuditados,
        pendientes:       totalPendientes,
        conComunicacion:  totalConComunicacion
      },
      porAsistencia:      asistenciaArr,
      porResultado:       resultadoArr,
      porOrganismo:       porOrganismo,
      porEstablecimiento: porEstablecimiento,
      contratosDisponibles,
      contratoActual:     contratoFiltro || ''
    };

    return resultado;

  } catch (e) {
    registrarEnBitacora('ERROR_DASHBOARD', { error: e.message });
    return { error: 'Error: ' + e.toString() };
  }
}

// ==================================================================
// COBERTURA — agrupado por Organismo → Establecimiento
// ==================================================================
function obtenerDatosCobertura() {
  try {
    const datos = cargarDatosBase();

    const EN_GESTION = new Set(['EN PROCESO', 'TURNO ASIGNADO', 'PENDIENTE DE TURNO', 'AUSENTE']);
    const AUSENTES_F = new Set(['AUSENTE FINAL', 'NO REALIZA GESTIÓN']);

    // Excluir "sin estudios"
    const rows = datos.filter(r =>
      !String(r.Estudio_a_Realizar).includes('No posee')
    );

    // Map: "org|||est" → acumuladores con Sets
    const grupos = new Map();

    rows.forEach(r => {
      const org  = r.Nombre_Organismo  || '(sin organismo)';
      const est  = r.Establecimiento   || '(sin establecimiento)';
      const key  = org + '|||' + est;
      const cuil = r.CUIL_Empleado || r.CUIL_Original;
      const asist = r.Asistencia;

      if (!grupos.has(key)) {
        grupos.set(key, {
          organismo:       org,
          establecimiento: est,
          empleados:       new Set(),
          estudios:        0,
          presentes:       new Set(),
          enGestion:       new Set(),
          ausentes:        new Set(),
          sinEstado:       new Set(),
          conResultado:    new Set()
        });
      }

      const g = grupos.get(key);
      g.empleados.add(cuil);
      g.estudios++;

      if      (asist === 'PRESENTE')      g.presentes.add(cuil);
      else if (EN_GESTION.has(asist))     g.enGestion.add(cuil);
      else if (AUSENTES_F.has(asist))     g.ausentes.add(cuil);
      else                                g.sinEstado.add(cuil);

      if (r.Resultado)                    g.conResultado.add(cuil);
    });

    // Agrupar por organismo y serializar
    const orgMap = new Map();
    grupos.forEach(g => {
      const org = g.organismo;
      if (!orgMap.has(org)) {
        orgMap.set(org, {
          nombre:          org,
          empleados:       new Set(),
          estudios:        0,
          presentes:       new Set(),
          enGestion:       new Set(),
          ausentes:        new Set(),
          sinEstado:       new Set(),
          conResultado:    new Set(),
          establecimientos: []
        });
      }
      const o = orgMap.get(org);
      const emp    = g.empleados.size;
      const conRes = g.conResultado.size;

      g.empleados.forEach(c    => o.empleados.add(c));
      g.presentes.forEach(c    => o.presentes.add(c));
      g.enGestion.forEach(c    => o.enGestion.add(c));
      g.ausentes.forEach(c     => o.ausentes.add(c));
      g.sinEstado.forEach(c    => o.sinEstado.add(c));
      g.conResultado.forEach(c => o.conResultado.add(c));
      o.estudios += g.estudios;

      o.establecimientos.push({
        nombre:       g.establecimiento,
        empleados:    emp,
        estudios:     g.estudios,
        presentes:    g.presentes.size,
        enGestion:    g.enGestion.size,
        ausentes:     g.ausentes.size,
        sinEstado:    g.sinEstado.size,
        conResultado: conRes,
        sinResultado: emp - conRes
      });
    });

    const result = Array.from(orgMap.values()).map(o => {
      const emp    = o.empleados.size;
      const conRes = o.conResultado.size;
      return {
        nombre:          o.nombre,
        empleados:       emp,
        estudios:        o.estudios,
        presentes:       o.presentes.size,
        enGestion:       o.enGestion.size,
        ausentes:        o.ausentes.size,
        sinEstado:       o.sinEstado.size,
        conResultado:    conRes,
        sinResultado:    emp - conRes,
        establecimientos: o.establecimientos.sort((a, b) => b.empleados - a.empleados)
      };
    }).sort((a, b) => b.empleados - a.empleados);

    // Totales globales
    const totEmp    = new Set(rows.map(r => r.CUIL_Empleado || r.CUIL_Original)).size;
    const totConRes = new Set(rows.filter(r => r.Resultado).map(r => r.CUIL_Empleado || r.CUIL_Original)).size;

    registrarEnBitacora('COBERTURA_CONSULTADO', { organismos: result.length, empleados: totEmp });
    return {
      data: result,
      totales: {
        organismos:   result.length,
        empleados:    totEmp,
        estudios:     rows.length,
        conResultado: totConRes,
        sinResultado: totEmp - totConRes
      }
    };

  } catch (e) {
    registrarEnBitacora('ERROR_COBERTURA', { error: e.message });
    return { error: 'Error: ' + e.toString() };
  }
}

// ==================================================================
// GESTIÓN DE HOJAS DEDICADAS: DESVÍOS, RECLAMOS, COMUNICACIONES, EMPLEADOS
// ==================================================================

/**
 * Inicializa encabezados y formato en la hoja dedicada "Desvios".
 * 1 fila = 1 empleado con desvíos/confirmaciones (15 columnas limpias).
 */
function inicializarHojaDesvios() {
  const ss = _getSpreadsheet_();
  let hoja = ss.getSheetByName(CONFIG.HOJA_DESVIOS);
  if (!hoja) {
    hoja = ss.insertSheet(CONFIG.HOJA_DESVIOS);
  }
  const headers = [
    "CUIL", "Apellido_y_Nombre", "Contrato", "Nombre_Contrato",
    "Codigo_SRT", "Establecimiento", "Nombre_Organismo", "Estado_Global",
    "Total_Estudios", "Total_Desvios", "Fecha_Auditoria", "Comunicacion",
    "Adjunto1", "Filas_Madre", "Detalle_Estudios", "Prestador", "Clave_Unica",
    "Prestador2", "Fecha_VE", "Resultado_Reconfirmacion", "Observacion_Reconfirmacion",
    "Estudios_Adicionales"
  ];
  if (hoja.getLastRow() === 0 || hoja.getRange(1, 1).getValue() !== "CUIL") {
    hoja.clear();
    hoja.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = hoja.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#d97706"); // Amber
    headerRange.setFontColor("#ffffff");
    hoja.setFrozenRows(1);
  } else {
    // Si la hoja ya tiene datos pero le faltan columnas, las agregamos dinámicamente por nombre
    const currentCols = hoja.getLastColumn();
    const existingHeaders = hoja.getRange(1, 1, 1, currentCols).getValues()[0].map(h => String(h || '').trim());
    headers.forEach(h => {
      if (existingHeaders.indexOf(h) === -1) {
        const nextCol = hoja.getLastColumn() + 1;
        hoja.getRange(1, nextCol).setValue(h).setFontWeight("bold").setBackground("#d97706").setFontColor("#ffffff");
      }
    });
  }
  return hoja;
}

/**
 * Inicializa encabezados y formato en la hoja dedicada "Reclamos".
 * 1 fila = 1 empleado con reclamos a prestadores (16 columnas limpias).
 */
function inicializarHojaReclamos() {
  const ss = _getSpreadsheet_();
  let hoja = ss.getSheetByName(CONFIG.HOJA_RECLAMOS);
  if (!hoja) {
    hoja = ss.insertSheet(CONFIG.HOJA_RECLAMOS);
  }
  const headers = [
    "CUIL", "Apellido_y_Nombre", "Contrato", "Nombre_Contrato",
    "Codigo_SRT", "Establecimiento", "Nombre_Organismo", "Motivo_Global",
    "Total_Reclamables", "Total_Estudios", "Fecha_Auditoria",
    "Adjunto1", "Filas_Madre", "Detalle_Estudios", "Prestador", "Clave_Unica"
  ];
  if (hoja.getLastRow() === 0 || hoja.getLastColumn() !== headers.length || hoja.getRange(1, 1).getValue() !== "CUIL") {
    hoja.clear();
    hoja.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = hoja.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#dc2626"); // Rojo
    headerRange.setFontColor("#ffffff");
    hoja.setFrozenRows(1);
  }
  return hoja;
}

/**
 * Inicializa encabezados y formato en la hoja dedicada "Comunicaciones".
 * 1 fila = 1 empleado apto para comunicación (16 columnas limpias).
 */
function inicializarHojaComunicaciones() {
  const ss = _getSpreadsheet_();
  let hoja = ss.getSheetByName(CONFIG.HOJA_COMUNICACIONES);
  if (!hoja) {
    hoja = ss.insertSheet(CONFIG.HOJA_COMUNICACIONES);
  }
  const headers = [
    "CUIL", "Apellido_y_Nombre", "Contrato", "Nombre_Contrato",
    "Codigo_SRT", "Establecimiento", "Nombre_Organismo", "Estado_Consolidado",
    "Total_Estudios", "Fecha_Auditoria", "Comunicacion",
    "Adjunto1", "Filas_Madre", "Detalle_Estudios", "Prestador", "Clave_Unica",
    "Frecuencia", "Lote_Key", "Canal_Comunicacion", "Destinatario", "Espera_Cierre", "Apto_Comunicar", "Estado_Lote",
    "Fecha_Comunicacion_Lote"
  ];
  if (hoja.getLastRow() === 0 || hoja.getRange(1, 1).getValue() !== "CUIL") {
    hoja.clear();
    hoja.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = hoja.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#2563eb"); // Azul
    headerRange.setFontColor("#ffffff");
    hoja.setFrozenRows(1);
  } else {
    // Si la hoja ya tiene datos pero le faltan columnas de Lote y Canal, agregarlas dinámicamente sin borrar datos
    const currentCols = hoja.getLastColumn();
    const existingHeaders = hoja.getRange(1, 1, 1, currentCols).getValues()[0].map(h => String(h || '').trim());
    headers.forEach(h => {
      if (existingHeaders.indexOf(h) === -1) {
        const nextCol = hoja.getLastColumn() + 1;
        hoja.getRange(1, nextCol).setValue(h).setFontWeight("bold").setBackground("#2563eb").setFontColor("#ffffff");
      }
    });
  }
  return hoja;
}

/**
 * Obtiene o crea la hoja Comunicaciones con búsqueda flexible.
 */
function _obtenerOCrearHojaComunicaciones_(ss) {
  const spreadsheet = ss || _getSpreadsheet_();
  let hoja = spreadsheet.getSheetByName(CONFIG.HOJA_COMUNICACIONES) ||
             spreadsheet.getSheetByName("Comunicaciones") ||
             spreadsheet.getSheetByName("Comunicacion") ||
             spreadsheet.getSheetByName("Comunicar");
  if (!hoja) {
    hoja = inicializarHojaComunicaciones();
  }
  return hoja;
}

function inicializarEncabezadosHojasEspeciales() {
  inicializarHojaDesvios();
  inicializarHojaReclamos();
  inicializarHojaComunicaciones();
  if (typeof inicializarHojaEmpleados === 'function') {
    inicializarHojaEmpleados();
  }
  return { status: "OK", mensaje: "Encabezados inicializados correctamente." };
}

function _normalizarTexto_(txt) {
  if (!txt) return "";
  return String(txt).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
}

function _esResultadoDesvio_(r) {
  if (!r) return false;
  const s = _normalizarTexto_(r);
  return s.includes("desvio") || s.includes("pendiente de confirmacion") || s.includes("enfermedad profesional");
}

function _esResultadoReclamo_(r) {
  if (!r) return false;
  const s = (typeof _normalizarTexto_ === 'function') ? _normalizarTexto_(r) : String(r).toLowerCase().trim();
  return s.includes("ausente") || s.includes("practica defectuosa") || s.includes("defecto") ||
         s.includes("sin realizar") || s.includes("no realizado") || s.includes("incompleto") ||
         s.includes("rechazado") || (typeof _esResultadoApto_ === 'function' ? !_esResultadoApto_(r) : false);
}

/**
 * Sincroniza en batch todas las filas desde la Hoja Madre hacia Desvios, Reclamos, Comunicaciones y Empleados.
 * Procesa en 1 sola lectura de Madre y agrupa por empleado (CUIL) para máxima velocidad.
 */
function sincronizarTodasLasHojasDesdeMadre(skipLock) {
  const lock = skipLock ? null : LockService.getScriptLock();
  try {
    if (lock && !lock.tryLock(15000)) {
      return { status: "Error", msg: "El servidor está procesando otra sincronización. Aguarda unos segundos y vuelve a intentar." };
    }

    const hojaEmp  = typeof inicializarHojaEmpleados === 'function' ? inicializarHojaEmpleados() : null;
    const hojaDesv = inicializarHojaDesvios();
    const hojaRecl = inicializarHojaReclamos();
    const hojaCom  = inicializarHojaComunicaciones();

    const sheetMadre = _getSheetMadre_();
    if (!sheetMadre) throw new Error("No se encontró la hoja Madre.");

    const ultFila = sheetMadre.getLastRow();
    if (ultFila < 2) {
      return { status: "OK", desvios: 0, reclamos: 0, comunicaciones: 0, empleados: 0, msg: "Madre vacía" };
    }

    // Mapa para preservar asignaciones previas (Prestador2 y Fecha_VE) en la hoja Desvios
    const prevDesvMap = {};
    if (hojaDesv.getLastRow() >= 2) {
      const hCols = hojaDesv.getLastColumn();
      const hHeaders = hojaDesv.getRange(1, 1, 1, hCols).getValues()[0].map(h => String(h || '').trim());
      const cClave = hHeaders.indexOf("Clave_Unica");
      const cCuil = hHeaders.indexOf("CUIL");
      const cSrt = hHeaders.indexOf("Codigo_SRT");
      const cP2 = hHeaders.indexOf("Prestador2");
      const cFve = hHeaders.indexOf("Fecha_VE");
      const cResReconf = hHeaders.indexOf("Resultado_Reconfirmacion");
      const cObsReconf = hHeaders.indexOf("Observacion_Reconfirmacion");
      const cEstAdic = hHeaders.indexOf("Estudios_Adicionales");
      const dRows = hojaDesv.getRange(2, 1, hojaDesv.getLastRow() - 1, hCols).getValues();
      dRows.forEach(r => {
        const k = (cClave !== -1 ? r[cClave] : '') || (cCuil !== -1 ? (r[cCuil] + '_' + (cSrt !== -1 ? r[cSrt] : '')) : '');
        if (k) {
          prevDesvMap[String(k).trim()] = {
            prestador2: cP2 !== -1 ? String(r[cP2] || '') : '',
            fechaVE: cFve !== -1 ? String(r[cFve] || '') : '',
            resultadoReconfirmacion: cResReconf !== -1 ? String(r[cResReconf] || '') : '',
            observacionReconfirmacion: cObsReconf !== -1 ? String(r[cObsReconf] || '') : '',
            estudiosAdicionales: cEstAdic !== -1 ? String(r[cEstAdic] || '') : ''
          };
        }
      });
    }

    // Mapa para preservar estado previo de Empleados y detectar reemplazo de PDFs por fileId
    const prevEmpMap = new Map();
    if (hojaEmp && hojaEmp.getLastRow() >= 2) {
      const eCols = hojaEmp.getLastColumn();
      const eRows = hojaEmp.getRange(2, 1, hojaEmp.getLastRow() - 1, Math.max(eCols, 20)).getValues();
      eRows.forEach(r => {
        const k = String(r[19] || (r[0] + '___' + (r[5] || ''))).trim();
        const obj = {
          fileId: String(r[13] || '').trim(),
          linkAdjunto: String(r[12] || '').trim(),
          guardado: String(r[10] || '').trim()
        };
        if (k) prevEmpMap.set(k, obj);
        if (r[0] && !prevEmpMap.has(String(r[0]).trim())) {
          prevEmpMap.set(String(r[0]).trim(), obj);
        }
      });
    }

    const numCols = Math.max(sheetMadre.getLastColumn(), 42);
    const headersMadreSync = sheetMadre.getRange(1, 1, 1, numCols).getValues()[0].map(h => String(h || '').trim().toLowerCase());
    let colPrestadorMadre = CONFIG.COL_PRESTADOR;
    const idxPrestFound = headersMadreSync.indexOf('prestador');
    if (idxPrestFound !== -1) {
      colPrestadorMadre = idxPrestFound;
    }

    // Lectura en 1 solo bloque continuo de la Hoja Madre (1 sola llamada a la API)
    const allMadreValues = sheetMadre.getRange(2, 1, ultFila - 1, numCols).getValues();

    const cuilsAuditados = new Set();
    const mapaLotesMadre = new Map();
    const colG = (CONFIG.COL_GUARDADO !== undefined ? CONFIG.COL_GUARDADO : 37);
    const colAdj1 = (CONFIG.COL_ADJUNTO_1 !== undefined ? CONFIG.COL_ADJUNTO_1 : 32);
    const colAdj2 = (CONFIG.COL_ADJUNTO_2 !== undefined ? CONFIG.COL_ADJUNTO_2 : 33);

    for (let i = 0; i < allMadreValues.length; i++) {
      const row = allMadreValues[i];
      const res = String(row[CONFIG.COL_RESULTADO] || '').trim();
      const gd = String(row[colG] || '').trim();
      const c = String(row[CONFIG.COL_CUIL] || '').trim();
      const a1 = String(row[colAdj1] || '').trim();
      const a2 = String(row[colAdj2] || '').trim();

      if (res !== '' || gd.toUpperCase() === 'SÍ' || gd.toUpperCase() === 'SI' || a1 !== '' || a2 !== '') {
        if (c) cuilsAuditados.add(c);
      }

      const estudio = String(row[CONFIG.COL_ESTUDIO] || '').trim();
      if (!estudio || estudio.includes("No posee")) continue;

      const cuit = String(row[CONFIG.COL_CUIT_CONTRATO] || '').trim();
      const srt = String(row[CONFIG.COL_SRT] || '').trim();
      const frec = String(row[CONFIG.COL_FRECUENCIA] || '').trim();
      const loteKey = `${cuit}___${srt}___${frec}`;

      if (!mapaLotesMadre.has(loteKey)) {
        mapaLotesMadre.set(loteKey, {
          loteKey: loteKey,
          cuit: cuit,
          srt: srt,
          frecuencia: frec,
          totalEstudios: 0,
          auditadosEstudios: 0,
          totalAusentesODefecto: 0,
          totalAlertasMedicas: 0,
          totalNormalODesvio: 0,
          cuils: new Set()
        });
      }

      const lObj = mapaLotesMadre.get(loteKey);
      lObj.totalEstudios++;
      if (c) lObj.cuils.add(c);
      if (res !== '') {
        lObj.auditadosEstudios++;
        const resNorm = _normalizarTexto_(res);
        if (resNorm.includes('ausente') || resNorm.includes('defecto') || resNorm.includes('defectuosa')) {
          lObj.totalAusentesODefecto++;
        } else if (resNorm.includes('enfermedad profesional') || resNorm.includes('pendiente de confirmacion')) {
          lObj.totalAlertasMedicas++;
        } else if (resNorm.includes('normal') || resNorm.includes('desvio') || resNorm.includes('desistido')) {
          lObj.totalNormalODesvio++;
        }
      }
    }

    mapaLotesMadre.forEach(lObj => {
      lObj.es100Auditado = (lObj.totalEstudios > 0 && lObj.auditadosEstudios === lObj.totalEstudios);
      lObj.aptoEmpleador = (lObj.es100Auditado && lObj.totalAusentesODefecto === 0);
    });

    const rawRows = [];
    for (let k = 0; k < allMadreValues.length; k++) {
      const cuilRow = String(allMadreValues[k][CONFIG.COL_CUIL] || "").trim();
      if (cuilRow && cuilsAuditados.has(cuilRow)) {
        rawRows.push({
          row: allMadreValues[k],
          filaMadre: k + 2
        });
      }
    }

    const tz = Session.getScriptTimeZone();
    const mapa = new Map();
    const filasSinFechaAud = [];
    const filasInconsistentesAusentes = [];

    for (let i = 0; i < rawRows.length; i++) {
      const item = rawRows[i];
      const row = item.row;
      const filaMadre = item.filaMadre;
      const cuil = String(row[CONFIG.COL_CUIL] || "").trim();
      if (!cuil) continue;

      const estudio = String(row[CONFIG.COL_ESTUDIO] || "").trim();
      if (estudio.includes("No posee")) continue;

      const asistencia = String(row[CONFIG.COL_ASISTENCIA] || "").toUpperCase().trim();

      const res = String(row[CONFIG.COL_RESULTADO] || "").trim();
      let link1 = String(row[CONFIG.COL_ADJUNTO_1] || "").trim();
      let link2 = String(row[CONFIG.COL_ADJUNTO_2] || "").trim();
      let guardado = String(row[CONFIG.COL_GUARDADO] || "").trim();
      let linkCarpeta = String(row[CONFIG.COL_LINK_CARPETA] || "").trim();

      // Si el estudio es Ausente o asistencia es AUSENTE, no posee links ni debe considerarse guardado
      const esAusente = _normalizarTexto_(res).includes('ausente') || asistencia === 'AUSENTE';
      if (esAusente) {
        if (link1 !== "" || link2 !== "" || guardado.toUpperCase() === 'SÍ' || guardado.toUpperCase() === 'SI' || linkCarpeta !== "") {
          filasInconsistentesAusentes.push(filaMadre);
        }
        link1 = "";
        link2 = "";
        guardado = "NO";
        linkCarpeta = "";
      }
      let fechaAud = _parseFechaConTz_(row[CONFIG.COL_FECHA_AUDITORIA], tz);
      // Auto-completar fecha de auditoría si se cargó resultado directamente en Madre y la fecha quedó vacía
      if (res !== '' && !fechaAud) {
        fechaAud = _parseFechaConTz_(new Date(), tz);
        filasSinFechaAud.push(filaMadre);
      }
      const fechaEst = _parseFechaConTz_(row[CONFIG.COL_FECHA_ESTUDIO], tz);
      const obs = String(row[CONFIG.COL_OBSERVACION] || "").trim();
      const comunicacion = _parseFechaConTz_(row[CONFIG.COL_COMUNICACION], tz);
      const establecimiento = String(row[CONFIG.COL_NOMBRE_ESTABLECIMIENTO] || "").trim();
      const srt = String(row[CONFIG.COL_SRT] || "").trim();
      const prestador = String(row[colPrestadorMadre] || "").trim();
      const frecuencia = String(row[CONFIG.COL_FRECUENCIA] || "").trim();

      // Clave única desde columna AN de Madre (CUIL_CodigoSRT), o calculada si aún no está completa
      const colClaveAN = CONFIG.COL_CLAVE_EMPLEADO !== undefined ? CONFIG.COL_CLAVE_EMPLEADO : 39;
      const claveMadre = String(row[colClaveAN] || "").trim();
      const clave = claveMadre || `${cuil}_${srt || 'SIN_SRT'}`;

      if (!mapa.has(clave)) {
        mapa.set(clave, {
          clave: clave,
          cuil: cuil,
          apeNom: String(row[CONFIG.COL_APELLIDO] || "").trim() || "Sin Nombre",
          contrato: String(row[CONFIG.COL_CUIT_CONTRATO] || "").trim(),
          nombreContrato: String(row[CONFIG.COL_NOMBRE_CONTRATO] || "").trim(),
          srt: srt,
          establecimiento: establecimiento,
          prestador: prestador,
          frecuencia: frecuencia,
          organismo: String(row[CONFIG.COL_NOMBRE_ORGANISMO] || "").trim(),
          totalEstudios: 0,
          auditadosEstudios: 0,
          guardado: 'NO',
          linkCarpeta: '',
          linkAdjunto: '',
          fileId: '',
          fechaAuditoria: '',
          fechaEstudio: '',
          filasIndices: [],
          estudios: []
        });
      }

      const emp = mapa.get(clave);
      emp.totalEstudios++;
      if (prestador && !emp.prestador) emp.prestador = prestador;
      if (frecuencia && !emp.frecuencia) emp.frecuencia = frecuencia;

      if (res !== '') emp.auditadosEstudios++;
      if (_esGuardado_(guardado)) emp.guardado = 'SÍ';
      if (linkCarpeta && !emp.linkCarpeta) emp.linkCarpeta = linkCarpeta;

      const linkUsable = link1 || link2;
      if (linkUsable && !emp.linkAdjunto) {
        emp.linkAdjunto = linkUsable;
        emp.fileId = _extraerFileId_(linkUsable);
      }
      if (fechaAud && !emp.fechaAuditoria) emp.fechaAuditoria = fechaAud;
      if (fechaEst && !emp.fechaEstudio) emp.fechaEstudio = fechaEst;

      emp.filasIndices.push(filaMadre);
      emp.estudios.push({
        rowIndex: filaMadre,
        estudio: estudio,
        prestador: prestador,
        resultado: res,
        observacion: obs,
        fechaEstudio: fechaEst,
        fechaAuditoria: fechaAud,
        adjunto1: link1,
        adjunto2: link2,
        comunicacion: comunicacion,
        guardado: guardado,
        linkCarpeta: linkCarpeta,
        esDesvio: _esResultadoDesvio_(res),
        esReclamable: _esResultadoReclamo_(res)
      });
    }

    // Auto-estampar fecha de hoy en Hoja Madre para los estudios auditados que tenían fecha vacía
    if (filasSinFechaAud.length > 0) {
      try {
        const colFAud = CONFIG.COL_FECHA_AUDITORIA + 1;
        const hoy = new Date();
        let ini = filasSinFechaAud[0];
        let cant = 1;
        for (let s = 1; s < filasSinFechaAud.length; s++) {
          if (filasSinFechaAud[s] === ini + cant) {
            cant++;
          } else {
            sheetMadre.getRange(ini, colFAud, cant, 1).setValues(Array(cant).fill([hoy]));
            ini = filasSinFechaAud[s];
            cant = 1;
          }
        }
        sheetMadre.getRange(ini, colFAud, cant, 1).setValues(Array(cant).fill([hoy]));
        registrarEnBitacora('AUTO_COMPLETAR_FECHA_AUDITORIA', { filas: filasSinFechaAud.length });
      } catch (eAutoFecha) {
        console.warn('Aviso auto-completando fecha de auditoría en Madre: ' + eAutoFecha.message);
      }
    }

    const todosEmp = Array.from(mapa.values());

    // Detección estricta de guardado en Drive: proviene exclusivamente de la Hoja Madre (Columna AL: Guardado)
    todosEmp.forEach(e => {
      const algunEstGuardado = (e.estudios || []).some(est => _esGuardado_(est.guardado));
      if (algunEstGuardado) {
        e.guardado = 'SÍ';
        (e.estudios || []).forEach(est => {
          if (est && !_normalizarTexto_(est.resultado || '').includes('ausente') && _esGuardado_(est.guardado)) {
            est.guardado = 'SÍ';
          } else {
            est.guardado = '';
          }
        });
      } else {
        e.guardado = 'NO';
        (e.estudios || []).forEach(est => {
          est.guardado = '';
        });
      }
    });

    // 1) EMPLEADOS: empleados con estudios auditados, guardados en Drive o con links adjuntos cargados
    let filasEmp = [];
    if (hojaEmp) {
      const empsParaHojaEmp = todosEmp.filter(e => e.auditadosEstudios > 0 || _esGuardado_(e.guardado) || Boolean(e.linkAdjunto) || (e.estudios || []).some(est => est.adjunto1 || est.adjunto2));
      filasEmp = empsParaHojaEmp.map(e => {
        const tieneNoApto = e.estudios.some(est => est.resultado && !_esResultadoApto_(est.resultado));
        let estadoAud = "Sin auditar";
        if (tieneNoApto) {
          estadoAud = "Incompleto";
        } else if (e.totalEstudios > 0 && e.auditadosEstudios === e.totalEstudios) {
          estadoAud = "100% Auditado";
        } else if (e.auditadosEstudios > 0) {
          estadoAud = "Auditoría Parcial";
        }

        const compactos = e.estudios.map(est => ({
          rowIndex: est.rowIndex,
          estudio: est.estudio,
          prestador: est.prestador || e.prestador || '',
          resultado: est.resultado,
          observacion: est.observacion,
          fechaEstudio: est.fechaEstudio,
          fechaAuditoria: est.fechaAuditoria,
          guardado: est.guardado,
          linkCarpeta: est.linkCarpeta,
          adjunto1: est.adjunto1 || '',
          adjunto2: est.adjunto2 || '',
          comunicacion: est.comunicacion || ''
        }));

        return [
          e.cuil, e.apeNom, e.contrato, e.nombreContrato, e.srt, e.establecimiento, e.organismo,
          e.totalEstudios, e.auditadosEstudios, estadoAud, e.guardado, e.linkCarpeta, e.linkAdjunto, e.fileId,
          e.fechaAuditoria, e.fechaEstudio, e.filasIndices.join(","), JSON.stringify(compactos),
          e.prestador || '', e.clave || (e.cuil + '_' + e.srt)
        ];
      });
      _volcarFilasEnHoja_(hojaEmp, filasEmp);
    }

    // 2) DESVIOS: empleados con algún estudio en desvío / confirmación / enf profesional
    const empsParaDesv = todosEmp.filter(e => e.estudios.some(est => est.esDesvio));
    const filasDesv = empsParaDesv.map(e => {
      const estudiosDesv = e.estudios.filter(est => est.esDesvio);
      const tieneEnf = e.estudios.some(est => _normalizarTexto_(est.resultado).includes('enfermedad profesional'));
      const tieneDesv = e.estudios.some(est => _normalizarTexto_(est.resultado).includes('desvio'));
      const estadoGlobal = tieneEnf ? 'Enfermedad Profesional' : (tieneDesv ? 'Desvío' : 'Pendiente de Confirmación');

      const todosCom = e.estudios.every(est => est.comunicacion);
      const fechaCom = todosCom ? e.estudios[0].comunicacion : '';
      const ultFechaAud = e.fechaAuditoria || '';
      const primerAdj = e.estudios.map(est => est.adjunto1 || est.adjunto2).find(Boolean) || '';

      const compactos = e.estudios.map(est => ({
        rowIndex: est.rowIndex,
        estudio: est.estudio,
        prestador: est.prestador || e.prestador || '',
        resultado: est.resultado || 'Pendiente',
        observacion: est.observacion || '',
        fechaEstudio: est.fechaEstudio || '',
        fechaAuditoria: est.fechaAuditoria || '',
        adjunto1: est.adjunto1 || '',
        adjunto2: est.adjunto2 || '',
        comunicacion: est.comunicacion || '',
        esDesvio: est.esDesvio
      }));

      const k = e.clave || (e.cuil + '_' + e.srt);
      const prev = prevDesvMap[k] || {};

      return [
        e.cuil, e.apeNom, e.contrato, e.nombreContrato, e.srt, e.establecimiento, e.organismo,
        estadoGlobal, e.totalEstudios, estudiosDesv.length, ultFechaAud, fechaCom, primerAdj,
        e.filasIndices.join(","), JSON.stringify(compactos),
        e.prestador || '', k,
        prev.prestador2 || '', prev.fechaVE || '',
        prev.resultadoReconfirmacion || '', prev.observacionReconfirmacion || '',
        prev.estudiosAdicionales || ''
      ];
    });

    // 3) RECLAMOS: empleados con algún estudio reclamable (ausente, defecto)
    const empsParaRecl = todosEmp.filter(e => e.estudios.some(est => est.esReclamable));
    const filasRecl = empsParaRecl.map(e => {
      const estudiosRecl = e.estudios.filter(est => est.esReclamable);
      const tieneDefecto = e.estudios.some(est => {
        const s = _normalizarTexto_(est.resultado);
        return s.includes('defectuosa') || s.includes('defecto');
      });
      const tieneAusente = e.estudios.some(est => _normalizarTexto_(est.resultado).includes('ausente'));
      const motivoGlobal = (tieneDefecto && tieneAusente) ? 'Ausente y Defectuoso' :
                           tieneDefecto ? 'Práctica defectuosa' : 'Ausente / Sin realizar';

      const ultFechaAud = e.fechaAuditoria || '';
      const primerAdj = e.estudios.map(est => est.adjunto1 || est.adjunto2).find(Boolean) || '';

      const compactos = e.estudios.map(est => ({
        rowIndex: est.rowIndex,
        estudio: est.estudio,
        prestador: est.prestador || e.prestador || '',
        resultado: est.resultado || 'Pendiente',
        observacion: est.observacion || '',
        fechaEstudio: est.fechaEstudio || '',
        fechaAuditoria: est.fechaAuditoria || '',
        adjunto1: est.adjunto1 || '',
        adjunto2: est.adjunto2 || '',
        esReclamable: est.esReclamable
      }));

      return [
        e.cuil, e.apeNom, e.contrato, e.nombreContrato, e.srt, e.establecimiento, e.organismo,
        motivoGlobal, estudiosRecl.length, e.totalEstudios, ultFechaAud, primerAdj,
        e.filasIndices.join(","), JSON.stringify(compactos),
        e.prestador || '', e.clave || (e.cuil + '_' + e.srt)
      ];
    });

    // 4) COMUNICACIONES: clasificados según la matriz de negocio
    // Lote = CUIT_Contrato + Código_SRT + Frecuencia
    const filasCom = [];
    todosEmp.forEach(e => {
      // Solo evaluar empleados que tengan al menos 1 estudio auditado
      if (e.auditadosEstudios === 0) return;

      const loteKey = `${e.contrato}___${e.srt}___${e.frecuencia || ''}`;
      const lote = mapaLotesMadre.get(loteKey) || {
        totalEstudios: e.totalEstudios,
        auditadosEstudios: e.auditadosEstudios,
        totalAusentesODefecto: 0,
        es100Auditado: (e.auditadosEstudios === e.totalEstudios && e.totalEstudios > 0),
        aptoEmpleador: (e.auditadosEstudios === e.totalEstudios && e.totalEstudios > 0)
      };

      const tieneEnf    = e.estudios.some(est => _normalizarTexto_(est.resultado).includes('enfermedad profesional'));
      const tienePend   = e.estudios.some(est => _normalizarTexto_(est.resultado).includes('pendiente de confirmacion'));
      const tieneDesv   = e.estudios.some(est => _normalizarTexto_(est.resultado).includes('desvio'));
      const tieneDesist = e.estudios.some(est => _normalizarTexto_(est.resultado).includes('desistido'));
      const todosDesist = e.estudios.length > 0 && e.estudios.every(est => _normalizarTexto_(est.resultado).includes('desistido'));
      const tieneRecl   = e.estudios.some(est => est.esReclamable);

      let estadoCons = 'Normal Laboral';
      if (tieneEnf) estadoCons = 'Enfermedad Profesional';
      else if (tienePend) estadoCons = 'Pendiente de Confirmación';
      else if (tieneDesv) estadoCons = 'Desvío';
      else if (todosDesist) estadoCons = 'Desistido Total';
      else if (tieneDesist) estadoCons = 'Normal (c/ Desistido)';
      else if (tieneRecl) estadoCons = 'Retenido (Ausente / Defecto)';

      // Determinar Canal, Destinatario y Aptitud según la Matriz de Negocio
      let canal = '';
      let destinatario = '';
      let esperaLote = 'NO';
      let aptoComunicar = 'NO';
      let estadoLote = '';

      if (tieneEnf || tienePend) {
        // Alerta Médica Crítica -> Inmediato al empleado puntual (NO espera cierre de lote)
        canal = 'ALERTA_EMPLEADO';
        destinatario = 'Empleado puntual';
        esperaLote = 'NO';
        aptoComunicar = 'SÍ';
        estadoLote = 'Inmediato (Alerta médica)';
      } else if (tieneRecl) {
        // Caso retenido -> Requiere gestión (NO se informa)
        canal = 'RETENIDO';
        destinatario = '—';
        esperaLote = 'SÍ';
        aptoComunicar = 'NO';
        estadoLote = 'Retenido, requiere gestión';
      } else {
        // Normal Laboral, Desvío, Desistido -> Empleador (comunicación de lote, espera cierre 100%)
        canal = 'LOTE_EMPLEADOR';
        destinatario = 'Empleador (comunicación de lote)';
        esperaLote = 'SÍ';
        if (lote.aptoEmpleador) {
          aptoComunicar = 'SÍ';
          estadoLote = 'Lote 100% Cerrado (Apto)';
        } else if (!lote.es100Auditado) {
          aptoComunicar = 'NO';
          estadoLote = `En espera (${lote.auditadosEstudios}/${lote.totalEstudios} auditados)`;
        } else {
          aptoComunicar = 'NO';
          estadoLote = `Bloqueado (${lote.totalAusentesODefecto} ausente/defecto retenidos)`;
        }
      }

      const todosCom = e.estudios.every(est => est.comunicacion);
      const fechaCom = todosCom ? e.estudios[0].comunicacion : '';
      const ultFechaAud = e.fechaAuditoria || '';
      const primerAdj = e.estudios.map(est => est.adjunto1 || est.adjunto2).find(Boolean) || '';

      const compactos = e.estudios.map(est => ({
        rowIndex: est.rowIndex,
        estudio: est.estudio,
        prestador: est.prestador || e.prestador || '',
        resultado: est.resultado || '',
        observacion: est.observacion || '',
        fechaEstudio: est.fechaEstudio || '',
        fechaAuditoria: est.fechaAuditoria || '',
        adjunto1: est.adjunto1 || '',
        adjunto2: est.adjunto2 || '',
        comunicacion: est.comunicacion || ''
      }));

      filasCom.push([
        e.cuil, e.apeNom, e.contrato, e.nombreContrato, e.srt, e.establecimiento, e.organismo,
        estadoCons, e.totalEstudios, ultFechaAud, fechaCom, primerAdj,
        e.filasIndices.join(","), JSON.stringify(compactos),
        e.prestador || '', e.clave || (e.cuil + '_' + e.srt),
        e.frecuencia || '', loteKey, canal, destinatario, esperaLote, aptoComunicar, estadoLote
      ]);
    });

    _volcarFilasEnHoja_(hojaDesv, filasDesv);
    _volcarFilasEnHoja_(hojaRecl, filasRecl);
    _volcarFilasEnHoja_(hojaCom, filasCom);

    SpreadsheetApp.flush();
    _invalidarCache_();

    registrarEnBitacora('SINCRONIZACION_HOJAS_COMPLETA', {
      empleados: filasEmp.length,
      desvios: filasDesv.length,
      reclamos: filasRecl.length,
      comunicaciones: filasCom.length
    });

    return {
      status: "OK",
      empleados: filasEmp.length,
      desvios: filasDesv.length,
      reclamos: filasRecl.length,
      comunicaciones: filasCom.length
    };

  } catch (e) {
    console.error("Error en sincronizarTodasLasHojasDesdeMadre: " + e.message);
    registrarEnBitacora('ERROR_SINCRONIZACION_HOJAS', { error: e.message });
    return { status: "Error", msg: e.message };
  } finally {
    if (lock) lock.releaseLock();
  }
}

function _volcarFilasEnHoja_(hoja, filas) {
  if (!hoja) return;
  const ultFila = hoja.getLastRow();
  const currentMaxCols = hoja.getMaxColumns();
  if (ultFila > 1) {
    hoja.getRange(2, 1, ultFila - 1, currentMaxCols).clearContent();
  }
  if (filas && filas.length > 0) {
    const requiredCols = filas[0].length;
    if (currentMaxCols < requiredCols) {
      hoja.insertColumnsAfter(currentMaxCols, requiredCols - currentMaxCols);
    }
    hoja.getRange(2, 1, filas.length, requiredCols).setValues(filas);
  }
}

/**
 * Sincronización continua al guardar auditoría masiva o individual.
 */
function _sincronizarCambiosHojasEspeciales_(sheetMadre, cambios, fechaHoy, tz) {
  try {
    sincronizarTodasLasHojasDesdeMadre(true);
  } catch (err) {
    console.error("Error en _sincronizarCambiosHojasEspeciales_: " + err.message);
  }
}

/**
 * Actualiza la fecha de comunicación en las hojas secundarias.
 */
function _actualizarComunicacionEnHojasSecundarias_(indices, fecha) {
  try {
    const ss = _getSpreadsheet_();
    const setIndices = new Set(indices.map(Number));
    const hojas = [
      { nombre: CONFIG.HOJA_COMUNICACIONES, colCom: 11, colJson: 14, colFilas: 13 },
      { nombre: CONFIG.HOJA_DESVIOS,         colCom: 12, colJson: 15, colFilas: 14 },
      { nombre: CONFIG.HOJA_RECLAMOS,        colCom: null, colJson: 14, colFilas: 13 },
      { nombre: CONFIG.HOJA_EMPLEADOS,       colCom: null, colJson: 18, colFilas: 17 }
    ];

    hojas.forEach(h => {
      const hoja = ss.getSheetByName(h.nombre);
      if (!hoja || hoja.getLastRow() < 2) return;
      const ultFila = hoja.getLastRow();
      const numCols = Math.max(hoja.getLastColumn(), 18);
      const rows = hoja.getRange(2, 1, ultFila - 1, numCols).getValues();

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const filasStr = String(r[h.colFilas - 1] || "");
        if (!filasStr) continue;
        const filasIndices = filasStr.split(',').map(n => parseInt(n.trim(), 10));
        const tieneAlguno = filasIndices.some(idx => setIndices.has(idx));
        if (!tieneAlguno) continue;

        let estudios = [];
        try {
          estudios = typeof r[h.colJson - 1] === 'string' ? JSON.parse(r[h.colJson - 1]) : r[h.colJson - 1];
        } catch (e) { estudios = []; }

        let estudioModificado = false;
        estudios.forEach(est => {
          if (setIndices.has(Number(est.rowIndex))) {
            est.comunicacion = fecha;
            estudioModificado = true;
          }
        });

        if (estudioModificado) {
          r[h.colJson - 1] = JSON.stringify(estudios);
          if (h.colCom) {
            const todosCom = estudios.every(est => est.comunicacion);
            if (todosCom) {
              r[h.colCom - 1] = fecha;
            }
          }
          hoja.getRange(i + 2, 1, 1, r.length).setValues([r]);
        }
      }
    });
  } catch (e) {
    console.error("Error actualizando comunicación en hojas secundarias: " + e.message);
  }
}

// ==================================================================
// COMUNICACIÓN DE RESULTADOS (Lectura directa desde hoja Comunicaciones)
// ==================================================================

/**
 * Acumula y sincroniza exclusivamente la hoja independiente "Comunicaciones".
 * Procesa directamente desde la hoja "Empleados" agrupando por Lote (CUIT/SRT/Frecuencia).
 * Ultrarrápido (<0.2s) sin tocar Hoja Madre.
 */
function sincronizarHojaComunicaciones(skipLock) {
  const lock = skipLock ? null : LockService.getScriptLock();
  try {
    if (lock) lock.waitLock(CONFIG.TIMEOUT_LOCK);

    const ss = _getSpreadsheet_();

    // Búsqueda flexible de hoja Empleados
    let hojaEmp = ss.getSheetByName(CONFIG.HOJA_EMPLEADOS) ||
                  ss.getSheetByName("Empleados") ||
                  ss.getSheetByName("Gestión Empleados") ||
                  ss.getSheetByName("Gestion Empleados");
    if (!hojaEmp) {
      const allSheets = ss.getSheets();
      for (let s of allSheets) {
        const sName = _normalizarTexto_(s.getName()).replace(/[\s_]+/g, '');
        if (sName.includes('emplead')) {
          hojaEmp = s;
          break;
        }
      }
    }

    // Búsqueda flexible o inicialización de hoja Comunicaciones
    let hojaCom = ss.getSheetByName(CONFIG.HOJA_COMUNICACIONES) ||
                  ss.getSheetByName("Comunicaciones") ||
                  ss.getSheetByName("Comunicacion") ||
                  ss.getSheetByName("Comunicar");
    if (!hojaCom) {
      hojaCom = inicializarHojaComunicaciones();
    }

    // Si hoja Empleados no existe o está vacía, no tocar Madre: retornar vacío de inmediato
    if (!hojaEmp || hojaEmp.getLastRow() < 2) {
      _volcarFilasEnHoja_(hojaCom, []);
      return { status: "OK", comunicaciones: 0, msg: "Hoja Empleados sin registros para extraer" };
    }

    const ultFilaEmp = hojaEmp.getLastRow();
    const actualCols = hojaEmp.getLastColumn();
    const headers = hojaEmp.getRange(1, 1, 1, actualCols).getValues()[0].map(h => String(h || '').trim());
    const col = {};
    headers.forEach((h, idx) => {
      if (h) {
        col[h.toLowerCase()] = idx;
        col[_normalizarTexto_(h).replace(/[\s_]+/g, '')] = idx;
      }
    });

    const getV = (r, name, fallbackIdx) => {
      const norm = _normalizarTexto_(name).replace(/[\s_]+/g, '');
      let idx = col[norm];
      if (idx === undefined) idx = col[name.toLowerCase()];
      if (idx === undefined) idx = fallbackIdx;
      return (idx !== undefined && idx < r.length && r[idx] !== undefined && r[idx] !== null) ? r[idx] : '';
    };

    const rows = hojaEmp.getRange(2, 1, ultFilaEmp - 1, actualCols).getValues();
    const tz = Session.getScriptTimeZone();

    // Pase 1: Armar lotes e identificar empleados
    const mapaLotes = new Map();
    const listaEmpleados = [];

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const cuil = String(getV(r, "CUIL", 0)).trim();
      if (!cuil) continue;

      let estudios = [];
      const detRaw = getV(r, "Detalle_Estudios", 17);
      if (detRaw) {
        try {
          estudios = typeof detRaw === 'string' ? JSON.parse(detRaw) : detRaw;
        } catch (e) {
          estudios = [];
        }
      }

      let filasIndices = [];
      const filRaw = getV(r, "Filas_Madre", 16);
      if (filRaw) {
        filasIndices = String(filRaw).split(',').map(n => parseInt(n.trim(), 10)).filter(n => !isNaN(n));
      }

      const totalEstudios = Number(getV(r, "Total_Estudios", 7) || (Array.isArray(estudios) ? estudios.length : 0));
      const auditadosEstudios = Number(getV(r, "Auditados_Estudios", 8) || 0);

      let audCount = auditadosEstudios;
      if (audCount === 0 && Array.isArray(estudios)) {
        audCount = estudios.filter(est => est.resultado && String(est.resultado).trim() !== '').length;
      }

      let prestador = String(getV(r, "Prestador", 18) || '').trim();
      if (!prestador && Array.isArray(estudios) && estudios.length > 0) {
        const pList = estudios.map(est => String(est.prestador || '').trim()).filter(Boolean);
        if (pList.length > 0) {
          prestador = Array.from(new Set(pList)).join(', ');
        }
      }

      const contrato = String(getV(r, "Contrato", 2) || '').trim();
      const nombreContrato = String(getV(r, "Nombre_Contrato", 3) || '').trim();
      const srt = String(getV(r, "Codigo_SRT", 4) || '').trim();
      const establecimiento = String(getV(r, "Establecimiento", 5) || '').trim();
      const organismo = String(getV(r, "Nombre_Organismo", 6) || '').trim();
      const clave = String(getV(r, "Clave_Unica", 19) || (cuil + '_' + (srt || 'SIN_SRT'))).trim();
      const fechaAuditoria = _parseFechaConTz_(getV(r, "Fecha_Auditoria", 14), tz);
      const fechaEstudio = _parseFechaConTz_(getV(r, "Fecha_Estudio", 15), tz);
      const linkAdjunto = String(getV(r, "Link_Adjunto", 12) || '').trim();

      // Frecuencia
      let frecuencia = '';
      if (Array.isArray(estudios) && estudios.length > 0) {
        const f = estudios.map(est => est.frecuencia).find(Boolean);
        if (f) frecuencia = String(f).trim();
      }

      const loteKey = `${contrato}___${srt}___${frecuencia}`;
      if (!mapaLotes.has(loteKey)) {
        mapaLotes.set(loteKey, {
          loteKey: loteKey,
          contrato: contrato,
          srt: srt,
          frecuencia: frecuencia,
          totalEstudios: 0,
          auditadosEstudios: 0,
          totalAusentesODefecto: 0,
          totalAlertasMedicas: 0,
          totalPendientesConfirmacion: 0,
          totalNormalODesvio: 0,
          cuils: new Set()
        });
      }

      const lObj = mapaLotes.get(loteKey);
      lObj.totalEstudios += totalEstudios;
      lObj.auditadosEstudios += audCount;
      lObj.cuils.add(cuil);

      if (Array.isArray(estudios)) {
        estudios.forEach(est => {
          const res = String(est.resultado || '').trim();
          if (res) {
            const resNorm = _normalizarTexto_(res);
            if (resNorm.includes('ausente') || resNorm.includes('defecto') || resNorm.includes('defectuosa')) {
              lObj.totalAusentesODefecto++;
            } else if (resNorm.includes('enfermedad profesional')) {
              lObj.totalAlertasMedicas++;
            } else if (resNorm.includes('pendiente de confirmacion')) {
              lObj.totalPendientesConfirmacion = (lObj.totalPendientesConfirmacion || 0) + 1;
              lObj.totalAlertasMedicas++;
            } else if (resNorm.includes('normal') || resNorm.includes('desvio') || resNorm.includes('desistido')) {
              lObj.totalNormalODesvio++;
            }
          }
        });
      }

      listaEmpleados.push({
        cuil: cuil,
        apeNom: String(getV(r, "Apellido_y_Nombre", 1) || '').trim(),
        contrato: contrato,
        nombreContrato: nombreContrato,
        srt: srt,
        establecimiento: establecimiento,
        organismo: organismo,
        totalEstudios: totalEstudios,
        auditadosEstudios: audCount,
        fechaAuditoria: fechaAuditoria,
        fechaEstudio: fechaEstudio,
        linkAdjunto: linkAdjunto,
        filasIndices: filasIndices,
        estudios: estudios,
        prestador: prestador,
        clave: clave,
        frecuencia: frecuencia,
        loteKey: loteKey
      });
    }

    mapaLotes.forEach(lObj => {
      lObj.es100Auditado = (lObj.totalEstudios > 0 && lObj.auditadosEstudios >= lObj.totalEstudios);
      lObj.tienePendientesConfirmacion = ((lObj.totalPendientesConfirmacion || 0) > 0);
      lObj.aptoEmpleador = (lObj.es100Auditado && lObj.totalAusentesODefecto === 0 && !lObj.tienePendientesConfirmacion);
    });

    // Pase 2: Construir las 23 columnas exactas de Comunicaciones
    const filasCom = [];

    listaEmpleados.forEach(e => {
      if (e.auditadosEstudios === 0) return;

      const lote = mapaLotes.get(e.loteKey) || {
        totalEstudios: e.totalEstudios,
        auditadosEstudios: e.auditadosEstudios,
        totalAusentesODefecto: 0,
        totalPendientesConfirmacion: 0,
        es100Auditado: (e.auditadosEstudios >= e.totalEstudios && e.totalEstudios > 0),
        aptoEmpleador: (e.auditadosEstudios >= e.totalEstudios && e.totalEstudios > 0)
      };

      const tieneEnf    = e.estudios.some(est => _normalizarTexto_(est.resultado).includes('enfermedad profesional'));
      const tienePend   = e.estudios.some(est => _normalizarTexto_(est.resultado).includes('pendiente de confirmacion'));
      const tieneDesv   = e.estudios.some(est => _normalizarTexto_(est.resultado).includes('desvio'));
      const tieneDesist = e.estudios.some(est => _normalizarTexto_(est.resultado).includes('desistido'));
      const todosDesist = e.estudios.length > 0 && e.estudios.every(est => _normalizarTexto_(est.resultado).includes('desistido'));
      const tieneRecl   = e.estudios.some(est => _esResultadoReclamo_(est.resultado));

      let estadoCons = 'Normal Laboral';
      if (tieneEnf) estadoCons = 'Enfermedad Profesional';
      else if (tienePend) estadoCons = 'Pendiente de Confirmación';
      else if (tieneDesv) estadoCons = 'Desvío';
      else if (todosDesist) estadoCons = 'Desistido Total';
      else if (tieneDesist) estadoCons = 'Normal (c/ Desistido)';
      else if (tieneRecl) estadoCons = 'Retenido (Ausente / Defecto)';

      let canal = '';
      let destinatario = '';
      let esperaLote = 'NO';
      let aptoComunicar = 'NO';
      let estadoLote = '';

      if (tieneEnf) {
        canal = 'ALERTA_EMPLEADO';
        destinatario = 'Empleado puntual';
        esperaLote = 'NO';
        aptoComunicar = 'SÍ';
        estadoLote = 'Inmediato (Alerta médica)';
      } else if (tienePend) {
        canal = 'ALERTA_EMPLEADO';
        destinatario = 'Reconfirmación médica prioritaria';
        esperaLote = 'SÍ';
        aptoComunicar = 'NO'; // Prioridad reconfirmación médica: no emitir hasta confirmar
        estadoLote = 'Prioridad: Requiere confirmación médica';
      } else if (tieneRecl) {
        canal = 'RETENIDO';
        destinatario = '—';
        esperaLote = 'SÍ';
        aptoComunicar = 'NO';
        estadoLote = 'Retenido, requiere gestión';
      } else {
        canal = 'LOTE_EMPLEADOR';
        destinatario = 'Empleador (comunicación de lote)';
        esperaLote = 'SÍ';
        if (lote.aptoEmpleador) {
          aptoComunicar = 'SÍ';
          estadoLote = 'Lote 100% Cerrado (Apto)';
        } else if (lote.tienePendientesConfirmacion) {
          aptoComunicar = 'NO';
          estadoLote = `Bloqueado (${lote.totalPendientesConfirmacion} pendiente(s) de confirmación)`;
        } else if (!lote.es100Auditado) {
          aptoComunicar = 'NO';
          estadoLote = `En espera (${lote.auditadosEstudios}/${lote.totalEstudios} auditados)`;
        } else {
          aptoComunicar = 'NO';
          estadoLote = `Bloqueado (${lote.totalAusentesODefecto} ausente/defecto retenidos)`;
        }
      }

      const todosCom = e.estudios.length > 0 && e.estudios.every(est => est.comunicacion);
      const fechaCom = todosCom ? _parseFechaConTz_(e.estudios[0].comunicacion, tz) : '';
      const ultFechaAud = e.fechaAuditoria || '';
      const primerAdj = e.linkAdjunto || (e.estudios.map(est => est.adjunto1 || est.adjunto2).find(Boolean) || '');

      const compactos = e.estudios.map(est => ({
        rowIndex: est.rowIndex !== undefined ? Number(est.rowIndex) : '',
        estudio: String(est.estudio || '').trim(),
        prestador: String(est.prestador || e.prestador || '').trim(),
        resultado: String(est.resultado || '').trim(),
        observacion: String(est.observacion || '').trim(),
        fechaEstudio: _parseFechaConTz_(est.fechaEstudio, tz),
        fechaAuditoria: _parseFechaConTz_(est.fechaAuditoria, tz),
        adjunto1: String(est.adjunto1 || '').trim(),
        adjunto2: String(est.adjunto2 || '').trim(),
        comunicacion: _parseFechaConTz_(est.comunicacion, tz)
      }));

      // 23 columnas exactas según estructura de negocio
      filasCom.push([
        e.cuil,                                           // 1: CUIL
        e.apeNom,                                         // 2: Apellido_y_Nombre
        e.contrato,                                       // 3: Contrato
        e.nombreContrato,                                 // 4: Nombre_Contrato
        e.srt,                                            // 5: Codigo_SRT
        e.establecimiento,                                // 6: Establecimiento
        e.organismo,                                      // 7: Nombre_Organismo
        estadoCons,                                       // 8: Estado_Consolidado
        e.totalEstudios,                                  // 9: Total_Estudios
        ultFechaAud,                                      // 10: Fecha_Auditoria
        fechaCom,                                         // 11: Comunicacion
        primerAdj,                                        // 12: Adjunto1
        e.filasIndices.join(","),                         // 13: Filas_Madre
        JSON.stringify(compactos),                        // 14: Detalle_Estudios
        e.prestador || '',                                // 15: Prestador
        e.clave || (e.cuil + '_' + e.srt),                // 16: Clave_Unica
        e.frecuencia || '',                               // 17: Frecuencia
        e.loteKey,                                        // 18: Lote_Key
        canal,                                            // 19: Canal_Comunicacion
        destinatario,                                     // 20: Destinatario
        esperaLote,                                       // 21: Espera_Cierre
        aptoComunicar,                                    // 22: Apto_Comunicar
        estadoLote,                                       // 23: Estado_Lote
        e.fechaComunicacionLote || ''                     // 24: Fecha_Comunicacion_Lote
      ]);
    });

    _volcarFilasEnHoja_(hojaCom, filasCom);
    SpreadsheetApp.flush();
    _invalidarCache_();
    return { status: "OK", comunicaciones: filasCom.length };
  } catch (e) {
    console.error("Error en sincronizarHojaComunicaciones: " + e.message);
    return { status: "Error", msg: e.message };
  } finally {
    if (lock) lock.releaseLock();
  }
}

function obtenerRegistrosComunicables() {
  try {
    const ss = _getSpreadsheet_();

    // Búsqueda flexible de la hoja Comunicaciones
    let hoja = ss.getSheetByName(CONFIG.HOJA_COMUNICACIONES) ||
               ss.getSheetByName("Comunicaciones") ||
               ss.getSheetByName("Comunicacion") ||
               ss.getSheetByName("Comunicar");
    if (!hoja) {
      const allSheets = ss.getSheets();
      for (let s of allSheets) {
        const sName = _normalizarTexto_(s.getName()).replace(/[\s_]+/g, '');
        if (sName.includes('comunic')) {
          hoja = s;
          break;
        }
      }
    }

    // Si la hoja no existe o está vacía, hacer la sincronización rápida acumulando desde Empleados
    if (!hoja || hoja.getLastRow() < 2) {
      if (!hoja) hoja = inicializarHojaComunicaciones();
      sincronizarHojaComunicaciones(true);
      hoja = ss.getSheetByName(CONFIG.HOJA_COMUNICACIONES) ||
             ss.getSheetByName("Comunicaciones") ||
             ss.getSheetByName("Comunicacion") ||
             ss.getSheetByName("Comunicar");
    }

    if (!hoja || hoja.getLastRow() < 2) return [];

    const ultFila = hoja.getLastRow();
    const numCols = Math.max(hoja.getLastColumn(), 23);
    const headers = hoja.getRange(1, 1, 1, numCols).getValues()[0].map(h => String(h || '').trim());
    const colMap = {};
    headers.forEach((h, idx) => {
      if (h) {
        const raw = String(h).trim();
        colMap[raw] = idx;
        colMap[raw.toLowerCase()] = idx;
        colMap[_normalizarTexto_(raw).replace(/[\s_]+/g, '')] = idx;
      }
    });

    const getVal = (r, key, defIdx) => {
      const normKey = _normalizarTexto_(key).replace(/[\s_]+/g, '');
      let idx = colMap[normKey];
      if (idx === undefined) idx = colMap[key.toLowerCase()];
      if (idx === undefined) idx = colMap[key];
      if (idx === undefined) idx = defIdx;
      return (idx !== undefined && idx < r.length && r[idx] !== undefined && r[idx] !== null) ? r[idx] : '';
    };

    const rows = hoja.getRange(2, 1, ultFila - 1, numCols).getValues();
    const tz = Session.getScriptTimeZone();

    const resultadoEmpleados = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];

      // Mapeo robusto: 23 columnas exactas compartidas por el usuario
      const cuil = String(getVal(r, 'CUIL', 0) || '').trim();
      if (!cuil) continue;

      const empleado = String(getVal(r, 'Apellido_y_Nombre', 1) || '').trim();
      const contrato = String(getVal(r, 'Contrato', 2) || '').trim();
      const nombreContrato = String(getVal(r, 'Nombre_Contrato', 3) || '').trim();
      const srt = String(getVal(r, 'Codigo_SRT', 4) || '').trim();
      const establecimiento = String(getVal(r, 'Establecimiento', 5) || '').trim();
      const organismo = String(getVal(r, 'Nombre_Organismo', 6) || '').trim();
      const estadoCons = String(getVal(r, 'Estado_Consolidado', 7) || '').trim();
      const totalEstudios = Number(getVal(r, 'Total_Estudios', 8) || 0);

      const fechaAuditoria = _parseFechaConTz_(getVal(r, 'Fecha_Auditoria', 9), tz);
      const comunicacion = _parseFechaConTz_(getVal(r, 'Comunicacion', 10), tz);
      const adjunto1 = String(getVal(r, 'Adjunto1', 11) || '').trim();

      const filasRaw = getVal(r, 'Filas_Madre', 12);
      let filasMadre = [];
      if (filasRaw) {
        filasMadre = String(filasRaw).split(',').map(n => parseInt(n.trim(), 10)).filter(n => !isNaN(n));
      }

      const jsonRaw = getVal(r, 'Detalle_Estudios', 13);
      let estudios = [];
      if (jsonRaw) {
        try {
          estudios = typeof jsonRaw === 'string' ? JSON.parse(jsonRaw) : jsonRaw;
        } catch (e) {
          estudios = [];
        }
      }

      const cleanEstudios = (Array.isArray(estudios) ? estudios : []).map(est => ({
        rowIndex: est.rowIndex !== undefined ? Number(est.rowIndex) : '',
        estudio: String(est.estudio || '').trim(),
        prestador: String(est.prestador || '').trim(),
        resultado: String(est.resultado || '').trim(),
        observacion: String(est.observacion || '').trim(),
        fechaEstudio: _parseFechaConTz_(est.fechaEstudio, tz),
        fechaAuditoria: _parseFechaConTz_(est.fechaAuditoria, tz),
        adjunto1: String(est.adjunto1 || '').trim(),
        adjunto2: String(est.adjunto2 || '').trim(),
        comunicacion: _parseFechaConTz_(est.comunicacion, tz)
      }));

      const prestador = String(getVal(r, 'Prestador', 14) || '').trim();
      const clave = String(getVal(r, 'Clave_Unica', 15) || (cuil + '___' + (srt || 'SIN_SRT'))).trim();
      const frecuencia = String(getVal(r, 'Frecuencia', 16) || '').trim();
      const loteKey = String(getVal(r, 'Lote_Key', 17) || `${contrato}___${srt}___${frecuencia}`).trim();

      let canal = String(getVal(r, 'Canal_Comunicacion', 18) || '').trim();
      let destinatario = String(getVal(r, 'Destinatario', 19) || '').trim();
      let esperaCierre = String(getVal(r, 'Espera_Cierre', 20) || '').trim();
      let aptoComunicar = String(getVal(r, 'Apto_Comunicar', 21) || '').trim();
      let estadoLote = String(getVal(r, 'Estado_Lote', 22) || '').trim();

      // Fallback dinámico si la hoja fue generada antes de las columnas de canal/lote
      if (!canal || !aptoComunicar) {
        const tieneEnf = estadoCons.includes('Enfermedad Profesional') ||
                         cleanEstudios.some(e => _normalizarTexto_(e.resultado).includes('enfermedad profesional'));
        const tienePend = estadoCons.includes('Pendiente de Confirmación') ||
                          cleanEstudios.some(e => _normalizarTexto_(e.resultado).includes('pendiente de confirmacion'));
        const tieneRecl = estadoCons.includes('Retenido') ||
                          cleanEstudios.some(e => _esResultadoReclamo_(e.resultado));

        if (tieneEnf || tienePend) {
          canal = 'ALERTA_EMPLEADO';
          destinatario = 'Empleado puntual';
          esperaCierre = 'NO';
          aptoComunicar = 'SÍ';
          estadoLote = 'Inmediato (Alerta médica)';
        } else if (tieneRecl) {
          canal = 'RETENIDO';
          destinatario = '—';
          esperaCierre = 'SÍ';
          aptoComunicar = 'NO';
          estadoLote = 'Retenido, requiere gestión';
        } else {
          canal = 'LOTE_EMPLEADOR';
          destinatario = 'Empleador (comunicación de lote)';
          esperaCierre = 'SÍ';
          aptoComunicar = 'SÍ';
          estadoLote = 'Lote Cerrado (Apto)';
        }
      }

      resultadoEmpleados.push({
        cuil:              cuil,
        empleado:          empleado,
        contrato:          contrato,
        nombreContrato:    nombreContrato,
        srt:               srt,
        establecimiento:   establecimiento,
        organismo:         organismo,
        estadoConsolidado: estadoCons,
        totalEstudios:     totalEstudios,
        fechaAuditoria:    fechaAuditoria,
        comunicacion:      comunicacion,
        adjunto1:          adjunto1,
        filasMadre:        filasMadre,
        rowIndices:        filasMadre,
        estudios:          [], // Carga ultraligera; los detalles individuales se abren bajo demanda
        prestador:         prestador,
        clave:             clave,
        frecuencia:        frecuencia,
        loteKey:           loteKey,
        canal:             canal,
        destinatario:      destinatario,
        esperaCierre:      esperaCierre,
        aptoComunicar:     aptoComunicar,
        estadoLote:        estadoLote,
        fechaComunicacionLote: _parseFechaConTz_(getVal(r, 'Fecha_Comunicacion_Lote', 23), tz),
        fechaVEReconf:         _parseFechaConTz_(getVal(r, 'Fecha_VE_Reconfirmacion', 24), tz)
      });
    }

    // Garantizar prioridad de 'Pendiente de Confirmación':
    // Si un lote contiene algún resultado pendiente de confirmación médica, bloquea el lote completo para el empleador
    const lotesConPendiente = new Set();
    resultadoEmpleados.forEach(r => {
      const norm = _normalizarTexto_(r.estadoConsolidado || '');
      if (norm.includes('pendiente de confirmacion') || (r.canal === 'ALERTA_EMPLEADO' && norm.includes('pendiente'))) {
        lotesConPendiente.add(r.loteKey);
      }
    });

    if (lotesConPendiente.size > 0) {
      resultadoEmpleados.forEach(r => {
        if (lotesConPendiente.has(r.loteKey)) {
          if (r.canal === 'LOTE_EMPLEADOR') {
            r.aptoComunicar = 'NO';
            r.estadoLote = 'Bloqueado por Pendiente de Confirmación';
          } else if (_normalizarTexto_(r.estadoConsolidado || '').includes('pendiente de confirmacion')) {
            r.aptoComunicar = 'NO';
            r.estadoLote = 'Prioridad: Requiere confirmación médica';
          }
        }
      });
    }

    return resultadoEmpleados;
  } catch (e) {
    console.error("Error en obtenerRegistrosComunicables: " + e.message);
    return [];
  }
}

/**
 * Retorna los estudios individuales de un empleado específico bajo demanda.
 * Lectura ultrarrápida (<0.1s) directamente de la fila en hoja Comunicaciones.
 */
function obtenerEstudiosEmpleado(cuil) {
  try {
    if (!cuil) return [];
    const ss = _getSpreadsheet_();
    let hoja = ss.getSheetByName(CONFIG.HOJA_COMUNICACIONES) || ss.getSheetByName("Comunicaciones");
    if (!hoja || hoja.getLastRow() < 2) return [];

    const strCuil = String(cuil).trim();
    const ultFila = hoja.getLastRow();
    const cuils = hoja.getRange(2, 1, ultFila - 1, 1).getValues();
    let filaTarget = -1;
    for (let i = 0; i < cuils.length; i++) {
      if (String(cuils[i][0]).trim() === strCuil) {
        filaTarget = i + 2;
        break;
      }
    }
    if (filaTarget === -1) return [];

    const rawJson = hoja.getRange(filaTarget, 14).getValue(); // Col 14: Detalle_Estudios
    if (!rawJson) return [];
    const parsed = typeof rawJson === 'string' ? JSON.parse(rawJson) : rawJson;
    const tz = Session.getScriptTimeZone();
    return (Array.isArray(parsed) ? parsed : []).map(est => ({
      rowIndex: est.rowIndex !== undefined ? Number(est.rowIndex) : '',
      estudio: String(est.estudio || '').trim(),
      prestador: String(est.prestador || '').trim(),
      resultado: String(est.resultado || '').trim(),
      observacion: String(est.observacion || '').trim(),
      fechaEstudio: _parseFechaConTz_(est.fechaEstudio, tz),
      fechaAuditoria: _parseFechaConTz_(est.fechaAuditoria, tz),
      adjunto1: String(est.adjunto1 || '').trim(),
      adjunto2: String(est.adjunto2 || '').trim(),
      comunicacion: _parseFechaConTz_(est.comunicacion, tz)
    }));
  } catch (e) {
    console.error("Error en obtenerEstudiosEmpleado: " + e.message);
    return [];
  }
}

// ==================================================================
// DESVÍOS Y CONFIRMACIONES (Lectura directa desde hoja Desvios)
// ==================================================================
function obtenerRegistrosDesvios() {
  try {
    const ss = _getSpreadsheet_();
    const hoja = ss.getSheetByName(CONFIG.HOJA_DESVIOS);
    if (!hoja || hoja.getLastRow() < 2) return [];

    const ultFila = hoja.getLastRow();
    const numCols = hoja.getLastColumn();
    const headers = hoja.getRange(1, 1, 1, numCols).getValues()[0].map(h => String(h || '').trim());

    // Mapeo dinámico Nombre_Columna -> Índice (0-based)
    const col = {};
    headers.forEach((nombre, idx) => {
      if (nombre) col[nombre] = idx;
    });

    // Asegurar que existan las columnas de reconfirmación en el encabezado
    const asegurarEncabezado = (nombreCol) => {
      if (col[nombreCol] === undefined) {
        const nextCol = hoja.getLastColumn() + 1;
        hoja.getRange(1, nextCol).setValue(nombreCol).setFontWeight("bold").setBackground("#d97706").setFontColor("#ffffff");
        col[nombreCol] = nextCol - 1;
      }
    };

    asegurarEncabezado("Prestador2");
    asegurarEncabezado("Fecha_VE");
    asegurarEncabezado("Resultado_Reconfirmacion");
    asegurarEncabezado("Observacion_Reconfirmacion");
    asegurarEncabezado("Estudios_Adicionales");

    // Conteo de empleados auditados totales desde hoja Empleados
    const hojaEmp = ss.getSheetByName(CONFIG.HOJA_EMPLEADOS);
    const totalAuditados = (hojaEmp && hojaEmp.getLastRow() > 1) ? (hojaEmp.getLastRow() - 1) : 0;

    const actualCols = hoja.getLastColumn();
    const rows = hoja.getRange(2, 1, ultFila - 1, actualCols).getValues();

    const getV = (r, name, def = '') => {
      const i = col[name];
      return (i !== undefined && r[i] !== undefined && r[i] !== null) ? r[i] : def;
    };

    const normalizarFechaVal = (val) => {
      if (!val) return '';
      if (val instanceof Date) {
        if (isNaN(val.getTime())) return '';
        try {
          return Utilities.formatDate(val, Session.getScriptTimeZone() || 'America/Argentina/Buenos_Aires', 'dd-MM-yyyy');
        } catch (eFmt) {
          const d = String(val.getDate()).padStart(2, '0');
          const m = String(val.getMonth() + 1).padStart(2, '0');
          const y = val.getFullYear();
          return `${d}-${m}-${y}`;
        }
      }
      const s = String(val).trim();
      if (!s || s.toLowerCase() === 'mon' || s.toLowerCase() === 'invalid date') return '';
      return s;
    };

    const resultadoEmpleados = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const cuil = String(getV(r, "CUIL")).trim();
      if (!cuil) continue;

      let estudios = [];
      const detRaw = getV(r, "Detalle_Estudios");
      if (detRaw) {
        try {
          estudios = typeof detRaw === 'string' ? JSON.parse(detRaw) : detRaw;
        } catch (e) {
          estudios = [];
        }
      }

      let filasMadre = [];
      const filasRaw = getV(r, "Filas_Madre");
      if (filasRaw) {
        filasMadre = String(filasRaw).split(',').map(n => parseInt(n.trim(), 10)).filter(n => !isNaN(n));
      }

      resultadoEmpleados.push({
        cuil:                      cuil,
        empleado:                  String(getV(r, "Apellido_y_Nombre")),
        contrato:                  String(getV(r, "Contrato")),
        nombreContrato:            String(getV(r, "Nombre_Contrato")),
        srt:                       String(getV(r, "Codigo_SRT")),
        establecimiento:           String(getV(r, "Establecimiento")),
        organismo:                 String(getV(r, "Nombre_Organismo")),
        estadoGlobal:              String(getV(r, "Estado_Global")),
        totalEstudios:             Number(getV(r, "Total_Estudios", 0)),
        totalDesvios:              Number(getV(r, "Total_Desvios", 0)),
        fechaAuditoria:            normalizarFechaVal(getV(r, "Fecha_Auditoria")),
        comunicacion:              normalizarFechaVal(getV(r, "Comunicacion")),
        adjunto1:                  String(getV(r, "Adjunto1")),
        filasMadre:                filasMadre,
        rowIndices:                filasMadre,
        estudios:                  estudios,
        prestador:                 String(getV(r, "Prestador")),
        clave:                     String(getV(r, "Clave_Unica", cuil + '_' + getV(r, "Codigo_SRT"))),
        prestador2:                String(getV(r, "Prestador2")),
        fechaVE:                   normalizarFechaVal(getV(r, "Fecha_VE")),
        resultadoReconfirmacion:   String(getV(r, "Resultado_Reconfirmacion")),
        observacionReconfirmacion: String(getV(r, "Observacion_Reconfirmacion")),
        estudiosAdicionales:       String(getV(r, "Estudios_Adicionales", "")),
        totalAuditados:            totalAuditados
      });
    }

    return resultadoEmpleados;
  } catch (e) {
    console.error("Error en obtenerRegistrosDesvios: " + e.message);
    return [];
  }
}

/**
 * Inicializa y devuelve la hoja de mapeo 'Estudios_Adicionales'.
 * Si no existe, la crea con las columnas y valores solicitados.
 */
function inicializarHojaEstudiosAdicionales() {
  const ss = _getSpreadsheet_();
  let hoja = ss.getSheetByName(CONFIG.HOJA_ESTUDIOS_ADICIONALES || "Estudios_Adicionales");
  if (!hoja) {
    hoja = ss.insertSheet(CONFIG.HOJA_ESTUDIOS_ADICIONALES || "Estudios_Adicionales");
  }
  const headers = ["Estudio_Original", "Causa_Exposicion", "Estudios_Adicionales_Sugeridos", "Activo"];

  const filasBase = [
    ["Orina Completa", "Exposición a tóxicos o daño renal", "Ecografía renal y vesical, Urocultivo, Clearance de creatinina, Citología urinaria, Dosaje de biomarcadores específicos en sangre u orina", "SÍ"],
    ["Foniátrico", "Sobrecarga del uso de la voz", "Fibrolaringoscopia, Laringoestroboscopia", "SÍ"],
    ["Examen Clínico con Orientación a Várices", "Bipedestación prolongada", "Eco Doppler venoso de miembros inferiores", "SÍ"],
    ["Espirometría", "Exposición a polvos, humos, asbesto, sílice", "Radiografía de tórax con técnica OIT, TAC de tórax de alta resolución, Prueba con broncodilatador, Pletismografía, Test de difusión de monóxido de carbono (DLCO)", "SÍ"],
    ["Audiometría", "Exposición a ruido", "Audiometría tonal con 48 hs de reposo auditivo, Logoaudiometría, Impedanciometría, Otoemisiones acústicas, Potenciales evocados auditivos (BERA)", "SÍ"]
  ];

  if (hoja.getLastRow() === 0 || hoja.getRange(1, 1).getValue() !== "Estudio_Original") {
    hoja.clear();
    hoja.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = hoja.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#0f766e"); // Teal elegante
    headerRange.setFontColor("#ffffff");
    hoja.setFrozenRows(1);
    hoja.getRange(2, 1, filasBase.length, headers.length).setValues(filasBase);
    try {
      hoja.setColumnWidth(1, 230); // Estudio_Original
      hoja.setColumnWidth(2, 280); // Causa_Exposicion
      hoja.setColumnWidth(3, 550); // Estudios_Adicionales_Sugeridos
      hoja.setColumnWidth(4, 90);  // Activo
    } catch(eW) {}
  }
  return hoja;
}

/**
 * Obtiene el diccionario normalizado de estudios adicionales sugeridos desde Sheets.
 * Lee DIRECTAMENTE de la hoja 'Estudios_Adicionales' para reflejar cambios del usuario en tiempo real.
 */
function obtenerMapeoEstudiosAdicionales() {
  try {
    const hoja = inicializarHojaEstudiosAdicionales();
    const ultFila = hoja.getLastRow();
    if (ultFila < 2) return _mapeoEstudiosAdicionalesFallback_();
    const cols = hoja.getLastColumn();
    const headers = hoja.getRange(1, 1, 1, cols).getValues()[0].map(h => String(h || '').trim().toLowerCase());
    const cEst = headers.indexOf("estudio_original");
    const cAdic = headers.indexOf("estudios_adicionales_sugeridos");
    const cAct = headers.indexOf("activo");

    const colEstIdx = cEst !== -1 ? cEst : 0;
    const colAdicIdx = cAdic !== -1 ? cAdic : (cols >= 4 ? 2 : 1);
    const colActIdx = cAct !== -1 ? cAct : (cols - 1);

    const values = hoja.getRange(2, 1, ultFila - 1, cols).getValues();
    const mapa = {};

    values.forEach((r) => {
      const estOrig = String(r[colEstIdx] || '').trim();
      let adicStr = String(r[colAdicIdx] || '').trim();
      const activo = String(r[colActIdx] || 'SÍ').trim().toUpperCase();

      if (estOrig) {
        const normKey = (typeof _normalizarTexto_ === 'function')
          ? _normalizarTexto_(estOrig)
          : estOrig.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();

        if (activo === 'SÍ' || activo === 'SI' || activo === 'TRUE') {
          const arr = adicStr.split(',').map(s => s.trim()).filter(Boolean);
          mapa[normKey] = arr;
        }
      }
    });

    return mapa;
  } catch(e) {
    console.error("Error en obtenerMapeoEstudiosAdicionales: " + e.message);
    return _mapeoEstudiosAdicionalesFallback_();
  }
}

/**
 * Obtiene la lista oficial de estudios periódicos desde la hoja Estudios de la Hoja Madre.
 * Se utiliza para mantener la nomenclatura estándar en el catálogo ABM y en los desplegables.
 */
function obtenerNombresEstudios() {
  try {
    const ss = _getSpreadsheet_();
    let lista = _leerColumnaSencilla_(ss, CONFIG.HOJA_ESTUDIOS, "A2:A");
    if (!lista || lista.length === 0) {
      const hoja = ss.getSheetByName(CONFIG.HOJA_ESTUDIOS);
      if (hoja && hoja.getLastRow() >= 2) {
        lista = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues().map(r => String(r[0] || '').trim()).filter(Boolean);
      }
    }
    return Array.from(new Set(lista)).filter(Boolean).sort();
  } catch (e) {
    console.error("Error en obtenerNombresEstudios: " + e.message);
    return [];
  }
}

/**
 * Devuelve la lista completa de estudios adicionales para el ABM (Alta, Baja, Modificación).
 */
function obtenerEstudiosAdicionalesABM() {
  try {
    const hoja = inicializarHojaEstudiosAdicionales();
    const ultFila = hoja.getLastRow();
    const cols = Math.max(hoja.getLastColumn(), 4);
    if (ultFila < 2) {
      return { status: "OK", items: [] };
    }
    const headers = hoja.getRange(1, 1, 1, cols).getValues()[0].map(h => String(h || '').trim().toLowerCase());
    const cEst = headers.indexOf("estudio_original");
    const cCausa = headers.indexOf("causa_exposicion");
    const cAdic = headers.indexOf("estudios_adicionales_sugeridos");
    const cAct = headers.indexOf("activo");

    const colEstIdx = cEst !== -1 ? cEst : 0;
    const colCausaIdx = cCausa !== -1 ? cCausa : 1;
    const colAdicIdx = cAdic !== -1 ? cAdic : 2;
    const colActIdx = cAct !== -1 ? cAct : 3;

    const values = hoja.getRange(2, 1, ultFila - 1, cols).getValues();
    const items = [];

    values.forEach((r, idx) => {
      const estOrig = String(r[colEstIdx] || '').trim();
      const causa = String(r[colCausaIdx] || '').trim();
      const adicStr = String(r[colAdicIdx] || '').trim();
      const act = String(r[colActIdx] || 'SÍ').trim();

      if (estOrig || adicStr) {
        const arr = adicStr.split(',').map(s => s.trim()).filter(Boolean);
        items.push({
          fila: idx + 2, // Fila física en Sheets (1-based)
          estudioOriginal: estOrig,
          causaExposicion: causa,
          estudiosSugeridos: adicStr,
          listaAdicionales: arr,
          activo: act.toUpperCase() === 'NO' ? 'NO' : 'SÍ'
        });
      }
    });

    return { status: "OK", items: items };
  } catch (err) {
    console.error("Error en obtenerEstudiosAdicionalesABM: " + err.message);
    return { status: "ERROR", msg: err.message, items: [] };
  }
}

/**
 * Guarda (crea o modifica) un registro del catálogo de estudios adicionales en Sheets.
 */
function guardarEstudioAdicionalABM(item) {
  try {
    const hoja = inicializarHojaEstudiosAdicionales();
    const cols = Math.max(hoja.getLastColumn(), 4);
    const headers = hoja.getRange(1, 1, 1, cols).getValues()[0].map(h => String(h || '').trim().toLowerCase());
    
    let cEst = headers.indexOf("estudio_original") + 1;
    let cCausa = headers.indexOf("causa_exposicion") + 1;
    let cAdic = headers.indexOf("estudios_adicionales_sugeridos") + 1;
    let cAct = headers.indexOf("activo") + 1;

    if (!cEst) cEst = 1;
    if (!cCausa) cCausa = 2;
    if (!cAdic) cAdic = 3;
    if (!cAct) cAct = 4;

    const estOrig = String(item.estudioOriginal || '').trim();
    const causa = String(item.causaExposicion || '').trim();
    let adicStr = '';
    if (Array.isArray(item.listaAdicionales)) {
      adicStr = item.listaAdicionales.map(s => String(s).trim()).filter(Boolean).join(', ');
    } else {
      adicStr = String(item.estudiosSugeridos || '').trim();
    }
    const activo = (String(item.activo || 'SÍ').trim().toUpperCase() === 'NO') ? 'NO' : 'SÍ';

    if (!estOrig) {
      return { status: "ERROR", msg: "El nombre del estudio periódico no puede estar vacío." };
    }

    const fila = parseInt(item.fila, 10);
    if (fila && fila >= 2 && fila <= hoja.getLastRow()) {
      // Modificación
      hoja.getRange(fila, cEst).setValue(estOrig);
      hoja.getRange(fila, cCausa).setValue(causa);
      hoja.getRange(fila, cAdic).setValue(adicStr);
      hoja.getRange(fila, cAct).setValue(activo);
    } else {
      // Alta nueva: verificar si ya existe este estudio_original en el catálogo
      const ult = hoja.getLastRow();
      let filaExistente = 0;
      if (ult >= 2) {
        const estExistentes = hoja.getRange(2, cEst, ult - 1, 1).getValues();
        const estOrigNorm = estOrig.toLowerCase();
        for (let idx = 0; idx < estExistentes.length; idx++) {
          if (String(estExistentes[idx][0] || '').trim().toLowerCase() === estOrigNorm) {
            filaExistente = idx + 2;
            break;
          }
        }
      }

      if (filaExistente > 0) {
        // En lugar de duplicar la fila, unificamos adicionales y actualizamos la existente
        const adicActuales = String(hoja.getRange(filaExistente, cAdic).getValue() || '').split(',').map(s => s.trim()).filter(Boolean);
        const nuevosAdic = adicStr.split(',').map(s => s.trim()).filter(Boolean);
        nuevosAdic.forEach(ad => {
          if (!adicActuales.some(a => a.toLowerCase() === ad.toLowerCase())) {
            adicActuales.push(ad);
          }
        });
        hoja.getRange(filaExistente, cAdic).setValue(adicActuales.join(', '));
        if (causa) hoja.getRange(filaExistente, cCausa).setValue(causa);
        hoja.getRange(filaExistente, cAct).setValue(activo);
      } else {
        const nextRow = ult + 1;
        hoja.getRange(nextRow, cEst).setValue(estOrig);
        hoja.getRange(nextRow, cCausa).setValue(causa);
        hoja.getRange(nextRow, cAdic).setValue(adicStr);
        hoja.getRange(nextRow, cAct).setValue(activo);
      }
    }

    SpreadsheetApp.flush();
    const abmResult = obtenerEstudiosAdicionalesABM();
    const mapa = obtenerMapeoEstudiosAdicionales();

    return {
      status: "OK",
      msg: "Registro guardado correctamente en la hoja de Estudios Adicionales",
      items: abmResult.items,
      mapa: mapa
    };
  } catch (err) {
    console.error("Error en guardarEstudioAdicionalABM: " + err.message);
    return { status: "ERROR", msg: err.message };
  }
}

/**
 * Elimina una fila completa del catálogo en la hoja Estudios_Adicionales.
 */
function eliminarFilaEstudioAdicionalABM(fila) {
  try {
    const hoja = inicializarHojaEstudiosAdicionales();
    const rowNum = parseInt(fila, 10);
    if (!rowNum || rowNum < 2 || rowNum > hoja.getLastRow()) {
      return { status: "ERROR", msg: "Número de fila inválido." };
    }
    hoja.deleteRow(rowNum);
    SpreadsheetApp.flush();

    const abmResult = obtenerEstudiosAdicionalesABM();
    const mapa = obtenerMapeoEstudiosAdicionales();

    return {
      status: "OK",
      msg: "Estudio eliminado del catálogo correctamente.",
      items: abmResult.items,
      mapa: mapa
    };
  } catch (err) {
    console.error("Error en eliminarFilaEstudioAdicionalABM: " + err.message);
    return { status: "ERROR", msg: err.message };
  }
}

/**
 * Quita un estudio adicional específico de la fila del catálogo en Sheets.
 */
function quitarAdicionalDeCatalogoABM(fila, adicionalParaQuitar) {
  try {
    const hoja = inicializarHojaEstudiosAdicionales();
    const rowNum = parseInt(fila, 10);
    if (!rowNum || rowNum < 2 || rowNum > hoja.getLastRow()) {
      return { status: "ERROR", msg: "Fila inválida." };
    }
    const cols = Math.max(hoja.getLastColumn(), 4);
    const headers = hoja.getRange(1, 1, 1, cols).getValues()[0].map(h => String(h || '').trim().toLowerCase());
    let cAdic = headers.indexOf("estudios_adicionales_sugeridos") + 1;
    if (!cAdic) cAdic = 3;

    const actualStr = String(hoja.getRange(rowNum, cAdic).getValue() || '');
    const arr = actualStr.split(',').map(s => s.trim()).filter(Boolean);
    const normQuitar = String(adicionalParaQuitar || '').trim().toLowerCase();
    const filtrado = arr.filter(s => s.toLowerCase() !== normQuitar);

    hoja.getRange(rowNum, cAdic).setValue(filtrado.join(', '));
    SpreadsheetApp.flush();

    const abmResult = obtenerEstudiosAdicionalesABM();
    const mapa = obtenerMapeoEstudiosAdicionales();
    return {
      status: "OK",
      msg: "Estudio adicional removido del catálogo.",
      items: abmResult.items,
      mapa: mapa
    };
  } catch (err) {
    return { status: "ERROR", msg: err.message };
  }
}

/**
 * Agrega un nuevo estudio adicional a la fila del catálogo en Sheets.
 */
function agregarAdicionalACatalogoABM(fila, nuevoAdicional) {
  try {
    const nuevo = String(nuevoAdicional || '').trim();
    if (!nuevo) return { status: "ERROR", msg: "El nombre del estudio no puede estar vacío." };

    const hoja = inicializarHojaEstudiosAdicionales();
    const rowNum = parseInt(fila, 10);
    if (!rowNum || rowNum < 2 || rowNum > hoja.getLastRow()) {
      return { status: "ERROR", msg: "Fila inválida." };
    }
    const cols = Math.max(hoja.getLastColumn(), 4);
    const headers = hoja.getRange(1, 1, 1, cols).getValues()[0].map(h => String(h || '').trim().toLowerCase());
    let cAdic = headers.indexOf("estudios_adicionales_sugeridos") + 1;
    if (!cAdic) cAdic = 3;

    const actualStr = String(hoja.getRange(rowNum, cAdic).getValue() || '');
    const arr = actualStr.split(',').map(s => s.trim()).filter(Boolean);
    const normNuevo = nuevo.toLowerCase();
    if (!arr.some(s => s.toLowerCase() === normNuevo)) {
      arr.push(nuevo);
      hoja.getRange(rowNum, cAdic).setValue(arr.join(', '));
      SpreadsheetApp.flush();
    }

    const abmResult = obtenerEstudiosAdicionalesABM();
    const mapa = obtenerMapeoEstudiosAdicionales();
    return {
      status: "OK",
      msg: "Estudio adicional agregado al catálogo.",
      items: abmResult.items,
      mapa: mapa
    };
  } catch (err) {
    return { status: "ERROR", msg: err.message };
  }
}

function _mapeoEstudiosAdicionalesFallback_() {
  return {
    "orina completa": [
      "Ecografía renal y vesical",
      "Urocultivo",
      "Clearance de creatinina",
      "Citología urinaria",
      "Dosaje de biomarcadores específicos en sangre u orina"
    ],
    "orina": [
      "Ecografía renal y vesical",
      "Urocultivo",
      "Clearance de creatinina",
      "Citología urinaria",
      "Dosaje de biomarcadores específicos en sangre u orina"
    ],
    "foniatrico": [
      "Fibrolaringoscopia",
      "Laringoestroboscopia"
    ],
    "examen clinico con orientacion a varices": [
      "Eco Doppler venoso de miembros inferiores"
    ],
    "varices": [
      "Eco Doppler venoso de miembros inferiores"
    ],
    "espirometria": [
      "Radiografía de tórax con técnica OIT",
      "TAC de tórax de alta resolución",
      "Prueba con broncodilatador",
      "Pletismografía",
      "Test de difusión de monóxido de carbono (DLCO)"
    ],
    "audiometria": [
      "Audiometría tonal con 48 hs de reposo auditivo",
      "Logoaudiometría",
      "Impedanciometría",
      "Otoemisiones acústicas",
      "Potenciales evocados auditivos (BERA)"
    ],
    "audiometria tonal (vias aerea y osea)": [
      "Audiometría tonal con 48 hs de reposo auditivo",
      "Logoaudiometría",
      "Impedanciometría",
      "Otoemisiones acústicas",
      "Potenciales evocados auditivos (BERA)"
    ]
  };
}

/**
 * Guarda los estudios adicionales asignados a un legajo en la hoja Desvios.
 */
function guardarEstudiosAdicionalesEmpleado(claveOcuil, estudiosAdicionales) {
  try {
    const ss = _getSpreadsheet_();
    const hoja = ss.getSheetByName(CONFIG.HOJA_DESVIOS);
    if (!hoja || hoja.getLastRow() < 2) return { status: "ERROR", msg: "Hoja Desvios vacía" };

    const numCols = hoja.getLastColumn();
    const headers = hoja.getRange(1, 1, 1, numCols).getValues()[0].map(h => String(h || '').trim());
    let cAdic = headers.indexOf("Estudios_Adicionales") + 1;
    if (cAdic === 0) {
      cAdic = hoja.getLastColumn() + 1;
      hoja.getRange(1, cAdic).setValue("Estudios_Adicionales").setFontWeight("bold").setBackground("#d97706").setFontColor("#ffffff");
    }

    const cClave = headers.indexOf("Clave_Unica") !== -1 ? (headers.indexOf("Clave_Unica") + 1) : (headers.indexOf("CUIL") + 1);
    const cCuil = headers.indexOf("CUIL") + 1;
    const cSrt = headers.indexOf("Codigo_SRT") + 1;
    const values = hoja.getRange(2, 1, hoja.getLastRow() - 1, Math.max(numCols, cAdic)).getValues();

    const sKey = String(claveOcuil || '').trim();
    let targetRow = -1;
    for (let r = 0; r < values.length; r++) {
      const vClave = String(values[r][cClave - 1] || '').trim();
      const vCuil = String(values[r][cCuil - 1] || '').trim();
      const vSrt = String(values[r][cSrt - 1] || '').trim();
      if (vClave === sKey || vCuil === sKey || (vCuil + '_' + vSrt) === sKey) {
        targetRow = r + 2;
        break;
      }
    }

    if (targetRow === -1) return { status: "ERROR", msg: "No se encontró registro con clave " + claveOcuil };

    const strAdic = Array.isArray(estudiosAdicionales) ? estudiosAdicionales.join(', ') : String(estudiosAdicionales || '');
    hoja.getRange(targetRow, cAdic).setValue(strAdic);
    return { status: "OK", msg: "Estudios adicionales guardados" };
  } catch (e) {
    console.error("Error en guardarEstudiosAdicionalesEmpleado: " + e.message);
    return { status: "ERROR", msg: e.message };
  }
}

/**
 * Obtiene el listado de prestadores habilitados para reconfirmación desde la hoja 'prestador_reco'.
 * Columnas: Prestador, Domicilio, Telefono.
 */
function obtenerPrestadoresReconfirmacion() {
  try {
    const ss = _getSpreadsheet_();
    let hoja = ss.getSheetByName(CONFIG.HOJA_PRESTADOR_RECO || "prestador_reco");
    if (!hoja) {
      const sheets = ss.getSheets();
      for (let s of sheets) {
        const n = s.getName().trim().toLowerCase();
        if (n === 'prestador_reco' || n === 'prestadores_reco' || n === 'prestador_reconfirmacion' || n === 'prestadores_reconfirmacion') {
          hoja = s;
          break;
        }
      }
    }
    if (!hoja || hoja.getLastRow() < 2) return [];

    const ultFila = hoja.getLastRow();
    const ultCol = Math.max(hoja.getLastColumn(), 3);
    const data = hoja.getRange(1, 1, ultFila, ultCol).getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    let cPrestador = headers.findIndex(h => h.includes('prestador'));
    let cDomicilio = headers.findIndex(h => h.includes('domicilio') || h.includes('direccion'));
    let cTelefono  = headers.findIndex(h => h.includes('telefono') || h.includes('tel') || h.includes('contacto'));

    if (cPrestador === -1) cPrestador = 0;
    if (cDomicilio === -1) cDomicilio = 1;
    if (cTelefono === -1)  cTelefono = 2;

    const lista = [];

    for (let i = 1; i < data.length; i++) {
      const p = String(data[i][cPrestador] || '').trim();
      const d = String(data[i][cDomicilio] || '').trim();
      const t = String(data[i][cTelefono] || '').trim();
      if (p) {
        const exist = lista.find(item => item.prestador.toUpperCase() === p.toUpperCase());
        if (exist) {
          if (!exist.domicilio && d) exist.domicilio = d;
          if (!exist.telefono && t) exist.telefono = t;
        } else {
          lista.push({
            prestador: p,
            domicilio: d,
            telefono: t
          });
        }
      }
    }

    lista.sort((a, b) => a.prestador.localeCompare(b.prestador));
    return lista;
  } catch (e) {
    console.error("Error en obtenerPrestadoresReconfirmacion: " + e.message);
    return [];
  }
}

/**
 * Obtiene la fecha de estudio inicial (Columna AB de Hoja Madre) y el prestador inicial (Columna O)
 * buscando por los índices de filas o por CUIL.
 */
function obtenerDatosEstudioInicialMadre(indicesMadre, cuil) {
  try {
    const ss = _getSpreadsheet_();
    const hojaMadre = ss.getSheetByName(CONFIG.HOJA_MADRE_NAME || "Madre");
    if (!hojaMadre || hojaMadre.getLastRow() < 2) return { status: 'OK', fechaEstudio: '', prestador: '' };

    const tz = Session.getScriptTimeZone();
    let fechaEstudio = '';
    let prestador = '';

    const indices = Array.isArray(indicesMadre) ? indicesMadre : [];
    for (let i = 0; i < indices.length; i++) {
      const idx = parseInt(indices[i], 10);
      if (!isNaN(idx) && idx >= 2 && idx <= hojaMadre.getLastRow()) {
        const row = hojaMadre.getRange(idx, 1, 1, Math.max(CONFIG.COL_FECHA_ESTUDIO + 1, 30)).getValues()[0];
        const fRaw = row[CONFIG.COL_FECHA_ESTUDIO]; // Columna AB (idx 27)
        const pRaw = row[CONFIG.COL_PRESTADOR];     // Columna O (idx 14)

        if (fRaw && !fechaEstudio) {
          fechaEstudio = _parseFechaConTz_(fRaw, tz);
        }
        if (pRaw && !prestador) {
          prestador = String(pRaw || '').trim();
        }
        if (fechaEstudio && prestador) break;
      }
    }

    // Si no vinieron índices o quedaron vacíos, buscar por CUIL en Hoja Madre
    if ((!fechaEstudio || !prestador) && cuil) {
      const cleanCuil = String(cuil).replace(/\D/g, '');
      const ultRow = hojaMadre.getLastRow();
      if (ultRow >= 2) {
        const colCuils = hojaMadre.getRange(2, CONFIG.COL_CUIL + 1, ultRow - 1, 1).getValues();
        for (let r = 0; r < colCuils.length; r++) {
          const cVal = String(colCuils[r][0] || '').replace(/\D/g, '');
          if (cVal === cleanCuil) {
            const rowIdx = r + 2;
            const row = hojaMadre.getRange(rowIdx, 1, 1, Math.max(CONFIG.COL_FECHA_ESTUDIO + 1, 30)).getValues()[0];
            if (!fechaEstudio && row[CONFIG.COL_FECHA_ESTUDIO]) {
              fechaEstudio = _parseFechaConTz_(row[CONFIG.COL_FECHA_ESTUDIO], tz);
            }
            if (!prestador && row[CONFIG.COL_PRESTADOR]) {
              prestador = String(row[CONFIG.COL_PRESTADOR] || '').trim();
            }
            if (fechaEstudio && prestador) break;
          }
        }
      }
    }

    return { status: 'OK', fechaEstudio: fechaEstudio, prestador: prestador };
  } catch (e) {
    console.error("Error en obtenerDatosEstudioInicialMadre: " + e.message);
    return { status: 'ERROR', fechaEstudio: '', prestador: '' };
  }
}

/**
 * Guarda Prestador2, Fecha_VE, Resultado_Reconfirmacion y Observacion_Reconfirmacion
 * en la hoja Desvios buscando dinámicamente por nombre de columna.
 */
function guardarGestionReconfirmacion(claveOcuil, prestador2, fechaVE, resultadoReconfirmacion, observacionReconfirmacion, estudiosAdicionales) {
  try {
    const ss = _getSpreadsheet_();
    const hoja = ss.getSheetByName(CONFIG.HOJA_DESVIOS);
    if (!hoja) return { status: "ERROR", msg: "No existe hoja Desvios" };

    const ultFila = hoja.getLastRow();
    if (ultFila < 2) return { status: "ERROR", msg: "Hoja Desvios vacía" };

    const numCols = hoja.getLastColumn();
    const headers = hoja.getRange(1, 1, 1, numCols).getValues()[0].map(h => String(h || '').trim());

    // Mapeo dinámico Nombre_Columna -> Índice (1-based para getRange)
    const colIdx = {};
    headers.forEach((nombre, idx) => {
      if (nombre) colIdx[nombre] = idx + 1;
    });

    const asegurarCol = (nombre) => {
      if (!colIdx[nombre]) {
        const c = hoja.getLastColumn() + 1;
        hoja.getRange(1, c).setValue(nombre).setFontWeight("bold").setBackground("#d97706").setFontColor("#ffffff");
        colIdx[nombre] = c;
      }
    };

    asegurarCol("Prestador2");
    asegurarCol("Fecha_VE");
    asegurarCol("Resultado_Reconfirmacion");
    asegurarCol("Observacion_Reconfirmacion");
    asegurarCol("Estudios_Adicionales");

    const colClave = colIdx["Clave_Unica"] || colIdx["CUIL"];
    const colCuil = colIdx["CUIL"];
    const colSrt = colIdx["Codigo_SRT"] || 1;
    const values = hoja.getRange(2, 1, ultFila - 1, hoja.getLastColumn()).getValues();

    let targetRow = -1;
    const sKey = String(claveOcuil || '').trim();

    for (let r = 0; r < values.length; r++) {
      const valClave = String(values[r][colClave - 1] || '').trim();
      const valCuil = String(values[r][colCuil - 1] || '').trim();
      const valSrt = String(values[r][colSrt - 1] || '').trim();

      if (valClave === sKey || valCuil === sKey || (valCuil + '_' + valSrt) === sKey) {
        targetRow = r + 2;
        break;
      }
    }

    if (targetRow === -1) {
      return { status: "ERROR", msg: "No se encontró el registro con clave " + claveOcuil };
    }

    if (prestador2 !== undefined) {
      hoja.getRange(targetRow, colIdx["Prestador2"]).setValue(prestador2 || '');
    }
    if (fechaVE !== undefined) {
      hoja.getRange(targetRow, colIdx["Fecha_VE"]).setValue(fechaVE || '');
    }
    if (resultadoReconfirmacion !== undefined) {
      hoja.getRange(targetRow, colIdx["Resultado_Reconfirmacion"]).setValue(resultadoReconfirmacion || '');
    }
    if (observacionReconfirmacion !== undefined) {
      hoja.getRange(targetRow, colIdx["Observacion_Reconfirmacion"]).setValue(observacionReconfirmacion || '');
    }
    if (estudiosAdicionales !== undefined) {
      const strAdic = Array.isArray(estudiosAdicionales) ? estudiosAdicionales.join(', ') : String(estudiosAdicionales || '');
      hoja.getRange(targetRow, colIdx["Estudios_Adicionales"]).setValue(strAdic);
    }

    // IMPACTO EN HOJA MADRE:
    // Si se cargó un resultado de reconfirmación (no vacío), impactar en Hoja Madre Col AC, Col AB, Col AA, Col W y Col AD
    if (resultadoReconfirmacion !== undefined && resultadoReconfirmacion !== '') {
      try {
        const hojaMadre = _getSheetMadre_();
        if (hojaMadre && hojaMadre.getLastRow() >= 2) {
          const filasMadreRaw = colIdx["Filas_Madre"] ? hoja.getRange(targetRow, colIdx["Filas_Madre"]).getValue() : '';
          let fIndices = [];
          if (filasMadreRaw) {
            fIndices = String(filasMadreRaw).split(',').map(n => parseInt(n.trim(), 10)).filter(n => !isNaN(n));
          }

          const cuilVal = String(hoja.getRange(targetRow, colIdx["CUIL"]).getValue() || '').trim();
          const srtVal = colIdx["Codigo_SRT"] ? String(hoja.getRange(targetRow, colIdx["Codigo_SRT"]).getValue() || '').trim() : '';

          // Si no había filas directas, buscar por CUIL + SRT en Madre
          if (fIndices.length === 0 && cuilVal) {
            const ultM = hojaMadre.getLastRow();
            const cuilsMadre = hojaMadre.getRange(2, CONFIG.COL_CUIL + 1, ultM - 1, 1).getValues();
            const srtsMadre = hojaMadre.getRange(2, CONFIG.COL_SRT + 1, ultM - 1, 1).getValues();
            for (let m = 0; m < cuilsMadre.length; m++) {
              if (String(cuilsMadre[m][0] || '').trim() === cuilVal && (!srtVal || String(srtsMadre[m][0] || '').trim() === srtVal)) {
                fIndices.push(m + 2);
              }
            }
          }

          const fVeVal = fechaVE || (colIdx["Fecha_VE"] ? hoja.getRange(targetRow, colIdx["Fecha_VE"]).getValue() : '') || formatearFecha(new Date());
          const pRecoVal = prestador2 || (colIdx["Prestador2"] ? hoja.getRange(targetRow, colIdx["Prestador2"]).getValue() : '') || '';
          const fechaHoy = formatearFecha(new Date());

          fIndices.forEach(idx => {
            if (idx >= 2 && idx <= hojaMadre.getLastRow()) {
              const resOrig = String(hojaMadre.getRange(idx, CONFIG.COL_RESULTADO + 1).getValue() || '').trim();
              const normRes = (typeof _normalizarTexto_ === 'function')
                ? _normalizarTexto_(resOrig)
                : resOrig.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();

              // Solo reconfirmar filas que estaban en desvío, pendiente o enf profesional
              if (normRes.includes('desvio') || normRes.includes('pendiente') || normRes.includes('profesional') || normRes === '') {
                // 1) Columna AC: Resultado definitivo (ej. Normal Laboral)
                const resLimpio = (resultadoReconfirmacion === 'Normal Laboral (Reconfirmado)') ? 'Normal Laboral' : resultadoReconfirmacion;
                hojaMadre.getRange(idx, CONFIG.COL_RESULTADO + 1).setValue(resLimpio);

                // 2) Columna AB: Fecha Estudio (reemplazada por la de la reconfirmación)
                if (fVeVal) {
                  hojaMadre.getRange(idx, CONFIG.COL_FECHA_ESTUDIO + 1).setValue(fVeVal);
                }

                // 3) Columna AA: Fecha Auditoría queda la original que está (no se pisa)
                // (se preserva la fecha original en Col AA)

                // 4) Columna W: Marca histórica permanente
                hojaMadre.getRange(idx, CONFIG.COL_RESULTADOS_RAR + 1).setValue('Reconfirmado: ' + (resOrig || 'Desvío Original'));

                // 5) Columna AD: Bitácora en Observaciones
                const prevObs = String(hojaMadre.getRange(idx, CONFIG.COL_OBSERVACION + 1).getValue() || '').trim();
                const notaReconf = `[RECONF: ${resLimpio} el ${fechaHoy}${pRecoVal ? ' - Prestador: ' + pRecoVal : ''}]`;
                if (!prevObs.includes('[RECONF:')) {
                  hojaMadre.getRange(idx, CONFIG.COL_OBSERVACION + 1).setValue((prevObs ? prevObs + ' ' : '') + notaReconf);
                }
              }
            }
          });

          SpreadsheetApp.flush();
          _invalidarCache_();
          _invalidarCacheDashboard_();
        }
      } catch (eMadre) {
        console.error("Error impactando reconfirmación en Hoja Madre: " + eMadre.message);
      }
    }

    return { status: "OK", msg: "Gestión guardada exitosamente", targetRow: targetRow };
  } catch(e) {
    console.error("Error en guardarGestionReconfirmacion: " + e.message);
    return { status: "ERROR", msg: e.message };
  }
}

/**
 * Guarda Prestador2 y Fecha_VE a nivel Establecimiento / Código SRT.
 * Si se pasa una lista de cuils, solo actualiza esos CUILs.
 * Si no se pasa o está vacía, actualiza los registros del Código SRT.
 */
function guardarGestionEstablecimiento(codigoSrt, prestador2, fechaVE, cuils) {
  try {
    const ss = _getSpreadsheet_();
    const hoja = ss.getSheetByName(CONFIG.HOJA_DESVIOS);
    if (!hoja) return { status: "ERROR", msg: "No existe hoja Desvios" };

    const ultFila = hoja.getLastRow();
    if (ultFila < 2) return { status: "ERROR", msg: "Hoja Desvios vacía" };

    const numCols = hoja.getLastColumn();
    const headers = hoja.getRange(1, 1, 1, numCols).getValues()[0].map(h => String(h || '').trim());

    const colIdx = {};
    headers.forEach((nombre, idx) => {
      if (nombre) colIdx[nombre] = idx + 1;
    });

    const asegurarCol = (nombre) => {
      if (!colIdx[nombre]) {
        const c = hoja.getLastColumn() + 1;
        hoja.getRange(1, c).setValue(nombre).setFontWeight("bold").setBackground("#d97706").setFontColor("#ffffff");
        colIdx[nombre] = c;
      }
    };

    asegurarCol("Prestador2");
    asegurarCol("Fecha_VE");

    const colCuil = colIdx["CUIL"];
    const colSrt = colIdx["Codigo_SRT"] || 1;
    const values = hoja.getRange(2, 1, ultFila - 1, hoja.getLastColumn()).getValues();

    const cuilsSet = (Array.isArray(cuils) && cuils.length > 0) ? new Set(cuils.map(c => String(c).trim())) : null;
    const sSrt = String(codigoSrt || '').trim();

    let actualizados = 0;
    for (let r = 0; r < values.length; r++) {
      const valCuil = String(values[r][colCuil - 1] || '').trim();
      const valSrt = String(values[r][colSrt - 1] || '').trim();

      let coincide = false;
      if (cuilsSet) {
        if (cuilsSet.has(valCuil) && (!sSrt || valSrt === sSrt)) {
          coincide = true;
        }
      } else {
        if (sSrt && valSrt === sSrt) {
          coincide = true;
        }
      }

      if (coincide) {
        const targetRow = r + 2;
        if (prestador2 !== undefined) {
          hoja.getRange(targetRow, colIdx["Prestador2"]).setValue(prestador2 || '');
        }
        if (fechaVE !== undefined) {
          hoja.getRange(targetRow, colIdx["Fecha_VE"]).setValue(fechaVE || '');
        }
        if (fechaVE || prestador2) {
          try {
            const filasMadreRaw = colIdx["Filas_Madre"] ? hoja.getRange(targetRow, colIdx["Filas_Madre"]).getValue() : '';
            if (filasMadreRaw) {
              const hojaMadre = _getSheetMadre_();
              const fIndices = String(filasMadreRaw).split(',').map(n => parseInt(n.trim(), 10)).filter(n => !isNaN(n));
              fIndices.forEach(idx => {
                if (idx >= 2 && idx <= hojaMadre.getLastRow()) {
                  const wVal = String(hojaMadre.getRange(idx, CONFIG.COL_RESULTADOS_RAR + 1).getValue() || '').trim();
                  if (!wVal) {
                    hojaMadre.getRange(idx, CONFIG.COL_RESULTADOS_RAR + 1).setValue('Pendiente Reconf.');
                  }
                }
              });
            }
          } catch(eW) {}
        }
        actualizados++;
      }
    }

    return { status: "OK", msg: `Se actualizaron ${actualizados} registros del establecimiento`, actualizados: actualizados };
  } catch(e) {
    console.error("Error en guardarGestionEstablecimiento: " + e.message);
    return { status: "ERROR", msg: e.message };
  }
}

// Alias para compatibilidad con llamadas existentes
function guardarAsignacionDesvio(claveOcuil, prestador2, fechaVE, resultadoReconfirmacion, observacionReconfirmacion) {
  return guardarGestionReconfirmacion(claveOcuil, prestador2, fechaVE, resultadoReconfirmacion, observacionReconfirmacion);
}

// ==================================================================
// RECLAMOS A PRESTADORES (Lectura directa desde hoja Reclamos)
// ==================================================================
function obtenerRegistrosReclamos() {
  try {
    const ss = _getSpreadsheet_();

    // 1) Intentar recuperar directamente desde HOJA_EMPLEADOS para asegurar paridad 100% con la vista Empleados
    const hojaEmp = ss.getSheetByName(CONFIG.HOJA_EMPLEADOS);
    if (hojaEmp && hojaEmp.getLastRow() >= 2) {
      const empsRes = cargarDatosEmpleados();
      const listEmp = Array.isArray(empsRes) ? empsRes : (empsRes && empsRes.empleados ? empsRes.empleados : []);

      const empsIncompletos = listEmp.filter(e => {
        return e.tieneNoApto || (Array.isArray(e.estudios) && e.estudios.some(est => est.resultado && !_esResultadoApto_(est.resultado)));
      });

      if (empsIncompletos.length > 0) {
        return empsIncompletos.map(e => {
          const estudiosRecl = (e.estudios || []).filter(est => est.resultado && !_esResultadoApto_(est.resultado));
          const tieneDefecto = (e.estudios || []).some(est => {
            const s = _normalizarTexto_(est.resultado);
            return s.includes('defectuosa') || s.includes('defecto');
          });
          const tieneAusente = (e.estudios || []).some(est => _normalizarTexto_(est.resultado).includes('ausente'));
          const motivoGlobal = (tieneDefecto && tieneAusente) ? 'Ausente y Defectuoso' :
                               tieneDefecto ? 'Práctica defectuosa' : 'Ausente / Sin realizar';

          return {
            cuil:             e.cuil,
            empleado:         e.apeNom,
            contrato:         e.contrato,
            nombreContrato:   e.nombreContrato,
            srt:              e.srt,
            establecimiento:  e.establecimiento,
            organismo:        e.organismo,
            motivoGlobal:     motivoGlobal,
            totalReclamables: estudiosRecl.length,
            totalEstudios:    e.totalEstudios,
            fechaAuditoria:   e.fechaAuditoria,
            adjunto1:         e.linkAdjunto,
            filasMadre:       e.filasIndices || [],
            rowIndices:       e.filasIndices || [],
            estudios:         (e.estudios || []).map(est => ({
              ...est,
              esReclamable: est.resultado ? !_esResultadoApto_(est.resultado) : false
            })),
            prestador:        e.prestador || '',
            clave:            e.clave || (e.cuil + '___' + e.establecimiento)
          };
        });
      }
    }

    // 2) Fallback a HOJA_RECLAMOS si Empleados estuviera vacía
    const hoja = ss.getSheetByName(CONFIG.HOJA_RECLAMOS);
    if (!hoja || hoja.getLastRow() < 2) return [];

    const ultFila = hoja.getLastRow();
    const numCols = Math.max(hoja.getLastColumn(), 16);
    const rows = hoja.getRange(2, 1, ultFila - 1, numCols).getValues();

    const resultadoEmpleados = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const cuil = String(r[0] || '').trim();
      if (!cuil) continue;

      let estudios = [];
      if (r[13]) {
        try {
          estudios = typeof r[13] === 'string' ? JSON.parse(r[13]) : r[13];
        } catch (e) {
          estudios = [];
        }
      }

      let filasMadre = [];
      if (r[12]) {
        filasMadre = String(r[12]).split(',').map(n => parseInt(n.trim(), 10)).filter(n => !isNaN(n));
      }

      resultadoEmpleados.push({
        cuil:             cuil,
        empleado:         String(r[1] || ''),
        contrato:         String(r[2] || ''),
        nombreContrato:   String(r[3] || ''),
        srt:              String(r[4] || ''),
        establecimiento:  String(r[5] || ''),
        organismo:        String(r[6] || ''),
        motivoGlobal:     String(r[7] || ''),
        totalReclamables: Number(r[8] || 0),
        totalEstudios:    Number(r[9] || 0),
        fechaAuditoria:   String(r[10] || ''),
        adjunto1:         String(r[11] || ''),
        filasMadre:       filasMadre,
        rowIndices:       filasMadre,
        estudios:         estudios,
        prestador:        String(r[14] || ''),
        clave:            String(r[15] || (cuil + '___' + (r[5] || '')))
      });
    }

    return resultadoEmpleados;
  } catch (e) {
    console.error("Error en obtenerRegistrosReclamos: " + e.message);
    return [];
  }
}

function comunicarRegistros(params) {
  // params = { indices: [rowIndex, ...], fecha: 'dd/MM/yyyy' }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.TIMEOUT_LOCK);
    if (!validarFecha_(params.fecha)) throw new Error(`Fecha inválida: "${params.fecha}"`);
    if (!params.indices || params.indices.length === 0) throw new Error('Sin índices');

    const sheet = _getSheetMadre_();
    const col   = CONFIG.COL_COMUNICACION + 1; // col 26 = Z

    params.indices.forEach(idx => {
      sheet.getRange(idx, col).setValue(params.fecha);
    });

    // Actualizar en hojas secundarias (Comunicaciones, Desvíos, Reclamos)
    _actualizarComunicacionEnHojasSecundarias_(params.indices, params.fecha);

    SpreadsheetApp.flush();
    _invalidarCache_();
    registrarEnBitacora('COMUNICACION_REGISTRADA', { cantidad: params.indices.length, fecha: params.fecha });
    return { status: 'OK', cantidad: params.indices.length };
  } catch (e) {
    registrarEnBitacora('ERROR_COMUNICACION', { error: e.message });
    return { status: 'Error', msg: String(e) };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Registra la comunicación oficial del lote (VE), actualiza la nueva columna
 * Fecha_Comunicacion_Lote en la hoja Comunicaciones y retorna los datos completos
 * de las 14 columnas para la generación de la Nota VE en el cliente.
 */
function comunicarLotesVE(params) {
  // params = { lotesKeys: ['...'], fecha: 'dd/MM/yyyy' }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.TIMEOUT_LOCK);
    if (!params.fecha || !validarFecha_(params.fecha)) {
      throw new Error(`Fecha de comunicación inválida: "${params.fecha}"`);
    }
    if (!params.lotesKeys || !Array.isArray(params.lotesKeys) || params.lotesKeys.length === 0) {
      throw new Error('Debe seleccionar al menos un lote para comunicar.');
    }

    const ss = _getSpreadsheet_();
    const sheetMadre = _getSheetMadre_();
    const sheetCom = _obtenerOCrearHojaComunicaciones_(ss);
    const tz = Session.getScriptTimeZone();

    // 1. Asegurar la columna "Fecha_Comunicacion_Lote" en hoja Comunicaciones
    const lastColCom = sheetCom.getLastColumn();
    const headersCom = sheetCom.getRange(1, 1, 1, lastColCom).getValues()[0].map(h => String(h || '').trim());
    let colFechaLoteIdx = headersCom.indexOf("Fecha_Comunicacion_Lote") + 1;
    if (colFechaLoteIdx === 0) {
      colFechaLoteIdx = lastColCom + 1;
      sheetCom.getRange(1, colFechaLoteIdx)
        .setValue("Fecha_Comunicacion_Lote")
        .setFontWeight("bold")
        .setBackground("#0f6b52")
        .setFontColor("#ffffff");
    }

    const colComunIdx = headersCom.indexOf("Comunicacion") + 1 || 11;
    const colEstadoLoteIdx = headersCom.indexOf("Estado_Lote") + 1 || 23;
    const colLoteKeyIdx = headersCom.indexOf("Lote_Key") + 1 || 18;
    const colFilasMadreIdx = headersCom.indexOf("Filas_Madre") + 1 || 13;

    // 2. Localizar y actualizar las filas en Comunicaciones
    const dataCom = sheetCom.getDataRange().getValues();
    const lotesSet = new Set(params.lotesKeys.map(k => String(k).trim()));
    const indicesMadre = [];

    for (let i = 1; i < dataCom.length; i++) {
      const rowLoteKey = String(dataCom[i][colLoteKeyIdx - 1] || '').trim();
      const rowEstab = String(dataCom[i][5] || '').trim().toUpperCase();
      const combinedKey = `${rowLoteKey}___${rowEstab}`;

      if (lotesSet.has(rowLoteKey) || lotesSet.has(combinedKey)) {
        sheetCom.getRange(i + 1, colFechaLoteIdx).setValue(params.fecha);
        sheetCom.getRange(i + 1, colComunIdx).setValue(params.fecha);
        sheetCom.getRange(i + 1, colEstadoLoteIdx).setValue("Lote Comunicado (VE)");

        const fMadreStr = String(dataCom[i][colFilasMadreIdx - 1] || '').trim();
        if (fMadreStr) {
          fMadreStr.split(',').forEach(idx => {
            const num = parseInt(idx.trim(), 10);
            if (!isNaN(num) && num > 1) indicesMadre.push(num);
          });
        }
      }
    }

    // 3. Actualizar en Hoja Madre columna Z (Comunicación de Resultados)
    if (indicesMadre.length > 0) {
      const colMadreCom = CONFIG.COL_COMUNICACION + 1;
      indicesMadre.forEach(rowIdx => {
        sheetMadre.getRange(rowIdx, colMadreCom).setValue(params.fecha);
      });

      // Actualizar también en hojas secundarias (Desvíos, Reclamos, Empleados)
      _actualizarComunicacionEnHojasSecundarias_(indicesMadre, params.fecha);
    }

    SpreadsheetApp.flush();
    _invalidarCache_();

    // 4. Preparar datos completos de las 14 columnas de los lotes seleccionados para generar el PDF en cliente
    const establecimientosPdf = _obtenerDatosDetalleParaVE_(sheetMadre, indicesMadre, params.fecha, tz);

    registrarEnBitacora('COMUNICACION_VE_REGISTRADA', {
      lotes: params.lotesKeys.length,
      estudios: indicesMadre.length,
      fecha: params.fecha
    });

    return {
      status: 'OK',
      fecha: params.fecha,
      cantidadLotes: params.lotesKeys.length,
      cantidadEstudios: indicesMadre.length,
      establecimientos: establecimientosPdf
    };
  } catch (e) {
    console.error("Error en comunicarLotesVE: " + e.message);
    registrarEnBitacora('ERROR_COMUNICACION_VE', { error: e.message });
    return { status: 'Error', msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Genera la Nota VE exclusiva para los casos pendientes de reconfirmación de un establecimiento.
 * Registra Fecha_VE en la hoja Desvios sin cerrar el lote completo ni informar el resto de los estudios.
 */
function comunicarLoteReconfirmacionVE(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.TIMEOUT_LOCK);
    if (!params.fecha || !validarFecha_(params.fecha)) {
      throw new Error(`Fecha de comunicación inválida: "${params.fecha}"`);
    }
    if (!params.loteKey && !params.srtCode) {
      throw new Error('Debe especificar el lote o Código SRT para la VE de Reconfirmación.');
    }

    const ss = _getSpreadsheet_();
    const sheetMadre = _getSheetMadre_();
    const sheetCom = _obtenerOCrearHojaComunicaciones_(ss);
    const hojaDesv = ss.getSheetByName(CONFIG.HOJA_DESVIOS);
    const tz = Session.getScriptTimeZone();

    const dataCom = sheetCom.getDataRange().getValues();
    const headersCom = dataCom[0].map(h => String(h || '').trim());
    const colLoteKeyIdx = headersCom.indexOf("Lote_Key") + 1 || 18;
    const colFilasMadreIdx = headersCom.indexOf("Filas_Madre") + 1 || 13;
    const colJsonIdx = headersCom.indexOf("Detalle_Estudios") + 1 || 14;
    const colSrtIdx = headersCom.indexOf("Codigo_SRT") + 1 || 4;

    const sLoteKey = String(params.loteKey || '').trim();
    const sSrtCode = String(params.srtCode || '').trim();
    const normalizarCuil = c => String(c || '').replace(/\D/g, '');
    const cuilsFiltro = Array.isArray(params.cuils) && params.cuils.length > 0
      ? new Set(params.cuils.map(normalizarCuil).filter(Boolean))
      : null;

    const indicesMadreReco = [];
    const cuilsReco = new Set();

    if (Array.isArray(params.indicesMadre) && params.indicesMadre.length > 0) {
      params.indicesMadre.forEach(idx => {
        const num = parseInt(idx, 10);
        if (!isNaN(num) && num > 1) indicesMadreReco.push(num);
      });
      if (cuilsFiltro) cuilsFiltro.forEach(c => cuilsReco.add(c));
    }

    // Si no vinieron índices Madre directamente, extraerlos de Comunicaciones para los CUILs seleccionados
    if (indicesMadreReco.length === 0) {
      for (let i = 1; i < dataCom.length; i++) {
        const rowLoteKey = String(dataCom[i][colLoteKeyIdx - 1] || '').trim();
        const rowEstab = String(dataCom[i][5] || '').trim().toUpperCase();
        const combinedKey = `${rowLoteKey}___${rowEstab}`;
        const rowSrt = String(dataCom[i][colSrtIdx - 1] || '').trim();
        const cuilLimpio = normalizarCuil(dataCom[i][0]);

        const coincideLote = sLoteKey && (rowLoteKey === sLoteKey || combinedKey === sLoteKey);
        const coincideSrt = sSrtCode && (rowSrt === sSrtCode);

        if (coincideLote || coincideSrt) {
          if (cuilsFiltro && !cuilsFiltro.has(cuilLimpio)) continue;

          const jsonRaw = dataCom[i][colJsonIdx - 1];
          let estudios = [];
          try { estudios = typeof jsonRaw === 'string' ? JSON.parse(jsonRaw) : jsonRaw; } catch (e) {}

          if (Array.isArray(estudios)) {
            estudios.forEach(est => {
              const rNorm = _normalizarTexto_(est.resultado);
              if (rNorm.includes('pendiente de confirmacion') || rNorm.includes('desvio') || rNorm.includes('desvío')) {
                const num = parseInt(est.rowIndex, 10);
                if (!isNaN(num) && num > 1 && !indicesMadreReco.includes(num)) {
                  indicesMadreReco.push(num);
                  if (cuilLimpio) cuilsReco.add(cuilLimpio);
                }
              }
            });
          }
        }
      }
    }

    if (indicesMadreReco.length === 0 && (!cuilsFiltro || cuilsFiltro.size === 0)) {
      throw new Error('No se encontraron estudios pendientes de confirmación para este lote/establecimiento.');
    }

    // Actualizar Fecha_VE y Prestador2 en hoja Desvios para estos empleados
    if (hojaDesv && hojaDesv.getLastRow() > 1) {
      const numColsD = hojaDesv.getLastColumn();
      const headersD = hojaDesv.getRange(1, 1, 1, numColsD).getValues()[0].map(h => String(h || '').trim());
      const cCuilD = headersD.indexOf("CUIL") + 1 || 1;
      let cFechaVeD = headersD.indexOf("Fecha_VE") + 1;
      let cPrestador2D = headersD.indexOf("Prestador2") + 1;

      if (cFechaVeD === 0) {
        cFechaVeD = numColsD + 1;
        hojaDesv.getRange(1, cFechaVeD).setValue("Fecha_VE").setFontWeight("bold").setBackground("#d97706").setFontColor("#ffffff");
      }
      if (cPrestador2D === 0) {
        cPrestador2D = hojaDesv.getLastColumn() + 1;
        hojaDesv.getRange(1, cPrestador2D).setValue("Prestador2").setFontWeight("bold").setBackground("#d97706").setFontColor("#ffffff");
      }

      const totalColsD = Math.max(numColsD, cFechaVeD, cPrestador2D);
      const dataD = hojaDesv.getRange(2, 1, hojaDesv.getLastRow() - 1, totalColsD).getValues();
      const fechaParaVE = params.fechaCitacion || params.fecha;

      for (let k = 0; k < dataD.length; k++) {
        const cD = normalizarCuil(dataD[k][cCuilD - 1]);
        if (cuilsReco.has(cD) || (cuilsFiltro && cuilsFiltro.has(cD))) {
          hojaDesv.getRange(k + 2, cFechaVeD).setValue(fechaParaVE);
          if (params.prestador) {
            hojaDesv.getRange(k + 2, cPrestador2D).setValue(params.prestador);
          }
        }
      }
    }

    // Actualizar columna Fecha_VE_Reconfirmacion en hoja Comunicaciones para los CUILs de este lote
    const lastColCom = sheetCom.getLastColumn();
    const headersComActuales = sheetCom.getRange(1, 1, 1, lastColCom).getValues()[0].map(h => String(h || '').trim());
    let colFechaRecoComIdx = headersComActuales.indexOf("Fecha_VE_Reconfirmacion") + 1;
    if (colFechaRecoComIdx === 0) {
      colFechaRecoComIdx = lastColCom + 1;
      sheetCom.getRange(1, colFechaRecoComIdx)
        .setValue("Fecha_VE_Reconfirmacion")
        .setFontWeight("bold")
        .setBackground("#d97706")
        .setFontColor("#ffffff");
    }
    const cCuilComIdx = headersComActuales.indexOf("CUIL") + 1 || 1;
    for (let i = 1; i < dataCom.length; i++) {
      const cCom = normalizarCuil(dataCom[i][cCuilComIdx - 1]);
      if (cuilsReco.has(cCom) || (cuilsFiltro && cuilsFiltro.has(cCom))) {
        sheetCom.getRange(i + 1, colFechaRecoComIdx).setValue(params.fecha);
      }
    }

    SpreadsheetApp.flush();
    _invalidarCache_();

    const establecimientosPdf = _obtenerDatosDetalleParaVE_(sheetMadre, indicesMadreReco, params.fecha, tz);

    registrarEnBitacora('COMUNICACION_VE_RECONFIRMACION_REGISTRADA', {
      lote: params.loteKey || params.srtCode,
      estudios: indicesMadreReco.length,
      fecha: params.fecha,
      prestador: params.prestador
    });

    return {
      status: 'OK',
      fecha: params.fecha,
      cantidadEstudios: indicesMadreReco.length,
      establecimientos: establecimientosPdf,
      metaReconfirmacion: {
        prestador: params.prestador || '',
        domicilio: params.domicilio || '',
        telefono: params.telefono || '',
        prestadorInicial: params.prestadorInicial || '',
        textoPersonalizado: params.textoPersonalizado || '',
        fechaCitacion: params.fechaCitacion || '',
        horaCitacion: params.horaCitacion || '',
        estudiosAdicionales: params.estudiosAdicionales || '',
        examenInicial: params.examenInicial || '',
        fechaExamenInicial: params.fechaExamenInicial || '',
        contrato: params.contrato || '',
        poliza: params.poliza || ''
      }
    };
  } catch (e) {
    console.error("Error en comunicarLoteReconfirmacionVE: " + e.message);
    return { status: 'Error', msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Registra en lote masivo la Fecha_VE y Prestador2 en hoja Desvios y Comunicaciones
 * para multiples establecimientos en una sola transaccion con lock.
 */
function comunicarLotesReconfirmacionVEMasivo(lotes) {
  if (!Array.isArray(lotes) || lotes.length === 0) {
    return { status: 'Error', msg: 'No se recibieron lotes para procesar' };
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.TIMEOUT_LOCK);
    const ss = _getSpreadsheet_();
    const sheetCom = _obtenerOCrearHojaComunicaciones_(ss);
    const hojaDesv = ss.getSheetByName(CONFIG.HOJA_DESVIOS);

    const normalizarCuil = c => String(c || '').replace(/\D/g, '');

    // Construir mapa de CUIL -> { fecha, prestador }
    const mapaCuilDatos = new Map();
    lotes.forEach(lote => {
      const f = lote.fechaCitacion || lote.fecha;
      const p = lote.prestador || '';
      if (Array.isArray(lote.cuils)) {
        lote.cuils.forEach(c => {
          const cNorm = normalizarCuil(c);
          if (cNorm) mapaCuilDatos.set(cNorm, { fecha: f, prestador: p });
        });
      }
    });

    if (mapaCuilDatos.size === 0) {
      return { status: 'OK', procesados: 0 };
    }

    // Actualizar hoja Desvios
    if (hojaDesv && hojaDesv.getLastRow() > 1) {
      const numColsD = hojaDesv.getLastColumn();
      const headersD = hojaDesv.getRange(1, 1, 1, numColsD).getValues()[0].map(h => String(h || '').trim());
      const cCuilD = headersD.indexOf("CUIL") + 1 || 1;
      let cFechaVeD = headersD.indexOf("Fecha_VE") + 1;
      let cPrestador2D = headersD.indexOf("Prestador2") + 1;

      if (cFechaVeD === 0) {
        cFechaVeD = numColsD + 1;
        hojaDesv.getRange(1, cFechaVeD).setValue("Fecha_VE").setFontWeight("bold").setBackground("#d97706").setFontColor("#ffffff");
      }
      if (cPrestador2D === 0) {
        cPrestador2D = hojaDesv.getLastColumn() + 1;
        hojaDesv.getRange(1, cPrestador2D).setValue("Prestador2").setFontWeight("bold").setBackground("#d97706").setFontColor("#ffffff");
      }

      const totalColsD = Math.max(numColsD, cFechaVeD, cPrestador2D);
      const dataD = hojaDesv.getRange(2, 1, hojaDesv.getLastRow() - 1, totalColsD).getValues();

      for (let k = 0; k < dataD.length; k++) {
        const cD = normalizarCuil(dataD[k][cCuilD - 1]);
        if (mapaCuilDatos.has(cD)) {
          const info = mapaCuilDatos.get(cD);
          hojaDesv.getRange(k + 2, cFechaVeD).setValue(info.fecha);
          if (info.prestador) {
            hojaDesv.getRange(k + 2, cPrestador2D).setValue(info.prestador);
          }
        }
      }
    }

    // Actualizar columna Fecha_VE_Reconfirmacion en hoja Comunicaciones
    if (sheetCom && sheetCom.getLastRow() > 1) {
      const lastColCom = sheetCom.getLastColumn();
      const headersComActuales = sheetCom.getRange(1, 1, 1, lastColCom).getValues()[0].map(h => String(h || '').trim());
      let colFechaRecoComIdx = headersComActuales.indexOf("Fecha_VE_Reconfirmacion") + 1;
      if (colFechaRecoComIdx === 0) {
        colFechaRecoComIdx = lastColCom + 1;
        sheetCom.getRange(1, colFechaRecoComIdx)
          .setValue("Fecha_VE_Reconfirmacion")
          .setFontWeight("bold")
          .setBackground("#d97706")
          .setFontColor("#ffffff");
      }
      const cCuilComIdx = headersComActuales.indexOf("CUIL") + 1 || 1;
      const dataCom = sheetCom.getDataRange().getValues();
      for (let i = 1; i < dataCom.length; i++) {
        const cCom = normalizarCuil(dataCom[i][cCuilComIdx - 1]);
        if (mapaCuilDatos.has(cCom)) {
          sheetCom.getRange(i + 1, colFechaRecoComIdx).setValue(mapaCuilDatos.get(cCom).fecha);
        }
      }
    }

    SpreadsheetApp.flush();
    _invalidarCache_();

    registrarEnBitacora('COMUNICACION_VE_RECONFIRMACION_MASIVA', {
      lotesCount: lotes.length,
      trabajadores: mapaCuilDatos.size
    });

    return {
      status: 'OK',
      lotesProcesados: lotes.length,
      trabajadoresActualizados: mapaCuilDatos.size
    };
  } catch (e) {
    console.error("Error en comunicarLotesReconfirmacionVEMasivo: " + e.message);
    return { status: 'Error', msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

function _obtenerDatosDetalleParaVE_(sheetMadre, indicesMadre, fechaEmision, tz) {
  const dataMadre = sheetMadre.getDataRange().getValues();
  const mapaEstablecimientos = {};

  let mapaVinculaciones = {};
  try {
    const ss = sheetMadre.getParent();
    const vincs = _leerVinculaciones_(ss);
    vincs.forEach(v => {
      if (v.srt && v.contrato) mapaVinculaciones[v.srt] = v.contrato;
    });
  } catch(e) {}

  indicesMadre.forEach(rowIdx => {
    if (rowIdx < 2 || rowIdx > dataMadre.length) return;
    const r = dataMadre[rowIdx - 1];
    if (!r) return;

    const nombreEst = String(r[CONFIG.COL_NOMBRE_ESTABLECIMIENTO] || '').trim() || '(sin nombre)';
    const srt = String(r[CONFIG.COL_SRT] || '').trim();
    const cuitContrato = String(r[CONFIG.COL_CUIT_CONTRATO] || '').trim();
    const nombreContrato = String(r[CONFIG.COL_NOMBRE_CONTRATO] || '').trim();
    const direccionEst = String(r[CONFIG.COL_DIRECCION] || '').trim();
    const localidadEst = String(r[CONFIG.COL_LOCALIDAD] || '').trim();
    const provinciaEst = String(r[CONFIG.COL_PROVINCIA] || '').trim();
    const contratoVal = mapaVinculaciones[srt] || cuitContrato;
    const key = `${cuitContrato}___${srt}___${nombreEst.toUpperCase()}`;

    if (!mapaEstablecimientos[key]) {
      mapaEstablecimientos[key] = {
        key: key,
        nombre: nombreEst,
        codigoSrt: srt,
        cuitContrato: cuitContrato,
        nombreContrato: nombreContrato,
        contrato: contratoVal,
        direccion: direccionEst,
        localidad: localidadEst,
        provincia: provinciaEst,
        fechaEmision: fechaEmision,
        filas: []
      };
    }

    let fEstVal = _parseFechaConTz_(r[CONFIG.COL_FECHA_ESTUDIO], tz);
    let resEmiVal = String(r[CONFIG.COL_RESULTADO] || '').trim();
    if (/^\d{1,4}[\/\-]\d{1,2}[\/\-]\d{1,4}/.test(resEmiVal) && !/^\d{1,4}[\/\-]\d{1,2}[\/\-]\d{1,4}/.test(fEstVal)) {
      const tmp = resEmiVal;
      resEmiVal = fEstVal;
      fEstVal = tmp;
    }
    if (/^\d{1,4}[\/\-]\d{1,2}[\/\-]\d{1,4}/.test(resEmiVal)) resEmiVal = '';

    mapaEstablecimientos[key].filas.push({
      nombre: String(r[CONFIG.COL_APELLIDO] || '').trim(),
      cuil: String(r[CONFIG.COL_CUIL] || '').trim(),
      fechaIngreso: _parseFechaConTz_(r[CONFIG.COL_FECHA_INGRESO], tz),
      fechaInicioPuesto: _parseFechaConTz_(r[CONFIG.COL_INICIO_EXPOSICION], tz),
      puesto: String(r[CONFIG.COL_PUESTO] || '').trim(),
      sector: String(r[CONFIG.COL_SECTOR] || '').trim(),
      riesgoCodigo: String(r[CONFIG.COL_CODIGO_RIESGO] || '').trim(),
      examen: String(r[CONFIG.COL_ESTUDIO] || '').trim(),
      fechaRealizacion: fEstVal,
      resultadoEmi: resEmiVal,
      fechaRealizacionEmr: '',
      resultadoEmr: '',
      observacion: String(r[CONFIG.COL_OBSERVACION] || '').trim()
    });
  });

  return Object.values(mapaEstablecimientos);
}

/**
 * Retorna los datos de la nota VE para reimprimir sin alterar fechas ni estados de comunicación.
 */
function obtenerDatosDetalleLotesVE(params) {
  try {
    if (!params.lotesKeys || !Array.isArray(params.lotesKeys) || params.lotesKeys.length === 0) {
      throw new Error('Debe seleccionar al menos un lote.');
    }
    const ss = _getSpreadsheet_();
    const sheetMadre = _getSheetMadre_();
    const sheetCom = _obtenerOCrearHojaComunicaciones_(ss);
    const tz = Session.getScriptTimeZone();

    const dataCom = sheetCom.getDataRange().getValues();
    const headersCom = dataCom[0].map(h => String(h || '').trim());
    const colLoteKeyIdx = headersCom.indexOf("Lote_Key") + 1 || 18;
    const colFilasMadreIdx = headersCom.indexOf("Filas_Madre") + 1 || 13;
    const lotesSet = new Set(params.lotesKeys.map(k => String(k).trim()));
    const indicesMadre = [];

    for (let i = 1; i < dataCom.length; i++) {
      const rowLoteKey = String(dataCom[i][colLoteKeyIdx - 1] || '').trim();
      const rowEstab = String(dataCom[i][5] || '').trim().toUpperCase();
      const combinedKey = `${rowLoteKey}___${rowEstab}`;

      if (lotesSet.has(rowLoteKey) || lotesSet.has(combinedKey)) {
        const fMadreStr = String(dataCom[i][colFilasMadreIdx - 1] || '').trim();
        if (fMadreStr) {
          fMadreStr.split(',').forEach(idx => {
            const num = parseInt(idx.trim(), 10);
            if (!isNaN(num) && num > 1) indicesMadre.push(num);
          });
        }
      }
    }

    const establecimientosPdf = _obtenerDatosDetalleParaVE_(sheetMadre, indicesMadre, params.fecha || '', tz);
    return {
      status: 'OK',
      establecimientos: establecimientosPdf
    };
  } catch (e) {
    console.error("Error en obtenerDatosDetalleLotesVE: " + e.message);
    return { status: 'Error', msg: e.message };
  }
}

// ==================================================================
// ARCHIVADO DE PDFs EN GOOGLE DRIVE
// ==================================================================
function _extraerFileId_(urlOId) {
  if (!urlOId) return null;
  const str = String(urlOId).trim();
  if (!str) return null;
  // Si coincide con patrón directo de Google Drive ID o URL
  const match = str.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function _formatoAAAAMMDD_(val, tz) {
  if (!val) return "";
  let d = val;
  if (typeof val === 'string') {
    const parts = val.trim().split('/');
    if (parts.length === 3) {
      const dia  = parts[0].padStart(2, '0');
      const mes  = parts[1].padStart(2, '0');
      const anio = parts[2];
      return `${anio}${mes}${dia}`;
    }
    const partsIso = val.trim().split('-');
    if (partsIso.length === 3) {
      const dia = partsIso[2].substring(0, 2).padStart(2, '0');
      return `${partsIso[0]}${partsIso[1].padStart(2, '0')}${dia}`;
    }
    d = new Date(val);
  }
  if (d instanceof Date && !isNaN(d.getTime())) {
    return Utilities.formatDate(d, tz || Session.getScriptTimeZone(), "yyyyMMdd");
  }
  return String(val).replace(/[^0-9]/g, '');
}

function _obtenerOCrearSubcarpeta_(carpetaPadre, nombre) {
  const nombreLimpio = String(nombre || "Sin Nombre").trim().replace(/[/\\?%*:|"<>]/g, '-');
  const carpetas = carpetaPadre.getFoldersByName(nombreLimpio);
  if (carpetas.hasNext()) {
    return carpetas.next();
  }
  return carpetaPadre.createFolder(nombreLimpio);
}

/**
 * Obtiene o crea la carpeta del Establecimiento en Drive según la jerarquía del proyecto:
 * Carpeta Raíz -> CUIT - Nombre Contrato (C-D) -> SRT - Establecimiento (H-G)
 */
function _obtenerCarpetaEstablecimientoDrive_(cuitContrato, nombreContrato, srt, establecimiento) {
  if (!CONFIG.CARPETA_PDFS_ID) return null;
  const carpetaRaiz = DriveApp.getFolderById(CONFIG.CARPETA_PDFS_ID);

  const cuit = String(cuitContrato || '').trim();
  const nomC = String(nombreContrato || '').trim();
  const codSrt = String(srt || '').trim();
  const est = String(establecimiento || '').trim();

  // Nivel 1: CUIT - Nombre Contrato (C-D)
  const rutaL1 = `${cuit}-${nomC}`.replace(/^-|-$/g, '').trim() || 'Contrato Sin Definir';
  const carpetaL1 = _obtenerOCrearSubcarpeta_(carpetaRaiz, rutaL1);

  // Nivel 2: SRT - Establecimiento (H-G)
  const rutaL2 = `${codSrt}-${est}`.replace(/^-|-$/g, '').trim() || 'Establecimiento Sin Definir';
  const carpetaDestino = _obtenerOCrearSubcarpeta_(carpetaL1, rutaL2);

  return carpetaDestino;
}

/**
 * Retorna la URL de la carpeta del establecimiento en Drive para abrirla desde la UI.
 */
function obtenerUrlCarpetaEstablecimiento(cuit, contrato, srt, estab) {
  try {
    const carpeta = _obtenerCarpetaEstablecimientoDrive_(cuit || contrato, '', srt, estab);
    return carpeta ? carpeta.getUrl() : null;
  } catch (e) {
    console.error("Error en obtenerUrlCarpetaEstablecimiento: " + e.message);
    return null;
  }
}

/**
 * Guarda automáticamente la Nota Oficial VE en la carpeta de Drive del establecimiento
 * donde residen los estudios auditados.
 */
function guardarNotaVEEnDrive(params) {
  try {
    const carpeta = _obtenerCarpetaEstablecimientoDrive_(params.cuitContrato, params.nombreContrato, params.srt, params.establecimiento);
    if (!carpeta) {
      return { status: "Error", msg: "No se pudo acceder a la carpeta de Drive del establecimiento" };
    }

    let base64Data = String(params.pdfBase64 || '');
    if (base64Data.indexOf('base64,') !== -1) {
      base64Data = base64Data.split('base64,')[1];
    }
    if (!base64Data) {
      return { status: "Error", msg: "Contenido de PDF vacío" };
    }

    const bytes = Utilities.base64Decode(base64Data);
    const nombreArchivo = String(params.nombreArchivo || `Nota_VE_${params.srt || ''}.pdf`).trim();
    const blob = Utilities.newBlob(bytes, 'application/pdf', nombreArchivo);

    // Reemplazo inteligente: enviar a papelera versiones previas de la Nota VE en esta carpeta
    const files = carpeta.getFilesByName(nombreArchivo);
    while (files.hasNext()) {
      try {
        files.next().setTrashed(true);
      } catch (eTrash) {
        console.warn("Aviso enviando VE previa a papelera: " + eTrash.message);
      }
    }

    const archivoGuardado = carpeta.createFile(blob);
    archivoGuardado.setDescription(`Nota Oficial VE - Notificación EMP - Fecha: ${params.fecha || ''}`);

    registrarEnBitacora('NOTA_VE_GUARDADA_DRIVE', {
      establecimiento: params.establecimiento,
      archivo: nombreArchivo,
      url: archivoGuardado.getUrl(),
      carpetaUrl: carpeta.getUrl()
    });

    return {
      status: "OK",
      archivoUrl: archivoGuardado.getUrl(),
      carpetaUrl: carpeta.getUrl(),
      archivoId: archivoGuardado.getId()
    };
  } catch (err) {
    console.error("Error en guardarNotaVEEnDrive: " + err.message);
    registrarEnBitacora('ERROR_NOTA_VE_DRIVE', { error: err.message, establecimiento: params.establecimiento });
    return { status: "Error", msg: err.message };
  }
}

/**
 * Retorna las Ventanillas Electrónicas (Notas VE) generadas y disponibles en Drive
 * para un establecimiento específico, junto con el resumen de archivos encontrados.
 */
function listarVentanillasEstablecimiento(params) {
  try {
    const cuit = String(params.cuitContrato || params.contrato || '').trim();
    const nomC = String(params.nombreContrato || '').trim();
    const srt = String(params.srt || '').trim();
    const estab = String(params.establecimiento || '').trim();

    const carpeta = _obtenerCarpetaEstablecimientoDrive_(cuit, nomC, srt, estab);
    const archivos = [];
    let carpetaUrl = null;

    if (carpeta) {
      carpetaUrl = carpeta.getUrl();
      const files = carpeta.getFiles();
      while (files.hasNext()) {
        const f = files.next();
        const fn = f.getName();
        const fnNorm = fn.toLowerCase();

        // Identificar si es un archivo de Ventanilla Electrónica / Comunicación oficial
        if (fnNorm.includes('comunicacion') || fnNorm.includes('nota_ve') || fnNorm.includes('ventanilla') || fnNorm.includes('ve -') || fnNorm.includes('ve_') || (fnNorm.includes('.pdf') && (fnNorm.includes('reconfirmacion') || fnNorm.includes('resultados')))) {
          let tipo = 'Nota VE Lote Completo (Resultados EMP)';
          let icon = 'bi-file-earmark-pdf-fill';
          let color = '#2563eb';

          if (fnNorm.includes('reconfirmacion') || fnNorm.includes('reco')) {
            tipo = 'Nota VE Reconfirmación / Interconsulta';
            icon = 'bi-file-earmark-medical-fill';
            color = '#b45309';
          }

          archivos.push({
            id: f.getId(),
            nombre: fn,
            tipo: tipo,
            icon: icon,
            color: color,
            url: f.getUrl(),
            downloadUrl: 'https://drive.google.com/uc?export=download&id=' + f.getId(),
            fechaMod: Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
            tamanio: (f.getSize() / 1024).toFixed(1) + ' KB'
          });
        }
      }
    }

    return {
      status: "OK",
      cuit: cuit,
      nombreContrato: nomC,
      srt: srt,
      establecimiento: estab,
      carpetaUrl: carpetaUrl,
      archivos: archivos
    };
  } catch (err) {
    console.error("Error en listarVentanillasEstablecimiento: " + err.message);
    return { status: "Error", msg: err.message };
  }
}

function _obtenerPendientesArchivarDeduplicados_(sheet, cambios) {
  const mapaEmps = new Map();
  const numCols = Math.max(sheet.getLastColumn(), (CONFIG.COL_LINK_CARPETA || 38) + 1);

  cambios.forEach(c => {
    // Si el estudio es Ausente, no tiene PDF a archivar
    if (_normalizarTexto_(c.res).includes('ausente')) return;

    const fila = sheet.getRange(c.idx, 1, 1, numCols).getValues()[0];
    const cuil = String(fila[CONFIG.COL_CUIL] || '').trim();
    const linkAdj = String(fila[CONFIG.COL_ADJUNTO_1] || '').trim();
    const colG = CONFIG.COL_GUARDADO !== undefined ? CONFIG.COL_GUARDADO : 37;
    const guardado = String(fila[colG] || '').trim();
    const estab = String(fila[CONFIG.COL_NOMBRE_ESTABLECIMIENTO] || '').trim();
    const prest = String(fila[CONFIG.COL_PRESTADOR] || '').trim();
    const clave = `${cuil}___${estab || 'SIN_ESTABLECIMIENTO'}`;

    if (cuil && linkAdj && !_esGuardado_(guardado)) {
      const fileId = _extraerFileId_(linkAdj);
      if (fileId) {
        if (!mapaEmps.has(clave)) {
          mapaEmps.set(clave, {
            clave: clave,
            cuil: cuil,
            apeNom: String(fila[CONFIG.COL_APELLIDO] || '').trim(),
            contrato: String(fila[CONFIG.COL_CUIT_CONTRATO] || '').trim(),
            nombreContrato: String(fila[CONFIG.COL_NOMBRE_CONTRATO] || '').trim(),
            srt: String(fila[CONFIG.COL_SRT] || '').trim(),
            establecimiento: estab,
            prestador: prest,
            linkAdjunto: linkAdj,
            fileId: fileId,
            fechaEstudio: c.fechaEstudio || fila[CONFIG.COL_FECHA_ESTUDIO],
            fechaAuditoria: new Date(),
            filasIndices: []
          });
        }
        mapaEmps.get(clave).filasIndices.push(c.idx);
      }
    }
  });

  return Array.from(mapaEmps.values());
}

function archivarLotePDFs(lote) {
  if (!Array.isArray(lote) || lote.length === 0) return { status: "OK", procesados: 0, resultados: [] };
  if (!CONFIG.CARPETA_PDFS_ID) return { status: "Error", msg: "No está configurada CARPETA_PDFS_ID en CONFIG" };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.TIMEOUT_LOCK);
    const carpetaRaiz = DriveApp.getFolderById(CONFIG.CARPETA_PDFS_ID);
    const tz = Session.getScriptTimeZone();
    const ss = _getSpreadsheet_();
    const sheet = ss.getSheetByName(CONFIG.HOJA_MADRE_NAME);
    const hojaEmp = ss.getSheetByName(CONFIG.HOJA_EMPLEADOS);

    const mapaClaveFilaEmp = new Map();
    if (hojaEmp && hojaEmp.getLastRow() > 1) {
      const ultFilaEmp = hojaEmp.getLastRow();
      const numColsEmp = hojaEmp.getLastColumn();
      const dataEmp = hojaEmp.getRange(2, 1, ultFilaEmp - 1, numColsEmp).getValues();
      for (let k = 0; k < dataEmp.length; k++) {
        const c = String(dataEmp[k][0] || '').trim();
        const srt = String(dataEmp[k][4] || '').trim();
        const est = String(dataEmp[k][5] || '').trim();
        const clave19 = String(dataEmp[k][19] || '').trim();
        const rowNum = k + 2;
        if (clave19) mapaClaveFilaEmp.set(clave19, rowNum);
        if (c && est) mapaClaveFilaEmp.set(c + '___' + est, rowNum);
        if (c && srt) mapaClaveFilaEmp.set(c + '_' + srt, rowNum);
        if (c && !mapaClaveFilaEmp.has(c)) mapaClaveFilaEmp.set(c, rowNum);
      }
    }

    const colGuardado = (CONFIG.COL_GUARDADO !== undefined ? CONFIG.COL_GUARDADO : 37) + 1; // Col AL (38)
    const colLinkCarpeta = (CONFIG.COL_LINK_CARPETA !== undefined ? CONFIG.COL_LINK_CARPETA : 38) + 1; // Col AM (39)

    const resultados = [];

    lote.forEach(emp => {
      const res = _archivarPDFEmpleado_(emp, carpetaRaiz, tz);
      if (res.status === "OK") {
        if (Array.isArray(emp.filasIndices)) {
          emp.filasIndices.forEach(idx => {
            const rowVals = sheet.getRange(idx, 1, 1, Math.max(sheet.getLastColumn(), 39)).getValues()[0];
            const rRes = String(rowVals[CONFIG.COL_RESULTADO] || '').trim();
            const rAsis = String(rowVals[CONFIG.COL_ASISTENCIA] || '').toUpperCase().trim();
            const rAdj = String(rowVals[CONFIG.COL_ADJUNTO_1] || '').trim() || String(rowVals[CONFIG.COL_ADJUNTO_2] || '').trim();
            const esAusente = _normalizarTexto_(rRes).includes('ausente') || rAsis === 'AUSENTE';

            if (!esAusente && (rAdj || emp.linkAdjunto || emp.fileId)) {
              sheet.getRange(idx, colGuardado).setValue("SÍ");
              sheet.getRange(idx, colLinkCarpeta).setValue(res.urlCarpeta);
            } else if (esAusente) {
              sheet.getRange(idx, colGuardado).setValue("");
              sheet.getRange(idx, colLinkCarpeta).setValue("");
              sheet.getRange(idx, CONFIG.COL_ADJUNTO_1 + 1, 1, 2).setValues([["", ""]]);
            }
          });
        }
        const claveBuscada = emp.clave || (emp.cuil + '___' + (emp.establecimiento || ''));
        const filaEmp = (emp.clave && mapaClaveFilaEmp.get(emp.clave)) ||
                        (emp.cuil && emp.establecimiento && mapaClaveFilaEmp.get(emp.cuil + '___' + emp.establecimiento)) ||
                        (emp.cuil && emp.srt && mapaClaveFilaEmp.get(emp.cuil + '_' + emp.srt)) ||
                        (emp.cuil && mapaClaveFilaEmp.get(emp.cuil));
        if (hojaEmp && filaEmp) {
          hojaEmp.getRange(filaEmp, 11).setValue("SÍ"); // Col 11: Guardado
          hojaEmp.getRange(filaEmp, 12).setValue(res.urlCarpeta); // Col 12: Link_Carpeta
          try {
            const rawDet = hojaEmp.getRange(filaEmp, 18).getValue();
            if (rawDet) {
              const arrDet = typeof rawDet === 'string' ? JSON.parse(rawDet) : rawDet;
              if (Array.isArray(arrDet)) {
                arrDet.forEach(d => {
                  if (d && !_normalizarTexto_(d.resultado || '').includes('ausente')) {
                    d.guardado = "SÍ";
                    d.linkCarpeta = res.urlCarpeta;
                  }
                });
                hojaEmp.getRange(filaEmp, 18).setValue(JSON.stringify(arrDet));
              }
            }
          } catch(eDet) {}
        }
        resultados.push({ cuil: emp.cuil, safeKey: emp.safeKey, clave: claveBuscada, estab: emp.establecimiento, status: "OK", urlCarpeta: res.urlCarpeta, filas: emp.filasIndices });
      } else {
        const errorMsg = "ERROR: " + (res.msg || "Fallo en Drive");
        const claveBuscada = emp.clave || (emp.cuil + '___' + (emp.establecimiento || ''));
        const filaEmp = (emp.clave && mapaClaveFilaEmp.get(emp.clave)) ||
                        (emp.cuil && emp.establecimiento && mapaClaveFilaEmp.get(emp.cuil + '___' + emp.establecimiento)) ||
                        (emp.cuil && emp.srt && mapaClaveFilaEmp.get(emp.cuil + '_' + emp.srt)) ||
                        (emp.cuil && mapaClaveFilaEmp.get(emp.cuil));
        if (hojaEmp && filaEmp) {
          hojaEmp.getRange(filaEmp, 11).setValue(errorMsg); // Col 11: Guardado
        }
        resultados.push({ cuil: emp.cuil, safeKey: emp.safeKey, clave: claveBuscada, estab: emp.establecimiento, status: "Error", msg: res.msg });
      }
    });

    SpreadsheetApp.flush();
    _invalidarCache_();
    return { status: "OK", procesados: resultados.length, resultados: resultados };
  } catch (e) {
    console.error("Error en archivarLotePDFs: " + e.message);
    return { status: "Error", msg: e.message };
  } finally {
    lock.releaseLock();
  }
}

function _archivarPDFEmpleado_(emp, carpetaRaiz, tz) {
  try {
    let fileId = emp.fileId || _extraerFileId_(emp.linkAdjunto);
    if (!fileId && Array.isArray(emp.estudios)) {
      for (let s = 0; s < emp.estudios.length; s++) {
        const est = emp.estudios[s];
        const fId = _extraerFileId_(est.adjunto1 || est.adjunto2);
        if (fId) {
          fileId = fId;
          break;
        }
      }
    }
    if (!fileId) return { status: "Error", msg: "Sin ID de archivo PDF o enlace no válido" };

    let file;
    try {
      file = DriveApp.getFileById(fileId);
    } catch (eFile) {
      return { status: "Error", msg: "Archivo no encontrado o sin permisos en Drive (" + eFile.message + ")" };
    }

    const cuitContrato    = String(emp.contrato || '').trim();
    const nombreContrato  = String(emp.nombreContrato || '').trim();
    const srt             = String(emp.srt || '').trim();
    const establecimiento = String(emp.establecimiento || '').trim();
    const cuil            = String(emp.cuil || '').trim();
    const apeNom          = String(emp.apeNom || '').trim().replace(/[/\\?%*:|"<>]/g, ' ');

    // Nivel 1: C-D
    const rutaL1 = `${cuitContrato}-${nombreContrato}`.replace(/^-|-$/g, '').trim() || 'Contrato Sin Definir';
    const carpetaL1 = _obtenerOCrearSubcarpeta_(carpetaRaiz, rutaL1);

    // Nivel 2: H-G
    const rutaL2 = `${srt}-${establecimiento}`.replace(/^-|-$/g, '').trim() || 'Establecimiento Sin Definir';
    const carpetaDestino = _obtenerOCrearSubcarpeta_(carpetaL1, rutaL2);

    // [2026-09-13] Nomenclatura simplificada: solo CUIL y Nombre (sin fechas)
    const apeNomLimpio = apeNom.trim();
    const nombreArchivo = `${cuil}_${apeNomLimpio}.pdf`;

    // Reemplazo inteligente: buscar y enviar a papelera archivos previos de este CUIL en la carpeta de destino
    // (tanto con el nuevo formato como con nombres legacy con fechas), garantizando un único archivo por empleado
    const todosArchivos = carpetaDestino.getFiles();
    while (todosArchivos.hasNext()) {
      const f = todosArchivos.next();
      const fn = f.getName();
      if ((fn === nombreArchivo || (fn.includes(cuil) && fn.toLowerCase().endsWith('.pdf'))) && f.getId() !== fileId) {
        try {
          f.setTrashed(true);
        } catch (eTrash) {
          console.warn("Aviso enviando archivo anterior a papelera: " + eTrash.message);
        }
      }
    }

    file.makeCopy(nombreArchivo, carpetaDestino);
    const urlCarpeta = carpetaDestino.getUrl();

    registrarEnBitacora('PDF_ARCHIVADO_EMPLEADO', { cuil: cuil, archivo: nombreArchivo, ruta: `${rutaL1} > ${rutaL2}` });
    return { status: "OK", urlCarpeta: urlCarpeta };
  } catch (err) {
    console.error(`Error archivando PDF de ${emp.cuil}: ${err.message}`);
    registrarEnBitacora('ERROR_PDF_ARCHIVADO_EMPLEADO', { cuil: emp.cuil, error: err.message });
    return { status: "Error", msg: err.message };
  }
}

// ==================================================================
// HOJA DEDICADA: EMPLEADOS AUDITADOS / EN PROCESO
// ==================================================================

/**
 * Inicializa encabezados y estilo en la hoja dedicada "Empleados".
 */
function inicializarHojaEmpleados() {
  const ss = _getSpreadsheet_();
  let hoja = ss.getSheetByName(CONFIG.HOJA_EMPLEADOS);
  if (!hoja) {
    hoja = ss.insertSheet(CONFIG.HOJA_EMPLEADOS);
  }
  const headers = [
    "CUIL", "Apellido_y_Nombre", "Contrato", "Nombre_Contrato",
    "Codigo_SRT", "Establecimiento", "Nombre_Organismo",
    "Total_Estudios", "Auditados_Estudios", "Estado_Auditoria",
    "Guardado", "Link_Carpeta", "Link_Adjunto", "File_Id",
    "Fecha_Auditoria", "Fecha_Estudio", "Filas_Madre", "Detalle_Estudios",
    "Prestador", "Clave_Unica"
  ];
  if (hoja.getLastRow() === 0 || hoja.getRange(1, 1).getValue() !== "CUIL") {
    hoja.clear();
    hoja.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = hoja.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#0d9488"); // Teal elegante
    headerRange.setFontColor("#ffffff");
    hoja.setFrozenRows(1);
  } else {
    // Si la hoja ya tiene datos pero le faltan columnas (ej: Prestador, Clave_Unica), agregarlas dinámicamente sin borrar datos
    const currentCols = hoja.getLastColumn();
    const existingHeaders = hoja.getRange(1, 1, 1, currentCols).getValues()[0].map(h => String(h || '').trim());
    headers.forEach(h => {
      if (existingHeaders.indexOf(h) === -1) {
        const nextCol = hoja.getLastColumn() + 1;
        hoja.getRange(1, nextCol).setValue(h).setFontWeight("bold").setBackground("#0d9488").setFontColor("#ffffff");
      }
    });
  }
  return hoja;
}

/**
 * Identifica los casos relevantes de la Hoja Madre y los copia exclusivamente a la hoja "Empleados".
 * Es ultra-rápida (2-3s) porque procesa en memoria y no toca las hojas de Desvíos, Reclamos ni Comunicaciones.
 */
function sincronizarHojaEmpleados(skipLock) {
  const lock = skipLock ? null : LockService.getScriptLock();
  try {
    if (lock && !lock.tryLock(12000)) {
      return { status: "Error", msg: "El servidor está ocupado procesando otra tarea. Aguarda unos segundos." };
    }

    const hojaEmp = typeof inicializarHojaEmpleados === 'function' ? inicializarHojaEmpleados() : null;
    if (!hojaEmp) throw new Error("No se pudo inicializar la hoja Empleados.");

    const sheetMadre = _getSheetMadre_();
    if (!sheetMadre) throw new Error("No se encontró la hoja Madre.");

    const ultFila = sheetMadre.getLastRow();
    if (ultFila < 2) {
      _volcarFilasEnHoja_(hojaEmp, []);
      return { status: "OK", empleados: 0, msg: "Madre vacía" };
    }

    const numCols = Math.max(sheetMadre.getLastColumn(), 42);
    const headersMadreSync = sheetMadre.getRange(1, 1, 1, numCols).getValues()[0].map(h => String(h || '').trim().toLowerCase());
    let colPrestadorMadre = CONFIG.COL_PRESTADOR;
    const idxPrestFound = headersMadreSync.indexOf('prestador');
    if (idxPrestFound !== -1) {
      colPrestadorMadre = idxPrestFound;
    }

    const allMadreValues = sheetMadre.getRange(2, 1, ultFila - 1, numCols).getValues();

    const cuilsAuditados = new Set();
    const colG = (CONFIG.COL_GUARDADO !== undefined ? CONFIG.COL_GUARDADO : 37);
    const colAdj1 = (CONFIG.COL_ADJUNTO_1 !== undefined ? CONFIG.COL_ADJUNTO_1 : 32);
    const colAdj2 = (CONFIG.COL_ADJUNTO_2 !== undefined ? CONFIG.COL_ADJUNTO_2 : 33);

    for (let i = 0; i < allMadreValues.length; i++) {
      const row = allMadreValues[i];
      const res = String(row[CONFIG.COL_RESULTADO] || '').trim();
      const gd = String(row[colG] || '').trim();
      const c = String(row[CONFIG.COL_CUIL] || '').trim();
      const a1 = String(row[colAdj1] || '').trim();
      const a2 = String(row[colAdj2] || '').trim();

      if (res !== '' || gd.toUpperCase() === 'SÍ' || gd.toUpperCase() === 'SI' || a1 !== '' || a2 !== '') {
        if (c) cuilsAuditados.add(c);
      }
    }

    const tz = Session.getScriptTimeZone();
    const mapa = new Map();

    for (let k = 0; k < allMadreValues.length; k++) {
      const row = allMadreValues[k];
      const cuil = String(row[CONFIG.COL_CUIL] || "").trim();
      if (!cuil || !cuilsAuditados.has(cuil)) continue;

      const estudio = String(row[CONFIG.COL_ESTUDIO] || "").trim();
      if (estudio.includes("No posee")) continue;

      const filaMadre = k + 2;
      const asistencia = String(row[CONFIG.COL_ASISTENCIA] || "").toUpperCase().trim();
      const res = String(row[CONFIG.COL_RESULTADO] || "").trim();
      let link1 = String(row[colAdj1] || "").trim();
      let link2 = String(row[colAdj2] || "").trim();
      let guardado = String(row[colG] || "").trim();
      let linkCarpeta = String(row[CONFIG.COL_LINK_CARPETA] || "").trim();

      const esAusente = _normalizarTexto_(res).includes('ausente') || asistencia === 'AUSENTE';
      if (esAusente) {
        link1 = "";
        link2 = "";
        guardado = "NO";
        linkCarpeta = "";
      }

      let fechaAud = _parseFechaConTz_(row[CONFIG.COL_FECHA_AUDITORIA], tz);
      if (res !== '' && !fechaAud) {
        fechaAud = _parseFechaConTz_(new Date(), tz);
      }
      const fechaEst = _parseFechaConTz_(row[CONFIG.COL_FECHA_ESTUDIO], tz);
      const obs = String(row[CONFIG.COL_OBSERVACION] || "").trim();
      const comunicacion = _parseFechaConTz_(row[CONFIG.COL_COMUNICACION], tz);
      const establecimiento = String(row[CONFIG.COL_NOMBRE_ESTABLECIMIENTO] || "").trim();
      const srt = String(row[CONFIG.COL_SRT] || "").trim();
      const prestador = String(row[colPrestadorMadre] || "").trim();
      const frecuencia = String(row[CONFIG.COL_FRECUENCIA] || "").trim();

      const colClaveAN = CONFIG.COL_CLAVE_EMPLEADO !== undefined ? CONFIG.COL_CLAVE_EMPLEADO : 39;
      const claveMadre = String(row[colClaveAN] || "").trim();
      const clave = claveMadre || `${cuil}_${srt || 'SIN_SRT'}`;

      if (!mapa.has(clave)) {
        mapa.set(clave, {
          clave: clave,
          cuil: cuil,
          apeNom: String(row[CONFIG.COL_APELLIDO] || "").trim() || "Sin Nombre",
          contrato: String(row[CONFIG.COL_CUIT_CONTRATO] || "").trim(),
          nombreContrato: String(row[CONFIG.COL_NOMBRE_CONTRATO] || "").trim(),
          srt: srt,
          establecimiento: establecimiento,
          prestador: prestador,
          frecuencia: frecuencia,
          organismo: String(row[CONFIG.COL_NOMBRE_ORGANISMO] || "").trim(),
          totalEstudios: 0,
          auditadosEstudios: 0,
          guardado: 'NO',
          linkCarpeta: '',
          linkAdjunto: '',
          fileId: '',
          fechaAuditoria: '',
          fechaEstudio: '',
          filasIndices: [],
          estudios: []
        });
      }

      const emp = mapa.get(clave);
      emp.totalEstudios++;
      if (prestador && !emp.prestador) emp.prestador = prestador;
      if (frecuencia && !emp.frecuencia) emp.frecuencia = frecuencia;
      if (res !== '') emp.auditadosEstudios++;
      if (_esGuardado_(guardado)) emp.guardado = 'SÍ';
      if (linkCarpeta && !emp.linkCarpeta) emp.linkCarpeta = linkCarpeta;

      const linkUsable = link1 || link2;
      if (linkUsable && !emp.linkAdjunto) {
        emp.linkAdjunto = linkUsable;
        emp.fileId = _extraerFileId_(linkUsable);
      }
      if (fechaAud && !emp.fechaAuditoria) emp.fechaAuditoria = fechaAud;
      if (fechaEst && !emp.fechaEstudio) emp.fechaEstudio = fechaEst;

      emp.filasIndices.push(filaMadre);
      emp.estudios.push({
        rowIndex: filaMadre,
        estudio: estudio,
        prestador: prestador,
        resultado: res,
        observacion: obs,
        fechaEstudio: fechaEst,
        fechaAuditoria: fechaAud,
        adjunto1: link1,
        adjunto2: link2,
        comunicacion: comunicacion,
        guardado: guardado,
        linkCarpeta: linkCarpeta,
        esDesvio: _esResultadoDesvio_(res),
        esReclamable: _esResultadoReclamo_(res)
      });
    }

    const todosEmp = Array.from(mapa.values());
    todosEmp.forEach(e => {
      const algunEstGuardado = (e.estudios || []).some(est => _esGuardado_(est.guardado));
      if (algunEstGuardado) {
        e.guardado = 'SÍ';
        (e.estudios || []).forEach(est => {
          if (est && !_normalizarTexto_(est.resultado || '').includes('ausente') && _esGuardado_(est.guardado)) {
            est.guardado = 'SÍ';
          } else {
            est.guardado = '';
          }
        });
      } else {
        e.guardado = 'NO';
        (e.estudios || []).forEach(est => {
          est.guardado = '';
        });
      }
    });

    const empsParaHojaEmp = todosEmp.filter(e => e.auditadosEstudios > 0 || _esGuardado_(e.guardado) || Boolean(e.linkAdjunto) || (e.estudios || []).some(est => est.adjunto1 || est.adjunto2));
    const filasEmp = empsParaHojaEmp.map(e => {
      const tieneNoApto = e.estudios.some(est => est.resultado && !_esResultadoApto_(est.resultado));
      let estadoAud = "Sin auditar";
      if (tieneNoApto) {
        estadoAud = "Incompleto";
      } else if (e.totalEstudios > 0 && e.auditadosEstudios === e.totalEstudios) {
        estadoAud = "100% Auditado";
      } else if (e.auditadosEstudios > 0) {
        estadoAud = "Auditoría Parcial";
      }

      const compactos = e.estudios.map(est => ({
        rowIndex: est.rowIndex,
        estudio: est.estudio,
        prestador: est.prestador || e.prestador || '',
        resultado: est.resultado,
        observacion: est.observacion,
        fechaEstudio: est.fechaEstudio,
        fechaAuditoria: est.fechaAuditoria,
        guardado: est.guardado,
        linkCarpeta: est.linkCarpeta,
        adjunto1: est.adjunto1 || '',
        adjunto2: est.adjunto2 || '',
        comunicacion: est.comunicacion || ''
      }));

      return [
        e.cuil, e.apeNom, e.contrato, e.nombreContrato, e.srt, e.establecimiento, e.organismo,
        e.totalEstudios, e.auditadosEstudios, estadoAud, e.guardado, e.linkCarpeta, e.linkAdjunto, e.fileId,
        e.fechaAuditoria, e.fechaEstudio, e.filasIndices.join(","), JSON.stringify(compactos),
        e.prestador || '', e.clave || (e.cuil + '_' + e.srt)
      ];
    });

    _volcarFilasEnHoja_(hojaEmp, filasEmp);
    SpreadsheetApp.flush();

    _invalidarCache_();
    try {
      CacheService.getScriptCache().remove('metricas_links_empleados_v3');
      PropertiesService.getScriptProperties().deleteProperty('METRICAS_LINKS_EMPLEADOS_V3');
    } catch(eProp) {}

    registrarEnBitacora('SINCRONIZACION_HOJA_EMPLEADOS', { empleados: filasEmp.length });

    return {
      status: "OK",
      empleados: filasEmp.length
    };
  } catch (err) {
    console.error("Error en sincronizarHojaEmpleados: " + err.message);
    return { status: "Error", msg: err.message };
  } finally {
    if (lock) lock.releaseLock();
  }
}

/**
 * Calcula las métricas globales consolidadas de empleados con link,
 * faltantes de auditar y presentes en asistencia sin link.
 * Utiliza lectura columnar ultrarrápida (1-1.5s para 90k filas) y cache multinivel.
 */
function obtenerMetricasLinksEmpleados(forceRefresh) {
  try {
    const cache = CacheService.getScriptCache();
    if (!forceRefresh) {
      const cached = cache.get('metricas_links_empleados_v3');
      if (cached) {
        try { return JSON.parse(cached); } catch(e) {}
      }
      const prop = PropertiesService.getScriptProperties().getProperty('METRICAS_LINKS_EMPLEADOS_V3');
      if (prop) {
        try {
          const parsed = JSON.parse(prop);
          cache.put('metricas_links_empleados_v3', prop, 1800); // 30 min
          return parsed;
        } catch(e) {}
      }
    }

    const ss = _getSpreadsheet_();
    const sheetMadre = ss.getSheetByName(CONFIG.HOJA_MADRE_NAME);
    if (!sheetMadre || sheetMadre.getLastRow() < 2) {
      return {
        totalConLink: 0,
        auditadosConLink: 0,
        faltanAuditar: 0,
        presentesSinLink: 0,
        totalAuditados: 0,
        totalArchivados: 0,
        pendArchivar: 0
      };
    }

    const ultFila = sheetMadre.getLastRow();
    const numRows = ultFila - 1;

    // Lectura columnar segura (evita excepciones si la hoja tiene menos columnas que las configuradas)
    const maxCols = sheetMadre.getLastColumn();
    const getColVals = (col1Based) => {
      if (col1Based <= maxCols) {
        try {
          return sheetMadre.getRange(2, col1Based, numRows, 1).getValues();
        } catch(e) {
          return null;
        }
      }
      return null;
    };

    const colClaveAN = (CONFIG.COL_CLAVE_EMPLEADO !== undefined ? CONFIG.COL_CLAVE_EMPLEADO : 39) + 1;
    const colG       = (CONFIG.COL_GUARDADO !== undefined ? CONFIG.COL_GUARDADO : 37) + 1;
    const colSRT     = (CONFIG.COL_SRT !== undefined ? CONFIG.COL_SRT : 4) + 1;

    const cuilVals  = getColVals(CONFIG.COL_CUIL + 1);
    const srtVals   = getColVals(colSRT);
    const claveVals = getColVals(colClaveAN);
    const estVals   = getColVals(CONFIG.COL_ESTUDIO + 1);
    const asistVals = getColVals(CONFIG.COL_ASISTENCIA + 1);
    const resVals   = getColVals(CONFIG.COL_RESULTADO + 1);
    const adj1Vals  = getColVals(CONFIG.COL_ADJUNTO_1 + 1);
    const adj2Vals  = getColVals(CONFIG.COL_ADJUNTO_2 + 1);
    const gdVals    = getColVals(colG);

    // Mapeo consolidado por legajo único (Columna AN / Clave SRT)
    const empMap = new Map();

    for (let i = 0; i < numRows; i++) {
      const cuil = cuilVals ? String(cuilVals[i][0] || '').trim() : '';
      if (!cuil) continue;

      const estudio = estVals ? String(estVals[i][0] || '').trim() : '';
      if (estudio.includes('No posee')) continue;

      const srt = srtVals ? String(srtVals[i][0] || '').trim() : '';
      const claveMadre = (claveVals && claveVals[i]) ? String(claveVals[i][0] || '').trim() : '';
      const clave = claveMadre || `${cuil}_${srt || 'SIN_SRT'}`;

      const link1 = (adj1Vals && adj1Vals[i]) ? String(adj1Vals[i][0] || '').trim() : '';
      const link2 = (adj2Vals && adj2Vals[i]) ? String(adj2Vals[i][0] || '').trim() : '';
      const tieneLink = Boolean(link1 || link2);

      const res = (resVals && resVals[i]) ? String(resVals[i][0] || '').trim() : '';
      const tieneRes = Boolean(res);

      const asist = (asistVals && asistVals[i]) ? String(asistVals[i][0] || '').trim().toUpperCase() : '';
      const esPresente = asist.includes('PRESENTE');

      const gd = (gdVals && gdVals[i]) ? gdVals[i][0] : '';
      const esGd = _esGuardado_(gd);

      if (!empMap.has(clave)) {
        empMap.set(clave, {
          totalEstudios: 0,
          auditadosEstudios: 0,
          tieneLink: false,
          tienePresente: false,
          tieneGuardado: false,
          tieneNoApto: false
        });
      }

      const item = empMap.get(clave);
      item.totalEstudios++;
      if (tieneRes) {
        item.auditadosEstudios++;
        if (!_esResultadoApto_(res)) {
          item.tieneNoApto = true;
        }
      }
      if (asist === 'AUSENTE') item.tieneNoApto = true;
      if (tieneLink) item.tieneLink = true;
      if (esPresente) item.tienePresente = true;
      if (esGd) item.tieneGuardado = true;
    }

    let totalEmpleados = empMap.size;
    let totalConLink = 0;
    let auditadosConLink = 0;
    let faltanAuditar = 0;             // Todos los AN con link que NO tienen resultado
    let enProcesoConLink = 0;          // Con link pero con resultado parcial
    let presentesSinLink = 0;
    let totalAuditados = 0;
    let totalArchivados = 0;
    let pendArchivar = 0;
    let totalIncompletos = 0;
    let totalParciales = 0;

    empMap.forEach((item) => {
      const es100Auditado = (item.totalEstudios > 0 && item.auditadosEstudios >= item.totalEstudios);
      const esParcial = (item.auditadosEstudios > 0 && item.auditadosEstudios < item.totalEstudios);
      const esPendAuditar = item.tieneLink && (item.auditadosEstudios === 0);

      if (item.tieneLink) totalConLink++;
      if (esPendAuditar) faltanAuditar++;
      if (es100Auditado) totalAuditados++;

      if (item.tieneNoApto) {
        totalIncompletos++;
      } else if (esParcial) {
        totalParciales++;
        if (item.tieneLink) enProcesoConLink++;
      } else if (es100Auditado && item.tieneLink) {
        auditadosConLink++;
        if (item.tieneGuardado) {
          totalArchivados++;
        } else {
          pendArchivar++;
        }
      }

      if (!item.tieneLink && item.tienePresente) {
        presentesSinLink++;
      }
    });

    const resObj = {
      totalEmpleados: totalEmpleados,
      totalConLink: totalConLink,
      auditadosConLink: auditadosConLink,
      faltanAuditar: faltanAuditar,
      enProcesoConLink: enProcesoConLink,
      presentesSinLink: presentesSinLink,
      totalAuditados: totalAuditados,
      totalArchivados: totalArchivados,
      pendArchivar: pendArchivar,
      totalIncompletos: totalIncompletos,
      totalParciales: totalParciales,
      actualizadoEl: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss")
    };

    const jsonStr = JSON.stringify(resObj);
    try {
      cache.put('metricas_links_empleados_v3', jsonStr, 1800);
      PropertiesService.getScriptProperties().setProperty('METRICAS_LINKS_EMPLEADOS_V3', jsonStr);
    } catch(e) {}

    return resObj;
  } catch (err) {
    console.error("Error en obtenerMetricasLinksEmpleados: " + err.message);
    return {
      totalConLink: 0,
      auditadosConLink: 0,
      faltanAuditar: 0,
      presentesSinLink: 0,
      totalAuditados: 0,
      totalArchivados: 0,
      pendArchivar: 0,
      error: err.message
    };
  }
}

function _obtenerMetricasLinksCacheRapido_() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('metricas_links_empleados_v3');
    if (cached) return JSON.parse(cached);
    const prop = PropertiesService.getScriptProperties().getProperty('METRICAS_LINKS_EMPLEADOS_V3');
    if (prop) {
      try {
        const parsed = JSON.parse(prop);
        cache.put('metricas_links_empleados_v3', prop, 21600); // 6 horas
        return parsed;
      } catch(e) {}
    }
  } catch(e) {}
  return null;
}

/**
 * Carga los datos de empleados leyendo DIRECTAMENTE de la hoja dedicada "Empleados".
 * Es ultra-rápida (menos de 0.3s) y de tamaño liviano (evita bloqueos en pantalla).
 */
function cargarDatosEmpleados() {
  const ss = _getSpreadsheet_();
  let hoja = ss.getSheetByName(CONFIG.HOJA_EMPLEADOS);

  const metricasLinks = _obtenerMetricasLinksCacheRapido_();

  if (!hoja || hoja.getLastRow() < 2) {
    return {
      empleados: [],
      metricasLinks: metricasLinks
    };
  }

  // Asegurar que las columnas Prestador y Clave_Unica existan en el encabezado
  if (typeof inicializarHojaEmpleados === 'function') {
    try {
      inicializarHojaEmpleados();
    } catch(eInit) {
      console.warn("Aviso inicializando columnas en Empleados: " + eInit.message);
    }
  }

  // Auto-sincronización única correctiva para purgar registros residuales de versiones anteriores
  const propSync = PropertiesService.getScriptProperties().getProperty('SYNC_EMPLEADOS_CLEAN_V5');
  if (!propSync) {
    try {
      sincronizarTodasLasHojasDesdeMadre(true);
      PropertiesService.getScriptProperties().setProperty('SYNC_EMPLEADOS_CLEAN_V5', 'true');
      hoja = ss.getSheetByName(CONFIG.HOJA_EMPLEADOS);
    } catch(eSync) {
      console.warn("Auto-sync correctivo omitido: " + eSync.message);
    }
  }

  const ultFila = hoja.getLastRow();
  const actualCols = hoja.getLastColumn();
  const headers = hoja.getRange(1, 1, 1, actualCols).getValues()[0].map(h => String(h || '').trim());
  const col = {};
  headers.forEach((h, idx) => {
    if (h) col[h.toLowerCase()] = idx;
  });

  const getV = (r, name, fallbackIdx) => {
    const idx = col[name.toLowerCase()] !== undefined ? col[name.toLowerCase()] : fallbackIdx;
    return (idx !== undefined && idx < r.length && r[idx] !== undefined && r[idx] !== null) ? r[idx] : '';
  };

  const rows = hoja.getRange(2, 1, ultFila - 1, actualCols).getValues();

  const empleados = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const cuil = String(getV(r, "CUIL", 0)).trim();
    if (!cuil) continue;

    let estudios = [];
    const detRaw = getV(r, "Detalle_Estudios", 17);
    if (detRaw) {
      try {
        estudios = typeof detRaw === 'string' ? JSON.parse(detRaw) : detRaw;
      } catch (e) {
        estudios = [];
      }
    }

    let filasIndices = [];
    const filRaw = getV(r, "Filas_Madre", 16);
    if (filRaw) {
      filasIndices = String(filRaw).split(',').map(n => parseInt(n.trim(), 10)).filter(n => !isNaN(n));
    }

    const establecimiento = String(getV(r, "Establecimiento", 5) || '');
    let prestador = String(getV(r, "Prestador", 18) || '').trim();

    // Fallback inteligente: si la celda Prestador está vacía, recuperar desde el JSON de estudios
    if (!prestador && Array.isArray(estudios) && estudios.length > 0) {
      const pList = estudios.map(est => String(est.prestador || '').trim()).filter(Boolean);
      if (pList.length > 0) {
        prestador = Array.from(new Set(pList)).join(', ');
      }
    }

    const totalEst = (Array.isArray(estudios) && estudios.length > 0)
      ? estudios.length
      : (Number(getV(r, "Total_Estudios", 7)) || 0);

    const audEst = (Array.isArray(estudios) && estudios.length > 0)
      ? estudios.filter(est => est && est.resultado && String(est.resultado).trim() !== '').length
      : (Number(getV(r, "Auditados_Estudios", 8)) || 0);

    const tieneNoApto = estudios.some(est => est.resultado && !_esResultadoApto_(est.resultado));
    const ausentesCount = estudios.filter(est => est.resultado && String(est.resultado).toLowerCase().includes('ausente')).length;
    const realizadosCount = (Array.isArray(estudios) && estudios.length > 0)
      ? estudios.filter(est => est.resultado && _esResultadoApto_(est.resultado)).length
      : audEst;

    const faltanEst = Math.max(0, totalEst - realizadosCount);

    let estadoAud = String(getV(r, "Estado_Auditoria", 9) || '');
    if (tieneNoApto) {
      estadoAud = 'Incompleto';
    } else if (totalEst > 0 && audEst === totalEst) {
      estadoAud = '100% Auditado';
    } else if (audEst > 0) {
      estadoAud = 'Auditoría Parcial';
    } else {
      estadoAud = 'Sin auditar';
    }

    const srtVal = String(getV(r, "Codigo_SRT", 4) || '').trim();
    const rawClave = String(getV(r, "Clave_Unica", 19) || '').trim();
    const claveVal = rawClave || `${cuil}_${srtVal || 'SIN_SRT'}`;

    const rawGuardado = String(getV(r, "Guardado", 10) || '').trim();
    let esGdEmp = _esGuardado_(rawGuardado);
    let errorArchivado = '';
    if (rawGuardado.toUpperCase().startsWith('ERROR')) {
      errorArchivado = rawGuardado.replace(/^ERROR:?\s*/i, '') || 'Error al archivar en Drive';
      esGdEmp = false;
    }
    if (!esGdEmp && !errorArchivado && Array.isArray(estudios)) {
      esGdEmp = estudios.some(est => _esGuardado_(est.guardado));
    }

    empleados.push({
      cuil: cuil,
      apeNom: String(getV(r, "Apellido_y_Nombre", 1) || ''),
      contrato: String(getV(r, "Contrato", 2) || ''),
      nombreContrato: String(getV(r, "Nombre_Contrato", 3) || ''),
      srt: srtVal,
      establecimiento: establecimiento,
      organismo: String(getV(r, "Nombre_Organismo", 6) || ''),
      totalEstudios: totalEst,
      auditadosEstudios: audEst,
      faltanEstudios: faltanEst,
      realizadosEstudios: realizadosCount,
      ausentesEstudios: ausentesCount,
      tieneAusente: tieneNoApto,
      tieneNoApto: tieneNoApto,
      estadoAuditoria: estadoAud,
      guardado: esGdEmp ? 'SÍ' : (errorArchivado ? 'ERROR' : 'NO'),
      errorArchivado: errorArchivado,
      linkCarpeta: String(getV(r, "Link_Carpeta", 11) || ''),
      linkAdjunto: String(getV(r, "Link_Adjunto", 12) || ''),
      fileId: String(getV(r, "File_Id", 13) || ''),
      fechaAuditoria: _parseFechaConTz_(getV(r, "Fecha_Auditoria", 14)),
      fechaEstudio: _parseFechaConTz_(getV(r, "Fecha_Estudio", 15)),
      filasIndices: filasIndices,
      estudios: estudios,
      prestador: prestador,
      clave: claveVal
    });
  }

  // Auto-recuperación ultrarrápida: si la hoja Empleados es legacy y le faltan prestadores o nombre de empleador
  const faltanDatos = empleados.some(e => !e.prestador || !e.nombreContrato);
  if (faltanDatos && empleados.length > 0) {
    try {
      const sheetMadre = _getSheetMadre_();
      if (sheetMadre && sheetMadre.getLastRow() >= 2) {
        const ultM = sheetMadre.getLastRow();
        const numColsM = Math.max(sheetMadre.getLastColumn(), 40);
        const headersM = sheetMadre.getRange(1, 1, 1, numColsM).getValues()[0].map(h => _normalizarTexto_(h));

        let colPrest = CONFIG.COL_PRESTADOR;
        for (let c = 0; c < headersM.length; c++) {
          if (headersM[c] === 'prestador' || headersM[c].includes('prestador')) {
            colPrest = c; break;
          }
        }

        let colNomContr = CONFIG.COL_NOMBRE_CONTRATO;
        for (let c = 0; c < headersM.length; c++) {
          if (headersM[c].includes('nombre') && (headersM[c].includes('contrato') || headersM[c].includes('empleador'))) {
            colNomContr = c; break;
          }
        }

        const cuilsM = sheetMadre.getRange(2, CONFIG.COL_CUIL + 1, ultM - 1, 1).getValues();
        const prestsM = sheetMadre.getRange(2, colPrest + 1, ultM - 1, 1).getValues();
        const nomContrM = sheetMadre.getRange(2, colNomContr + 1, ultM - 1, 1).getValues();

        const mapaPrestMadre = new Map();
        const mapaNomContrMadre = new Map();
        for (let i = 0; i < cuilsM.length; i++) {
          const c = String(cuilsM[i][0] || '').trim();
          const p = String(prestsM[i][0] || '').trim();
          const nc = String(nomContrM[i][0] || '').trim();
          if (c && p && !mapaPrestMadre.has(c)) mapaPrestMadre.set(c, p);
          if (c && nc && !mapaNomContrMadre.has(c)) mapaNomContrMadre.set(c, nc);
        }

        empleados.forEach(e => {
          if (!e.prestador && mapaPrestMadre.has(e.cuil)) {
            e.prestador = mapaPrestMadre.get(e.cuil);
          }
          if (!e.nombreContrato && mapaNomContrMadre.has(e.cuil)) {
            e.nombreContrato = mapaNomContrMadre.get(e.cuil);
          }
        });
      }
    } catch (eAutoP) {
      console.warn("Aviso auto-recuperando datos de Madre: " + eAutoP.message);
    }
  }

  return {
    empleados: empleados,
    metricasLinks: metricasLinks
  };
}

/**
 * Actualiza en tiempo real los cambios de auditoría en la hoja dedicada "Empleados".
 */
function _actualizarAuditoriaEnHojaEmpleados_(sheetMadre, cambios, fechaHoy, tz) {
  try {
    const ss = _getSpreadsheet_();
    const hojaEmp = ss.getSheetByName(CONFIG.HOJA_EMPLEADOS);
    if (!hojaEmp || hojaEmp.getLastRow() < 2) return;

    const ultFila = hojaEmp.getLastRow();
    const numCols = Math.max(hojaEmp.getLastColumn(), 20);
    const rows = hojaEmp.getRange(2, 1, ultFila - 1, numCols).getValues();

    // Mapear cada fila de la hoja Empleados por los índices de filas Madre que contiene
    const mapaFilaMadreAIdxHoja = new Map();
    for (let i = 0; i < rows.length; i++) {
      const strFilas = String(rows[i][16] || '');
      if (strFilas) {
        strFilas.split(',').forEach(n => {
          const idx = parseInt(n.trim(), 10);
          if (!isNaN(idx)) mapaFilaMadreAIdxHoja.set(idx, i);
        });
      }
    }

    const filasModificadasHoja = new Set();
    const cuilesNuevosProcesar = new Set();
    const fHoyStr = _parseFechaConTz_(fechaHoy, tz);

    cambios.forEach(c => {
      const rIdx = mapaFilaMadreAIdxHoja.get(Number(c.idx));
      if (rIdx === undefined) {
        const cuilMadre = String(sheetMadre.getRange(Number(c.idx), CONFIG.COL_CUIL + 1).getValue() || '').trim();
        if (cuilMadre) cuilesNuevosProcesar.add(cuilMadre);
        return;
      }

      const rowEmp = rows[rIdx];
      let estudios = [];
      try {
        estudios = typeof rowEmp[17] === 'string' ? JSON.parse(rowEmp[17]) : rowEmp[17];
      } catch (e) {
        estudios = [];
      }

      let cambioAplicado = false;
      estudios.forEach(est => {
        if (Number(est.rowIndex) === Number(c.idx)) {
          est.resultado = c.res || "";
          est.observacion = c.obs || "";
          if (c.fechaEstudio) est.fechaEstudio = c.fechaEstudio;
          est.fechaAuditoria = fHoyStr;
          cambioAplicado = true;
        }
      });

      if (cambioAplicado) {
        rowEmp[17] = JSON.stringify(estudios);
        const auditados = estudios.filter(e => e.resultado && e.resultado.trim() !== '').length;
        rowEmp[8] = auditados; // Col 9 (idx 8): Auditados_Estudios
        const total = Number(rowEmp[7]) || estudios.length;
        const tieneAusente = estudios.some(e => e.resultado && _normalizarTexto_(e.resultado).includes('ausente'));
        let estAud = "Sin auditar";
        if (total > 0 && auditados === total) {
          estAud = tieneAusente ? "Incompleto (con Ausentes)" : "100% Auditado";
        } else if (auditados > 0) {
          estAud = tieneAusente ? "Incompleto (con Ausentes)" : "Auditoría Parcial";
        }
        rowEmp[9] = estAud;
        rowEmp[14] = fHoyStr; // Col 15 (idx 14): Fecha_Auditoria
        if (c.fechaEstudio) rowEmp[15] = c.fechaEstudio; // Col 16 (idx 15): Fecha_Estudio

        // Si la columna Prestador (idx 18) está vacía, poblarla desde los estudios
        if (!rowEmp[18] && estudios.length > 0) {
          const p = estudios.map(x => x.prestador).filter(Boolean)[0];
          if (p) rowEmp[18] = p;
        }

        filasModificadasHoja.add(rIdx);
      }
    });

    filasModificadasHoja.forEach(rIdx => {
      const filaSheet = rIdx + 2;
      hojaEmp.getRange(filaSheet, 1, 1, numCols).setValues([rows[rIdx]]);
    });

    // Si hubo CUILs que no estaban en la hoja Empleados, agregarlos
    cuilesNuevosProcesar.forEach(cuil => {
      _agregarEmpleadoAHojaEmpleados_(sheetMadre, cuil, hojaEmp, tz);
    });
  } catch (err) {
    console.error("Error en _actualizarAuditoriaEnHojaEmpleados_: " + err.message);
  }
}

function _agregarEmpleadoAHojaEmpleados_(sheetMadre, cuil, hojaEmp, tz) {
  try {
    if (!cuil || !hojaEmp) return;
    const ultFila = sheetMadre.getLastRow();
    if (ultFila < 2) return;

    const numCols = Math.max(sheetMadre.getLastColumn(), 40);
    const headersMadre = sheetMadre.getRange(1, 1, 1, numCols).getValues()[0].map(h => String(h || '').trim().toLowerCase());
    let colPrestador = CONFIG.COL_PRESTADOR;
    const idxP = headersMadre.indexOf('prestador');
    if (idxP !== -1) colPrestador = idxP;

    const cuilValues = sheetMadre.getRange(2, CONFIG.COL_CUIL + 1, ultFila - 1, 1).getValues();
    const indicesEmp = [];
    for (let i = 0; i < cuilValues.length; i++) {
      if (String(cuilValues[i][0] || '').trim() === cuil) {
        indicesEmp.push(i + 2);
      }
    }
    if (indicesEmp.length === 0) return;

    let empData = null;
    const estudiosCompactos = [];
    let auditados = 0;
    let guardado = 'NO';
    let linkCarpeta = '';
    let linkAdjunto = '';
    let fileId = '';
    let fechaAud = '';
    let fechaEst = '';

    indicesEmp.forEach(rIdx => {
      const row = sheetMadre.getRange(rIdx, 1, 1, numCols).getValues()[0];
      const estudio = String(row[CONFIG.COL_ESTUDIO] || '').trim();
      if (estudio.includes("No posee")) return;

      const res = String(row[CONFIG.COL_RESULTADO] || '').trim();
      const lk1 = String(row[CONFIG.COL_ADJUNTO_1] || '').trim();
      const lk2 = String(row[CONFIG.COL_ADJUNTO_2] || '').trim();
      const gd  = String(row[CONFIG.COL_GUARDADO] || '').trim();
      const lc  = String(row[CONFIG.COL_LINK_CARPETA] || '').trim();
      const fa  = _parseFechaConTz_(row[CONFIG.COL_FECHA_AUDITORIA], tz);
      const fe  = _parseFechaConTz_(row[CONFIG.COL_FECHA_ESTUDIO], tz);
      const obs = String(row[CONFIG.COL_OBSERVACION] || '').trim();
      const prest = String(row[colPrestador] || '').trim();

      if (!empData) {
        empData = {
          cuil: cuil,
          apeNom: String(row[CONFIG.COL_APELLIDO] || '').trim() || 'Sin Nombre',
          contrato: String(row[CONFIG.COL_CUIT_CONTRATO] || '').trim(),
          nombreContrato: String(row[CONFIG.COL_NOMBRE_CONTRATO] || '').trim(),
          srt: String(row[CONFIG.COL_SRT] || '').trim(),
          establecimiento: String(row[CONFIG.COL_NOMBRE_ESTABLECIMIENTO] || '').trim(),
          organismo: String(row[CONFIG.COL_NOMBRE_ORGANISMO] || '').trim(),
          prestador: prest
        };
      }
      if (prest && !empData.prestador) empData.prestador = prest;

      if (res !== '') auditados++;
      if (_esGuardado_(gd)) guardado = 'SÍ';
      if (lc && !linkCarpeta) linkCarpeta = lc;
      const lk = lk1 || lk2;
      if (lk && !linkAdjunto) {
        linkAdjunto = lk;
        fileId = _extraerFileId_(lk);
      }
      if (fa && !fechaAud) fechaAud = fa;
      if (fe && !fechaEst) fechaEst = fe;

      estudiosCompactos.push({
        rowIndex: rIdx,
        estudio: estudio,
        prestador: prest,
        resultado: res,
        observacion: obs,
        fechaEstudio: fe,
        fechaAuditoria: fa,
        guardado: gd,
        linkCarpeta: lc
      });
    });

    if (!empData || estudiosCompactos.length === 0) return;
    if (auditados === 0 && !_esGuardado_(guardado)) return;

    const total = estudiosCompactos.length;
    const tieneAusente = estudiosCompactos.some(e => e.resultado && _normalizarTexto_(e.resultado).includes('ausente'));
    let estadoAud = "Sin auditar";
    if (total > 0 && auditados === total) {
      estadoAud = tieneAusente ? "Incompleto (con Ausentes)" : "100% Auditado";
    } else if (auditados > 0) {
      estadoAud = tieneAusente ? "Incompleto (con Ausentes)" : "Auditoría Parcial";
    }

    const nuevaFila = [
      empData.cuil,
      empData.apeNom,
      empData.contrato,
      empData.nombreContrato,
      empData.srt,
      empData.establecimiento,
      empData.organismo,
      total,
      auditados,
      estadoAud,
      guardado,
      linkCarpeta,
      linkAdjunto,
      fileId,
      fechaAud,
      fechaEst,
      indicesEmp.join(","),
      JSON.stringify(estudiosCompactos),
      empData.prestador || '',
      empData.cuil + '_' + (empData.srt || '')
    ];

    hojaEmp.appendRow(nuevaFila);
  } catch (err) {
    console.error("Error agregando nuevo empleado a hoja Empleados: " + err.message);
  }
}

// [2026-09-06] Obtener URL de la carpeta del establecimiento en Drive (retorna null si no está creada)
function obtenerUrlCarpetaEstablecimiento(cuit, contrato, srt, estab) {
  try {
    if (!CONFIG.CARPETA_PDFS_ID) return null;
    const cache = CacheService.getScriptCache();
    const rawKey = `${cuit || ''}_${contrato || ''}_${srt || ''}_${estab || ''}`;
    const cacheKey = 'fld_' + Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, rawKey)).slice(0, 30);
    const cachedUrl = cache.get(cacheKey);
    if (cachedUrl) return cachedUrl === 'NONE' ? null : cachedUrl;

    const carpetaRaiz = DriveApp.getFolderById(CONFIG.CARPETA_PDFS_ID);
    
    // Nivel 1: C-D
    const rutaL1 = `${cuit || ''}-${contrato || ''}`.replace(/^-|-$/g, '').trim();
    if (!rutaL1) { cache.put(cacheKey, 'NONE', 600); return null; }
    const carpetasL1 = carpetaRaiz.getFoldersByName(rutaL1);
    if (!carpetasL1.hasNext()) { cache.put(cacheKey, 'NONE', 600); return null; }
    const carpetaL1 = carpetasL1.next();

    // Nivel 2: H-G
    const rutaL2 = `${srt || ''}-${estab || ''}`.replace(/^-|-$/g, '').trim();
    if (!rutaL2) { cache.put(cacheKey, 'NONE', 600); return null; }
    const carpetasL2 = carpetaL1.getFoldersByName(rutaL2);
    if (!carpetasL2.hasNext()) {
      cache.put(cacheKey, 'NONE', 600);
      return null; // [2026-09-06] Si no está creada, devuelve null y no abre nada
    }

    const url = carpetasL2.next().getUrl();
    cache.put(cacheKey, url, 21600); // Guardar en caché 6 horas
    return url;
  } catch (e) {
    console.error('Error buscando carpeta en Drive: ' + e.message);
    return null;
  }
}

// ==================================================================
// REPORTES DE AVANCE · BASE TABLA AUDITORÍAS (MADRE)
// ==================================================================

/**
 * Calcula las métricas consolidadas de avance por Prestador y globales:
 * - 90 días hábiles hasta el 13/11
 * - Prestadores: establecimientos totales, con al menos 1 presente, sin presentes
 * - Empleados asignados vs realizados (asistencia presente)
 * - Asistencia presente con adjunto vs sin adjunto
 * - Estudios asignados
 */
function obtenerDatosAvance() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('datosAvance_cache_v2');
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch(e) {}
    }

    const ss = _getSpreadsheet_();
    const sheetMadre = ss.getSheetByName(CONFIG.HOJA_MADRE_NAME);
    if (!sheetMadre || sheetMadre.getLastRow() < 2) {
      return {
        prestadores: [],
        kpis: {
          prestadores: 0,
          establecimientosTotal: 0,
          estabConPresente: 0,
          estabSinPresente: 0,
          pctEstabConPresente: 0,
          empleadosAsignados: 0,
          empleadosRealizados: 0,
          pctAvanceEmpleados: 0,
          estudiosAsignados: 0,
          presentesConAdjunto: 0,
          presentesSinAdjunto: 0,
          pctPresentesConAdjunto: 0
        },
        cronograma: _calcularCronograma90Dias_()
      };
    }

    const ultFila = sheetMadre.getLastRow();
    const numRows = ultFila - 1;

    // Lectura columnar rápida para evitar timeouts de memoria y tiempo en 90k filas
    const prestVals = sheetMadre.getRange(2, CONFIG.COL_PRESTADOR + 1, numRows, 1).getValues();
    const estabVals = sheetMadre.getRange(2, CONFIG.COL_NOMBRE_ESTABLECIMIENTO + 1, numRows, 1).getValues();
    const srtVals   = sheetMadre.getRange(2, CONFIG.COL_SRT + 1, numRows, 1).getValues();
    const cuilVals  = sheetMadre.getRange(2, CONFIG.COL_CUIL + 1, numRows, 1).getValues();
    const estVals   = sheetMadre.getRange(2, CONFIG.COL_ESTUDIO + 1, numRows, 1).getValues();
    const asistVals = sheetMadre.getRange(2, CONFIG.COL_ASISTENCIA + 1, numRows, 1).getValues();
    const adj1Vals  = sheetMadre.getRange(2, CONFIG.COL_ADJUNTO_1 + 1, numRows, 1).getValues();
    const adj2Vals  = sheetMadre.getRange(2, CONFIG.COL_ADJUNTO_2 + 1, numRows, 1).getValues();

    const prestMap = new Map();

    for (let i = 0; i < numRows; i++) {
      const cuil = String(cuilVals[i][0] || '').trim();
      if (!cuil) continue;

      const estudio = String(estVals[i][0] || '').trim();
      if (estudio.includes('No posee')) continue;

      const asist = String(asistVals[i][0] || '').trim().toUpperCase();

      const prest = String(prestVals[i][0] || '').trim() || 'Sin Prestador';
      const estab = String(estabVals[i][0] || '').trim() || 'Sin Establecimiento';
      const srt   = String(srtVals[i][0] || '').trim();
      const estabKey = estab + (srt ? ' (SRT: ' + srt + ')' : '');

      const esPresente = asist.includes('PRESENTE');
      const tieneAdjunto = Boolean(String(adj1Vals[i][0] || '').trim() || String(adj2Vals[i][0] || '').trim());

      if (!prestMap.has(prest)) {
        prestMap.set(prest, {
          nombre: prest,
          estudiosAsignados: 0,
          empleadosMap: new Map(), // cuil -> { presente: bool, conAdjunto: bool }
          estabMap: new Map()      // estabKey -> { tienePresente: bool, cuils: Set }
        });
      }

      const pData = prestMap.get(prest);
      pData.estudiosAsignados++;

      // Tracking empleado por prestador
      if (!pData.empleadosMap.has(cuil)) {
        pData.empleadosMap.set(cuil, { presente: false, conAdjunto: false });
      }
      const empData = pData.empleadosMap.get(cuil);
      if (esPresente) empData.presente = true;
      if (tieneAdjunto) empData.conAdjunto = true;

      // Tracking establecimiento por prestador
      if (!pData.estabMap.has(estabKey)) {
        pData.estabMap.set(estabKey, { tienePresente: false, cuils: new Set() });
      }
      const estObj = pData.estabMap.get(estabKey);
      estObj.cuils.add(cuil);
      if (esPresente) estObj.tienePresente = true;
    }

    const prestadoresList = [];
    let totEstudiosGlobal = 0;

    const estabsGlobalesTodos = new Set();
    const estabsGlobalesConPresente = new Set();
    const cuilsGlobalesAsignados = new Set();
    const cuilsGlobalesRealizados = new Set();
    const cuilsGlobalesConAdjunto = new Set();

    prestMap.forEach((pData, nombre) => {
      let empAsignados = pData.empleadosMap.size;
      let empRealizados = 0;
      let empConAdjunto = 0;
      let empSinAdjunto = 0;

      pData.empleadosMap.forEach((val, c) => {
        cuilsGlobalesAsignados.add(c);
        if (val.presente) {
          empRealizados++;
          cuilsGlobalesRealizados.add(c);
          if (val.conAdjunto) {
            empConAdjunto++;
            cuilsGlobalesConAdjunto.add(c);
          } else {
            empSinAdjunto++;
          }
        }
      });

      let estTotal = pData.estabMap.size;
      let estConPres = 0;
      let estSinPres = 0;

      pData.estabMap.forEach((val, eKey) => {
        estabsGlobalesTodos.add(eKey);
        if (val.tienePresente) {
          estConPres++;
          estabsGlobalesConPresente.add(eKey);
        } else {
          estSinPres++;
        }
      });

      const pctAvance = empAsignados > 0 ? Number(((empRealizados / empAsignados) * 100).toFixed(1)) : 0;
      const pctEstab = estTotal > 0 ? Number(((estConPres / estTotal) * 100).toFixed(1)) : 0;
      const pctAdj = empRealizados > 0 ? Number(((empConAdjunto / empRealizados) * 100).toFixed(1)) : 0;

      prestadoresList.push({
        nombre: nombre,
        establecimientosTotal: estTotal,
        estabConPresente: estConPres,
        estabSinPresente: estSinPres,
        pctEstabConPresente: pctEstab,
        empleadosAsignados: empAsignados,
        empleadosRealizados: empRealizados,
        pctAvanceEmpleados: pctAvance,
        estudiosAsignados: pData.estudiosAsignados,
        presentesConAdjunto: empConAdjunto,
        presentesSinAdjunto: empSinAdjunto,
        pctPresentesConAdjunto: pctAdj
      });

      totEstudiosGlobal += pData.estudiosAsignados;
    });

    prestadoresList.sort((a, b) => b.empleadosAsignados - a.empleadosAsignados);

    const cronograma = _calcularCronograma90Dias_();

    const totalEstabs = estabsGlobalesTodos.size;
    const totalEstabsConPres = estabsGlobalesConPresente.size;
    const totalEstabsSinPres = Math.max(0, totalEstabs - totalEstabsConPres);
    const totalEmpAsig = cuilsGlobalesAsignados.size;
    const totalEmpReal = cuilsGlobalesRealizados.size;
    const totalAdjCon = cuilsGlobalesConAdjunto.size;
    const totalAdjSin = Math.max(0, totalEmpReal - totalAdjCon);

    return {
      prestadores: prestadoresList,
      kpis: {
        prestadores: prestadoresList.length,
        establecimientosTotal: totalEstabs,
        estabConPresente: totalEstabsConPres,
        estabSinPresente: totalEstabsSinPres,
        pctEstabConPresente: totalEstabs > 0 ? Number(((totalEstabsConPres / totalEstabs) * 100).toFixed(1)) : 0,
        empleadosAsignados: totalEmpAsig,
        empleadosRealizados: totalEmpReal,
        pctAvanceEmpleados: totalEmpAsig > 0 ? Number(((totalEmpReal / totalEmpAsig) * 100).toFixed(1)) : 0,
        estudiosAsignados: totEstudiosGlobal,
        presentesConAdjunto: totalAdjCon,
        presentesSinAdjunto: totalAdjSin,
        pctPresentesConAdjunto: totalEmpReal > 0 ? Number(((totalAdjCon / totalEmpReal) * 100).toFixed(1)) : 0
      },
      cronograma: cronograma
    };

    try {
      cache.put('datosAvance_cache_v2', JSON.stringify(resObj), 1800);
    } catch(e) {}

    return resObj;

  } catch (err) {
    console.error("Error en obtenerDatosAvance: " + err.message);
    return {
      prestadores: [],
      kpis: {
        prestadores: 0,
        establecimientosTotal: 0,
        estabConPresente: 0,
        estabSinPresente: 0,
        pctEstabConPresente: 0,
        empleadosAsignados: 0,
        empleadosRealizados: 0,
        pctAvanceEmpleados: 0,
        estudiosAsignados: 0,
        presentesConAdjunto: 0,
        presentesSinAdjunto: 0,
        pctPresentesConAdjunto: 0
      },
      cronograma: _calcularCronograma90Dias_(),
      error: err.message
    };
  }
}

/**
 * Calcula el avance temporal de 90 días hábiles hasta el vencimiento improrrogable (13/11/2026).
 */
function _calcularCronograma90Dias_() {
  const feriados2026 = ['2026-07-09', '2026-08-17', '2026-10-12'];

  const esHabil = (d) => {
    const day = d.getDay();
    if (day === 0 || day === 6) return false;
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const dayNum = String(d.getDate()).padStart(2, '0');
    return !feriados2026.includes(y + '-' + m + '-' + dayNum);
  };

  const fechaFin = new Date(2026, 10, 13); // 13 Noviembre 2026
  let habiles = 0;
  let fechaInicio = new Date(fechaFin);

  while (habiles < 90) {
    if (esHabil(fechaInicio)) habiles++;
    if (habiles < 90) fechaInicio.setDate(fechaInicio.getDate() - 1);
  }

  const hoy = new Date();
  let transcurridos = 0;
  let restantes = 0;

  let cur = new Date(fechaInicio);
  while (cur <= hoy && cur <= fechaFin) {
    if (esHabil(cur)) transcurridos++;
    cur.setDate(cur.getDate() + 1);
  }

  cur = new Date(hoy);
  cur.setDate(cur.getDate() + 1);
  while (cur <= fechaFin) {
    if (esHabil(cur)) restantes++;
    cur.setDate(cur.getDate() + 1);
  }

  if (transcurridos > 90) transcurridos = 90;
  if (restantes < 0) restantes = 0;

  const pct = Number(((transcurridos / 90) * 100).toFixed(1));

  const fmt = (d) => {
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    return dd + '/' + mm + '/' + d.getFullYear();
  };

  return {
    diasTotales: 90,
    diasTranscurridos: transcurridos,
    diasRestantes: restantes,
    pctTranscurrido: pct,
    fechaInicio: fmt(fechaInicio),
    fechaFin: '13/11/2026'
  };
}

/**
 * Inicializa y da formato a la hoja física "Reportes Avance" en el Spreadsheet.
 */
function inicializarHojaAvance() {
  const ss = _getSpreadsheet_();
  const nombreHoja = CONFIG.HOJA_AVANCE || "Reportes Avance";
  let hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    hoja = ss.insertSheet(nombreHoja);
  }

  const headers = [
    "Prestador",
    "Establecimientos Totales",
    "Establecimientos c/ Presente",
    "Establecimientos s/ Presente",
    "% Cobertura Estab.",
    "Empleados Asignados",
    "Empleados Realizados (Presentes)",
    "% Avance Realización",
    "Estudios Asignados",
    "Presentes con Adjunto",
    "Presentes sin Adjunto",
    "% Presentes con Adjunto"
  ];

  if (hoja.getLastRow() < 1) {
    hoja.getRange(1, 1, 1, headers.length).setValues([headers]);
    hoja.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#1e293b")
      .setFontColor("#ffffff");
    hoja.setFrozenRows(1);
  }

  return hoja;
}

/**
 * Sincroniza y vuelca el resumen de avance en la pestaña física "Reportes Avance" de Google Sheets.
 */
function sincronizarHojaAvance() {
  try {
    CacheService.getScriptCache().remove('datosAvance_cache_v2');
    const data = obtenerDatosAvance();
    if (!data || !data.prestadores) {
      return { status: "Error", msg: "No se pudieron obtener datos de avance" };
    }

    const hoja = inicializarHojaAvance();
    const ultFila = hoja.getLastRow();
    if (ultFila > 1) {
      hoja.getRange(2, 1, ultFila - 1, hoja.getLastColumn()).clearContent();
    }

    const filas = [];
    data.prestadores.forEach(p => {
      filas.push([
        p.nombre,
        p.establecimientosTotal,
        p.estabConPresente,
        p.estabSinPresente,
        (p.pctEstabConPresente / 100),
        p.empleadosAsignados,
        p.empleadosRealizados,
        (p.pctAvanceEmpleados / 100),
        p.estudiosAsignados,
        p.presentesConAdjunto,
        p.presentesSinAdjunto,
        (p.pctPresentesConAdjunto / 100)
      ]);
    });

    if (filas.length > 0) {
      const k = data.kpis;
      filas.push([
        "TOTAL CONSOLIDADO",
        k.establecimientosTotal,
        k.estabConPresente,
        k.estabSinPresente,
        (k.pctEstabConPresente / 100),
        k.empleadosAsignados,
        k.empleadosRealizados,
        (k.pctAvanceEmpleados / 100),
        k.estudiosAsignados,
        k.presentesConAdjunto,
        k.presentesSinAdjunto,
        (k.pctPresentesConAdjunto / 100)
      ]);

      hoja.getRange(2, 1, filas.length, filas[0].length).setValues(filas);

      // Formato de porcentajes en columnas 5, 8 y 12
      hoja.getRange(2, 5, filas.length, 1).setNumberFormat("0.0%");
      hoja.getRange(2, 8, filas.length, 1).setNumberFormat("0.0%");
      hoja.getRange(2, 12, filas.length, 1).setNumberFormat("0.0%");

      // Resaltar fila de total consolidado
      const filaTot = filas.length + 1;
      hoja.getRange(filaTot, 1, 1, filas[0].length)
        .setFontWeight("bold")
        .setBackground("#f1f5f9");
    }

    return { status: "OK", prestadores: data.prestadores.length };
  } catch (err) {
    console.error("Error en sincronizarHojaAvance: " + err.message);
    return { status: "Error", msg: err.message };
  }
}

/**
 * Escanea la Hoja Madre completa y limpia cualquier fila donde Resultado sea Ausente/Sin realizar
 * o Asistencia sea AUSENTE, pero que erróneamente tenga links de adjunto o marca de Guardado / Carpeta.
 */
function limpiarInconsistenciasAusentesEnMadre(skipLock) {
  const lock = skipLock ? null : LockService.getScriptLock();
  try {
    if (lock) lock.waitLock(CONFIG.TIMEOUT_LOCK);
    const ss = _getSpreadsheet_();
    const sheetMadre = ss.getSheetByName(CONFIG.HOJA_MADRE_NAME);
    if (!sheetMadre) return { status: "Error", msg: "No se encontró la Hoja Madre" };

    const ultFila = sheetMadre.getLastRow();
    if (ultFila < 2) return { status: "OK", limpiadas: 0 };

    const colRes = CONFIG.COL_RESULTADO + 1;
    const colAsis = CONFIG.COL_ASISTENCIA + 1;
    const colAdj1 = CONFIG.COL_ADJUNTO_1 + 1;
    const colAdj2 = CONFIG.COL_ADJUNTO_2 + 1;
    const colG = (CONFIG.COL_GUARDADO !== undefined ? CONFIG.COL_GUARDADO : 37) + 1;
    const colCarpeta = (CONFIG.COL_LINK_CARPETA !== undefined ? CONFIG.COL_LINK_CARPETA : 38) + 1;

    const n = ultFila - 1;
    const resVals = sheetMadre.getRange(2, colRes, n, 1).getValues();
    const asisVals = sheetMadre.getRange(2, colAsis, n, 1).getValues();
    const adj1Vals = sheetMadre.getRange(2, colAdj1, n, 1).getValues();
    const adj2Vals = sheetMadre.getRange(2, colAdj2, n, 1).getValues();
    const guarVals = sheetMadre.getRange(2, colG, n, 1).getValues();
    const carpVals = sheetMadre.getRange(2, colCarpeta, n, 1).getValues();

    const filasALimpiar = [];
    for (let i = 0; i < n; i++) {
      const res = String(resVals[i][0] || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
      const asis = String(asisVals[i][0] || '').toUpperCase().trim();
      const esAusente = res.includes('ausente') || asis === 'AUSENTE';

      if (esAusente) {
        const tieneAdj = Boolean(adj1Vals[i][0] || adj2Vals[i][0]);
        const tieneGuar = Boolean(guarVals[i][0] || carpVals[i][0]);
        if (tieneAdj || tieneGuar) {
          filasALimpiar.push(i + 2);
        }
      }
    }

    if (filasALimpiar.length > 0) {
      filasALimpiar.forEach(idx => {
        sheetMadre.getRange(idx, colAdj1, 1, 2).setValues([["", ""]]); // Limpia Link1 y Link2
        sheetMadre.getRange(idx, colG, 1, 2).setValues([["", ""]]);    // Limpia Guardado y Link_Carpeta
      });
      SpreadsheetApp.flush();
      _invalidarCache_();
      registrarEnBitacora('LIMPIEZA_INCONSISTENCIAS_AUSENTES', { filas: filasALimpiar.length });
    }

    return { status: "OK", limpiadas: filasALimpiar.length };
  } catch (e) {
    console.error("Error en limpiarInconsistenciasAusentesEnMadre: " + e.message);
    return { status: "Error", msg: e.message };
  } finally {
    lock.releaseLock();
  }
}


// ==================================================================
// KPIS Y MÉTRICAS CONSOLIDADAS (ESTUDIOS, AN, H, C, CON/SIN PRESTADOR)
// ==================================================================

/**
 * Obtiene las métricas consolidadas de KPIs:
 * - Cantidad de estudios (filas totales)
 * - Cantidad de empleados totales según columna AN
 * - Cantidad de establecimientos (H) de esos AN
 * - Cantidad de empresas (C) de esos AN
 * - Segmentación con y sin prestador (O)
 * - Grado de cobertura de empleados (100% asignados, mixtos, sin asignar)
 * - Desglose detallado por prestador
 */
function obtenerMetricasKPIs(forzar) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'metricas_kpis_v2';

    if (!forzar) {
      // 1. Intentar memoria rápida de ScriptCache (~0.05s)
      const cached = cache.get(cacheKey);
      if (cached) {
        try {
          const obj = JSON.parse(cached);
          obj.origen = 'cache';
          return obj;
        } catch(e) {}
      }

      // 2. Intentar leer desde la hoja física "KPIs" (~0.15s)
      const desdeHoja = _leerMetricasKPIsDesdeHoja_();
      if (desdeHoja && desdeHoja.totales && desdeHoja.totales.estudios > 0) {
        desdeHoja.origen = 'hoja_kpis';
        try { cache.put(cacheKey, JSON.stringify(desdeHoja), 1800); } catch(e) {}
        return desdeHoja;
      }
    }

    // 3. Si se fuerza o no hay caché/hoja, escanear Madre
    return _calcularMetricasKPIsDesdeMadre_(true);
  } catch (err) {
    console.error("Error en obtenerMetricasKPIs: " + err.message);
    return {
      error: err.message,
      timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
      totales: { estudios: 0, empleadosAN: 0, establecimientosH: 0, empresasC: 0 },
      conPrestador: { estudios: 0, pctEstudios: 0, empleadosAN: 0, pctEmpleados: 0, establecimientosH: 0, pctEstablecimientos: 0, empresasC: 0, pctEmpresas: 0 },
      sinPrestador: { estudios: 0, pctEstudios: 0, empleadosAN: 0, pctEmpleados: 0, establecimientosH: 0, pctEstablecimientos: 0, empresasC: 0, pctEmpresas: 0 },
      exclusividadEmpleados: { asignados100: 0, pctAsignados100: 0, mixtos: 0, pctMixtos: 0, sinAsignar100: 0, pctSinAsignar100: 0 },
      prestadores: []
    };
  }
}

/**
 * Escanea la hoja Madre y calcula con total exactitud todas las cantidades,
 * cuadrando al 100% con segmentación exclusiva (100% con prestador, parciales y 100% sin prestador).
 */
function _calcularMetricasKPIsDesdeMadre_(guardarEnHoja) {
  const t0 = new Date().getTime();
  const ss = _getSpreadsheet_();
  const sheetMadre = ss.getSheetByName(CONFIG.HOJA_MADRE_NAME);
  if (!sheetMadre || sheetMadre.getLastRow() < 2) {
    return _obtenerObjetoKPIsVacio_();
  }

  const ultFila = sheetMadre.getLastRow();
  const numRows = ultFila - 1;

  // Bloque 1: Columnas 3 (C) a 15 (O) -> 13 columnas contiguas
  // Col 3:  Cuit_Contrato (offset 0)
  // Col 4:  Nombre_del_contrato (offset 1)
  // Col 7:  Nombre_del_Establecimiento (offset 4)
  // Col 8:  Codigo_SRT (offset 5)
  // Col 13: CUIL (offset 10)
  // Col 14: Apellido_y_Nombre (offset 11)
  // Col 15: Prestador (offset 12)
  const bloque1 = sheetMadre.getRange(2, 3, numRows, 13).getValues();

  // Bloque 2: Columna 40 (AN: Clave_Empleado_SRT)
  const colAN = (CONFIG.COL_CLAVE_EMPLEADO !== undefined ? CONFIG.COL_CLAVE_EMPLEADO : 39) + 1;
  const bloque2 = sheetMadre.getRange(2, colAN, numRows, 1).getValues();

  let totEstudios = 0;
  let conPrestEstudios = 0;
  let sinPrestEstudios = 0;

  const setEmpleadosTotal = new Set();
  const setEstabsTotal = new Set();
  const setEmpresasTotal = new Set();

  const setEmpleadosConPrest = new Set();
  const setEstabsConPrest = new Set();
  const setEmpresasConPrest = new Set();

  const setEmpleadosSinPrest = new Set();
  const setEstabsSinPrest = new Set();
  const setEmpresasSinPrest = new Set();

  // Mapas por entidad para evaluar grado de asignación exacto
  const mapaEmpresaGrado = new Map();
  const mapaEstabGrado = new Map();
  const mapaEmpleadoGrado = new Map();

  // Mapa de desglose por cada prestador
  const mapaPrestadores = new Map();

  for (let i = 0; i < numRows; i++) {
    const rawAN = String(bloque2[i][0] || '').trim();
    const cuit  = String(bloque1[i][0] || '').trim();
    const nombreEmpresa = String(bloque1[i][1] || '').trim() || cuit;
    const nombreEstab   = String(bloque1[i][4] || '').trim();
    const srt   = String(bloque1[i][5] || '').trim();
    const cuil  = String(bloque1[i][10] || '').trim();
    const nombreEmpleado = String(bloque1[i][11] || '').trim();
    const prestadorRaw   = String(bloque1[i][12] || '').trim();

    if (!rawAN && !cuil && !cuit && !srt) continue;
    const an = rawAN || (cuil && srt ? (cuil + '_' + srt) : cuil);
    if (!an) continue;

    const pUpper = prestadorRaw.toUpperCase().trim();
    const tienePrestador = Boolean(
      prestadorRaw &&
      pUpper !== '-' &&
      pUpper !== 'S/P' &&
      pUpper !== 'SIN PRESTADOR' &&
      pUpper !== 'SIN ASIGNAR' &&
      pUpper !== 'PENDIENTE' &&
      pUpper !== 'NO ASIGNADO' &&
      pUpper !== 'NO' &&
      pUpper !== 'N/A' &&
      pUpper !== 'VACIO' &&
      pUpper !== 'VACÍO'
    );

    const prestadorNombre = tienePrestador ? prestadorRaw : 'Sin Prestador Asignado';

    // 1. Totales generales
    totEstudios++;
    setEmpleadosTotal.add(an);
    if (srt) setEstabsTotal.add(srt);
    if (cuit) setEmpresasTotal.add(cuit);

    // 2. Grado Empresa (C)
    if (cuit) {
      if (!mapaEmpresaGrado.has(cuit)) {
        mapaEmpresaGrado.set(cuit, { cuit: cuit, nombre: nombreEmpresa, total: 0, con: 0, sin: 0, prestadoresSet: new Set() });
      }
      const eObj = mapaEmpresaGrado.get(cuit);
      eObj.total++;
      if (tienePrestador) {
        eObj.con++;
        eObj.prestadoresSet.add(prestadorRaw);
      } else {
        eObj.sin++;
      }
    }

    // 3. Grado Establecimiento (H)
    if (srt) {
      if (!mapaEstabGrado.has(srt)) {
        mapaEstabGrado.set(srt, { srt: srt, nombre: nombreEstab, cuit: cuit, nombreEmpresa: nombreEmpresa, total: 0, con: 0, sin: 0, prestadoresSet: new Set() });
      }
      const sObj = mapaEstabGrado.get(srt);
      sObj.total++;
      if (tienePrestador) {
        sObj.con++;
        sObj.prestadoresSet.add(prestadorRaw);
      } else {
        sObj.sin++;
      }
    }

    // 4. Grado Empleado (AN)
    if (!mapaEmpleadoGrado.has(an)) {
      mapaEmpleadoGrado.set(an, { an: an, cuil: cuil, nombre: nombreEmpleado, srt: srt, nombreEstab: nombreEstab, cuit: cuit, nombreEmpresa: nombreEmpresa, total: 0, con: 0, sin: 0, prestadoresSet: new Set() });
    }
    const empG = mapaEmpleadoGrado.get(an);
    empG.total++;
    if (tienePrestador) {
      empG.con++;
      empG.prestadoresSet.add(prestadorRaw);
    } else {
      empG.sin++;
    }

    // 5. Segmentación a nivel Fila Estudio
    if (tienePrestador) {
      conPrestEstudios++;
      setEmpleadosConPrest.add(an);
      if (srt) setEstabsConPrest.add(srt);
      if (cuit) setEmpresasConPrest.add(cuit);

      if (!mapaPrestadores.has(prestadorNombre)) {
        mapaPrestadores.set(prestadorNombre, {
          nombre: prestadorNombre,
          estudios: 0,
          empleadosSet: new Set(),
          estabsSet: new Set(),
          empresasSet: new Set()
        });
      }
      const pObj = mapaPrestadores.get(prestadorNombre);
      pObj.estudios++;
      pObj.empleadosSet.add(an);
      if (srt) pObj.estabsSet.add(srt);
      if (cuit) pObj.empresasSet.add(cuit);
    } else {
      sinPrestEstudios++;
      setEmpleadosSinPrest.add(an);
      if (srt) setEstabsSinPrest.add(srt);
      if (cuit) setEmpresasSinPrest.add(cuit);
    }
  }

  // A. Exclusividad Empresas (C)
  let empresas100Con = 0;
  let empresasParciales = 0;
  let empresas100Sin = 0;
  const casosEmpresasParciales = [];

  mapaEmpresaGrado.forEach((val) => {
    if (val.con === val.total) {
      empresas100Con++;
    } else if (val.sin === val.total) {
      empresas100Sin++;
    } else {
      empresasParciales++;
      casosEmpresasParciales.push({
        cuit: val.cuit,
        nombre: val.nombre,
        total: val.total,
        con: val.con,
        sin: val.sin,
        pctCon: Number(((val.con / val.total) * 100).toFixed(1)),
        prestadores: Array.from(val.prestadoresSet).join(', ')
      });
    }
  });
  casosEmpresasParciales.sort((a, b) => b.total - a.total);

  // B. Exclusividad Establecimientos (H)
  let estabs100Con = 0;
  let estabsParciales = 0;
  let estabs100Sin = 0;
  const casosEstabsParciales = [];

  mapaEstabGrado.forEach((val) => {
    if (val.con === val.total) {
      estabs100Con++;
    } else if (val.sin === val.total) {
      estabs100Sin++;
    } else {
      estabsParciales++;
      casosEstabsParciales.push({
        srt: val.srt,
        nombre: val.nombre,
        cuit: val.cuit,
        nombreEmpresa: val.nombreEmpresa,
        total: val.total,
        con: val.con,
        sin: val.sin,
        pctCon: Number(((val.con / val.total) * 100).toFixed(1)),
        prestadores: Array.from(val.prestadoresSet).join(', ')
      });
    }
  });
  casosEstabsParciales.sort((a, b) => b.total - a.total);

  // C. Exclusividad Empleados (AN)
  let empAsignados100 = 0;
  let empMixtos = 0;
  let empSinAsignar100 = 0;
  const casosEmpleadosParciales = [];

  mapaEmpleadoGrado.forEach((val) => {
    if (val.con === val.total) {
      empAsignados100++;
    } else if (val.sin === val.total) {
      empSinAsignar100++;
    } else {
      empMixtos++;
      if (casosEmpleadosParciales.length < 500) {
        casosEmpleadosParciales.push({
          an: val.an,
          cuil: val.cuil,
          nombre: val.nombre,
          srt: val.srt,
          nombreEstab: val.nombreEstab,
          cuit: val.cuit,
          nombreEmpresa: val.nombreEmpresa,
          total: val.total,
          con: val.con,
          sin: val.sin,
          pctCon: Number(((val.con / val.total) * 100).toFixed(1)),
          prestadores: Array.from(val.prestadoresSet).join(', ')
        });
      }
    }
  });
  casosEmpleadosParciales.sort((a, b) => b.total - a.total);

  const totEmp = setEmpleadosTotal.size;
  const totEst = setEstabsTotal.size;
  const totCuit = setEmpresasTotal.size;

  // Lista de prestadores individuales
  const prestadoresList = [];
  mapaPrestadores.forEach((p) => {
    prestadoresList.push({
      nombre: p.nombre,
      estudios: p.estudios,
      pctEstudios: totEstudios > 0 ? Number(((p.estudios / totEstudios) * 100).toFixed(1)) : 0,
      empleadosAN: p.empleadosSet.size,
      pctEmpleados: totEmp > 0 ? Number(((p.empleadosSet.size / totEmp) * 100).toFixed(1)) : 0,
      establecimientosH: p.estabsSet.size,
      empresasC: p.empresasSet.size,
      promEstudiosPorEmpleado: p.empleadosSet.size > 0 ? Number((p.estudios / p.empleadosSet.size).toFixed(2)) : 0
    });
  });
  prestadoresList.sort((a, b) => b.estudios - a.estudios);

  const t1 = new Date().getTime();

  const resObj = {
    origen: 'hoja_madre',
    timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
    tiempoMs: t1 - t0,
    totales: {
      estudios: totEstudios,
      empleadosAN: totEmp,
      establecimientosH: totEst,
      empresasC: totCuit
    },
    // Filas para la tabla 100% cuadrada:
    asignados100: {
      estudios: conPrestEstudios,
      pctEstudios: totEstudios > 0 ? Number(((conPrestEstudios / totEstudios) * 100).toFixed(1)) : 0,
      empleadosAN: empAsignados100,
      pctEmpleados: totEmp > 0 ? Number(((empAsignados100 / totEmp) * 100).toFixed(1)) : 0,
      establecimientosH: estabs100Con,
      pctEstablecimientos: totEst > 0 ? Number(((estabs100Con / totEst) * 100).toFixed(1)) : 0,
      empresasC: empresas100Con,
      pctEmpresas: totCuit > 0 ? Number(((empresas100Con / totCuit) * 100).toFixed(1)) : 0
    },
    parciales: {
      estudios: 0,
      pctEstudios: 0,
      empleadosAN: empMixtos,
      pctEmpleados: totEmp > 0 ? Number(((empMixtos / totEmp) * 100).toFixed(1)) : 0,
      establecimientosH: estabsParciales,
      pctEstablecimientos: totEst > 0 ? Number(((estabsParciales / totEst) * 100).toFixed(1)) : 0,
      empresasC: empresasParciales,
      pctEmpresas: totCuit > 0 ? Number(((empresasParciales / totCuit) * 100).toFixed(1)) : 0
    },
    sinAsignar100: {
      estudios: sinPrestEstudios,
      pctEstudios: totEstudios > 0 ? Number(((sinPrestEstudios / totEstudios) * 100).toFixed(1)) : 0,
      empleadosAN: empSinAsignar100,
      pctEmpleados: totEmp > 0 ? Number(((empSinAsignar100 / totEmp) * 100).toFixed(1)) : 0,
      establecimientosH: estabs100Sin,
      pctEstablecimientos: totEst > 0 ? Number(((estabs100Sin / totEst) * 100).toFixed(1)) : 0,
      empresasC: empresas100Sin,
      pctEmpresas: totCuit > 0 ? Number(((empresas100Sin / totCuit) * 100).toFixed(1)) : 0
    },
    // Compatibilidad y detalle por fila
    conPrestador: {
      estudios: conPrestEstudios,
      pctEstudios: totEstudios > 0 ? Number(((conPrestEstudios / totEstudios) * 100).toFixed(1)) : 0,
      empleadosAN: setEmpleadosConPrest.size,
      pctEmpleados: totEmp > 0 ? Number(((setEmpleadosConPrest.size / totEmp) * 100).toFixed(1)) : 0,
      establecimientosH: setEstabsConPrest.size,
      pctEstablecimientos: totEst > 0 ? Number(((setEstabsConPrest.size / totEst) * 100).toFixed(1)) : 0,
      empresasC: setEmpresasConPrest.size,
      pctEmpresas: totCuit > 0 ? Number(((setEmpresasConPrest.size / totCuit) * 100).toFixed(1)) : 0
    },
    sinPrestador: {
      estudios: sinPrestEstudios,
      pctEstudios: totEstudios > 0 ? Number(((sinPrestEstudios / totEstudios) * 100).toFixed(1)) : 0,
      empleadosAN: setEmpleadosSinPrest.size,
      pctEmpleados: totEmp > 0 ? Number(((setEmpleadosSinPrest.size / totEmp) * 100).toFixed(1)) : 0,
      establecimientosH: setEstabsSinPrest.size,
      pctEstablecimientos: totEst > 0 ? Number(((setEstabsSinPrest.size / totEst) * 100).toFixed(1)) : 0,
      empresasC: setEmpresasSinPrest.size,
      pctEmpresas: totCuit > 0 ? Number(((setEmpresasSinPrest.size / totCuit) * 100).toFixed(1)) : 0
    },
    exclusividadEmpleados: {
      asignados100: empAsignados100,
      pctAsignados100: totEmp > 0 ? Number(((empAsignados100 / totEmp) * 100).toFixed(1)) : 0,
      mixtos: empMixtos,
      pctMixtos: totEmp > 0 ? Number(((empMixtos / totEmp) * 100).toFixed(1)) : 0,
      sinAsignar100: empSinAsignar100,
      pctSinAsignar100: totEmp > 0 ? Number(((empSinAsignar100 / totEmp) * 100).toFixed(1)) : 0
    },
    casosParciales: {
      empresas: casosEmpresasParciales,
      establecimientos: casosEstabsParciales,
      empleados: casosEmpleadosParciales
    },
    prestadores: prestadoresList
  };

  // Guardar en ScriptCache (30 minutos)
  try {
    CacheService.getScriptCache().put('metricas_kpis_v2', JSON.stringify(resObj), 1800);
  } catch(e) {}

  if (guardarEnHoja !== false) {
    try {
      _volcarMetricasEnHojaKPIs_(resObj);
    } catch(eHoja) {
      console.warn("Aviso al volcar en pestaña física KPIs: " + eHoja.message);
    }
  }

  return resObj;
}

/**
 * Devuelve un objeto de KPIs vacío limpio.
 */
function _obtenerObjetoKPIsVacio_() {
  return {
    totales: { estudios: 0, empleadosAN: 0, establecimientosH: 0, empresasC: 0 },
    asignados100: { estudios: 0, pctEstudios: 0, empleadosAN: 0, pctEmpleados: 0, establecimientosH: 0, pctEstablecimientos: 0, empresasC: 0, pctEmpresas: 0 },
    parciales: { estudios: 0, pctEstudios: 0, empleadosAN: 0, pctEmpleados: 0, establecimientosH: 0, pctEstablecimientos: 0, empresasC: 0, pctEmpresas: 0 },
    sinAsignar100: { estudios: 0, pctEstudios: 0, empleadosAN: 0, pctEmpleados: 0, establecimientosH: 0, pctEstablecimientos: 0, empresasC: 0, pctEmpresas: 0 },
    conPrestador: { estudios: 0, pctEstudios: 0, empleadosAN: 0, pctEmpleados: 0, establecimientosH: 0, pctEstablecimientos: 0, empresasC: 0, pctEmpresas: 0 },
    sinPrestador: { estudios: 0, pctEstudios: 0, empleadosAN: 0, pctEmpleados: 0, establecimientosH: 0, pctEstablecimientos: 0, empresasC: 0, pctEmpresas: 0 },
    exclusividadEmpleados: { asignados100: 0, pctAsignados100: 0, mixtos: 0, pctMixtos: 0, sinAsignar100: 0, pctSinAsignar100: 0 },
    casosParciales: { empresas: [], establecimientos: [], empleados: [] },
    prestadores: [],
    timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
    tiempoMs: 0
  };
}

/**
 * Inicializa y obtiene la hoja física "KPIs".
 */
function inicializarHojaKPIs() {
  const ss = _getSpreadsheet_();
  const nombreHoja = (CONFIG && CONFIG.HOJA_KPIS) ? CONFIG.HOJA_KPIS : "KPIs";
  let hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    hoja = ss.insertSheet(nombreHoja);
  }
  return hoja;
}

/**
 * Vuelca las métricas consolidadas en la hoja física "KPIs".
 */
function _volcarMetricasEnHojaKPIs_(data) {
  if (!data || !data.totales) return;
  const hoja = inicializarHojaKPIs();
  hoja.clear();

  const tot = data.totales || {};
  const asig = data.asignados100 || {};
  const parc = data.parciales || {};
  const sin = data.sinAsignar100 || {};
  const casos = data.casosParciales || {};

  // Título
  hoja.getRange(1, 1).setValue("PANEL CONSOLIDADO DE MÉTRICAS & KPIS (100% CUADRADO)").setFontWeight("bold").setFontSize(12);
  hoja.getRange(2, 1).setValue("Última actualización: " + (data.timestamp || '')).setFontColor("#64748b").setFontSize(9);

  // TABLA 1: RESUMEN COMPARATIVO CONSOLIDADO (100% Cuadrado)
  const headersT1 = [
    "Segmento / Condición",
    "Estudios (Filas)",
    "% Estudios",
    "Empleados (AN)",
    "% Empleados",
    "Establecimientos (H)",
    "% Establecimientos",
    "Empresas (C)",
    "% Empresas"
  ];

  const filasT1 = [
    [
      "TOTAL GENERAL CONSOLIDADO",
      tot.estudios || 0,
      1,
      tot.empleadosAN || 0,
      1,
      tot.establecimientosH || 0,
      1,
      tot.empresasC || 0,
      1
    ],
    [
      "100% CON PRESTADOR (Totalmente Asignados)",
      asig.estudios || 0,
      (asig.pctEstudios || 0) / 100,
      asig.empleadosAN || 0,
      (asig.pctEmpleados || 0) / 100,
      asig.establecimientosH || 0,
      (asig.pctEstablecimientos || 0) / 100,
      asig.empresasC || 0,
      (asig.pctEmpresas || 0) / 100
    ],
    [
      "PARCIALES / MIXTOS (Con y Sin Prestador)",
      parc.estudios || 0,
      (parc.pctEstudios || 0) / 100,
      parc.empleadosAN || 0,
      (parc.pctEmpleados || 0) / 100,
      parc.establecimientosH || 0,
      (parc.pctEstablecimientos || 0) / 100,
      parc.empresasC || 0,
      (parc.pctEmpresas || 0) / 100
    ],
    [
      "100% SIN PRESTADOR (Totalmente Sin Asignar)",
      sin.estudios || 0,
      (sin.pctEstudios || 0) / 100,
      sin.empleadosAN || 0,
      (sin.pctEmpleados || 0) / 100,
      sin.establecimientosH || 0,
      (sin.pctEstablecimientos || 0) / 100,
      sin.empresasC || 0,
      (sin.pctEmpresas || 0) / 100
    ]
  ];

  hoja.getRange(4, 1, 1, headersT1.length).setValues([headersT1])
    .setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
  hoja.getRange(5, 1, filasT1.length, headersT1.length).setValues(filasT1);

  // Formato porcentajes en cols 3, 5, 7, 9
  hoja.getRange(5, 3, filasT1.length, 1).setNumberFormat("0.0%");
  hoja.getRange(5, 5, filasT1.length, 1).setNumberFormat("0.0%");
  hoja.getRange(5, 7, filasT1.length, 1).setNumberFormat("0.0%");
  hoja.getRange(5, 9, filasT1.length, 1).setNumberFormat("0.0%");

  // Colores de filas
  hoja.getRange(5, 1, 1, headersT1.length).setFontWeight("bold").setBackground("#e2e8f0");
  hoja.getRange(6, 1, 1, headersT1.length).setBackground("#ecfdf5");
  hoja.getRange(7, 1, 1, headersT1.length).setBackground("#eff6ff");
  hoja.getRange(8, 1, 1, headersT1.length).setBackground("#fffbeb");

  // TABLA 2: ESTABLECIMIENTOS PARCIALES (DETALLE DE CASOS)
  let curRow = 11;
  hoja.getRange(curRow - 1, 1).setValue("ESTABLECIMIENTOS PARCIALES (CON ESTUDIOS CON Y SIN PRESTADOR)").setFontWeight("bold").setFontSize(10);
  const headersEstParc = ["Código SRT", "Establecimiento", "CUIT Empresa", "Empresa / Contrato", "Estudios Totales", "Con Prestador", "Sin Prestador", "% Asignado", "Prestadores"];
  hoja.getRange(curRow, 1, 1, headersEstParc.length).setValues([headersEstParc])
    .setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");

  if (casos.establecimientos && casos.establecimientos.length > 0) {
    const filasEst = casos.establecimientos.map(e => [
      e.srt,
      e.nombre,
      e.cuit,
      e.nombreEmpresa,
      e.total,
      e.con,
      e.sin,
      (e.pctCon / 100),
      e.prestadores
    ]);
    hoja.getRange(curRow + 1, 1, filasEst.length, headersEstParc.length).setValues(filasEst);
    hoja.getRange(curRow + 1, 8, filasEst.length, 1).setNumberFormat("0.0%");
    curRow += filasEst.length + 3;
  } else {
    curRow += 3;
  }

  // TABLA 3: EMPRESAS PARCIALES
  hoja.getRange(curRow - 1, 1).setValue("EMPRESAS PARCIALES (CON ESTUDIOS CON Y SIN PRESTADOR)").setFontWeight("bold").setFontSize(10);
  const headersEmpParc = ["CUIT Contrato", "Empresa / Contrato", "Estudios Totales", "Con Prestador", "Sin Prestador", "% Asignado", "Prestadores"];
  hoja.getRange(curRow, 1, 1, headersEmpParc.length).setValues([headersEmpParc])
    .setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");

  if (casos.empresas && casos.empresas.length > 0) {
    const filasEmp = casos.empresas.map(e => [
      e.cuit,
      e.nombre,
      e.total,
      e.con,
      e.sin,
      (e.pctCon / 100),
      e.prestadores
    ]);
    hoja.getRange(curRow + 1, 1, filasEmp.length, headersEmpParc.length).setValues(filasEmp);
    hoja.getRange(curRow + 1, 6, filasEmp.length, 1).setNumberFormat("0.0%");
    curRow += filasEmp.length + 3;
  } else {
    curRow += 3;
  }

  // TABLA 4: DETALLE POR PRESTADOR
  hoja.getRange(curRow - 1, 1).setValue("DETALLE POR PRESTADOR ASIGNADO").setFontWeight("bold").setFontSize(10);
  const headersT3 = [
    "Prestador",
    "Estudios (Filas)",
    "% Estudios",
    "Empleados (AN)",
    "% Empleados",
    "Establecimientos (H)",
    "Empresas (C)",
    "Promedio Estudios/Empleado"
  ];
  hoja.getRange(curRow, 1, 1, headersT3.length).setValues([headersT3])
    .setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");

  if (data.prestadores && data.prestadores.length > 0) {
    const filasT3 = data.prestadores.map(p => [
      p.nombre,
      p.estudios || 0,
      (p.pctEstudios || 0) / 100,
      p.empleadosAN || 0,
      (p.pctEmpleados || 0) / 100,
      p.establecimientosH || 0,
      p.empresasC || 0,
      p.promEstudiosPorEmpleado || 0
    ]);
    hoja.getRange(curRow + 1, 1, filasT3.length, headersT3.length).setValues(filasT3);
    hoja.getRange(curRow + 1, 3, filasT3.length, 1).setNumberFormat("0.0%");
    hoja.getRange(curRow + 1, 5, filasT3.length, 1).setNumberFormat("0.0%");
    hoja.getRange(curRow + 1, 8, filasT3.length, 1).setNumberFormat("0.00");
  }
}

/**
 * Lee las métricas consolidadas directamente desde la hoja física "KPIs" en <0.2s.
 */
function _leerMetricasKPIsDesdeHoja_() {
  try {
    const ss = _getSpreadsheet_();
    const nombreHoja = (CONFIG && CONFIG.HOJA_KPIS) ? CONFIG.HOJA_KPIS : "KPIs";
    const hoja = ss.getSheetByName(nombreHoja);
    if (!hoja || hoja.getLastRow() < 8) return null;

    const data = hoja.getDataRange().getValues();
    const timestamp = String(data[1][0] || '').replace('Última actualización: ', '').trim();

    // Filas 4, 5, 6, 7 (base 0)
    const totRow  = data[4];
    const asigRow = data[5];
    const parcRow = data[6];
    const sinRow  = data[7];

    if (!totRow || !asigRow || !parcRow || !sinRow) return null;

    // Buscar inicio de tablas de casos y prestadores
    const casosEst = [];
    const casosEmp = [];
    const prestadores = [];
    let seccion = '';

    for (let r = 8; r < data.length; r++) {
      const celda = String(data[r][0] || '').trim();
      if (celda.startsWith('ESTABLECIMIENTOS PARCIALES')) {
        seccion = 'estabs_parc';
        r++; // saltar encabezado
        continue;
      }
      if (celda.startsWith('EMPRESAS PARCIALES')) {
        seccion = 'empresas_parc';
        r++; // saltar encabezado
        continue;
      }
      if (celda.startsWith('DETALLE POR PRESTADOR')) {
        seccion = 'prestadores';
        r++; // saltar encabezado
        continue;
      }

      if (seccion === 'estabs_parc' && celda) {
        casosEst.push({
          srt: celda,
          nombre: String(data[r][1] || ''),
          cuit: String(data[r][2] || ''),
          nombreEmpresa: String(data[r][3] || ''),
          total: Number(data[r][4]) || 0,
          con: Number(data[r][5]) || 0,
          sin: Number(data[r][6]) || 0,
          pctCon: Number(((Number(data[r][7]) || 0) * 100).toFixed(1)),
          prestadores: String(data[r][8] || '')
        });
      } else if (seccion === 'empresas_parc' && celda) {
        casosEmp.push({
          cuit: celda,
          nombre: String(data[r][1] || ''),
          total: Number(data[r][2]) || 0,
          con: Number(data[r][3]) || 0,
          sin: Number(data[r][4]) || 0,
          pctCon: Number(((Number(data[r][5]) || 0) * 100).toFixed(1)),
          prestadores: String(data[r][6] || '')
        });
      } else if (seccion === 'prestadores' && celda) {
        prestadores.push({
          nombre: celda,
          estudios: Number(data[r][1]) || 0,
          pctEstudios: Number(((Number(data[r][2]) || 0) * 100).toFixed(1)),
          empleadosAN: Number(data[r][3]) || 0,
          pctEmpleados: Number(((Number(data[r][4]) || 0) * 100).toFixed(1)),
          establecimientosH: Number(data[r][5]) || 0,
          empresasC: Number(data[r][6]) || 0,
          promEstudiosPorEmpleado: Number(data[r][7]) || 0
        });
      }
    }

    return {
      timestamp: timestamp,
      totales: {
        estudios: Number(totRow[1]) || 0,
        empleadosAN: Number(totRow[3]) || 0,
        establecimientosH: Number(totRow[5]) || 0,
        empresasC: Number(totRow[7]) || 0
      },
      asignados100: {
        estudios: Number(asigRow[1]) || 0,
        pctEstudios: Number(((Number(asigRow[2]) || 0) * 100).toFixed(1)),
        empleadosAN: Number(asigRow[3]) || 0,
        pctEmpleados: Number(((Number(asigRow[4]) || 0) * 100).toFixed(1)),
        establecimientosH: Number(asigRow[5]) || 0,
        pctEstablecimientos: Number(((Number(asigRow[6]) || 0) * 100).toFixed(1)),
        empresasC: Number(asigRow[7]) || 0,
        pctEmpresas: Number(((Number(asigRow[8]) || 0) * 100).toFixed(1))
      },
      parciales: {
        estudios: Number(parcRow[1]) || 0,
        pctEstudios: Number(((Number(parcRow[2]) || 0) * 100).toFixed(1)),
        empleadosAN: Number(parcRow[3]) || 0,
        pctEmpleados: Number(((Number(parcRow[4]) || 0) * 100).toFixed(1)),
        establecimientosH: Number(parcRow[5]) || 0,
        pctEstablecimientos: Number(((Number(parcRow[6]) || 0) * 100).toFixed(1)),
        empresasC: Number(parcRow[7]) || 0,
        pctEmpresas: Number(((Number(parcRow[8]) || 0) * 100).toFixed(1))
      },
      sinAsignar100: {
        estudios: Number(sinRow[1]) || 0,
        pctEstudios: Number(((Number(sinRow[2]) || 0) * 100).toFixed(1)),
        empleadosAN: Number(sinRow[3]) || 0,
        pctEmpleados: Number(((Number(sinRow[4]) || 0) * 100).toFixed(1)),
        establecimientosH: Number(sinRow[5]) || 0,
        pctEstablecimientos: Number(((Number(sinRow[6]) || 0) * 100).toFixed(1)),
        empresasC: Number(sinRow[7]) || 0,
        pctEmpresas: Number(((Number(sinRow[8]) || 0) * 100).toFixed(1))
      },
      conPrestador: {
        estudios: Number(asigRow[1]) || 0,
        pctEstudios: Number(((Number(asigRow[2]) || 0) * 100).toFixed(1)),
        empleadosAN: Number(asigRow[3]) + Number(parcRow[3]),
        pctEmpleados: 0,
        establecimientosH: Number(asigRow[5]) + Number(parcRow[5]),
        pctEstablecimientos: 0,
        empresasC: Number(asigRow[7]) + Number(parcRow[7]),
        pctEmpresas: 0
      },
      sinPrestador: {
        estudios: Number(sinRow[1]) || 0,
        pctEstudios: Number(((Number(sinRow[2]) || 0) * 100).toFixed(1)),
        empleadosAN: Number(sinRow[3]) + Number(parcRow[3]),
        pctEmpleados: 0,
        establecimientosH: Number(sinRow[5]) + Number(parcRow[5]),
        pctEstablecimientos: 0,
        empresasC: Number(sinRow[7]) + Number(parcRow[7]),
        pctEmpresas: 0
      },
      exclusividadEmpleados: {
        asignados100: Number(asigRow[3]) || 0,
        pctAsignados100: Number(((Number(asigRow[4]) || 0) * 100).toFixed(1)),
        mixtos: Number(parcRow[3]) || 0,
        pctMixtos: Number(((Number(parcRow[4]) || 0) * 100).toFixed(1)),
        sinAsignar100: Number(sinRow[3]) || 0,
        pctSinAsignar100: Number(((Number(sinRow[4]) || 0) * 100).toFixed(1))
      },
      casosParciales: {
        empresas: casosEmp,
        establecimientos: casosEst,
        empleados: []
      },
      prestadores: prestadores
    };
  } catch(e) {
    console.warn("No se pudo leer desde hoja física KPIs: " + e.message);
    return null;
  }
}

/**
 * Recalcula directamente desde Madre y sincroniza la hoja física.
 */
function recalcularKPIsDesdeMadre() {
  CacheService.getScriptCache().remove('metricas_kpis_v2');
  return _calcularMetricasKPIsDesdeMadre_(true);
}

/**
 * Función pública para sincronizar hoja física.
 */
function sincronizarHojaKPIs() {
  try {
    const res = recalcularKPIsDesdeMadre();
    return { status: "OK", totales: res.totales };
  } catch(e) {
    return { status: "Error", msg: e.message };
  }
}

/**
 * Normaliza todas las fechas de la Columna AB (Fecha Estudio, col 28) en la Hoja Madre
 * asegurando el formato estricto dd/MM/yyyy y previniendo formatos inconsistentes.
 */
function normalizarFormatoFechasColumnaABMadre() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.HOJA_MADRE_ID);
    const hojaMadre = ss.getSheetByName(CONFIG.HOJA_MADRE_NAME || "Madre");
    if (!hojaMadre) return { status: 'ERROR', msg: 'No se encontró la hoja Madre' };

    const ultFila = hojaMadre.getLastRow();
    if (ultFila < 2) return { status: 'OK', totalFilas: 0, corregidos: 0 };

    const colAB = CONFIG.COL_FECHA_ESTUDIO + 1; // Columna 28 = AB
    const numFilas = ultFila - 1;
    const rango = hojaMadre.getRange(2, colAB, numFilas, 1);
    const valores = rango.getValues();
    const tz = Session.getScriptTimeZone() || "America/Argentina/Buenos_Aires";

    let corregidos = 0;
    const nuevosValores = valores.map(row => {
      const v = row[0];
      if (!v && v !== 0) return [""];
      const fmt = formatearFecha(v, tz);
      if (fmt !== String(v).trim()) corregidos++;
      return [fmt];
    });

    if (corregidos > 0) {
      rango.setValues(nuevosValores);
    }
    rango.setNumberFormat("@");
    SpreadsheetApp.flush();

    return { status: 'OK', totalFilas: numFilas, corregidos: corregidos };
  } catch (e) {
    console.error("Error en normalizarFormatoFechasColumnaABMadre: " + e.message);
    return { status: 'ERROR', msg: e.message };
  }
}