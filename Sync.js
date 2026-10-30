// ==================================================================
// NÚCLEO DE SINCRONIZACIÓN
// ==================================================================
function _ejecutarSincronizacion_(opciones) {
  opciones = opciones || {};
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) {
      registrarEnBitacora('SYNC_OMITIDA', { razon: 'Lock en uso' });
      return { success: false, msg: 'La sincronización está siendo ejecutada por otro proceso. Intente nuevamente en unos segundos.' };
    }

    // [2026-09-06] openById para ejecución multiusuario segura en Web App
    let ssMadre;
    try {
      ssMadre = SpreadsheetApp.openById(CONFIG.HOJA_MADRE_ID);
    } catch (e) {
      throw new Error("No se puede acceder al archivo Hoja Madre: " + e.message);
    }

    const hojaMadre = ssMadre.getSheetByName(CONFIG.HOJA_MADRE_NAME);
    if (!hojaMadre) throw new Error(`Pestaña "${CONFIG.HOJA_MADRE_NAME}" no existe en el archivo Hoja Madre.`);
    const ultimaFilaMadre = hojaMadre.getLastRow();
    if (ultimaFilaMadre < 2) return { success: true, registros: 0 };

    // 1. PRIORIDAD: Buscar pestaña "Links" o "Gestion_Links" dentro de la misma Hoja Madre (Ultrarrápido, local)
    let hojaLinksLocal = ssMadre.getSheetByName(CONFIG.HOJA_LINKS || "Links") ||
                         ssMadre.getSheetByName("Gestion_Links") ||
                         ssMadre.getSheetByName("Gestion Links");
    let datosExternos = null;

    if (!opciones.recargarDesdeExterna && hojaLinksLocal && hojaLinksLocal.getLastRow() >= 2) {
      // Leer DIRECTAMENTE de la Hoja Madre (sin llamadas a planillas externas)
      const uFilaL = hojaLinksLocal.getLastRow();
      const nColsL = Math.max(hojaLinksLocal.getLastColumn(), 5);
      datosExternos = hojaLinksLocal.getRange(2, 1, uFilaL - 1, nColsL).getValues();
    } else {
      // Fallback: si aún no existe localmente o se fuerza recarga, leer de la planilla externa y volcar a Links
      let ssExterna;
      try {
        ssExterna = SpreadsheetApp.openById(CONFIG.EXTERNA_ID);
      } catch (e) {
        throw new Error("No se pudo acceder al archivo externo (" + CONFIG.EXTERNA_ID + "): " + e.message);
      }

      let hojaExterna = null;
      // 1. Búsqueda directa por GID (ej. 1054481966)
      if (CONFIG.EXTERNA_GID) {
        hojaExterna = ssExterna.getSheets().find(s => s.getSheetId() === Number(CONFIG.EXTERNA_GID));
      }
      // 2. Búsqueda por nombres configurados
      if (!hojaExterna) {
        hojaExterna = ssExterna.getSheetByName(CONFIG.EXTERNA_NAME) ||
                      ssExterna.getSheetByName("Gestion_Links") ||
                      ssExterna.getSheetByName("Gestion Links") ||
                      ssExterna.getSheetByName("Gestión Empleados") ||
                      ssExterna.getSheetByName("Gestion Empleados");
      }
      // 3. Búsqueda flexible (insensible a mayúsculas, tildes y guiones bajos/espacios)
      if (!hojaExterna) {
        hojaExterna = ssExterna.getSheets().find(s => {
          const norm = s.getName().trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\s_-]+/g, '');
          return norm === "gestionlinks" || norm === "gestionempleados" || norm.includes("link");
        });
      }

      if (!hojaExterna) {
        const pestañasDisponibles = ssExterna.getSheets().map(s => `"${s.getName()}" (gid=${s.getSheetId()})`).join(", ");
        throw new Error(`Pestaña "${CONFIG.EXTERNA_NAME}" (gid: ${CONFIG.EXTERNA_GID}) no existe en hoja externa. Pestañas disponibles en el archivo: [${pestañasDisponibles}]`);
      }

      const uFilaExt = hojaExterna.getLastRow();
      if (uFilaExt < 2) return { success: true, registros: 0, msg: "Hoja externa sin datos" };

      const maxColExt = Math.max(hojaExterna.getLastColumn(), 5);
      // Auto-detección inteligente de columnas por encabezados
      let idxClave = CONFIG.EXT_COL_CLAVE !== undefined ? CONFIG.EXT_COL_CLAVE : 0;
      let idxAsist = CONFIG.EXT_COL_ASISTENCIA !== undefined ? CONFIG.EXT_COL_ASISTENCIA : 1;
      let idxAdj1  = CONFIG.EXT_COL_ADJUNTO_1 !== undefined ? CONFIG.EXT_COL_ADJUNTO_1 : 2;
      let idxAdj2  = CONFIG.EXT_COL_ADJUNTO_2 !== undefined ? CONFIG.EXT_COL_ADJUNTO_2 : 3;
      let idxFecha = CONFIG.EXT_COL_FECHA_ESTUDIO !== undefined ? CONFIG.EXT_COL_FECHA_ESTUDIO : 4;

      const encabezadosExt = hojaExterna.getRange(1, 1, 1, maxColExt).getValues()[0];
      encabezadosExt.forEach((h, colIdx) => {
        const hNorm = String(h || '').trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\s_-]+/g, '_');
        if (hNorm === 'clave' || hNorm === 'clave_cruce' || hNorm === 'id') idxClave = colIdx;
        else if (hNorm.includes('asistencia') || hNorm === 'estado') idxAsist = colIdx;
        else if (hNorm.includes('adjunto_1') || hNorm === 'link1' || hNorm === 'link_1' || hNorm === 'link_adjunto_1') idxAdj1 = colIdx;
        else if (hNorm.includes('adjunto_2') || hNorm === 'link2' || hNorm === 'link_2' || hNorm === 'link_adjunto_2') idxAdj2 = colIdx;
        else if (hNorm.includes('fecha_examen') || hNorm.includes('fecha_estudio') || hNorm === 'fecha') idxFecha = colIdx;
      });

      const colsNec = Math.max(idxAdj2 + 1, idxAdj1 + 1, idxFecha + 1, idxAsist + 1, idxClave + 1, 5);
      const matrizCrudaExt = hojaExterna.getRange(2, 1, uFilaExt - 1, colsNec).getValues();

      datosExternos = matrizCrudaExt.map(fila => [
        String(fila[idxClave] || '').trim(),
        String(fila[idxAsist] || '').trim(),
        String(fila[idxAdj1] || '').trim(),
        String(fila[idxAdj2] || '').trim(),
        formatearFecha(fila[idxFecha])
      ]);

      // Guardar copia local en la pestaña Links
      try {
        if (!hojaLinksLocal) hojaLinksLocal = ssMadre.insertSheet(CONFIG.HOJA_LINKS || "Links");
        const tsSync = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        const headersLinks = ["Clave", "Asistencia_Calculada", "Link_Adjunto_1", "Link_Adjunto_2", "Fecha_Examen", "Ultima_Actualizacion"];
        const filasVolcado = datosExternos.map(fila => [
          fila[0],
          fila[1],
          fila[2],
          fila[3],
          fila[4],
          tsSync
        ]);
        hojaLinksLocal.clearContents();
        hojaLinksLocal.getRange(1, 1, 1, headersLinks.length).setValues([headersLinks]);
        hojaLinksLocal.getRange(1, 1, 1, headersLinks.length).setFontWeight("bold").setBackground("#0052cc").setFontColor("#ffffff");
        if (filasVolcado.length > 0) {
          hojaLinksLocal.getRange(2, 1, filasVolcado.length, headersLinks.length).setValues(filasVolcado);
        }
      } catch (eVolc) {
        console.warn("Aviso al crear copia local en Links: " + eVolc.message);
      }
    }

    if (!datosExternos || datosExternos.length === 0) {
      return { success: true, registros: 0, msg: "Sin datos en hoja Links" };
    }

    const clavesExternas = new Set();
    const mapas = { asistencia: {}, fechaEstudio: {}, adjunto1: {}, adjunto2: {} };

    datosExternos.forEach(fila => {
      const clave = String(fila[0] || '').trim();
      if (!clave) return;
      clavesExternas.add(clave);
      mapas.asistencia[clave] = String(fila[1] || '').trim().toUpperCase();
      const a1 = String(fila[2] || '').trim();
      if (a1) mapas.adjunto1[clave] = a1;
      const a2 = String(fila[3] || '').trim();
      if (a2) mapas.adjunto2[clave] = a2;
      const f = String(fila[4] || '').trim();
      if (f) mapas.fechaEstudio[clave] = formatearFecha(f);
    });

    // 2. LECTURA EN UN SOLO BLOQUE DE LA HOJA MADRE (1 sola llamada a la API)
    const nMadre = ultimaFilaMadre - 1;
    const maxColsMadre = Math.max(hojaMadre.getLastColumn(), 42);
    const matrizMadre = hojaMadre.getRange(2, 1, nMadre, maxColsMadre).getValues();

    // Detección de casos conflictivos y métricas de cambio entre Madre y Links
    const casosImpactoVacio = [];
    let linksNuevos = 0;
    let linksModificados = 0;
    let eliminadosNoAuditados = 0;
    let eliminadosAuditados = 0;

    for (let i = 0; i < nMadre; i++) {
      const filaM = matrizMadre[i];
      const c = String(filaM[CONFIG.COL_CLAVE] !== undefined ? filaM[CONFIG.COL_CLAVE] : filaM[0] || '').trim();
      if (!c) continue;

      const estudio = String(filaM[CONFIG.COL_ESTUDIO] || '').trim();
      if (estudio.includes('No posee')) continue;

      const link1M = String(filaM[CONFIG.COL_ADJUNTO_1] || '').trim();
      const link2M = String(filaM[CONFIG.COL_ADJUNTO_2] || '').trim();
      const teniaLinkMadre = Boolean(link1M || link2M);

      const link1E = mapas.adjunto1[c] || "";
      const link2E = mapas.adjunto2[c] || "";
      const vieneConLinkExt = Boolean(link1E || link2E);

      const resM = String(filaM[CONFIG.COL_RESULTADO] || '').trim();
      const resMNorm = (typeof _normalizarTexto_ === 'function')
        ? _normalizarTexto_(resM)
        : resM.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();

      const esGuardadoM = typeof _esGuardado_ === 'function' ? _esGuardado_(filaM[CONFIG.COL_GUARDADO]) : Boolean(filaM[CONFIG.COL_GUARDADO]);
      const esAuditadoM = Boolean(resM && !resMNorm.includes('ausente'));

      // DETECCIÓN: Tenía link en Madre pero ahora se eliminó en Links
      if (teniaLinkMadre && !vieneConLinkExt) {
        if (esAuditadoM || esGuardadoM) {
          casosImpactoVacio.push({
            idxMadre: i,
            filaHoja: i + 2,
            clave: c,
            cuil: String(filaM[CONFIG.COL_CUIL] || '').trim(),
            empleado: String(filaM[CONFIG.COL_APELLIDO] || '').trim(),
            estudio: estudio,
            establecimiento: String(filaM[CONFIG.COL_NOMBRE_ESTABLECIMIENTO] || '').trim(),
            srt: String(filaM[CONFIG.COL_SRT] || '').trim(),
            resultado: resM,
            guardado: esGuardadoM,
            link1: link1M,
            link2: link2M
          });
          eliminadosAuditados++;
        } else {
          eliminadosNoAuditados++;
        }
      } else if (!teniaLinkMadre && vieneConLinkExt) {
        linksNuevos++;
      } else if (teniaLinkMadre && vieneConLinkExt) {
        if (link1M !== link1E || link2M !== link2E) {
          linksModificados++;
        }
      }
    }

    // Si hay conflictos y no se ha tomado una decisión explícita, solicitar confirmación al usuario
    if (casosImpactoVacio.length > 0 && !opciones.confirmarImpactoVacio && !opciones.ignorarConflictos) {
      lock.releaseLock();
      return {
        success: true,
        requiereConfirmacion: true,
        totalCasos: casosImpactoVacio.length,
        casos: casosImpactoVacio.slice(0, 10).map(c => ({
          cuil: c.cuil,
          empleado: c.empleado,
          estudio: c.estudio,
          establecimiento: c.establecimiento,
          resultado: c.resultado,
          guardado: c.guardado ? 'SÍ' : 'NO'
        })),
        msg: `Se detectaron ${casosImpactoVacio.length} estudios auditados o guardados cuyos enlaces vienen vacíos en la hoja externa.`
      };
    }

    const indicesConfirmadosVacios = new Set();
    const eventosBitacora = [];

    // Si el usuario confirmó el impacto de vacío: borrar resultado, asistencia, guardado, renombrar Drive y bitácora
    if (opciones.confirmarImpactoVacio && casosImpactoVacio.length > 0) {
      casosImpactoVacio.forEach(caso => {
        indicesConfirmadosVacios.add(caso.idxMadre);

        // 1. Limpiar datos de auditoría en memoria
        matrizMadre[caso.idxMadre][CONFIG.COL_RESULTADO] = "";
        matrizMadre[caso.idxMadre][CONFIG.COL_GUARDADO] = "";
        matrizMadre[caso.idxMadre][CONFIG.COL_ASISTENCIA] = "";

        // 2. Renombrar archivo(s) en Google Drive
        let renomb1 = "";
        let renomb2 = "";
        if (caso.link1) renomb1 = _renombrarArchivoBajaDrive_(caso.link1);
        if (caso.link2) renomb2 = _renombrarArchivoBajaDrive_(caso.link2);
        const renombTotal = [renomb1, renomb2].filter(Boolean).join("; ") || "Sin archivo / No modificable";

        // 3. Registrar evento en Bitácora con links eliminados
        eventosBitacora.push([
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
          Session.getActiveUser().getEmail() || "Usuario Web",
          caso.clave,
          caso.cuil,
          caso.empleado,
          caso.estudio,
          `${caso.establecimiento} (SRT: ${caso.srt})`,
          caso.resultado,
          caso.guardado ? 'SÍ' : 'NO',
          caso.link1 || '-',
          caso.link2 || '-',
          renombTotal,
          'Baja confirmada por impacto de vacío desde Externa (des-auditado)'
        ]);
      });

      _registrarEnBitacoraBajas_(eventosBitacora);

      // Persistir columnas Resultado y Guardado limpiadas en Hoja Madre
      const resLimpios = matrizMadre.map((fila, idx) => indicesConfirmadosVacios.has(idx) ? [""] : [fila[CONFIG.COL_RESULTADO]]);
      const guarLimpios = matrizMadre.map((fila, idx) => indicesConfirmadosVacios.has(idx) ? [""] : [fila[CONFIG.COL_GUARDADO]]);
      hojaMadre.getRange(2, CONFIG.COL_RESULTADO + 1, nMadre, 1).setValues(resLimpios);
      hojaMadre.getRange(2, CONFIG.COL_GUARDADO + 1, nMadre, 1).setValues(guarLimpios);
    }

    const out = { asistencia: [], fechaEstudio: [], adjunto1: [], adjunto2: [] };

    matrizMadre.forEach((filaM, i) => {
      const c = String(filaM[CONFIG.COL_CLAVE] !== undefined ? filaM[CONFIG.COL_CLAVE] : filaM[0] || '').trim();
      const link1M = String(filaM[CONFIG.COL_ADJUNTO_1] || '').trim();
      const link2M = String(filaM[CONFIG.COL_ADJUNTO_2] || '').trim();
      const teniaLinkMadre = Boolean(link1M || link2M);

      const estaEnExterna = clavesExternas.has(c);
      const link1E = mapas.adjunto1[c] || "";
      const link2E = mapas.adjunto2[c] || "";

      let asis = estaEnExterna ? (mapas.asistencia[c] || "") : String(filaM[CONFIG.COL_ASISTENCIA] || "");
      let fechaEst = estaEnExterna ? (formatearFecha(mapas.fechaEstudio[c]) || "") : formatearFecha(filaM[CONFIG.COL_FECHA_ESTUDIO] || "");
      let a1 = estaEnExterna ? link1E : (teniaLinkMadre ? "" : "");
      let a2 = estaEnExterna ? link2E : "";

      const resMadre = String(filaM[CONFIG.COL_RESULTADO] || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();

      // Si fue confirmado para borrado total (baja confirmada de caso auditado/guardado)
      if (indicesConfirmadosVacios.has(i)) {
        out.asistencia.push([""]);
        out.fechaEstudio.push([fechaEst]);
        out.adjunto1.push([""]);
        out.adjunto2.push([""]);
        return;
      }

      // Si es un conflicto (tenía link auditado/guardado y ahora no viene en externa),
      // pero el usuario no confirmó el impacto de vacío (conservar auditorías):
      const esConflicto = casosImpactoVacio.some(cv => cv.idxMadre === i);
      if (esConflicto && (opciones.ignorarConflictos || !opciones.confirmarImpactoVacio)) {
        a1 = link1M;
        a2 = link2M;
        asis = String(filaM[CONFIG.COL_ASISTENCIA] || "");
        fechaEst = formatearFecha(filaM[CONFIG.COL_FECHA_ESTUDIO] || "");
      } else if (!estaEnExterna && !teniaLinkMadre) {
        // Caso normal: no estaba en externa y nunca tuvo link -> preservar datos preexistentes de Madre
        a1 = "";
        a2 = "";
        asis = String(filaM[CONFIG.COL_ASISTENCIA] || "");
        fechaEst = formatearFecha(filaM[CONFIG.COL_FECHA_ESTUDIO] || "");
      }

      // REGLA 1: Si fue dictaminado Ausente o asistencia es Ausente sin link
      if (asis === "AUSENTE" || resMadre.includes("ausente")) {
        a1 = "";
        a2 = "";
      }

      // REGLA 2: "si tiene LINK tiene que ser SIEMPRE PRESENTE"
      if (Boolean(a1 || a2)) {
        asis = "PRESENTE";
      }

      out.asistencia.push([asis]);
      out.fechaEstudio.push([fechaEst]);
      out.adjunto1.push([a1]);
      out.adjunto2.push([a2]);
    });

    if (out.asistencia.length > 0) {
      const n = out.asistencia.length;
      hojaMadre.getRange(2, CONFIG.COL_ASISTENCIA + 1,    n, 1).setValues(out.asistencia);
      hojaMadre.getRange(2, CONFIG.COL_FECHA_ESTUDIO + 1,  n, 1).setValues(out.fechaEstudio);
      hojaMadre.getRange(2, CONFIG.COL_FECHA_ESTUDIO + 1,  n, 1).setNumberFormat("@");
      hojaMadre.getRange(2, CONFIG.COL_ADJUNTO_1 + 1,      n, 1).setValues(out.adjunto1);
      hojaMadre.getRange(2, CONFIG.COL_ADJUNTO_2 + 1,      n, 1).setValues(out.adjunto2);
      SpreadsheetApp.flush();

      try {
        sincronizarHojaEmpleados(true);
      } catch (eEmp) {
        console.error("Error sincronizando hoja Empleados en sync: " + eEmp.message);
      }
      _invalidarCache_();           // invalida caché de datos base (datosBase_v3, chunked)
      _invalidarCacheDashboard_();  // dashboard se recalcula fresco en la próxima visita

      // Auditoría ultrarrápida de inconsistencias en memoria (0 llamadas a API)
      let inconsistencias = 0;
      try {
        for (let i = 0; i < n; i++) {
          const filaM = matrizMadre[i];
          const est = String(filaM[CONFIG.COL_ESTUDIO] || '');
          if (est.includes('No posee')) continue;

          const asis = String(out.asistencia[i] ? out.asistencia[i][0] : '').toUpperCase().trim();
          const res  = String(filaM[CONFIG.COL_RESULTADO] || '').trim();
          const obs  = String(filaM[CONFIG.COL_OBSERVACION] || '').trim();
          const adj  = Boolean((out.adjunto1[i] && out.adjunto1[i][0]) || (out.adjunto2[i] && out.adjunto2[i][0]));

          const resNorm = (typeof _normalizarTexto_ === 'function')
            ? _normalizarTexto_(res)
            : String(res).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();

          const esAusenteAsis = asis.includes('AUSENTE');
          const esAusenteRes  = resNorm.includes('ausente');

          let inc = false;
          if (esAusenteAsis && res !== '' && !esAusenteRes) inc = true;
          else if (esAusenteRes && asis.includes('PRESENTE')) inc = true;
          else if ((esAusenteAsis || esAusenteRes) && adj) inc = true;
          else if (res === 'Desistido' && !obs) inc = true;
          else if (res !== '' && !esAusenteRes && !asis) inc = true;

          if (inc) inconsistencias++;
        }
      } catch (eInc) {
        console.error("Error auditando inconsistencias en sync: " + eInc.message);
      }

      if (inconsistencias > 0) {
        registrarEnBitacora('SYNC_INCONSISTENCIAS', { total: inconsistencias });
      }

      registrarEnBitacora('SINCRONIZACION_OK', { registros: n, inconsistencias: inconsistencias, bajasConfirmadas: indicesConfirmadosVacios.size });
      _guardarUltimaSincronizacion_();
      const totalEliminados = eliminadosAuditados + eliminadosNoAuditados;
      return {
        success: true,
        registros: n,
        linksNuevos: linksNuevos,
        linksModificados: linksModificados,
        eliminadosTotal: totalEliminados,
        eliminadosAuditados: eliminadosAuditados,
        eliminadosNoAuditados: eliminadosNoAuditados,
        inconsistencias: inconsistencias,
        bajasConfirmadas: indicesConfirmadosVacios.size,
        msg: `Sync exitosa: +${linksNuevos} nuevos, ${linksModificados} actualizados, ${totalEliminados} eliminados detectados.`
      };
    }
    return { success: true, registros: 0, inconsistencias: 0 };

  } catch (e) {
    registrarEnBitacora('SINCRONIZACION_ERROR', { error: e.message });
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Renombra en Google Drive un archivo asociado a un enlace dado de baja,
 * agregándole el prefijo [DESVINCULADO_YYYYMMDD_HHmm]_
 */
function _renombrarArchivoBajaDrive_(link) {
  if (!link) return null;
  let fileId = null;
  if (typeof _extraerFileId_ === 'function') {
    fileId = _extraerFileId_(link);
  } else {
    const m = String(link).match(/[-\w]{25,}/);
    fileId = m ? m[0] : null;
  }
  if (!fileId) return null;

  try {
    const file = DriveApp.getFileById(fileId);
    const nombreActual = file.getName();
    if (!nombreActual.startsWith('[DESVINCULADO') && !nombreActual.startsWith('[ELIMINADO')) {
      const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
      const nuevoNombre = `[DESVINCULADO_${ts}]_${nombreActual}`;
      file.setName(nuevoNombre);
      return nuevoNombre;
    }
    return nombreActual;
  } catch (err) {
    console.warn("No se pudo renombrar en Drive archivo (" + fileId + "): " + err.message);
    return null;
  }
}

/**
 * Registra en la pestaña 'Bitacora' los eventos de baja/impacto de vacío
 */
function _registrarEnBitacoraBajas_(eventos) {
  if (!CONFIG.HABILITAR_BITACORA) return;
  if (!eventos || eventos.length === 0) return;
  try {
    const ss = SpreadsheetApp.openById(CONFIG.HOJA_MADRE_ID);
    const nombreHoja = (CONFIG.HOJA_BITACORA && CONFIG.HOJA_BITACORA.trim()) ? CONFIG.HOJA_BITACORA : "Bitacora";
    let sheetBitacora = ss.getSheetByName(nombreHoja);
    if (!sheetBitacora) {
      sheetBitacora = ss.insertSheet(nombreHoja);
      const encabezados = [
        "Fecha/Hora", "Usuario / Auditor", "Clave", "CUIL", "Empleado",
        "Estudio", "Establecimiento / SRT", "Resultado Previo", "Guardado Previo",
        "Link 1 Eliminado", "Link 2 Eliminado", "Archivos Renombrados Drive", "Motivo / Acción"
      ];
      sheetBitacora.getRange(1, 1, 1, encabezados.length).setValues([encabezados])
        .setFontWeight("bold")
        .setBackground("#f1f5f9");
      sheetBitacora.setFrozenRows(1);
    }
    const ultFila = Math.max(1, sheetBitacora.getLastRow());
    sheetBitacora.getRange(ultFila + 1, 1, eventos.length, eventos[0].length).setValues(eventos);
  } catch (e) {
    console.error("Error registrando en Bitácora: " + e.message);
  }
}

function _guardarUltimaSincronizacion_() {
  PropertiesService.getScriptProperties().setProperty('lastSyncTime', Date.now().toString());
}

// Llamada por trigger automático (no interactivo, preserva auditorías ante conflictos)
function sincronizarAutomatico() {
  try { _ejecutarSincronizacion_({ ignorarConflictos: true }); }
  catch (e) { console.error("Error sync automático: " + e.message); }
}

// Llamada desde el botón "Actualizar" del frontend
function sincronizarForzado(opciones) {
  try { return _ejecutarSincronizacion_(opciones); }
  catch (e) { throw new Error("Error al sincronizar: " + e.message); }
}

// Llamada para forzar la recarga desde la planilla externa Gestion_Links hacia Links y Hoja Madre
function recargarLinksDesdeExterna(opciones) {
  opciones = opciones || {};
  opciones.recargarDesdeExterna = true;
  try {
    const res = _ejecutarSincronizacion_(opciones);
    try {
      repararYCompletarTablaGestionLinks();
    } catch (eRep) {
      console.warn("Aviso al sincronizar columnas en tabla externa: " + eRep.message);
    }
    return res;
  }
  catch (e) { throw new Error("Error al recargar links externos: " + e.message); }
}

// ==================================================================
// GESTIÓN DE TRIGGERS
// ==================================================================
function instalarTrigger() {
  _eliminarTriggerPorNombre_('sincronizarAutomatico');
  ScriptApp.newTrigger('sincronizarAutomatico').timeBased().everyMinutes(30).create();
  SpreadsheetApp.getUi().alert("✅ Trigger instalado: sincronizarAutomatico cada 30 min.");
}

function desinstalarTrigger() {
  const n = _eliminarTriggerPorNombre_('sincronizarAutomatico');
  SpreadsheetApp.getUi().alert(`Trigger eliminado (${n} instancias).`);
}

function instalarTriggerCache() {
  _eliminarTriggerPorNombre_('calentarCacheDatos');
  ScriptApp.newTrigger('calentarCacheDatos').timeBased().everyMinutes(10).create();
  SpreadsheetApp.getUi().alert('✅ Trigger de caché instalado: cada 10 minutos.');
}

function instalarTriggerDashboard() {
  _eliminarTriggerPorNombre_('calentarCacheDashboard');
  ScriptApp.newTrigger('calentarCacheDashboard').timeBased().everyMinutes(30).create();
  SpreadsheetApp.getUi().alert('✅ Trigger de dashboard instalado: calentarCacheDashboard cada 30 min.');
}

function desinstalarTriggerDashboard() {
  const n = _eliminarTriggerPorNombre_('calentarCacheDashboard');
  SpreadsheetApp.getUi().alert(`Trigger dashboard eliminado (${n} instancias).`);
}

function _eliminarTriggerPorNombre_(nombre) {
  let count = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === nombre) { ScriptApp.deleteTrigger(t); count++; }
  });
  return count;
}

// ==================================================================
// DIAGNÓSTICO DE CONEXIÓN CON PLANILLA EXTERNA
// ==================================================================
function diagnosticarConexionExterna() {
  const diag = {
    externaId: CONFIG.EXTERNA_ID,
    externaGid: CONFIG.EXTERNA_GID,
    tituloPlanilla: "",
    pestañas: [],
    pestañaSeleccionada: "",
    gidSeleccionado: null,
    totalFilas: 0,
    totalColumnas: 0,
    encabezados: [],
    muestraFila2: [],
    coincidenciasConMadre: 0,
    totalFilasMadre: 0,
    error: null
  };

  try {
    const ssExterna = SpreadsheetApp.openById(CONFIG.EXTERNA_ID);
    diag.tituloPlanilla = ssExterna.getName();
    diag.pestañas = ssExterna.getSheets().map(s => ({ nombre: s.getName(), gid: s.getSheetId() }));

    let hoja = null;
    if (CONFIG.EXTERNA_GID) {
      hoja = ssExterna.getSheets().find(s => s.getSheetId() === Number(CONFIG.EXTERNA_GID));
    }
    if (!hoja) {
      hoja = ssExterna.getSheetByName(CONFIG.EXTERNA_NAME) ||
             ssExterna.getSheetByName("Gestion_Links") ||
             ssExterna.getSheetByName("Gestion Links") ||
             ssExterna.getSheetByName("Gestión Empleados");
    }
    if (!hoja) {
      hoja = ssExterna.getSheets().find(s => {
        const norm = s.getName().trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\s_-]+/g, '');
        return norm === "gestionlinks" || norm.includes("link");
      });
    }

    if (!hoja) {
      diag.error = "No se pudo encontrar la pestaña ni por GID (" + CONFIG.EXTERNA_GID + ") ni por nombre. Revisa la lista de pestañas.";
      Logger.log(JSON.stringify(diag, null, 2));
      return diag;
    }

    diag.pestañaSeleccionada = hoja.getName();
    diag.gidSeleccionado = hoja.getSheetId();
    diag.totalFilas = hoja.getLastRow();
    diag.totalColumnas = hoja.getLastColumn();

    if (diag.totalFilas >= 1) {
      diag.encabezados = hoja.getRange(1, 1, 1, Math.min(diag.totalColumnas, 15)).getValues()[0];
    }
    if (diag.totalFilas >= 2) {
      diag.muestraFila2 = hoja.getRange(2, 1, 1, Math.min(diag.totalColumnas, 15)).getValues()[0];
    }

    // Probar cruce con Madre
    const ssMadre = SpreadsheetApp.openById(CONFIG.HOJA_MADRE_ID);
    const hojaMadre = ssMadre.getSheetByName(CONFIG.HOJA_MADRE_NAME);
    if (hojaMadre) {
      diag.totalFilasMadre = hojaMadre.getLastRow() - 1;
      const limite = Math.min(200, diag.totalFilasMadre);
      const clavesMadre = new Set(hojaMadre.getRange(2, CONFIG.COL_CLAVE + 1, limite, 1).getValues().map(r => String(r[0] || '').trim()).filter(Boolean));
      let matches = 0;
      if (diag.totalFilas >= 2) {
        const limiteExt = Math.min(200, diag.totalFilas - 1);
        const clavesExt = hoja.getRange(2, 1, limiteExt, 1).getValues().map(r => String(r[0] || '').trim()).filter(Boolean);
        clavesExt.forEach(k => { if (clavesMadre.has(k)) matches++; });
      }
      diag.coincidenciasConMadre = matches;
    }

    Logger.log("=== RESULTADO DIAGNÓSTICO ===");
    Logger.log(JSON.stringify(diag, null, 2));
    return diag;
  } catch (err) {
    diag.error = err.message;
    Logger.log("Error en diagnóstico: " + err.message);
    return diag;
  }
}

// ==================================================================
// REPARACIÓN Y AJUSTE DE COLUMNAS EN TABLA GESTION_LINKS
// ==================================================================
/**
 * Ajusta y repara la tabla Gestion_Links en la planilla externa:
 * Localiza las columnas 'Exámenes a Realizar', 'Fecha Realización (emi)' y 'Resultado (emi)'.
 * Elimina errores de fórmula (#REF!, #N/A, #VALUE!) cruzando directamente por Clave
 * con la Hoja Madre y volcando los valores limpios y formateados.
 */
function repararYCompletarTablaGestionLinks() {
  const resultado = {
    status: "OK",
    filasActualizadas: 0,
    columnasDetectadas: {},
    msg: ""
  };

  try {
    const ssExterna = SpreadsheetApp.openById(CONFIG.EXTERNA_ID);
    let hoja = null;
    if (CONFIG.EXTERNA_GID) {
      hoja = ssExterna.getSheets().find(s => s.getSheetId() === Number(CONFIG.EXTERNA_GID));
    }
    if (!hoja) {
      hoja = ssExterna.getSheetByName(CONFIG.EXTERNA_NAME) ||
             ssExterna.getSheetByName("Gestion_Links") ||
             ssExterna.getSheetByName("Gestion Links") ||
             ssExterna.getSheetByName("Gestión Empleados");
    }
    if (!hoja) {
      throw new Error("No se encontró la pestaña en la planilla externa.");
    }

    const ultFila = hoja.getLastRow();
    const ultCol  = hoja.getLastColumn();
    if (ultFila < 2 || ultCol < 1) {
      return { status: "OK", msg: "La hoja externa no contiene datos." };
    }

    // 1. Detectar índices de columnas en fila 1
    const encabezados = hoja.getRange(1, 1, 1, ultCol).getValues()[0];
    let colClave    = -1;
    let colEstudio  = -1;
    let colFechaEmi = -1;
    let colResEmi   = -1;

    encabezados.forEach((h, idx) => {
      const hNorm = String(h || '').trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\s_-]+/g, ' ');
      if (hNorm === 'clave' || hNorm === 'clave cruce' || hNorm === 'id') {
        colClave = idx;
      } else if (hNorm.includes('examen') || hNorm.includes('estudio a realizar') || hNorm === 'estudio') {
        colEstudio = idx;
      } else if (hNorm.includes('fecha realiz') || hNorm.includes('fecha estudio') || hNorm === 'fecha' || hNorm.includes('fecha examen')) {
        if (!hNorm.includes('ingreso') && !hNorm.includes('deriv') && !hNorm.includes('inicio') && !hNorm.includes('rar') && !hNorm.includes('vigencia')) {
          colFechaEmi = idx;
        }
      } else if (hNorm.includes('resultado')) {
        if (!hNorm.includes('rar')) {
          colResEmi = idx;
        }
      }
    });

    if (colClave === -1) colClave = 0; // Columna A por defecto

    resultado.columnasDetectadas = {
      colClave: colClave + 1,
      colEstudio: colEstudio !== -1 ? colEstudio + 1 : "No encontrada",
      colFechaEmi: colFechaEmi !== -1 ? colFechaEmi + 1 : "No encontrada",
      colResEmi: colResEmi !== -1 ? colResEmi + 1 : "No encontrada"
    };

    // 2. Leer datos de la Hoja Madre para cruce con detección dinámica de columnas
    const ssMadre = SpreadsheetApp.openById(CONFIG.HOJA_MADRE_ID);
    const hojaMadre = ssMadre.getSheetByName(CONFIG.HOJA_MADRE_NAME);
    if (!hojaMadre) throw new Error("No se encontró la pestaña Madre en el archivo principal.");
    const ultFilaMadre = hojaMadre.getLastRow();
    if (ultFilaMadre < 2) throw new Error("La Hoja Madre está vacía.");

    const ultColMadre = hojaMadre.getLastColumn();
    const encabezadosMadre = hojaMadre.getRange(1, 1, 1, ultColMadre).getValues()[0];

    let idxClaveMadre = CONFIG.COL_CLAVE !== undefined ? CONFIG.COL_CLAVE : 0;
    let idxEstudioMadre = CONFIG.COL_ESTUDIO !== undefined ? CONFIG.COL_ESTUDIO : 19;
    let idxFechaEstudioMadre = CONFIG.COL_FECHA_ESTUDIO !== undefined ? CONFIG.COL_FECHA_ESTUDIO : 27;
    let idxResultadoMadre = CONFIG.COL_RESULTADO !== undefined ? CONFIG.COL_RESULTADO : 28;

    encabezadosMadre.forEach((h, idx) => {
      const hNorm = String(h || '').trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\s_-]+/g, ' ');
      if (hNorm === 'clave') idxClaveMadre = idx;
      else if (hNorm === 'estudio a realizar' || hNorm === 'estudio') idxEstudioMadre = idx;
      else if (hNorm === 'fecha' || hNorm === 'fecha estudio') idxFechaEstudioMadre = idx;
      else if (hNorm === 'resultado' || hNorm === 'resultado auditoria') idxResultadoMadre = idx;
    });

    const maxColsMadre = Math.max(idxClaveMadre + 1, idxEstudioMadre + 1, idxFechaEstudioMadre + 1, idxResultadoMadre + 1, 35);
    const matrizMadre = hojaMadre.getRange(2, 1, ultFilaMadre - 1, maxColsMadre).getValues();

    // Discriminadores estrictos de tipo para evitar mezclar Fechas con Resultados Clínicos
    function esFormatoFecha(val) {
      if (!val) return false;
      if (val instanceof Date) return true;
      const s = String(val).trim();
      return /^\d{1,4}[\/\-]\d{1,2}[\/\-]\d{1,4}/.test(s);
    }

    function esResultadoClinico(val) {
      if (!val) return false;
      const s = String(val).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      return s.includes('normal') ||
             s.includes('desvio') ||
             s.includes('pendiente') ||
             s.includes('ausente') ||
             s.includes('enfermedad') ||
             s.includes('defectuosa') ||
             s.includes('desistido');
    }

    const mapaMadre = new Map();
    matrizMadre.forEach(row => {
      const clave = String(row[idxClaveMadre] || '').trim();
      if (!clave) return;

      const estudioCrudo = String(row[idxEstudioMadre] || '').trim();
      let fEstCruda = row[idxFechaEstudioMadre];
      let resCrudo = String(row[idxResultadoMadre] || '').trim();

      let fEstStr = formatearFecha(fEstCruda);

      // Si por desfasaje de columna en Madre la Fecha cayó en Resultado:
      if (esFormatoFecha(resCrudo) && !esFormatoFecha(fEstStr)) {
        const aux = resCrudo;
        resCrudo = fEstStr;
        fEstStr = formatearFecha(aux);
      }

      // Si el resultado sigue pareciendo una fecha, NUNCA permitirlo como resultado clínico
      if (esFormatoFecha(resCrudo) && !esResultadoClinico(resCrudo)) {
        if (!fEstStr) fEstStr = formatearFecha(resCrudo);
        resCrudo = '';
      }

      // Limpieza de errores de fórmula tipo #N/A o #REF!
      if (fEstStr.startsWith('#')) fEstStr = '';
      if (resCrudo.startsWith('#')) resCrudo = '';
      const estudioLimpio = estudioCrudo.startsWith('#') ? '' : estudioCrudo;

      mapaMadre.set(clave, {
        estudio: estudioLimpio,
        fechaEstudio: fEstStr,
        resultado: resCrudo
      });
    });

    // 3. Leer Claves de la hoja externa
    const nFilasExt = ultFila - 1;
    const clavesExt = hoja.getRange(2, colClave + 1, nFilasExt, 1).getValues();

    // 4. Preparar matrices de reemplazo limpias sin errores
    const valoresEstudio = [];
    const valoresFecha   = [];
    const valoresRes     = [];

    for (let i = 0; i < nFilasExt; i++) {
      const c = String(clavesExt[i][0] || '').trim();
      const datosM = mapaMadre.get(c);

      if (datosM) {
        valoresEstudio.push([datosM.estudio || '']);
        valoresFecha.push([datosM.fechaEstudio || '']);
        valoresRes.push([datosM.resultado || '']);
      } else {
        valoresEstudio.push(['']);
        valoresFecha.push(['']);
        valoresRes.push(['']);
      }
    }

    // 5. Volcar en la hoja externa en las columnas identificadas
    if (colEstudio !== -1) {
      hoja.getRange(2, colEstudio + 1, nFilasExt, 1).setValues(valoresEstudio);
    }
    if (colFechaEmi !== -1) {
      hoja.getRange(2, colFechaEmi + 1, nFilasExt, 1).setValues(valoresFecha);
    }
    if (colResEmi !== -1) {
      hoja.getRange(2, colResEmi + 1, nFilasExt, 1).setValues(valoresRes);
    }

    SpreadsheetApp.flush();
    resultado.filasActualizadas = nFilasExt;
    resultado.msg = `Se repararon y actualizaron con éxito ${nFilasExt} filas en la planilla externa.`;
    Logger.log("=== RESULTADO REPARACIÓN GESTION_LINKS ===");
    Logger.log(JSON.stringify(resultado, null, 2));
    return resultado;
  } catch (err) {
    resultado.status = "ERROR";
    resultado.msg = err.message;
    Logger.log("Error en repararYCompletarTablaGestionLinks: " + err.message);
    return resultado;
  }
}
