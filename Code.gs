/**
 * DOMUS MÓVIL v3 — Buzón central (Google Sheets + Apps Script)
 * Une Domus ID (PC) con la app de los prestadores. La PC es la fuente de verdad.
 * CLAVE_SISTEMA: la misma de domus_movil.py (nadie la tipea nunca).
 * Al actualizar: Implementar → Administrar implementaciones → Editar → Nueva versión (la URL no cambia).
 * v3: la PC puede eliminar visitas y partes del buzón (borrar_directivas / borrar_informes).
 */
const CLAVE_SISTEMA = 'Domus-Central-2026-Movil-47-Lanterna-Verde-8392';
const VERSION = '3';

const HOJAS = {
  Prestadores: ['token','nombre','rol','telefono','activo','alta','ultimo_acceso'],
  Directivas:  ['id','token','paciente','domicilio','fecha','hora','tareas','notas','estado','creada','vista',
                'paciente_id','agenda_id','dni','tel_afiliado','responsable','form','form_titulo'],
  Informes:    ['id','directiva_id','token','llegada','lat','lng','salida','ta','fc','fr','temp','sat',
                'glucemia','dolor','procedimientos','evolucion','insumos','fotos','recibido','importado',
                'resultado','motivo','campos']
};
const DATOS_DIRECTIVA = HOJAS.Directivas.filter(k => !['id','token','estado','creada','vista'].includes(k));

/* ---------- Estructura (se ejecuta sola al cambiar de versión; no toca datos) ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HOJAS).forEach(n => {
    const h = ss.getSheetByName(n) || ss.insertSheet(n);
    h.getRange('A:Z').setNumberFormat('@');
    h.getRange(1, 1, 1, HOJAS[n].length).setValues([HOJAS[n]]).setFontWeight('bold');
    h.setFrozenRows(1);
  });
  carpetaFotos();
  PropertiesService.getScriptProperties().setProperty('VERSION', VERSION);
}
function asegurar() { if (PropertiesService.getScriptProperties().getProperty('VERSION') !== VERSION) setup(); }

/* ---------- Entradas web ---------- */
function doGet(e)  { return responder(e.parameter); }
function doPost(e) { return responder(JSON.parse(e.postData.contents || '{}')); }

function responder(p) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    asegurar();
    const r = p.admin ? central(p) : prestador(p);
    return json(Object.assign({ ok: true }, r));
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  } finally { lock.releaseLock(); }
}

/* ---------- Lado prestador (celular) ---------- */
function prestador(p) {
  const pr = filas('Prestadores').find(x => x.token === p.t && x.activo === 'si');
  if (!pr) throw new Error('Vínculo inválido o desactivado. Pedí un nuevo link a Domus.');
  poner('Prestadores', pr._fila, 'ultimo_acceso', ahora());

  if (p.a === 'yo') {
    const dirs = filas('Directivas')
      .filter(d => d.token === p.t && (d.estado === 'pendiente' || d.estado === 'vista')).map(limpio);
    return { nombre: pr.nombre, rol: pr.rol, directivas: dirs };
  }

  if (p.a === 'historial') {   // partes enviados por este prestador (quedan aunque cambie de teléfono)
    const dirs = mapa('Directivas', 'id');
    const lista = filas('Informes').filter(x => x.token === p.t).slice(-40).reverse().map(x => {
      const d = dirs[x.directiva_id] || {};
      return { id: x.id, paciente: d.paciente || '', domicilio: d.domicilio || '', titulo: d.form_titulo || '',
        llegada: x.llegada, salida: x.salida, resultado: x.resultado || 'realizada', motivo: x.motivo,
        evolucion: x.evolucion, ta: x.ta, fc: x.fc, temp: x.temp, sat: x.sat, campos: x.campos,
        insumos: x.insumos, fotos: x.fotos ? x.fotos.split(' ').length : 0, importado: x.importado ? 'si' : '' };
    });
    return { partes: lista };
  }

  if (p.a === 'visto') {
    const d = filas('Directivas').find(x => x.id === p.id && x.token === p.t);
    if (d && d.estado === 'pendiente') {
      poner('Directivas', d._fila, 'estado', 'vista');
      poner('Directivas', d._fila, 'vista', ahora());
    }
    return {};
  }

  if (p.a === 'informe') {
    const inf = p.informe || {};
    if (filas('Informes').some(x => x.id === inf.id)) return { id: inf.id, repetido: true }; // reintento offline
    const d = filas('Directivas').find(x => x.id === inf.directiva_id && x.token === p.t);
    if (!d) throw new Error('La visita no pertenece a este prestador o fue eliminada por Domus.');
    const urls = (inf.fotos || []).slice(0, 8).map((b64, i) => guardarFoto(b64, d.paciente + '_' + inf.id + '_' + (i + 1)));
    const campos = typeof inf.campos === 'string' ? inf.campos : JSON.stringify(inf.campos || {});
    agregar('Informes', Object.assign({}, inf, { token: p.t, fotos: urls.join(' '), campos,
      resultado: inf.resultado || 'realizada', recibido: ahora(), importado: '' }));
    poner('Directivas', d._fila, 'estado', inf.resultado === 'no_realizada' ? 'no realizada' : 'realizada');
    return { id: inf.id };
  }
  throw new Error('Acción desconocida');
}

/* ---------- Lado central (Domus ID) ---------- */
function central(p) {
  if (p.admin !== CLAVE_SISTEMA) throw new Error('No autorizado');

  switch (p.a) {
    case 'alta_prestador': {
      const token = Utilities.getUuid().replace(/-/g, '');
      agregar('Prestadores', { token, nombre: p.nombre, rol: p.rol, telefono: p.telefono, activo: 'si', alta: ahora() });
      return { token };
    }
    case 'actualizar_prestador': {   // espejo: manda la ficha de la PC
      const pr = filas('Prestadores').find(x => x.token === p.token);
      if (pr) ['nombre', 'rol', 'telefono'].forEach(k => poner('Prestadores', pr._fila, k, p[k] || ''));
      return {};
    }
    case 'baja_prestador': {
      const pr = filas('Prestadores').find(x => x.token === p.token);
      if (pr) poner('Prestadores', pr._fila, 'activo', 'no');
      return {};
    }
    case 'directiva': {
      const id = Utilities.getUuid().slice(0, 8), o = { id, token: p.token, estado: 'pendiente', creada: ahora() };
      DATOS_DIRECTIVA.forEach(k => o[k] = p[k]);
      agregar('Directivas', o);
      return { id };
    }
    case 'actualizar_directiva': {   // cambió domicilio, teléfono, nota, fecha, etc.
      const d = filas('Directivas').find(x => x.id === p.id);
      if (d && (d.estado === 'pendiente' || d.estado === 'vista'))
        DATOS_DIRECTIVA.forEach(k => { if (p[k] != null && String(p[k]) !== d[k]) poner('Directivas', d._fila, k, p[k]); });
      return {};
    }
    case 'cancelar': {
      const d = filas('Directivas').find(x => x.id === p.id);
      if (d && (d.estado === 'pendiente' || d.estado === 'vista')) poner('Directivas', d._fila, 'estado', 'cancelada');
      return {};
    }
    case 'borrar_directivas': {      // la PC elimina visitas del buzón (desaparecen del celular)
      const ids = new Set(p.ids || []);
      return { borradas: borrarFilas('Directivas', filas('Directivas').filter(x => ids.has(x.id))) };
    }
    case 'borrar_informes': {        // la PC elimina partes (también del historial del prestador)
      const ids = new Set(p.ids || []);
      return { borrados: borrarFilas('Informes', filas('Informes').filter(x => ids.has(x.id))) };
    }
    case 'informes_nuevos': {
      const pres = mapa('Prestadores', 'token'), dirs = mapa('Directivas', 'id');
      const lista = filas('Informes').filter(x => !x.importado).map(x => {
        const d = dirs[x.directiva_id] || {};
        return Object.assign(limpio(x), { prestador: (pres[x.token] || {}).nombre || '', paciente: d.paciente || '',
          domicilio: d.domicilio || '', paciente_id: d.paciente_id || '', agenda_id: d.agenda_id || '' });
      });
      return { informes: lista };
    }
    case 'marcar_importados': {
      const ids = new Set(p.ids || []);
      filas('Informes').filter(x => ids.has(x.id)).forEach(x => poner('Informes', x._fila, 'importado', ahora()));
      return {};
    }
    case 'tablero': {
      return {
        prestadores: filas('Prestadores').map(limpio),
        directivas: filas('Directivas').slice(-80).reverse().map(d => { const c = limpio(d); delete c.form; return c; }),
        sin_importar: filas('Informes').filter(x => !x.importado).length
      };
    }
  }
  throw new Error('Acción desconocida');
}

/* ---------- Utilidades ---------- */
function hoja(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function filas(n) {
  const v = hoja(n).getDataRange().getValues(), cab = v.shift();
  return v.map((r, i) => { const o = { _fila: i + 2 }; cab.forEach((k, j) => { if (k) o[k] = String(r[j]); }); return o; });
}
function mapa(n, k) { const m = {}; filas(n).forEach(x => m[x[k]] = x); return m; }
function agregar(n, o) { hoja(n).appendRow(HOJAS[n].map(k => o[k] == null ? '' : String(o[k]))); }
function poner(n, fila, campo, val) { hoja(n).getRange(fila, HOJAS[n].indexOf(campo) + 1).setValue(String(val)); }
function borrarFilas(n, lista) {  // de abajo hacia arriba para no correr los números de fila
  const h = hoja(n); lista.map(x => x._fila).sort((a, b) => b - a).forEach(f => h.deleteRow(f)); return lista.length;
}
function limpio(o) { const c = Object.assign({}, o); delete c._fila; return c; }
function ahora() { return Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm:ss'); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function carpetaFotos() {
  const pr = PropertiesService.getScriptProperties();
  let id = pr.getProperty('CARPETA');
  if (!id) { id = DriveApp.createFolder('Domus Móvil - Fotos de visitas').getId(); pr.setProperty('CARPETA', id); }
  return DriveApp.getFolderById(id);
}
function guardarFoto(dataUrl, nombre) {  // fotos (jpg/png) o PDF adjuntado desde el celular
  const s = String(dataUrl), mime = (s.match(/^data:([^;]+);/) || [, 'image/jpeg'])[1];
  const ext = mime === 'application/pdf' ? '.pdf' : mime === 'image/png' ? '.png' : '.jpg';
  const b = Utilities.base64Decode(s.split(',')[1] || '');
  return carpetaFotos().createFile(Utilities.newBlob(b, mime, nombre + ext)).getUrl(); // privada
}
