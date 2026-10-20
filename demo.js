// Google "de mentira" para el modo de prueba de la web (?demo).
// construir.py lo junta con Codigo.gs y salen en web/demo.js: así la prueba usa EL MISMO cerebro
// que Google, pero guardando en la memoria del navegador (al recargar, vuelve a empezar).
// Los datos de ejemplo son inventados: esta web es pública y no debe llevar nada real.
(function () {
  function hojaFalsa(nombre) {
    const filas = [];
    const h = {
      nombre: nombre,
      setName: function (n) { h.nombre = n; return h; },
      getDataRange: function () { return { getValues: function () { return filas.map(function (f) { return f.slice(); }); } }; },
      getLastRow: function () { return filas.length; },
      appendRow: function (f) { filas.push(f.slice()); },
      setFrozenRows: function () {},
      getRange: function (fila, col, nf, nc) {
        return {
          setValues: function (v) { v.forEach(function (f, i) { filas[fila - 1 + i] = f.slice(); }); return { setFontWeight: function () {} }; },
          setValue: function (v) { filas[fila - 1][col - 1] = v; },
          getValues: function () { return filas.slice(fila - 1, fila - 1 + (nf || 1)).map(function (f) { return f.slice(col - 1, col - 1 + (nc || 1)); }); }
        };
      }
    };
    return h;
  }
  let libro = null;
  const props = { CLAVE: 'demo' }, cache = {};
  const encuesta = hojaFalsa('respuestas');
  encuesta.appendRow(['Marca temporal']);
  for (let i = 0; i < 8; i++) encuesta.appendRow(['fila ' + i]);
  // Huella sencilla (en Google es SHA-256; aquí basta con que sea siempre la misma)
  function huellaDemo(t) { let h = 2166136261; for (let i = 0; i < t.length; i++) { h ^= t.charCodeAt(i); h = Math.imul(h, 16777619); } return [h & 255, (h >> 8) & 255, (h >> 16) & 255, (h >>> 24)]; }

  const google = {
    SpreadsheetApp: {
      create: function () {
        const hojas = [hojaFalsa('Hoja 1')];
        libro = { getId: function () { return 'DEMO'; }, getUrl: function () { return '#'; }, getSheets: function () { return hojas; },
          insertSheet: function (n) { const x = hojaFalsa(n); hojas.push(x); return x; },
          getSheetByName: function (n) { return hojas.find(function (x) { return x.nombre === n; }) || null; } };
        return libro;
      },
      openById: function (id) {
        if (id === 'DEMO') return libro;
        if (id === 'ENCUESTA-DEMO') return { getSheets: function () { return [encuesta]; } };
        throw new Error('sin permiso');
      }
    },
    PropertiesService: { getScriptProperties: function () { return { getProperty: function (k) { return props[k] || null; }, setProperty: function (k, v) { props[k] = v; } }; } },
    CacheService: { getScriptCache: function () { return { get: function (k) { return cache[k] || null; }, put: function (k, v) { cache[k] = v; } }; } },
    LockService: { getScriptLock: function () { return { waitLock: function () {}, releaseLock: function () {} }; } },
    ContentService: { MimeType: { JSON: 'j' }, createTextOutput: function (t) { return { setMimeType: function () { return JSON.parse(t); } }; } },
    HtmlService: { createHtmlOutput: function (h) { return h; } },
    Session: { getScriptTimeZone: function () { return 'Atlantic/Canary'; } },
    Logger: { log: function () {} },
    Utilities: {
      sleep: function () {},
      formatDate: function (d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); },
      DigestAlgorithm: { SHA_256: 1 },
      computeDigest: function (a, t) { return huellaDemo(t); }
    }
  };

  // Aquí construir.py mete el cerebro (Codigo.gs) tal cual.
  const cerebro = function (SpreadsheetApp, PropertiesService, CacheService, LockService, ContentService, HtmlService, Session, Logger, Utilities) {
// 🧭 Centro de mando (20/10/2026)
// El "cerebro" de la web del centro de mando. Guarda todo en una hoja de Google
// ("Centro de mando · datos") que se crea sola la primera vez.
// Dos puertas:
//   - el JEFE (Emmanuel) entra con su contraseña: ve y cambia todo;
//   - cada persona del EQUIPO entra con su código: solo ve SUS tareas y puede marcarlas hechas.
// Nunca se borra nada: lo quitado se marca en la columna "quitado" y se puede recuperar.

// La pantalla vive en GitHub Pages (WEB). Esta dirección de Google no enseña datos.
const WEB = 'https://manuromher21canario-sudo.github.io/centro-mando/';
const MAX_FALLOS = 10;          // intentos con contraseña o código malos antes de bloquear
const BLOQUEO_SEGUNDOS = 900;   // 15 minutos

// Cada pestaña de la hoja: sus columnas y qué se puede escribir en cada una.
// tipo: texto (con largo máximo), num, opcion (de una lista), fecha (la pone el código).
const TABLAS = {
  Tareas: {
    titulo: { tipo: 'texto', max: 120, obligatorio: true },
    porque: { tipo: 'texto', max: 600 },
    como: { tipo: 'texto', max: 1500 },
    para: { tipo: 'texto', max: 40 },            // id de la persona del equipo, o "jefe"
    idea: { tipo: 'texto', max: 60 },
    prioridad: { tipo: 'opcion', opciones: ['principal', 'secundaria', 'terciaria'] },
    estado: { tipo: 'opcion', opciones: ['pendiente', 'en curso', 'hecha'] },
    para_cuando: { tipo: 'texto', max: 10 },     // aaaa-mm-dd
    hecha_el: { tipo: 'fecha' },
    comentario: { tipo: 'texto', max: 1000 }     // lo que cuenta quien la hace
  },
  Equipo: {
    nombre: { tipo: 'texto', max: 40, obligatorio: true },
    rol: { tipo: 'texto', max: 80 },
    huella: { tipo: 'interno' },                 // huella del código (el código no se guarda)
    ultima_vez: { tipo: 'fecha' }
  },
  Dinero: {
    tipo: { tipo: 'opcion', opciones: ['ingreso', 'gasto'], obligatorio: true },
    concepto: { tipo: 'texto', max: 80, obligatorio: true },
    categoria: { tipo: 'texto', max: 40 },
    importe: { tipo: 'num', min: 0.01, max: 100000, obligatorio: true },
    idea: { tipo: 'texto', max: 60 },
    dia: { tipo: 'texto', max: 10 }              // aaaa-mm-dd (si no, el de hoy)
  },
  Clientes: {
    nombre: { tipo: 'texto', max: 60, obligatorio: true },
    negocio: { tipo: 'texto', max: 80 },
    clase: { tipo: 'opcion', opciones: ['servicio', 'app'], obligatorio: true },
    que: { tipo: 'texto', max: 80 },             // qué servicio o qué app
    estado: { tipo: 'opcion', opciones: ['posible', 'hablando', 'cliente', 'ya no'] },
    contacto: { tipo: 'texto', max: 80 },        // por dónde se habla (WhatsApp, en persona…)
    importe: { tipo: 'num', min: 0, max: 100000 },
    notas: { tipo: 'texto', max: 1000 }
  },
  Ideas: {
    nombre: { tipo: 'texto', max: 60, obligatorio: true },
    paso: { tipo: 'opcion', opciones: ['idea', 'lienzo', 'marca', 'prototipo', 'presentación', 'probando con clientes', 'vendiendo', 'en pausa', 'descartada'] },
    siguiente: { tipo: 'texto', max: 200 },
    enlace: { tipo: 'texto', max: 300 },
    notas: { tipo: 'texto', max: 1000 }
  },
  Talleres: {
    nombre: { tipo: 'texto', max: 60, obligatorio: true },
    estado: { tipo: 'opcion', opciones: ['en marcha', 'en pausa', 'terminado'] },
    hechas: { tipo: 'num', min: 0, max: 999 },
    total: { tipo: 'num', min: 0, max: 999 },
    enlace: { tipo: 'texto', max: 300 },
    notas: { tipo: 'texto', max: 1000 }
  },
  Encuestas: {
    nombre: { tipo: 'texto', max: 60, obligatorio: true },
    idea: { tipo: 'texto', max: 60 },
    enlace: { tipo: 'texto', max: 300 },         // el que se reparte (forms.gle/…)
    hoja: { tipo: 'texto', max: 100 },           // ID de la hoja de respuestas (para contarlas)
    meta: { tipo: 'num', min: 0, max: 10000 }
  }
};
const COMUNES = ['id', 'creado', 'quitado'];
const columnas_ = (tabla) => COMUNES.concat(Object.keys(TABLAS[tabla]));

function doGet() {
  return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">El centro de mando está en <a href="' + WEB + '" target="_top">' + WEB + '</a></p>');
}

function doPost(e) {
  let cuerpo = {};
  try { cuerpo = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, error: 'peticion' }); }
  const cache = CacheService.getScriptCache();
  const fallos = Number(cache.get('fallos') || 0);
  if (fallos >= MAX_FALLOS) return json_({ ok: false, error: 'bloqueado' });
  const fallo = (error) => {
    cache.put('fallos', String(fallos + 1), BLOQUEO_SEGUNDOS);
    Utilities.sleep(1000);  // frena a quien pruebe contraseñas una tras otra
    return json_({ ok: false, error: error });
  };
  try {
    if (cuerpo.codigo !== undefined) {
      const persona = personaDelCodigo_(String(cuerpo.codigo));
      if (!persona) return fallo('codigo');
      return json_(puertaEquipo_(persona, cuerpo));
    }
    // La contraseña está en Configuración del proyecto > Propiedades del script > CLAVE
    // (la pone Emmanuel; no está en el código).
    const clave = PropertiesService.getScriptProperties().getProperty('CLAVE');
    if (!clave) return json_({ ok: false, error: 'sin-clave' });
    if (String(cuerpo.clave || '') !== clave) return fallo('clave');
    return json_(puertaJefe_(cuerpo));
  } catch (err) {
    return json_({ ok: false, error: 'interno', detalle: String(err && err.message || err).slice(0, 200) });
  }
}

// ---------- Jefe ----------
function puertaJefe_(c) {
  let extra = {};
  if (c.accion === 'guardar') {
    const r = conCerrojo_(() => guardar_(c.tabla, c.id, c.campos || {}));
    if (r.error) return { ok: false, error: r.error };
    extra = { id: r.id };
    if (c.tabla === 'Equipo' && !c.id) extra.codigo = darCodigo_(r.id);
  } else if (c.accion === 'quitar' || c.accion === 'recuperar') {
    const r = conCerrojo_(() => marcarQuitado_(c.tabla, c.id, c.accion === 'quitar'));
    if (r.error) return { ok: false, error: r.error };
  } else if (c.accion === 'codigo-nuevo') {
    // Si alguien pierde su código, se le da otro (el viejo deja de valer).
    if (!buscarFila_('Equipo', c.id)) return { ok: false, error: 'no-existe' };
    extra = { codigo: conCerrojo_(() => darCodigo_(c.id)) };
  } else if (c.accion === 'aviso') {
    // Un aviso fijo que el equipo ve arriba al entrar (vacío = sin aviso)
    const texto = String(c.texto || '').trim().slice(0, 500).replace(/^[=+\-@]+/, '');
    PropertiesService.getScriptProperties().setProperty('AVISO_EQUIPO', texto);
  } else if (c.accion && c.accion !== 'datos') {
    return { ok: false, error: 'accion' };
  }
  const datos = {};
  Object.keys(TABLAS).forEach((t) => { datos[t] = leer_(t).map(sinInterno_); });
  datos.Encuestas.forEach((x) => { x.respuestas = contarRespuestas_(x.hoja); });
  return Object.assign({ ok: true, quien: 'jefe', datos: datos, aviso: aviso_(), ahora: Date.now() }, extra);
}

// ---------- Equipo ----------
function puertaEquipo_(persona, c) {
  if (c.accion === 'hecha' || c.accion === 'deshacer' || c.accion === 'comentar' || c.accion === 'empezar') {
    const tarea = leer_('Tareas').find((t) => t.id === c.id && t.para === persona.id && !t.quitado);
    if (!tarea) return { ok: false, error: 'no-existe' };
    const campos = {};
    if (c.accion === 'hecha') { campos.estado = 'hecha'; campos.hecha_el = new Date(); }
    if (c.accion === 'deshacer') { campos.estado = 'pendiente'; campos.hecha_el = ''; }
    if (c.accion === 'empezar') campos.estado = 'en curso';
    if (c.comentario !== undefined) campos.comentario = c.comentario;
    const r = conCerrojo_(() => guardar_('Tareas', c.id, campos, true));
    if (r.error) return { ok: false, error: r.error };
  } else if (c.accion && c.accion !== 'datos') {
    return { ok: false, error: 'accion' };
  }
  conCerrojo_(() => guardar_('Equipo', persona.id, { ultima_vez: new Date() }, true));
  const mias = leer_('Tareas').filter((t) => t.para === persona.id && !t.quitado)
    .map((t) => ({ id: t.id, titulo: t.titulo, porque: t.porque, como: t.como, idea: t.idea, prioridad: t.prioridad,
                   estado: t.estado || 'pendiente', para_cuando: t.para_cuando, hecha_el: t.hecha_el, comentario: t.comentario }));
  return { ok: true, quien: 'equipo', nombre: persona.nombre, tareas: mias, aviso: aviso_(), ahora: Date.now() };
}

function aviso_() {
  return PropertiesService.getScriptProperties().getProperty('AVISO_EQUIPO') || '';
}

// ---------- La hoja ----------
function libro_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('HOJA');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('Centro de mando · datos');
  Object.keys(TABLAS).forEach((t, i) => {
    const hoja = i === 0 ? ss.getSheets()[0].setName(t) : ss.insertSheet(t);
    const cols = columnas_(t);
    hoja.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    hoja.setFrozenRows(1);
  });
  props.setProperty('HOJA', ss.getId());
  return ss;
}

function hoja_(tabla) {
  if (!TABLAS[tabla]) throw new Error('tabla');
  return libro_().getSheetByName(tabla);
}

function leer_(tabla) {
  const valores = hoja_(tabla).getDataRange().getValues();
  const cab = valores[0] || [];
  return valores.slice(1).filter((f) => f[0]).map((f) => {
    const o = {};
    cab.forEach((k, i) => { o[k] = f[i] instanceof Date ? f[i].getTime() : f[i]; });
    o.id = String(o.id);
    o.quitado = String(o.quitado) === 'sí';
    return o;
  });
}

function sinInterno_(o) {
  const copia = Object.assign({}, o);
  delete copia.huella;
  return copia;
}

function buscarFila_(tabla, id) {
  if (!id) return 0;
  const ids = hoja_(tabla).getRange(1, 1, Math.max(hoja_(tabla).getLastRow(), 1), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  return 0;
}

// Comprueba y limpia lo que llega de la web. Devuelve { campos } o { error }.
function limpiar_(tabla, campos, interno) {
  const esquema = TABLAS[tabla];
  const limpio = {};
  for (const k of Object.keys(campos)) {
    const regla = esquema[k];
    if (!regla) return { error: 'campo:' + k };
    if ((regla.tipo === 'interno' || regla.tipo === 'fecha') && !interno) return { error: 'campo:' + k };
    let v = campos[k];
    if (regla.tipo === 'texto') {
      // los textos no pueden empezar por "=", "+", "-" o "@": Sheets los tomaría por fórmulas
      v = String(v === null || v === undefined ? '' : v).trim().slice(0, regla.max).replace(/^[=+\-@]+/, '');
    } else if (regla.tipo === 'num') {
      if (v === '' || v === null || v === undefined) { v = ''; } else {
        v = Math.round(Number(String(v).replace(',', '.')) * 100) / 100;
        if (!(v >= regla.min && v <= regla.max)) return { error: 'numero:' + k };
      }
    } else if (regla.tipo === 'opcion') {
      v = String(v || '');
      if (v && regla.opciones.indexOf(v) === -1) return { error: 'opcion:' + k };
    }
    limpio[k] = v;
  }
  return { campos: limpio };
}

function guardar_(tabla, id, campos, interno) {
  if (!TABLAS[tabla]) return { error: 'tabla' };
  const r = limpiar_(tabla, campos, interno);
  if (r.error) return r;
  const hoja = hoja_(tabla);
  const cols = columnas_(tabla);
  let fila = buscarFila_(tabla, id);
  if (id && !fila) return { error: 'no-existe' };
  if (!fila) {
    const faltan = Object.keys(TABLAS[tabla]).filter((k) => TABLAS[tabla][k].obligatorio && (r.campos[k] === undefined || r.campos[k] === ''));
    if (faltan.length) return { error: 'falta:' + faltan.join(',') };
    id = tabla.slice(0, 1).toLowerCase() + Date.now().toString(36) + Math.floor(Math.random() * 1296).toString(36);
    const nueva = cols.map((k) => k === 'id' ? id : k === 'creado' ? new Date() : k === 'quitado' ? '' : (r.campos[k] !== undefined ? r.campos[k] : ''));
    if (tabla === 'Dinero' && !r.campos.dia) nueva[cols.indexOf('dia')] = hoyTexto_();
    hoja.appendRow(nueva);
    return { id: id };
  }
  Object.keys(r.campos).forEach((k) => {
    if (TABLAS[tabla][k].obligatorio && r.campos[k] === '') return;  // un dato obligatorio no se vacía
    hoja.getRange(fila, cols.indexOf(k) + 1).setValue(r.campos[k]);
  });
  return { id: id };
}

function marcarQuitado_(tabla, id, quitar) {
  if (!TABLAS[tabla]) return { error: 'tabla' };
  const fila = buscarFila_(tabla, id);
  if (!fila) return { error: 'no-existe' };
  hoja_(tabla).getRange(fila, columnas_(tabla).indexOf('quitado') + 1).setValue(quitar ? 'sí' : '');
  return { id: id };
}

function conCerrojo_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function hoyTexto_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// ---------- Códigos del equipo ----------
// El código (6 cifras) solo se enseña una vez al jefe; en la hoja se guarda su huella.
function huella_(texto) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'centro-mando:' + texto);
  return bytes.map((b) => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}

function darCodigo_(id) {
  const usadas = leer_('Equipo').map((p) => p.huella);
  let codigo;
  do { codigo = String(Math.floor(100000 + Math.random() * 900000)); } while (usadas.indexOf(huella_(codigo)) !== -1);
  guardar_('Equipo', id, { huella: huella_(codigo) }, true);
  return codigo;
}

function personaDelCodigo_(codigo) {
  if (!/^\d{6}$/.test(codigo)) return null;
  const h = huella_(codigo);
  return leer_('Equipo').find((p) => p.huella === h && !p.quitado) || null;
}

// ---------- Encuestas ----------
// Cuenta las filas de la hoja de respuestas (menos la de títulos). Si no se puede abrir, null.
function contarRespuestas_(idHoja) {
  if (!idHoja) return null;
  try {
    const hoja = SpreadsheetApp.openById(String(idHoja)).getSheets()[0];
    return Math.max(hoja.getLastRow() - 1, 0);
  } catch (e) { return null; }
}

function json_(objeto) {
  return ContentService.createTextOutput(JSON.stringify(objeto)).setMimeType(ContentService.MimeType.JSON);
}

// Para ejecutar UNA vez desde el editor (botón ▶): pide los permisos y crea la hoja.
function instalar() {
  const ss = libro_();
  const cargado = typeof cargarIniciales_ === 'function' && conCerrojo_(cargarIniciales_);
  Logger.log('Hoja lista: ' + ss.getUrl() + (cargado ? ' (con lo que ya teníamos)' : ''));
}

    return { doPost: doPost, guardar_: guardar_, huella_: huella_ };
  };
  const c = cerebro(google.SpreadsheetApp, google.PropertiesService, google.CacheService, google.LockService, google.ContentService,
    google.HtmlService, google.Session, google.Logger, google.Utilities);

  let sembrado = false;
  function sembrar() {
    sembrado = true;
    const jefe = function (x) { return c.doPost({ postData: { contents: JSON.stringify(Object.assign({ clave: 'demo' }, x)) } }); };
    const g = function (tabla, campos) { return jefe({ accion: 'guardar', tabla: tabla, campos: campos }).id; };
    const amigo = g('Equipo', { nombre: 'Álex', rol: 'Probar en el móvil' });
    c.guardar_('Equipo', amigo, { huella: c.huella_('123456') }, true);   // en la prueba, su código es 123456
    g('Ideas', { nombre: 'App de cuentas para manitas', paso: 'presentación', siguiente: 'Enseñársela a 3 personas' });
    g('Ideas', { nombre: 'Llamada de los buenos días', paso: 'prototipo', siguiente: 'Ponerle nombre' });
    g('Ideas', { nombre: 'Inventario de piezas', paso: 'prototipo', siguiente: 'Esperar a saber qué le molesta a la clienta' });
    g('Ideas', { nombre: 'Tablón de la clase', paso: 'idea' });
    g('Tareas', { titulo: 'Probar el prototipo en tu móvil', para: amigo, prioridad: 'principal',
      porque: 'Quiero saber si alguien que no lo ha visto nunca lo entiende sin ayuda.',
      como: '1. Abre el enlace que te mandé.\n2. Intenta apuntar un trabajo sin que nadie te explique.\n3. Apunta lo que no entiendas.' });
    g('Tareas', { titulo: 'Repartir la encuesta en 2 grupos', para: amigo, prioridad: 'secundaria', porque: 'Necesitamos 20 respuestas para decidir.',
      como: 'Mándala a un grupo de clase y a uno de la familia con el mensaje de «Mandar».' });
    g('Tareas', { titulo: 'Mirar si el nombre está libre en la oficina de marcas', para: 'jefe', prioridad: 'principal' });
    g('Tareas', { titulo: 'Escribir a la tienda de piezas', para: 'jefe', prioridad: 'secundaria' });
    g('Tareas', { titulo: 'Buscar 3 ideas nuevas', para: 'jefe', prioridad: 'terciaria' });
    g('Clientes', { nombre: 'Rosa', negocio: 'Tienda de piezas', clase: 'servicio', que: 'Inventario', estado: 'hablando', contacto: 'WhatsApp' });
    g('Clientes', { nombre: 'Peluquería Lola', negocio: 'Peluquería', clase: 'servicio', que: 'Caja del día', estado: 'cliente', importe: 75 });
    g('Clientes', { nombre: 'Juan', negocio: 'Trabajos en casas', clase: 'app', que: 'App de cuentas', estado: 'posible' });
    g('Dinero', { tipo: 'ingreso', concepto: 'Caja del día (Peluquería Lola)', categoria: 'Servicios', importe: 75 });
    g('Dinero', { tipo: 'gasto', concepto: 'Cartulinas para carteles', categoria: 'Material', importe: 4.5 });
    g('Encuestas', { nombre: 'Encuesta de manitas', idea: 'App de cuentas para manitas', enlace: 'https://forms.gle/ejemplo', hoja: 'ENCUESTA-DEMO', meta: 20 });
    g('Talleres', { nombre: 'Centro de mando', estado: 'en marcha', hechas: 3, total: 13 });
    g('Talleres', { nombre: 'Inventario de piezas', estado: 'en pausa', hechas: 3, total: 17 });
  }
  window.cerebroDemo = function (cuerpo) {
    if (!sembrado) sembrar();
    return c.doPost({ postData: { contents: JSON.stringify(cuerpo) } });
  };
})();
