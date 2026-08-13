/* ===========================================================================
   Visor CF · Tablero de Gerencia RENOBO 2026-2027
   ---------------------------------------------------------------------------
   §0  Configuración y semilla
   §1  Sesión (OAuth) y API de Google Sheets
   §2  Modelo: parseo, edición y guardado
   §3  Cálculos
   §4  Gráficos (SVG dibujado a mano)
   §5  Vistas
   §6  Arranque
   =========================================================================== */

'use strict';

/* ===========================================================================
   §0  CONFIGURACIÓN Y SEMILLA
   =========================================================================== */

const ALCANCE = 'https://www.googleapis.com/auth/spreadsheets';
const CLAVE_CONFIG = 'visorcf.config';
const CLAVE_PENDIENTE = 'visorcf.pendiente';

/* Horizonte del tablero: ago-2026 → ene-2028, igual que el Excel original. */
const MESES = (() => {
  const salida = [];
  let anio = 2026, mes = 8;
  for (let i = 0; i < 18; i++) {
    salida.push(`${anio}-${String(mes).padStart(2, '0')}`);
    if (++mes > 12) { mes = 1; anio++; }
  }
  return salida;
})();

const PILARES = [
  { id: 'operador-urbano', nombre: 'Operador urbano' },
  { id: 'gestor-suelo',    nombre: 'Gestor de suelo' },
  { id: 'gestor-inmob',    nombre: 'Gestor inmobiliario' },
  { id: 'gestor-proy',     nombre: 'Gestor de proyectos' },
  { id: 'institucion',     nombre: 'Institución' }
];

const RESPONSABLES = [
  'Gerencia',
  'Subgerencia planeamiento',
  'Subgerencia corporativa',
  'Subgerencia ejecución'
];

/* Rangos de lectura. Se piden holgados; Sheets solo devuelve lo que existe. */
const RANGOS = ['KPIs!A1:J200', 'Serie!A1:Z200', 'Hitos!A1:D500', 'Bitacora!A1:D2000', 'Config!A1:B50'];

/* --- Semilla derivada de Dashboard.xlsx (hoja «Dashboard RenoBo») ---------
   Las unidades salen de los formatos de celda del archivo original.
   Los porcentajes se guardan como decimales (0,15 = 15 %).
   La moneda está en millones de COP.                                       */
const SEMILLA = {
  config: {
    titulo: 'RENOBO · Tablero de Gerencia 2026–2027',
    mesCorte: '2026-08',
    umbralAtencion: 0.4,
    revision: 1
  },
  kpis: [
    /* KPI-01 no deriva de hitos: el Excel nombra las 7 AE pendientes pero no las
       3 ya adoptadas, así que la lista sirve de checklist, no de contador.     */
    { id: 'KPI-01', pilar: 'operador-urbano', nombre: 'AE adoptadas',                                        unidad: 'conteo',     avance: 3,     meta2026: 6,     meta2027: null, metaFinal: 10,     responsable: 'Subgerencia planeamiento', avanceDesdeHitos: false },
    { id: 'KPI-02', pilar: 'operador-urbano', nombre: 'Ciudad Florida adoptada',                             unidad: 'porcentaje', avance: 0.15,  meta2026: 1,     meta2027: null, metaFinal: 1,      responsable: 'Subgerencia planeamiento', avanceDesdeHitos: false },
    { id: 'KPI-03', pilar: 'operador-urbano', nombre: 'Bono TIF emitido',                                    unidad: 'porcentaje', avance: 0.15,  meta2026: 0.7,   meta2027: null, metaFinal: 1,      responsable: 'Gerencia',                 avanceDesdeHitos: false },
    { id: 'KPI-04', pilar: 'operador-urbano', nombre: 'Recursos subasta de certificados usados',             unidad: 'moneda',     avance: 0,     meta2026: 5000,  meta2027: null, metaFinal: 42000,  responsable: 'Gerencia',                 avanceDesdeHitos: false },
    { id: 'KPI-05', pilar: 'operador-urbano', nombre: 'Vivienda licenciada (meta PDD)',                      unidad: 'conteo',     avance: 2491,  meta2026: 4851,  meta2027: null, metaFinal: 9000,   responsable: 'Subgerencia planeamiento', avanceDesdeHitos: false },
    { id: 'KPI-06', pilar: 'gestor-suelo',    nombre: 'Recursos FCO comprometidos',                          unidad: 'moneda',     avance: 138,   meta2026: 115000,meta2027: null, metaFinal: 236819, responsable: 'Gerencia',                 avanceDesdeHitos: false },
    { id: 'KPI-07', pilar: 'gestor-suelo',    nombre: 'Predios adquiridos Potosí',                           unidad: 'conteo',     avance: 31,    meta2026: 99,    meta2027: null, metaFinal: 166,    responsable: 'Subgerencia planeamiento', avanceDesdeHitos: false },
    { id: 'KPI-08', pilar: 'gestor-suelo',    nombre: 'Predios adquiridos Edén',                             unidad: 'conteo',     avance: 31,    meta2026: 59,    meta2027: null, metaFinal: 123,    responsable: 'Subgerencia planeamiento', avanceDesdeHitos: false },
    { id: 'KPI-09', pilar: 'gestor-inmob',    nombre: 'Comisiones de proyectos de reuso (corretaje)',        unidad: 'moneda',     avance: 0,     meta2026: 100,   meta2027: null, metaFinal: 1000,   responsable: 'Subgerencia planeamiento', avanceDesdeHitos: false },
    { id: 'KPI-10', pilar: 'gestor-inmob',    nombre: 'Calle 26 adjudicado',                                 unidad: 'porcentaje', avance: 0.15,  meta2026: 0.6,   meta2027: null, metaFinal: 1,      responsable: 'Gerencia',                 avanceDesdeHitos: false },
    { id: 'KPI-11', pilar: 'gestor-inmob',    nombre: 'M2 de entidades del distrito gestionados en arriendo',unidad: 'm2',         avance: 5000,  meta2026: 15000, meta2027: null, metaFinal: 15000,  responsable: 'Subgerencia corporativa',  avanceDesdeHitos: false },
    { id: 'KPI-12', pilar: 'gestor-proy',     nombre: 'Estudios y diseños entregados',                       unidad: 'conteo',     avance: 0,     meta2026: 6,     meta2027: null, metaFinal: 7,      responsable: 'Subgerencia ejecución',    avanceDesdeHitos: true  },
    { id: 'KPI-13', pilar: 'gestor-proy',     nombre: 'Obras entregadas',                                    unidad: 'conteo',     avance: 0,     meta2026: 5,     meta2027: null, metaFinal: 12,     responsable: 'Subgerencia ejecución',    avanceDesdeHitos: true  },
    { id: 'KPI-14', pilar: 'institucion',     nombre: 'Implementación sistema misional: subprocesos en uso', unidad: 'conteo',     avance: 1,     meta2026: 5,     meta2027: null, metaFinal: 20,     responsable: 'Subgerencia corporativa',  avanceDesdeHitos: false },
    { id: 'KPI-15', pilar: 'institucion',     nombre: 'Convenios y contratos liquidados',                    unidad: 'conteo',     avance: 1,     meta2026: 5,     meta2027: null, metaFinal: 13,     responsable: 'Gerencia',                 avanceDesdeHitos: false }
  ],
  /* El anexo de entregas del Excel, convertido en hitos del KPI que lo mide. */
  hitos: [
    ['KPI-01', 'Reencuentro'], ['KPI-01', 'Chapinero'], ['KPI-01', 'Montevideo'], ['KPI-01', 'Fontibón'],
    ['KPI-01', 'Chucua'], ['KPI-01', 'Teleport'], ['KPI-01', 'Sevillana'],
    ['KPI-12', 'Liceo Femenino Mercedes Nariño'], ['KPI-12', 'Liceo Agustín Nieto Caballero'],
    ['KPI-12', 'Col. Francisco de Paula Santander'], ['KPI-12', 'Espacios Emblemáticos (CHSJD)'],
    ['KPI-12', 'Pabellones Franceses (CHSJD)'], ['KPI-12', 'Ancianato San Pedro Claver (CHSJD)'],
    ['KPI-12', 'Centro Cultural Juvenil (2027)'],
    ['KPI-13', 'Bronx Distrito Creativo · Etapa 1'], ['KPI-13', 'Col. San Francisco Sede B'],
    ['KPI-13', 'Laboratorio Santiago Samper (CHSJD)'], ['KPI-13', 'Laboratorio Enfermedades Tropicales (CHSJD)'],
    ['KPI-13', 'Col. Teresa M. de Varela (La Magdalena)'],
    ['KPI-13', 'Nodo La Gloria (2027)'], ['KPI-13', 'Rampa la Concordia (2027)'],
    ['KPI-13', 'Edificio de Laboratorios · U. Distrital (2027)'], ['KPI-13', 'CAPS Bravo Páez · Salud (2027)'],
    ['KPI-13', 'Ecopunto San Cristóbal · UAESP (2027)'], ['KPI-13', 'Primera Infancia Las Brisas (2027)'],
    ['KPI-13', 'Colegio Altos de Egipto (2027)']
  ].map(([kpiId, nombre]) => ({ kpiId, nombre, hecho: false, fecha: '' })),
  /* La columna «Detalle» del Excel, con fecha de corte y autor. */
  bitacora: [
    ['KPI-01', 'Pendientes Reencuentro, Chapinero, Montevideo, Fontibón, Chucua, Teleport y Sevillana.'],
    ['KPI-02', 'Adoptación Resolución Renobo.'],
    ['KPI-03', 'Aprobación Concejo en 2026.'],
    ['KPI-06', 'Revisar unidades: el avance registrado (138) y la meta (115.000) no parecen estar en la misma escala.'],
    ['KPI-07', 'Ofertas aceptadas o expropiaciones realizadas.'],
    ['KPI-08', 'Ofertas aceptadas o expropiaciones realizadas.'],
    ['KPI-10', 'Pliegos definitivos en 2026.'],
    ['KPI-11', 'Pliegos definitivos en 2027.'],
    ['KPI-15', 'Según listado definido por JD Ching.']
  ].map(([kpiId, texto]) => ({ kpiId, fecha: '2026-08-01', autor: 'Traslado desde Excel', texto }))
};

/* ===========================================================================
   §1  SESIÓN (OAuth) Y API DE GOOGLE SHEETS
   =========================================================================== */

const sesion = {
  clienteToken: null,
  token: null,
  expira: 0
};

function leerConfig() {
  try { return JSON.parse(localStorage.getItem(CLAVE_CONFIG)) || {}; }
  catch { return {}; }
}
function guardarConfig(cfg) {
  localStorage.setItem(CLAVE_CONFIG, JSON.stringify({ ...leerConfig(), ...cfg }));
}

function cargarScript(src) {
  return new Promise((ok, falla) => {
    const s = document.createElement('script');
    s.src = src; s.async = true; s.defer = true;
    s.onload = ok;
    s.onerror = () => falla(new Error('No se pudo cargar ' + src));
    document.head.appendChild(s);
  });
}

/* Síncrona a propósito: se llama desde el clic del usuario y cualquier `await`
   previo consumiría la activación del gesto, con lo que el navegador
   bloquearía la ventana emergente de Google. El script GIS se precarga en
   §6 antes de mostrar nada. */
function prepararSesion(clientId) {
  if (!window.google?.accounts?.oauth2) throw new Error('No se pudo cargar el conector de Google. Revisa la conexión.');
  sesion.clienteToken = google.accounts.oauth2.initTokenClient({
    client_id: clientId,
    scope: ALCANCE,
    callback: () => {}   // se reemplaza en cada petición
  });
}

/* `interactivo` abre la ventana de autorización; hay que llamarlo desde un
   clic. Sin él usa `prompt: ''`, que renueva en silencio una autorización ya
   concedida — pero si no hay nada que renovar no muestra nada NI llama a
   ningún callback, de ahí el guardia de tiempo. */
function pedirToken(interactivo = false) {
  return new Promise((ok, falla) => {
    if (!sesion.clienteToken) return falla(new Error('Sesión no inicializada'));

    let listo = false;
    const cerrar = (fn, arg) => { if (listo) return; listo = true; clearTimeout(guardia); fn(arg); };
    const guardia = setTimeout(() => cerrar(falla, new Error(
      interactivo
        ? 'Google no respondió. Lo más probable es que el navegador bloqueara la ventana emergente: permítela para este sitio y vuelve a intentarlo.'
        : 'sin sesión activa')),
      interactivo ? 90_000 : 8_000);

    sesion.clienteToken.callback = (resp) => {
      if (resp.error) return cerrar(falla, new Error(resp.error_description || resp.error));
      sesion.token = resp.access_token;
      sesion.expira = Date.now() + (Number(resp.expires_in) || 3600) * 1000;
      cerrar(ok, sesion.token);
    };
    sesion.clienteToken.error_callback = (err) =>
      cerrar(falla, new Error(err?.message || err?.type || 'No se pudo autorizar'));

    sesion.clienteToken.requestAccessToken({ prompt: interactivo ? 'consent' : '' });
  });
}

async function tokenVigente() {
  if (sesion.token && Date.now() < sesion.expira - 60_000) return sesion.token;
  return pedirToken(false);
}

async function api(ruta, opciones = {}, reintento = false) {
  const token = await tokenVigente();
  const resp = await fetch('https://sheets.googleapis.com/v4/spreadsheets' + ruta, {
    ...opciones,
    headers: {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json',
      ...(opciones.headers || {})
    }
  });
  if (resp.status === 401 && !reintento) {   // token vencido a mitad de camino
    sesion.token = null;
    await pedirToken();
    return api(ruta, opciones, true);
  }
  if (!resp.ok) {
    let detalle = '';
    try { detalle = (await resp.json()).error?.message || ''; } catch { /* respuesta sin cuerpo */ }
    throw new Error(detalle || `La API respondió ${resp.status}`);
  }
  return resp.json();
}

function leerHoja(sheetId) {
  const q = RANGOS.map(r => 'ranges=' + encodeURIComponent(r)).join('&');
  return api(`/${sheetId}/values:batchGet?${q}&valueRenderOption=UNFORMATTED_VALUE`);
}

function escribirHoja(sheetId, data) {
  return api(`/${sheetId}/values:batchUpdate`, {
    method: 'POST',
    body: JSON.stringify({ valueInputOption: 'RAW', data })
  });
}

function limpiarRangos(sheetId, ranges) {
  return api(`/${sheetId}/values:batchClear`, {
    method: 'POST',
    body: JSON.stringify({ ranges })
  });
}

async function crearHoja() {
  const creada = await api('', {
    method: 'POST',
    body: JSON.stringify({
      properties: { title: 'Tablero RENOBO 2026-2027', locale: 'es_CO' },
      sheets: ['KPIs', 'Serie', 'Hitos', 'Bitacora', 'Config']
        .map(title => ({ properties: { title } }))
    })
  });
  const datos = {
    config: { ...SEMILLA.config },
    kpis: SEMILLA.kpis.map(k => ({ ...k })),
    serie: Object.fromEntries(SEMILLA.kpis.map(k => [k.id, { '2026-08': k.avance }])),
    hitos: SEMILLA.hitos.map(h => ({ ...h })),
    bitacora: SEMILLA.bitacora.map(b => ({ ...b }))
  };
  await escribirHoja(creada.spreadsheetId, serializar(datos));
  return creada.spreadsheetId;
}

/* ===========================================================================
   §2  MODELO: PARSEO, EDICIÓN Y GUARDADO
   =========================================================================== */

const COLS_KPI = ['id', 'pilar', 'nombre', 'unidad', 'avance', 'meta2026', 'meta2027', 'metaFinal', 'responsable', 'avanceDesdeHitos'];

const numero = v => (v === '' || v === null || v === undefined) ? null : (Number.isFinite(Number(v)) ? Number(v) : null);
const booleano = v => v === true || String(v).toUpperCase() === 'TRUE' || String(v).toUpperCase() === 'VERDADERO';

/* El visor escribe los meses como texto («2026-08»), pero si alguien los teclea
   en la hoja, Sheets los convierte en fecha y los devuelve como número de serie.
   Esto los normaliza en ambos casos. */
function mesNormal(v) {
  if (typeof v === 'number' && v > 20000) {
    const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
  }
  return String(v ?? '').slice(0, 7);
}

function parsear(respuesta) {
  const porRango = {};
  for (const vr of respuesta.valueRanges || []) {
    const hoja = String(vr.range || '').split('!')[0].replace(/^'|'$/g, '');
    porRango[hoja] = vr.values || [];
  }

  const filasKpi = (porRango.KPIs || []).slice(1);
  const kpis = filasKpi.filter(f => f && f[0]).map(f => {
    const k = {};
    COLS_KPI.forEach((c, i) => { k[c] = f[i] ?? ''; });
    ['avance', 'meta2026', 'meta2027', 'metaFinal'].forEach(c => { k[c] = numero(k[c]); });
    k.avanceDesdeHitos = booleano(k.avanceDesdeHitos);
    return k;
  });

  const filasSerie = porRango.Serie || [];
  const cabecera = (filasSerie[0] || []).map((v, i) => i === 0 ? 'kpiId' : mesNormal(v));
  const serie = {};
  for (const fila of filasSerie.slice(1)) {
    if (!fila || !fila[0]) continue;
    const punto = {};
    for (let i = 1; i < cabecera.length; i++) {
      const v = numero(fila[i]);
      if (v !== null) punto[cabecera[i]] = v;
    }
    serie[fila[0]] = punto;
  }

  const hitos = (porRango.Hitos || []).slice(1)
    .filter(f => f && f[0])
    .map(f => ({ kpiId: f[0], nombre: f[1] ?? '', hecho: booleano(f[2]), fecha: f[3] ?? '' }));

  const bitacora = (porRango.Bitacora || []).slice(1)
    .filter(f => f && f[0])
    .map(f => ({ kpiId: f[0], fecha: f[1] ?? '', autor: f[2] ?? '', texto: f[3] ?? '' }));

  const config = { ...SEMILLA.config };
  for (const f of (porRango.Config || []).slice(1)) {
    if (!f || !f[0]) continue;
    const v = numero(f[1]);
    config[f[0]] = (f[0] === 'umbralAtencion' || f[0] === 'revision') ? (v ?? config[f[0]]) : f[1];
  }
  config.mesCorte = mesNormal(config.mesCorte) || SEMILLA.config.mesCorte;
  if (!MESES.includes(config.mesCorte)) config.mesCorte = SEMILLA.config.mesCorte;

  return { config, kpis, serie, hitos, bitacora };
}

function serializar(d) {
  const filasKpi = [COLS_KPI, ...d.kpis.map(k => COLS_KPI.map(c => {
    const v = k[c];
    if (c === 'avanceDesdeHitos') return v ? 'TRUE' : 'FALSE';
    return v === null || v === undefined ? '' : v;
  }))];

  const filasSerie = [['kpiId', ...MESES], ...d.kpis.map(k => {
    const punto = d.serie[k.id] || {};
    return [k.id, ...MESES.map(m => (punto[m] ?? ''))];
  })];

  const filasHitos = [['kpiId', 'nombre', 'hecho', 'fecha'],
    ...d.hitos.map(h => [h.kpiId, h.nombre, h.hecho ? 'TRUE' : 'FALSE', h.fecha || ''])];

  const filasBit = [['kpiId', 'fecha', 'autor', 'texto'],
    ...d.bitacora.map(b => [b.kpiId, b.fecha, b.autor, b.texto])];

  const filasCfg = [['clave', 'valor'],
    ...Object.entries(d.config).map(([k, v]) => [k, v ?? ''])];

  return [
    { range: 'KPIs!A1',     values: filasKpi },
    { range: 'Serie!A1',    values: filasSerie },
    { range: 'Hitos!A1',    values: filasHitos },
    { range: 'Bitacora!A1', values: filasBit },
    { range: 'Config!A1',   values: filasCfg }
  ];
}

/* --- Estado de la aplicación --------------------------------------------- */
const app = {
  sheetId: '',
  datos: null,
  cambios: [],          // etiquetas de los cambios sin guardar
  vista: 'resumen',
  kpiSel: null,
  filtroEstado: null,
  orden: { col: 'id', asc: true },
  guardando: false,
  ultimoGuardado: null
};

/* `parcial` repinta solo la cabecera. Se usa al editar en la tabla: un
   re-render completo le quitaría el foco al campo siguiente mientras tabula. */
function registrar(etiqueta, parcial = false) {
  app.cambios.push(etiqueta);
  try {
    localStorage.setItem(CLAVE_PENDIENTE, JSON.stringify({
      sheetId: app.sheetId, datos: app.datos, cambios: app.cambios, ts: Date.now()
    }));
  } catch { /* almacenamiento lleno o bloqueado: el aviso de cierre sigue activo */ }
  if (parcial) document.getElementById('cabecera').innerHTML = vistaCabecera();
  else pintar();
}

function limpiarPendiente() {
  app.cambios = [];
  localStorage.removeItem(CLAVE_PENDIENTE);
}

/* Los KPI marcados como `avanceDesdeHitos` derivan su avance de la lista. */
function recalcularDesdeHitos(kpiId) {
  const kpi = app.datos.kpis.find(k => k.id === kpiId);
  if (!kpi || !kpi.avanceDesdeHitos) return;
  kpi.avance = app.datos.hitos.filter(h => h.kpiId === kpiId && h.hecho).length;
}

async function guardar() {
  if (app.guardando || !app.cambios.length) return;
  app.guardando = true; pintar();
  try {
    // Antes de escribir, comprobar que la hoja no cambió por fuera.
    const actual = parsear(await leerHoja(app.sheetId));
    if (Number(actual.config.revision || 0) !== Number(app.datos.config.revision || 0)) {
      const seguir = confirm(
        'La hoja fue modificada por fuera del visor desde que abriste el tablero.\n\n' +
        'Si continúas, tus cambios reemplazarán los de la hoja. ' +
        'Cancela si prefieres recargar y perder los tuyos.'
      );
      if (!seguir) { app.guardando = false; pintar(); return; }
    }
    app.datos.config.revision = Number(actual.config.revision || 0) + 1;

    // Se escribe primero y solo después se borran las filas sobrantes: si algo
    // falla a mitad, nunca queda la hoja vacía.
    await escribirHoja(app.sheetId, serializar(app.datos));
    await limpiarRangos(app.sheetId, [
      `KPIs!A${app.datos.kpis.length + 2}:J`,
      `Serie!A${app.datos.kpis.length + 2}:S`,
      `Hitos!A${app.datos.hitos.length + 2}:D`,
      `Bitacora!A${app.datos.bitacora.length + 2}:D`
    ]);

    limpiarPendiente();
    app.ultimoGuardado = new Date();
  } catch (e) {
    alert('No se pudo guardar:\n\n' + e.message + '\n\nTus cambios siguen aquí; puedes reintentar.');
  } finally {
    app.guardando = false; pintar();
  }
}

async function recargar() {
  if (app.cambios.length && !confirm('Tienes cambios sin guardar. ¿Descartarlos y traer lo que hay en la hoja?')) return;
  try {
    app.datos = parsear(await leerHoja(app.sheetId));
    limpiarPendiente();
    pintar();
  } catch (e) {
    alert('No se pudo leer la hoja:\n\n' + e.message);
  }
}

/* Cierre de mes: congela el avance de cada KPI en la columna del mes de corte
   y adelanta el corte. Reemplaza el llenado manual del histórico del Excel. */
function cerrarMes() {
  const mes = app.datos.config.mesCorte;
  const i = MESES.indexOf(mes);
  if (i < 0) return alert('El mes de corte no está dentro del horizonte del tablero.');
  if (!confirm(`Se guardará el avance actual de los ${app.datos.kpis.length} KPIs en ${nombreMes(mes)} y el corte pasará a ${nombreMes(MESES[i + 1] || mes)}.\n\n¿Continuar?`)) return;

  for (const k of app.datos.kpis) {
    app.datos.serie[k.id] = app.datos.serie[k.id] || {};
    app.datos.serie[k.id][mes] = k.avance ?? 0;
  }
  if (i + 1 < MESES.length) app.datos.config.mesCorte = MESES[i + 1];
  registrar(`Cierre de ${nombreMes(mes)}`);
}

/* ===========================================================================
   §3  CÁLCULOS
   =========================================================================== */

const NOMBRES_MES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
const nombreMes = m => { const [a, s] = String(m).split('-'); return `${NOMBRES_MES[Number(s) - 1]}-${String(a).slice(2)}`; };
const anioDe = m => Number(String(m).split('-')[0]);

const fmt = (v, dec = 0) => v === null || v === undefined || !Number.isFinite(v)
  ? '—'
  : v.toLocaleString('es-CO', { minimumFractionDigits: dec, maximumFractionDigits: dec });

function fmtValor(v, unidad) {
  if (v === null || v === undefined || !Number.isFinite(v)) return '—';
  switch (unidad) {
    // `dec` tiene que ser entero: %1 devuelve la parte fraccionaria, no un 1.
    case 'porcentaje': return fmt(v * 100, (v * 100) % 1 === 0 ? 0 : 1) + ' %';
    case 'moneda':     return '$ ' + fmt(v) + ' M';
    case 'm2':         return fmt(v) + ' m²';
    default:           return fmt(v);
  }
}
const fmtPct = v => !Number.isFinite(v) ? '—'
  : fmt(v * 100, (v * 100) % 1 === 0 ? 0 : (Math.abs(v * 100) < 10 ? 1 : 0)) + ' %';

/* Meta vigente según el año del mes de corte. En 2027 y adelante se usa
   meta2027; si la gerencia aún no la definió, se cae a la meta final. */
function metaAnio(kpi, mesCorte) {
  return anioDe(mesCorte) <= 2026 ? kpi.meta2026 : (kpi.meta2027 ?? kpi.metaFinal);
}

/* Replica G6:G20 del Excel: avance / meta final, acotado a [0 ; 1,2]. */
function progresoFinal(kpi) {
  if (!kpi.metaFinal) return null;
  return Math.max(0, Math.min(1.2, (kpi.avance ?? 0) / kpi.metaFinal));
}

/* Replica H6:H20 del Excel, con la meta del año en curso. */
function estado(kpi, cfg) {
  const meta = metaAnio(kpi, cfg.mesCorte);
  if (kpi.avance === null || meta === null || meta === undefined) return 'sindato';
  if (kpi.avance >= meta) return 'meta';
  if (kpi.avance >= meta * (cfg.umbralAtencion ?? 0.4)) return 'atencion';
  return 'grave';
}

const ETIQUETA_ESTADO = { meta: 'En meta', atencion: 'Atención', grave: 'Grave', sindato: 'Sin dato' };
const ICONO_ESTADO    = { meta: '✓', atencion: '!', grave: '×', sindato: '·' };

/* El color nunca va solo: la pastilla lleva icono y etiqueta de texto. */
const pastilla = (est, extra = '') =>
  `<span class="pastilla ${extra} est-${est}"><i aria-hidden="true">${ICONO_ESTADO[est]}</i>${ETIQUETA_ESTADO[est]}</span>`;

function puntosSerie(kpiId, hastaMes) {
  const punto = app.datos.serie[kpiId] || {};
  const tope = MESES.indexOf(hastaMes);
  return MESES
    .filter((m, i) => (tope < 0 || i <= tope) && Number.isFinite(punto[m]))
    .map(m => ({ mes: m, valor: punto[m] }));
}

function deltaMes(kpi, cfg) {
  const p = puntosSerie(kpi.id, cfg.mesCorte);
  const previo = p.filter(x => x.mes !== cfg.mesCorte).pop();
  if (!previo) return null;
  return (kpi.avance ?? 0) - previo.valor;
}

/* Cuánto hay que avanzar cada mes para llegar a la meta del año. */
function ritmoRequerido(kpi, cfg) {
  const meta = metaAnio(kpi, cfg.mesCorte);
  if (meta === null || meta === undefined) return null;
  const anio = anioDe(cfg.mesCorte);
  const restantes = MESES.filter(m => anioDe(m) === anio && m >= cfg.mesCorte).length;
  if (restantes <= 0) return null;
  return Math.max(0, meta - (kpi.avance ?? 0)) / restantes;
}

/* Extrapolación lineal de los últimos tres puntos al cierre del año. */
function proyeccion(kpi, cfg) {
  const p = puntosSerie(kpi.id, cfg.mesCorte).slice(-3);
  if (p.length < 2) return null;
  const primero = p[0], ultimo = p[p.length - 1];
  const tramos = MESES.indexOf(ultimo.mes) - MESES.indexOf(primero.mes);
  if (tramos <= 0) return null;
  const pendiente = (ultimo.valor - primero.valor) / tramos;
  const anio = anioDe(cfg.mesCorte);
  const finAnio = MESES.filter(m => anioDe(m) === anio).pop();
  return ultimo.valor + pendiente * (MESES.indexOf(finAnio) - MESES.indexOf(ultimo.mes));
}

/* Los dos indicadores globales. El Excel solo tenía el segundo (12,8 %). */
function indicadoresGlobales() {
  const { kpis, config } = app.datos;
  const vsAnio = kpis.map(k => {
    const meta = metaAnio(k, config.mesCorte);
    return meta ? Math.max(0, Math.min(1.2, (k.avance ?? 0) / meta)) : null;
  }).filter(Number.isFinite);
  const vsFinal = kpis.map(progresoFinal).filter(Number.isFinite);
  const prom = a => a.length ? a.reduce((s, x) => s + x, 0) / a.length : null;
  return { vsAnio: prom(vsAnio), vsFinal: prom(vsFinal) };
}

const conteoEstados = () => app.datos.kpis.reduce((acc, k) => {
  const e = estado(k, app.datos.config);
  acc[e] = (acc[e] || 0) + 1;
  return acc;
}, {});

/* ===========================================================================
   §4  GRÁFICOS (SVG dibujado a mano, sin librerías)
   =========================================================================== */

const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

/* Paso de eje en números redondos (1, 2, 5 × potencia de 10): los ticks del
   eje cargan los valores que no se etiquetan directamente, así que tienen que
   leerse de un vistazo. */
function pasoBonito(max, divisiones = 4) {
  const bruto = max / divisiones;
  const magnitud = Math.pow(10, Math.floor(Math.log10(bruto)));
  const norm = bruto / magnitud;
  return (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10) * magnitud;
}

/* Barra bullet: relleno = avance, marca = meta del año, extremo = meta final. */
function barraBullet(kpi, cfg, ancho = 220, alto = 14) {
  const meta = metaAnio(kpi, cfg.mesCorte);
  const tope = Math.max(kpi.metaFinal || 0, kpi.avance || 0, meta || 0) || 1;
  const x = v => Math.max(0, Math.min(ancho, (v / tope) * ancho));
  const est = estado(kpi, cfg);
  const largo = x(kpi.avance ?? 0);
  const r = Math.min(4, largo / 2);

  const marca = Number.isFinite(meta) && meta > 0 && meta < tope
    ? `<rect x="${(x(meta) - 1).toFixed(1)}" y="-3" width="2" height="${alto + 6}" rx="1" class="g-marca"/>` : '';

  return `<svg class="g-bullet" viewBox="0 0 ${ancho} ${alto}" width="${ancho}" height="${alto}"
            role="img" aria-label="${esc(fmtValor(kpi.avance, kpi.unidad))} de ${esc(fmtValor(kpi.metaFinal, kpi.unidad))}">
    <rect x="0" y="0" width="${ancho}" height="${alto}" rx="2" class="g-pista est-${est}"/>
    ${largo > 0 ? `<path d="M0,0 H${(largo - r).toFixed(1)} a${r},${r} 0 0 1 ${r},${r} V${alto - r} a${r},${r} 0 0 1 -${r},${r} H0 Z" class="g-relleno est-${est}"/>` : ''}
    ${marca}
  </svg>`;
}

function sparkline(puntos, ancho = 88, alto = 26) {
  if (puntos.length < 2) return `<span class="g-vacio" title="Aún no hay serie histórica">—</span>`;
  const vals = puntos.map(p => p.valor);
  const min = Math.min(...vals), max = Math.max(...vals);
  const rango = (max - min) || 1;
  const px = i => 3 + (i / (puntos.length - 1)) * (ancho - 6);
  const py = v => (alto - 4) - ((v - min) / rango) * (alto - 8);
  const d = puntos.map((p, i) => `${i ? 'L' : 'M'}${px(i).toFixed(1)},${py(p.valor).toFixed(1)}`).join(' ');
  const ux = px(puntos.length - 1), uy = py(vals[vals.length - 1]);
  return `<svg class="g-spark" viewBox="0 0 ${ancho} ${alto}" width="${ancho}" height="${alto}" role="img"
            aria-label="Evolución: ${puntos.map(p => `${nombreMes(p.mes)} ${fmt(p.valor)}`).join(', ')}">
    <path d="${d}" class="g-linea"/>
    <circle cx="${ux.toFixed(1)}" cy="${uy.toFixed(1)}" r="4" class="g-punta"/>
  </svg>`;
}

/* Evolución mensual de un KPI, con la meta del año y la ruta que hace falta. */
function graficoEvolucion(kpi, cfg, ancho = 640, alto = 260) {
  const m = { arr: 18, der: 96, aba: 34, izq: 62 };
  const w = ancho - m.izq - m.der, h = alto - m.arr - m.aba;
  const puntos = puntosSerie(kpi.id, cfg.mesCorte);
  const anio = anioDe(cfg.mesCorte);
  const mesesAnio = MESES.filter(x => anioDe(x) === anio);
  const meta = metaAnio(kpi, cfg.mesCorte);

  if (!puntos.length) {
    return `<p class="aviso">Todavía no hay serie histórica para este KPI.
      Se construye sola: cada vez que uses <strong>Cerrar mes</strong> se guarda un punto.</p>`;
  }

  const ejeMeses = MESES.slice(0, Math.max(MESES.indexOf(mesesAnio[mesesAnio.length - 1]) + 1, MESES.indexOf(cfg.mesCorte) + 1));
  const cima = Math.max(...puntos.map(p => p.valor), meta || 0, kpi.avance || 0) || 1;
  const paso = pasoBonito(cima);
  const maxY = Math.ceil(cima / paso) * paso;
  const px = mes => m.izq + (ejeMeses.indexOf(mes) / Math.max(1, ejeMeses.length - 1)) * w;
  const py = v => m.arr + h - (v / maxY) * h;

  const ticks = [];
  for (let t = 0; t <= maxY + paso / 2; t += paso) ticks.push(t);
  const rejilla = ticks.map(t => `
    <line x1="${m.izq}" y1="${py(t).toFixed(1)}" x2="${m.izq + w}" y2="${py(t).toFixed(1)}" class="g-rejilla"/>
    <text x="${m.izq - 10}" y="${(py(t) + 4).toFixed(1)}" class="g-tick g-tick-y">${esc(fmtValor(t, kpi.unidad))}</text>`).join('');

  const cadaCuantos = Math.ceil(ejeMeses.length / 9);
  const ejeX = ejeMeses.map((mes, i) => i % cadaCuantos ? '' :
    `<text x="${px(mes).toFixed(1)}" y="${alto - 12}" class="g-tick g-tick-x">${nombreMes(mes)}</text>`).join('');

  const d = puntos.map((p, i) => `${i ? 'L' : 'M'}${px(p.mes).toFixed(1)},${py(p.valor).toFixed(1)}`).join(' ');
  const marcadores = puntos.map(p =>
    `<circle cx="${px(p.mes).toFixed(1)}" cy="${py(p.valor).toFixed(1)}" r="4" class="g-punto">
       <title>${nombreMes(p.mes)}: ${esc(fmtValor(p.valor, kpi.unidad))}</title></circle>`).join('');

  // Ruta a la meta: del último punto real al cierre del año.
  const finAnio = mesesAnio[mesesAnio.length - 1];
  const ult = puntos[puntos.length - 1];
  const ruta = Number.isFinite(meta) && ejeMeses.includes(finAnio)
    ? `<path d="M${px(ult.mes).toFixed(1)},${py(ult.valor).toFixed(1)} L${px(finAnio).toFixed(1)},${py(meta).toFixed(1)}" class="g-ruta"/>`
    : '';
  const lineaMeta = Number.isFinite(meta) ? `
    <line x1="${m.izq}" y1="${py(meta).toFixed(1)}" x2="${m.izq + w}" y2="${py(meta).toFixed(1)}" class="g-meta"/>
    <text x="${m.izq + w + 8}" y="${(py(meta) + 4).toFixed(1)}" class="g-etiq">Meta ${anio}</text>` : '';

  return `
  <svg class="g-evolucion" viewBox="0 0 ${ancho} ${alto}" role="img"
       aria-label="Evolución mensual de ${esc(kpi.nombre)}">
    ${rejilla}${ejeX}${lineaMeta}${ruta}
    <path d="${d}" class="g-linea-gruesa"/>${marcadores}
    <text x="${(px(ult.mes) + 8).toFixed(1)}" y="${(py(ult.valor) - 10).toFixed(1)}" class="g-etiq g-etiq-fuerte">${esc(fmtValor(ult.valor, kpi.unidad))}</text>
  </svg>
  <div class="leyenda">
    <span><i class="clave clave-avance"></i>Avance real</span>
    <span><i class="clave clave-ruta"></i>Ruta necesaria hasta la meta</span>
    <span><i class="clave clave-meta"></i>Meta ${anio}</span>
  </div>`;
}

/* ===========================================================================
   §5  VISTAS
   =========================================================================== */

function pintar() {
  if (!app.datos) return;
  document.getElementById('cabecera').innerHTML = vistaCabecera();
  document.getElementById('contenido').innerHTML =
      app.vista === 'tabla'   ? vistaTabla()
    : app.vista === 'detalle' ? vistaDetalle()
    : app.vista === 'comite'  ? vistaComite()
    :                           vistaResumen();
  document.body.dataset.vista = app.vista;
}

function vistaCabecera() {
  const { config } = app.datos;
  const pendientes = app.cambios.length;
  const estadoGuardado = app.guardando
    ? '<span class="chip chip-trabajando">Guardando…</span>'
    : pendientes
      ? `<span class="chip chip-pendiente">${pendientes} ${pendientes === 1 ? 'cambio' : 'cambios'} sin guardar</span>`
      : `<span class="chip chip-ok">Guardado${app.ultimoGuardado ? ' ' + app.ultimoGuardado.toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit' }) : ''}</span>`;

  const pestanas = [['resumen', 'Resumen'], ['tabla', 'Tabla'], ['comite', 'Comité']]
    .map(([id, txt]) => `<button class="pest ${app.vista === id ? 'activa' : ''}" data-vista="${id}">${txt}</button>`).join('');

  return `
    <div class="barra">
      <div class="titulo">
        <h1>${esc(config.titulo || 'Tablero de Gerencia')}</h1>
        <p class="corte">Mes de corte · <strong>${nombreMes(config.mesCorte)}</strong></p>
      </div>
      <div class="acciones">
        ${estadoGuardado}
        <button class="btn btn-primario" data-accion="guardar" ${pendientes && !app.guardando ? '' : 'disabled'}>Guardar</button>
        <button class="btn" data-accion="recargar"
          title="Descarta lo que tengas en pantalla y vuelve a traer los datos de la hoja">Recargar de la hoja</button>
        <button class="btn" data-accion="cerrar-mes">Cerrar mes</button>
        <button class="btn btn-sutil" data-accion="hoja" title="Abrir la hoja de cálculo">Hoja</button>
      </div>
    </div>
    <nav class="pestanas">${pestanas}</nav>`;
}

function fichasEstado() {
  const c = conteoEstados();
  return ['meta', 'atencion', 'grave'].map(e => `
    <button class="ficha est-${e} ${app.filtroEstado === e ? 'activa' : ''}" data-filtro="${e}">
      <span class="ficha-icono" aria-hidden="true">${ICONO_ESTADO[e]}</span>
      <span class="ficha-num">${c[e] || 0}</span>
      <span class="ficha-txt">${ETIQUETA_ESTADO[e]}</span>
    </button>`).join('');
}

function vistaResumen() {
  const { config, kpis } = app.datos;
  const g = indicadoresGlobales();
  const visibles = app.filtroEstado ? kpis.filter(k => estado(k, config) === app.filtroEstado) : kpis;

  const bloques = PILARES.map(p => {
    const delPilar = visibles.filter(k => k.pilar === p.id);
    if (!delPilar.length) return '';
    return `
      <section class="pilar">
        <h2>${esc(p.nombre)} <span class="cuenta">${delPilar.length}</span></h2>
        <div class="lista">${delPilar.map(k => tarjetaKpi(k, config)).join('')}</div>
      </section>`;
  }).join('');

  return `
    <div class="tablero">
      <div class="hero">
        <p class="hero-etiq">Avance promedio hacia la meta ${anioDe(config.mesCorte)}</p>
        <p class="hero-num">${fmtPct(g.vsAnio)}</p>
        <p class="hero-pie">Hacia la meta final del periodo: <strong>${fmtPct(g.vsFinal)}</strong></p>
      </div>
      <div class="fichas">${fichasEstado()}</div>
    </div>
    ${app.filtroEstado ? `<p class="filtro-activo">Mostrando solo <strong>${ETIQUETA_ESTADO[app.filtroEstado]}</strong>. <button class="enlace" data-filtro="">Ver todos</button></p>` : ''}
    ${bloques || '<p class="aviso">Ningún KPI en ese estado.</p>'}`;
}

function tarjetaKpi(k, cfg) {
  const est = estado(k, cfg);
  const d = deltaMes(k, cfg);
  const meta = metaAnio(k, cfg.mesCorte);
  const flecha = d === null || d === 0 ? '' : (d > 0 ? '↑' : '↓');
  return `
    <article class="kpi" data-kpi="${k.id}" tabindex="0" role="button">
      <div class="kpi-enc">
        <span class="kpi-id">${esc(k.id)}</span>
        <h3>${esc(k.nombre)}</h3>
        ${pastilla(est)}
      </div>
      <div class="kpi-cifras">
        <span class="cifra">${esc(fmtValor(k.avance, k.unidad))}</span>
        <span class="cifra-meta">meta ${anioDe(cfg.mesCorte)} · ${esc(fmtValor(meta, k.unidad))}</span>
      </div>
      <div class="kpi-graf">
        ${barraBullet(k, cfg)}
        ${sparkline(puntosSerie(k.id, cfg.mesCorte))}
      </div>
      <div class="kpi-pie">
        <span class="resp">${esc(k.responsable)}</span>
        ${d === null ? '<span class="delta neutro">sin mes previo</span>'
          : `<span class="delta ${d > 0 ? 'sube' : d < 0 ? 'baja' : 'neutro'}">${flecha} ${esc(fmtValor(Math.abs(d), k.unidad))} vs. mes anterior</span>`}
      </div>
    </article>`;
}

function vistaTabla() {
  const { config, kpis } = app.datos;
  const cols = [
    ['id', 'ID'], ['pilar', 'Pilar'], ['nombre', 'KPI'], ['avance', 'Avance'],
    ['metaAnio', `Meta ${anioDe(config.mesCorte)}`], ['metaFinal', 'Meta final'],
    ['progreso', 'Progreso'], ['estado', 'Estado'], ['responsable', 'Responsable']
  ];
  const valor = (k, c) =>
      c === 'metaAnio'  ? metaAnio(k, config.mesCorte)
    : c === 'progreso'  ? progresoFinal(k)
    : c === 'estado'    ? estado(k, config)
    : c === 'pilar'     ? (PILARES.find(p => p.id === k.pilar)?.nombre || k.pilar)
    : k[c];

  const filas = [...kpis].sort((a, b) => {
    const va = valor(a, app.orden.col), vb = valor(b, app.orden.col);
    const n = typeof va === 'number' && typeof vb === 'number' ? va - vb : String(va).localeCompare(String(vb), 'es');
    return app.orden.asc ? n : -n;
  }).map(k => {
    const est = estado(k, config);
    const esPct = k.unidad === 'porcentaje';
    const campo = (c, v) => `<span class="campo"><input class="celda-edit" type="number" step="any"
        data-kpi="${k.id}" data-campo="${c}" value="${v === null || v === undefined ? '' : (esPct ? +(v * 100).toFixed(4) : v)}"
        aria-label="${esc(c)} de ${esc(k.nombre)}">${esPct ? '<span class="sufijo">%</span>' : ''}</span>`;
    const anio = anioDe(config.mesCorte);
    return `<tr data-fila="${k.id}">
      <td class="col-id">${esc(k.id)}</td>
      <td class="col-pilar">${esc(PILARES.find(p => p.id === k.pilar)?.nombre || k.pilar)}</td>
      <td class="col-nombre"><button class="enlace" data-kpi="${k.id}">${esc(k.nombre)}</button>
          <span class="unidad">${k.unidad === 'moneda' ? 'millones COP' : k.unidad === 'm2' ? 'm²' : k.unidad}</span></td>
      <td class="col-num editable">${k.avanceDesdeHitos
          ? `<span class="derivado" title="Se calcula con los hitos marcados">${esc(fmtValor(k.avance, k.unidad))} <i>auto</i></span>`
          : campo('avance', k.avance)}</td>
      <td class="col-num editable">${campo(anio <= 2026 ? 'meta2026' : 'meta2027', metaAnio(k, config.mesCorte))}</td>
      <td class="col-num editable">${campo('metaFinal', k.metaFinal)}</td>
      <td class="col-num celda-progreso">${fmtPct(progresoFinal(k))}</td>
      <td class="celda-estado">${pastilla(est)}</td>
      <td class="editable"><select class="celda-edit" data-kpi="${k.id}" data-campo="responsable"
          title="${esc(k.responsable)}" aria-label="Responsable de ${esc(k.nombre)}">
        ${RESPONSABLES.map(r => `<option ${r === k.responsable ? 'selected' : ''}>${esc(r)}</option>`).join('')}
      </select></td>
    </tr>`;
  }).join('');

  return `
    <p class="aviso-sutil">Las celdas en amarillo son editables. Los cambios se acumulan arriba hasta que pulses <strong>Guardar</strong>.</p>
    <div class="tabla-env">
      <table class="tabla">
        <thead><tr>${cols.map(([c, t]) =>
          `<th><button class="orden ${app.orden.col === c ? 'activa' : ''}" data-orden="${c}">${esc(t)}${app.orden.col === c ? (app.orden.asc ? ' ▴' : ' ▾') : ''}</button></th>`).join('')}</tr></thead>
        <tbody>${filas}</tbody>
      </table>
    </div>`;
}

/* Actualiza en sitio las dos celdas derivadas de una fila de la tabla, sin
   volver a dibujarla entera. */
function refrescarFila(kpi) {
  const fila = document.querySelector(`.tabla tbody tr[data-fila="${kpi.id}"]`);
  if (!fila) return;
  const prog = fila.querySelector('.celda-progreso');
  const est = fila.querySelector('.celda-estado');
  if (prog) prog.textContent = fmtPct(progresoFinal(kpi));
  if (est) est.innerHTML = pastilla(estado(kpi, app.datos.config));
}

function vistaDetalle() {
  const { config } = app.datos;
  const k = app.datos.kpis.find(x => x.id === app.kpiSel);
  if (!k) { app.vista = 'resumen'; return vistaResumen(); }

  const est = estado(k, config);
  const meta = metaAnio(k, config.mesCorte);
  const ritmo = ritmoRequerido(k, config);
  const proy = proyeccion(k, config);
  const hitos = app.datos.hitos.filter(h => h.kpiId === k.id);
  const notas = app.datos.bitacora.filter(b => b.kpiId === k.id).sort((a, b) => String(b.fecha).localeCompare(String(a.fecha)));

  return `
    <button class="enlace volver" data-vista="resumen">← Volver al resumen</button>
    <header class="detalle-enc">
      <div>
        <span class="kpi-id">${esc(k.id)} · ${esc(PILARES.find(p => p.id === k.pilar)?.nombre || '')}</span>
        <h2>${esc(k.nombre)}</h2>
        <p class="resp">Responsable: <strong>${esc(k.responsable)}</strong></p>
      </div>
      ${pastilla(est, 'grande')}
    </header>

    <div class="cifras-detalle">
      ${[['Avance', fmtValor(k.avance, k.unidad)],
         [`Meta ${anioDe(config.mesCorte)}`, fmtValor(meta, k.unidad)],
         ['Meta final', fmtValor(k.metaFinal, k.unidad)],
         ['Ritmo requerido', ritmo === null ? '—' : fmtValor(ritmo, k.unidad) + ' / mes'],
         ['Proyección a fin de año', proy === null ? 'sin serie suficiente' : fmtValor(proy, k.unidad)]
        ].map(([t, v]) => `<div class="cifra-caja"><span class="cifra-etiq">${esc(t)}</span><span class="cifra-val">${esc(v)}</span></div>`).join('')}
    </div>

    <section class="panel">
      <h3>Evolución</h3>
      ${graficoEvolucion(k, config)}
    </section>

    <div class="dos-columnas">
      <section class="panel">
        <h3>Hitos <span class="cuenta">${hitos.filter(h => h.hecho).length}/${hitos.length}</span></h3>
        ${k.avanceDesdeHitos ? '<p class="aviso-sutil">El avance de este KPI se calcula con los hitos marcados.</p>' : ''}
        ${hitos.length ? `<ul class="hitos">${hitos.map((h, i) => `
          <li><label>
            <input type="checkbox" data-hito="${esc(k.id)}|${i}" ${h.hecho ? 'checked' : ''}>
            <span class="${h.hecho ? 'hecho' : ''}">${esc(h.nombre)}</span>
            ${h.fecha ? `<time>${esc(h.fecha)}</time>` : ''}
          </label></li>`).join('')}</ul>`
          : '<p class="aviso">Este KPI aún no tiene hitos definidos.</p>'}
        <form class="alta" data-alta="hito">
          <input name="texto" placeholder="Añadir un hito…" required>
          <button class="btn btn-sutil">Añadir</button>
        </form>
      </section>

      <section class="panel">
        <h3>Bitácora</h3>
        ${notas.length ? `<ul class="bitacora">${notas.map(n => `
          <li><div class="nota-enc"><time>${esc(n.fecha)}</time><span>${esc(n.autor)}</span></div>
              <p>${esc(n.texto)}</p></li>`).join('')}</ul>`
          : '<p class="aviso">Sin registros todavía.</p>'}
        <form class="alta" data-alta="nota">
          <textarea name="texto" rows="2" placeholder="Registrar una decisión, un compromiso, un bloqueo…" required></textarea>
          <button class="btn btn-sutil">Registrar</button>
        </form>
      </section>
    </div>`;
}

function vistaComite() {
  const { config, kpis } = app.datos;
  const g = indicadoresGlobales();
  const c = conteoEstados();
  const filas = kpis.map(k => {
    const est = estado(k, config);
    return `<tr class="est-fila-${est}">
      <td>${esc(k.id)}</td>
      <td>${esc(k.nombre)}</td>
      <td class="col-num">${esc(fmtValor(k.avance, k.unidad))}</td>
      <td class="col-num">${esc(fmtValor(metaAnio(k, config.mesCorte), k.unidad))}</td>
      <td class="col-bar">${barraBullet(k, config, 160, 12)}</td>
      <td>${pastilla(est)}</td>
    </tr>`;
  }).join('');

  return `
    <div class="comite">
      <header class="comite-enc">
        <div>
          <h2>${esc(config.titulo)}</h2>
          <p>Corte a ${nombreMes(config.mesCorte)} · ${kpis.length} indicadores</p>
        </div>
        <div class="comite-cifras">
          <div><span>${fmtPct(g.vsAnio)}</span><small>hacia la meta ${anioDe(config.mesCorte)}</small></div>
          <div><span>${fmtPct(g.vsFinal)}</span><small>hacia la meta final</small></div>
          <div><span>${c.meta || 0} · ${c.atencion || 0} · ${c.grave || 0}</span><small>en meta · atención · grave</small></div>
        </div>
      </header>
      <table class="tabla tabla-comite">
        <thead><tr><th>ID</th><th>Indicador</th><th class="col-num">Avance</th><th class="col-num">Meta ${anioDe(config.mesCorte)}</th><th>Progreso</th><th>Estado</th></tr></thead>
        <tbody>${filas}</tbody>
      </table>
    </div>`;
}

/* ===========================================================================
   §6  ARRANQUE E INTERACCIÓN
   =========================================================================== */

function pantalla(html) {
  document.getElementById('portada').innerHTML = html;
  document.getElementById('portada').hidden = false;
  document.getElementById('app').hidden = true;
}
function mostrarApp() {
  document.getElementById('portada').hidden = true;
  document.getElementById('app').hidden = false;
  pintar();
}

function pantallaConfig(mensaje = '') {
  const cfg = leerConfig();
  pantalla(`
    <div class="portada-caja">
      <h1>Visor CF · Tablero de Gerencia</h1>
      <p class="portada-sub">Conecta el visor con tu hoja de cálculo en Google Drive. Se hace una sola vez.</p>
      ${mensaje ? `<p class="portada-error">${esc(mensaje)}</p>` : ''}
      <form id="form-config">
        <label>ID de cliente OAuth
          <input name="clientId" value="${esc(cfg.clientId || '')}" placeholder="…apps.googleusercontent.com" required>
          <small>Se crea en Google Cloud. Ver el README del proyecto.</small>
        </label>
        <label>ID de la hoja de cálculo
          <input name="sheetId" value="${esc(cfg.sheetId || '')}" placeholder="Déjalo vacío para crear una nueva">
          <small>Está en la URL de la hoja, entre <code>/d/</code> y <code>/edit</code>.</small>
        </label>
        <button class="btn btn-primario" type="submit">Conectar con Google</button>
      </form>
    </div>`);

  document.getElementById('form-config').addEventListener('submit', async (ev) => {
    ev.preventDefault();
    const f = new FormData(ev.target);
    const clientId = String(f.get('clientId')).trim();
    let sheetId = String(f.get('sheetId')).trim();
    // Tolerar que peguen la URL completa de la hoja.
    const m = sheetId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (m) sheetId = m[1];

    try {
      // El orden importa: `requestAccessToken` tiene que salir del clic, sin
      // ningún `await` por delante, o la ventana emergente queda bloqueada.
      prepararSesion(clientId);
      const autorizacion = pedirToken(true);
      pantalla(`<div class="portada-caja">
          <h1>Conectando con Google…</h1>
          <p class="portada-sub">Se abrió una ventana para que autorices el acceso a tu hoja.
          Si no la ves, revisa si el navegador bloqueó las ventanas emergentes de este sitio.</p>
        </div>`);
      await autorizacion;

      pantalla('<div class="portada-caja"><h1>Preparando el tablero…</h1><p class="portada-sub">Creando la hoja y cargando los 15 KPIs.</p></div>');
      if (!sheetId) sheetId = await crearHoja();
      guardarConfig({ clientId, sheetId });
      await iniciar();
    } catch (e) {
      pantallaConfig(e.message);
    }
  });
}

/* La ventana de autorización de Google necesita un clic del usuario: al
   recargar la página no se puede pedir sola sin que el navegador la bloquee. */
function pantallaConectar(mensaje = '') {
  pantalla(`
    <div class="portada-caja">
      <h1>Visor CF · Tablero de Gerencia</h1>
      <p class="portada-sub">Tu sesión de Google expiró. Vuelve a entrar para seguir.</p>
      ${mensaje ? `<p class="portada-error">${esc(mensaje)}</p>` : ''}
      <button class="btn btn-primario" id="btn-conectar">Conectar con Google</button>
      <p><button class="enlace" id="btn-reconfig">Usar otra hoja o cambiar las credenciales</button></p>
    </div>`);
  document.getElementById('btn-conectar').addEventListener('click', async () => {
    const p = pedirToken(true);        // dentro del gesto, sin await previo
    try { await p; await iniciar(); }
    catch (e) { pantallaConectar(e.message); }
  });
  document.getElementById('btn-reconfig').addEventListener('click', () => pantallaConfig());
}

async function iniciar() {
  const cfg = leerConfig();
  if (!cfg.clientId || !cfg.sheetId) return pantallaConfig();

  try {
    if (!sesion.clienteToken) prepararSesion(cfg.clientId);
    if (!sesion.token) {
      // Renovación silenciosa; si no hay autorización viva, hace falta un clic.
      try { await pedirToken(false); }
      catch { return pantallaConectar(); }
    }

    app.sheetId = cfg.sheetId;
    app.datos = parsear(await leerHoja(app.sheetId));

    // ¿Quedaron cambios sin guardar de una sesión anterior?
    const pend = JSON.parse(localStorage.getItem(CLAVE_PENDIENTE) || 'null');
    if (pend && pend.sheetId === app.sheetId && pend.cambios?.length) {
      if (confirm(`La última vez quedaron ${pend.cambios.length} cambios sin guardar.\n\n¿Recuperarlos?`)) {
        app.datos = pend.datos;
        app.cambios = pend.cambios;
      } else {
        limpiarPendiente();
      }
    }
    mostrarApp();
  } catch (e) {
    pantallaConfig('No se pudo abrir la hoja: ' + e.message);
  }
}

/* --- Eventos -------------------------------------------------------------- */

document.addEventListener('click', (ev) => {
  const t = ev.target.closest('[data-accion],[data-vista],[data-kpi],[data-filtro],[data-orden]');
  if (!t || !app.datos) return;

  if (t.dataset.accion === 'guardar')   return guardar();
  if (t.dataset.accion === 'recargar')  return recargar();
  if (t.dataset.accion === 'cerrar-mes')return cerrarMes();
  if (t.dataset.accion === 'hoja')      return window.open(`https://docs.google.com/spreadsheets/d/${app.sheetId}/edit`, '_blank');

  if (t.dataset.vista)  { app.vista = t.dataset.vista; app.kpiSel = null; return pintar(); }
  if (t.dataset.filtro !== undefined && !t.dataset.kpi) { app.filtroEstado = t.dataset.filtro || null; return pintar(); }
  if (t.dataset.orden)  {
    app.orden = { col: t.dataset.orden, asc: app.orden.col === t.dataset.orden ? !app.orden.asc : true };
    return pintar();
  }
  if (t.dataset.kpi && !t.classList.contains('celda-edit')) {
    app.kpiSel = t.dataset.kpi; app.vista = 'detalle'; return pintar();
  }
});

document.addEventListener('keydown', (ev) => {
  if ((ev.ctrlKey || ev.metaKey) && ev.key.toLowerCase() === 's') { ev.preventDefault(); guardar(); }
  if (ev.key === 'Enter' && ev.target.matches('.kpi[data-kpi]')) {
    app.kpiSel = ev.target.dataset.kpi; app.vista = 'detalle'; pintar();
  }
});

/* Las ediciones se confirman al salir del campo, no al teclear: así el
   re-render no le quita el foco al gerente a mitad de un número. */
document.addEventListener('change', (ev) => {
  const el = ev.target;
  if (!app.datos) return;

  if (el.matches('.celda-edit')) {
    const kpi = app.datos.kpis.find(k => k.id === el.dataset.kpi);
    if (!kpi) return;
    const campo = el.dataset.campo;
    if (campo === 'responsable') {
      kpi.responsable = el.value;
      return registrar(`${kpi.id} · responsable`, true);
    }
    let v = el.value === '' ? null : Number(el.value);
    if (v !== null && !Number.isFinite(v)) { el.value = ''; return; }
    if (v !== null && kpi.unidad === 'porcentaje') v = v / 100;
    if (v !== null && v < 0) v = 0;
    if (v !== null && ['conteo', 'm2'].includes(kpi.unidad)) v = Math.round(v);
    kpi[campo] = v;
    refrescarFila(kpi);
    return registrar(`${kpi.id} · ${campo}`, true);
  }

  if (el.matches('[data-hito]')) {
    const [kpiId, idx] = el.dataset.hito.split('|');
    const propios = app.datos.hitos.filter(h => h.kpiId === kpiId);
    const h = propios[Number(idx)];
    if (!h) return;
    h.hecho = el.checked;
    h.fecha = el.checked ? new Date().toISOString().slice(0, 10) : '';
    recalcularDesdeHitos(kpiId);
    return registrar(`${kpiId} · hito «${h.nombre}»`);
  }
});

document.addEventListener('submit', (ev) => {
  const f = ev.target;
  if (!f.matches('[data-alta]')) return;
  ev.preventDefault();
  const texto = String(new FormData(f).get('texto')).trim();
  if (!texto || !app.kpiSel) return;

  if (f.dataset.alta === 'hito') {
    app.datos.hitos.push({ kpiId: app.kpiSel, nombre: texto, hecho: false, fecha: '' });
    registrar(`${app.kpiSel} · hito nuevo`);
  } else {
    app.datos.bitacora.push({
      kpiId: app.kpiSel,
      fecha: new Date().toISOString().slice(0, 10),
      autor: leerConfig().autor || 'Gerencia',
      texto
    });
    registrar(`${app.kpiSel} · nota`);
  }
});

window.addEventListener('beforeunload', (ev) => {
  if (app.cambios.length) { ev.preventDefault(); ev.returnValue = ''; }
});

/* El conector de Google se carga antes de mostrar nada, para que al pulsar
   «Conectar» no haya ningún `await` entre el clic y la ventana emergente. */
async function arrancar() {
  pantalla('<div class="portada-caja"><h1>Visor CF</h1><p class="portada-sub">Cargando…</p></div>');
  try {
    await cargarScript('https://accounts.google.com/gsi/client');
  } catch {
    return pantalla(`<div class="portada-caja">
      <h1>Sin conexión con Google</h1>
      <p class="portada-sub">No se pudo cargar el conector de Google. Revisa la conexión a internet
      y recarga la página. Si el problema persiste, puede que la red de la entidad esté bloqueando
      <code>accounts.google.com</code>.</p></div>`);
  }
  iniciar();
}

arrancar();
