/**
 * NUTRI-ROSI · Agenda online + recordatorios automáticos + dashboard
 * ---------------------------------------------------------------
 * Backend en Google Apps Script, usando esta misma Hoja de cálculo
 * como base de datos (pestañas "Reservas", "Config", "Servicios" y
 * "Usuarios").
 *
 * INSTALACIÓN NUEVA — hazlo en este orden:
 *   1. Ejecuta  configurarHojas  una vez (▶ en el editor). Pedirá
 *      autorización la primera vez: acéptala.
 *   2. Completa la pestaña "Config" con los datos de tu centro, y
 *      revisa que la pestaña "Servicios" tenga el Rol correcto para
 *      cada servicio (nutricionista / kinesiologo).
 *   3. Implementar > Nueva implementación > Tipo: Aplicación web.
 *      - Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario
 *      Copia la URL: esa es la página que subes a tu web.
 *      El dashboard de profesionales vive en esa misma URL, agregando
 *      "?page=dashboard" al final.
 *   4. Ejecuta  crearTriggerRecordatorios  una vez.
 *   5. Ejecuta  crearUsuario  una vez POR CADA profesional (edita los
 *      valores dentro de la función antes de cada ejecución).
 *
 * SI YA TENÍAS EL SISTEMA FUNCIONANDO (con reservas guardadas):
 *   No uses configurarHojas (borra todo). Usa en cambio, en este
 *   orden: agregarColumnaSucursal, configurarDashboard, y luego
 *   crearUsuario una vez por profesional.
 *
 * Todo el detalle paso a paso está en SETUP.md.
 */

const SHEET_RESERVAS = 'Reservas';
const SHEET_CONFIG = 'Config';

// ID de la hoja de cálculo donde viven las pestañas "Reservas" y "Config".
// Se usa así (en vez de "la hoja activa") para que el script funcione
// igual sea un proyecto independiente (script.google.com) o uno creado
// desde Extensiones > Apps Script dentro de la propia hoja.
// Cópialo de la URL de tu hoja: .../spreadsheets/d/ESTE_TROZO/edit
const SPREADSHEET_ID = '1dbovnsQO6JCBowjIZNbm6a-HsEgQADIhrv21cpIMH78';

function getSpreadsheet_() {
  return SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function avisar_(mensaje) {
  try {
    SpreadsheetApp.getUi().alert(mensaje);
  } catch (e) {
    Logger.log(mensaje);
  }
}

const COL = {
  TIMESTAMP: 1, NOMBRE: 2, EMAIL: 3, TELEFONO: 4, SERVICIO: 5,
  FECHA: 6, HORA: 7, ESTADO: 8, RECORDATORIO: 9, SUCURSAL: 10,
  PROFESIONAL: 11, NOTA: 12, TIPO_ATENCION: 13
};
const NUM_COLS = 13;

const SHEET_SERVICIOS = 'Servicios';
const SHEET_USUARIOS = 'Usuarios';
const SHEET_DISPONIBILIDAD = 'Disponibilidad';

// Cambia esto por cualquier texto propio antes de crear usuarios — es un
// ingrediente extra para las contraseñas guardadas, no hace falta
// recordarlo ni compartirlo con nadie.
const PEPPER_ = 'nutri-rosi-2026';

// ============================================================
// 1. CONFIGURACIÓN INICIAL — ejecutar una sola vez
// ============================================================

// ¡OJO! Esta función BORRA todo el contenido de "Reservas" y "Config".
// Úsala solo para una instalación nueva. Si ya tienes reservas guardadas,
// usa en cambio  agregarColumnaSucursal()  más abajo.
function configurarHojas() {
  const ss = getSpreadsheet_();

  let reservas = ss.getSheetByName(SHEET_RESERVAS);
  if (!reservas) reservas = ss.insertSheet(SHEET_RESERVAS);
  reservas.clear();
  reservas.getRange(1, 1, 1, NUM_COLS).setValues([[
    'Marca temporal', 'Nombre', 'Email', 'Teléfono', 'Servicio',
    'Fecha', 'Hora', 'Estado', 'Recordatorio enviado', 'Sucursal',
    'Profesional', 'Nota', 'Tipo de atención'
  ]]);
  reservas.setFrozenRows(1);
  reservas.autoResizeColumns(1, NUM_COLS);

  let config = ss.getSheetByName(SHEET_CONFIG);
  if (!config) config = ss.insertSheet(SHEET_CONFIG);
  config.clear();
  config.getRange('A1:B14').setValues([
    ['Campo', 'Valor'],
    ['Nombre del centro', 'Nutri-Rosi'],
    ['Email de contacto', Session.getActiveUser().getEmail() || 'contacto@nutri-rosi.cl'],
    ['Servicios (separados por coma)', 'Primera consulta nutricional, Control nutricional, Evaluación antropométrica, Evaluación kinésica, Sesión de kinesiología'],
    ['Sucursales (separadas por coma)', 'Las Compañías, San Ramón'],
    ['Duración de cada cita (minutos)', 45],
    ['Días de atención (0=Dom .. 6=Sáb, separados por coma)', '1,2,3,4,5'],
    ['Hora de inicio (HH:mm)', '09:00'],
    ['Hora de término (HH:mm)', '18:00'],
    ['Días de anticipación máxima para reservar', 30],
    ['Horas antes de la cita para enviar recordatorio', 24],
    ['Asunto email confirmación', 'Tu hora en Nutri-Rosi quedó confirmada'],
    ['Asunto email recordatorio', 'Recordatorio: tu hora en Nutri-Rosi es pronto'],
    ['Zona horaria', 'America/Santiago'],
  ]);
  config.setColumnWidth(1, 320);
  config.setColumnWidth(2, 320);

  crearHojaServicios_(ss);
  crearHojaUsuarios_(ss);
  crearHojaDisponibilidad_(ss);

  avisar_(
    'Listo. Ahora completa la pestaña "Config" con los datos de tu centro, ' +
    'y luego implementa este proyecto como Aplicación web (menú Implementar).'
  );
}

function crearHojaServicios_(ss) {
  let hoja = ss.getSheetByName(SHEET_SERVICIOS);
  if (hoja) return hoja;
  hoja = ss.insertSheet(SHEET_SERVICIOS);
  hoja.getRange(1, 1, 6, 2).setValues([
    ['Servicio', 'Rol'],
    ['Primera consulta nutricional', 'nutricionista'],
    ['Control nutricional', 'nutricionista'],
    ['Evaluación antropométrica', 'nutricionista'],
    ['Evaluación kinésica', 'kinesiologo'],
    ['Sesión de kinesiología', 'kinesiologo'],
  ]);
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, 2);
  return hoja;
}

function crearHojaUsuarios_(ss) {
  let hoja = ss.getSheetByName(SHEET_USUARIOS);
  if (hoja) return hoja;
  hoja = ss.insertSheet(SHEET_USUARIOS);
  hoja.getRange(1, 1, 1, 5).setValues([
    ['Usuario', 'ContraseñaHash', 'Rol', 'Nombre', 'Activo']
  ]);
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, 5);
  return hoja;
}

// Horarios que un profesional (o el admin, a nombre de un profesional)
// abrió a mano para que los pacientes puedan reservarlos. Sin una fila
// acá, ese horario NO aparece disponible — el sistema es "cerrado por
// defecto, se abre a mano" (ver getHorariosDisponibles más abajo).
function crearHojaDisponibilidad_(ss) {
  let hoja = ss.getSheetByName(SHEET_DISPONIBILIDAD);
  if (hoja) return hoja;
  hoja = ss.insertSheet(SHEET_DISPONIBILIDAD);
  hoja.getRange(1, 1, 1, 6).setValues([
    ['Fecha', 'Hora', 'Sucursal', 'Profesional', 'Abierto por', 'Marca temporal']
  ]);
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, 6);
  return hoja;
}

// Agrega la columna "Sucursal" y el campo "Sucursales" a un proyecto que
// YA tiene reservas guardadas, sin borrar nada. Ejecútala una sola vez.
function agregarColumnaSucursal() {
  const ss = getSpreadsheet_();

  const reservas = ss.getSheetByName(SHEET_RESERVAS);
  if (reservas.getRange(1, COL.SUCURSAL).getValue() !== 'Sucursal') {
    reservas.getRange(1, COL.SUCURSAL).setValue('Sucursal');
  }

  const config = ss.getSheetByName(SHEET_CONFIG);
  const campos = config.getRange(2, 1, config.getLastRow() - 1, 1).getValues().map(r => r[0]);
  if (campos.indexOf('Sucursales (separadas por coma)') === -1) {
    config.getRange(config.getLastRow() + 1, 1, 1, 2)
      .setValues([['Sucursales (separadas por coma)', 'Las Compañías, San Ramón']]);
  }

  avisar_('Listo: agregué la columna "Sucursal" en Reservas y el campo "Sucursales" en Config. Revisa que la lista de sucursales sea correcta.');
}

// Agrega el dashboard de profesionales a un proyecto que YA tiene
// reservas guardadas, sin borrar nada. Ejecútala una sola vez.
// Después de correrla, crea cada usuario a mano con crearUsuario()
// (ver instrucciones en SETUP.md).
function configurarDashboard() {
  const ss = getSpreadsheet_();

  const reservas = ss.getSheetByName(SHEET_RESERVAS);
  if (reservas.getRange(1, COL.PROFESIONAL).getValue() !== 'Profesional') {
    reservas.getRange(1, COL.PROFESIONAL).setValue('Profesional');
  }
  if (reservas.getRange(1, COL.NOTA).getValue() !== 'Nota') {
    reservas.getRange(1, COL.NOTA).setValue('Nota');
  }

  crearHojaServicios_(ss);
  crearHojaUsuarios_(ss);
  crearHojaDisponibilidad_(ss);

  avisar_(
    'Listo: se agregó la pestaña "Servicios" (revisa que el Rol de cada uno sea correcto), ' +
    'la pestaña "Usuarios" (vacía), la pestaña "Disponibilidad" y las columnas "Profesional"/"Nota" en Reservas. ' +
    'Ahora crea cada cuenta con la función crearUsuario() — instrucciones en SETUP.md.'
  );
}

// Agrega solo la pestaña "Disponibilidad" a un proyecto que YA tiene el
// dashboard funcionando (configurarDashboard ya corrida antes). Ejecútala
// una sola vez para pasar del modelo viejo (todo abierto, se bloquea) al
// nuevo (todo cerrado, se abre a mano).
function agregarHojaDisponibilidad() {
  crearHojaDisponibilidad_(getSpreadsheet_());
  avisar_('Listo: se agregó la pestaña "Disponibilidad". Ahora rosi y kin deben abrir sus horarios desde el dashboard antes de que los pacientes puedan reservarlos.');
}

// Agrega la columna "Tipo de atención" (presencial/online, solo se usa
// para servicios de nutrición) a un proyecto que YA tiene reservas
// guardadas, sin borrar nada. Ejecútala una sola vez.
function agregarColumnaTipoAtencion() {
  const ss = getSpreadsheet_();
  const reservas = ss.getSheetByName(SHEET_RESERVAS);
  if (reservas.getRange(1, COL.TIPO_ATENCION).getValue() !== 'Tipo de atención') {
    reservas.getRange(1, COL.TIPO_ATENCION).setValue('Tipo de atención');
  }
  avisar_('Listo: agregué la columna "Tipo de atención" en Reservas.');
}

// Crea o actualiza un usuario del dashboard. Ejecútala UNA VEZ POR CADA
// profesional, cambiando los valores de abajo antes de correrla. Vuelve a
// dejar el código como estaba (o simplemente ignóralo) después de crear
// cada cuenta — no hace falta borrar la función.
//
// rol debe ser exactamente: 'nutricionista', 'kinesiologo', o 'admin'
// ('admin' ve y gestiona las horas de todos los profesionales).
function crearUsuario() {
  const usuario = 'rosi';           // <-- cámbialo
  const password = 'CambiaEsto123'; // <-- cámbialo (díselo solo a esa persona)
  const rol = 'nutricionista';      // <-- 'nutricionista' | 'kinesiologo' | 'admin'
  const nombre = 'Rosi';            // <-- nombre para mostrar en el dashboard

  const hoja = crearHojaUsuarios_(getSpreadsheet_());
  const hash = hashPassword_(usuario, password);

  const filas = hoja.getLastRow() > 1
    ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues().map(r => r[0])
    : [];
  const idx = filas.indexOf(usuario);

  if (idx === -1) {
    hoja.appendRow([usuario, hash, rol, nombre, 'Sí']);
  } else {
    hoja.getRange(idx + 2, 1, 1, 5).setValues([[usuario, hash, rol, nombre, 'Sí']]);
  }

  avisar_(`Listo: usuario "${usuario}" (${rol}) guardado/actualizado.`);
}

// Crea de una vez las cuentas iniciales del dashboard. Pensada para
// correrla UNA sola vez al instalar; después de eso, usa crearUsuario()
// para agregar o cambiar cuentas de a una.
function crearUsuariosIniciales() {
  const cuentas = [
    { usuario: 'rosi', password: '123456789', rol: 'nutricionista', nombre: 'Rosi' },
    { usuario: 'kin', password: '123456789', rol: 'kinesiologo', nombre: 'kine' },
    { usuario: 'root', password: '123456789', rol: 'admin', nombre: 'admin' },
  ];

  const hoja = crearHojaUsuarios_(getSpreadsheet_());
  const filas = hoja.getLastRow() > 1
    ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues().map(r => r[0])
    : [];

  cuentas.forEach(c => {
    const hash = hashPassword_(c.usuario, c.password);
    const idx = filas.indexOf(c.usuario);
    if (idx === -1) {
      hoja.appendRow([c.usuario, hash, c.rol, c.nombre, 'Sí']);
      filas.push(c.usuario);
    } else {
      hoja.getRange(idx + 2, 1, 1, 5).setValues([[c.usuario, hash, c.rol, c.nombre, 'Sí']]);
    }
  });

  avisar_('Listo: usuarios rosi, kin y root creados/actualizados.');
}

function hashPassword_(usuario, password) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    usuario + ':' + password + ':' + PEPPER_
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function crearTriggerRecordatorios() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'enviarRecordatorios') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarRecordatorios').timeBased().everyHours(1).create();
  avisar_('Recordatorios automáticos activados: se revisarán cada hora.');
}

// ============================================================
// 2. LECTURA DE CONFIGURACIÓN
// ============================================================

function getConfigMap_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  const map = {};
  if (!sheet) {
    throw new Error('No existe la pestaña "Config". Ejecuta la función configurarHojas() una vez desde el editor de Apps Script.');
  }
  if (sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    rows.forEach(([campo, valor]) => { map[campo] = valor; });
  }
  const zona = map['Zona horaria'] || 'America/Santiago';
  return {
    nombreCentro: map['Nombre del centro'] || 'Nutri-Rosi',
    emailContacto: map['Email de contacto'],
    servicios: String(map['Servicios (separados por coma)'] || '')
      .split(',').map(s => s.trim()).filter(Boolean),
    sucursales: String(map['Sucursales (separadas por coma)'] || '')
      .split(',').map(s => s.trim()).filter(Boolean),
    duracionMin: Number(map['Duración de cada cita (minutos)']) || 45,
    diasAtencion: String(map['Días de atención (0=Dom .. 6=Sáb, separados por coma)'] || '1,2,3,4,5')
      .split(',').map(d => Number(d.trim())),
    horaInicio: normalizarHora_(map['Hora de inicio (HH:mm)'], '09:00', zona),
    horaTermino: normalizarHora_(map['Hora de término (HH:mm)'], '18:00', zona),
    diasAnticipacion: Number(map['Días de anticipación máxima para reservar']) || 30,
    horasRecordatorio: Number(map['Horas antes de la cita para enviar recordatorio']) || 24,
    asuntoConfirmacion: map['Asunto email confirmación'] || 'Tu hora quedó confirmada',
    asuntoRecordatorio: map['Asunto email recordatorio'] || 'Recordatorio de tu hora',
    zona: zona,
  };
}

// Google Sheets auto-convierte texto tipo "09:00" a un valor de hora/fecha
// interno. Esta función lo deja siempre como texto "HH:mm", venga como
// Date (por el auto-formato) o como string (si la celda estaba en texto plano).
function normalizarHora_(valor, porDefecto, zona) {
  if (!valor) return porDefecto;
  if (valor instanceof Date) return Utilities.formatDate(valor, zona, 'HH:mm');
  return String(valor);
}

// Servicios con su rol y, si corresponde, la lista de profesionales entre
// los que el paciente debe elegir (solo cuando hay más de uno activo para
// ese rol — así hoy aparece para nutrición, con kinesiología queda igual
// que antes, y si mañana hay dos kinesiólogos aparece ahí también, sin
// tocar el formulario).
function getServiciosPublicos_() {
  const hoja = getSpreadsheet_().getSheetByName(SHEET_SERVICIOS);
  if (hoja && hoja.getLastRow() >= 2) {
    const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues();
    return filas.filter(f => f[0]).map(f => {
      const rol = f[1];
      const profesionales = getProfesionalesPorRol_(rol);
      return {
        servicio: f[0],
        rol: rol,
        requiereProfesional: profesionales.length > 1,
        requiereTipoAtencion: rol === 'nutricionista',
        profesionales: profesionales,
      };
    });
  }
  // Respaldo si la hoja "Servicios" no existe: lista plana desde Config.
  return getConfigMap_().servicios.map(s => ({
    servicio: s, rol: '', requiereProfesional: false, requiereTipoAtencion: false, profesionales: [],
  }));
}

// Expuesto al formulario web (sin datos internos sensibles)
function getInfoPublica() {
  const cfg = getConfigMap_();
  return {
    nombreCentro: cfg.nombreCentro,
    servicios: getServiciosPublicos_(),
    sucursales: cfg.sucursales,
    diasAnticipacion: cfg.diasAnticipacion,
    zona: cfg.zona,
  };
}

// ============================================================
// 3. APLICACIÓN WEB
// ============================================================

function doGet(e) {
  const pagina = e && e.parameter && e.parameter.page;

  if (pagina === 'dashboard') {
    return HtmlService.createHtmlOutputFromFile('Dashboard')
      .setTitle('Dashboard — Nutri-Rosi')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Reservar hora — Nutri-Rosi')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ------------------------------------------------------------
// Login y sesiones (dashboard de profesionales)
// ------------------------------------------------------------

function login(usuario, password) {
  usuario = String(usuario || '').trim();
  const hoja = getSpreadsheet_().getSheetByName(SHEET_USUARIOS);
  if (!hoja || hoja.getLastRow() < 2) {
    return { ok: false, mensaje: 'Usuario o contraseña incorrectos.' };
  }

  const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 5).getValues();
  const fila = filas.find(f => f[0] === usuario);
  if (!fila) return { ok: false, mensaje: 'Usuario o contraseña incorrectos.' };

  const [ , hashGuardado, rol, nombre, activo] = fila;
  if (activo === 'No') return { ok: false, mensaje: 'Esta cuenta está desactivada.' };
  if (hashPassword_(usuario, password) !== hashGuardado) {
    return { ok: false, mensaje: 'Usuario o contraseña incorrectos.' };
  }

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(
    'sesion_' + token,
    JSON.stringify({ usuario, rol, nombre }),
    21600 // 6 horas, el máximo que permite CacheService
  );

  return { ok: true, token, rol, nombre };
}

function validarToken_(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get('sesion_' + token);
  return raw ? JSON.parse(raw) : null;
}

function cerrarSesion(token) {
  if (token) CacheService.getScriptCache().remove('sesion_' + token);
  return { ok: true };
}

// ============================================================
// 4. HORARIOS DISPONIBLES
// ============================================================

// La hoja "Servicios" dice qué profesional (rol) atiende cada servicio.
// Esto permite que nutricionista y kinesiólogo tengan horas a la misma
// hora en la misma sucursal sin pisarse — cada uno tiene su propia
// disponibilidad.
function getRolServicio_(servicio) {
  const hoja = getSpreadsheet_().getSheetByName(SHEET_SERVICIOS);
  if (!hoja || hoja.getLastRow() < 2) return '';
  const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues();
  const fila = filas.find(f => f[0] === servicio);
  return fila ? fila[1] : '';
}

// Profesionales activos (usuarios del dashboard) para un rol dado
// ('nutricionista' o 'kinesiologo'). Permite que un mismo rol tenga más
// de un profesional (ej. dos nutricionistas), cada uno con su propia
// agenda — y que se puedan agregar más cuando haga falta, sin tocar código.
function getProfesionalesPorRol_(rol) {
  const hoja = getSpreadsheet_().getSheetByName(SHEET_USUARIOS);
  if (!hoja || hoja.getLastRow() < 2) return [];
  const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 5).getValues();
  return filas
    .filter(f => f[2] === rol && f[4] !== 'No')
    .map(f => ({ usuario: f[0], nombre: f[3] }));
}

// Nombre para mostrar de un profesional, a partir de su usuario de login.
function nombrePorUsuario_(usuario) {
  if (!usuario) return '';
  const hoja = getSpreadsheet_().getSheetByName(SHEET_USUARIOS);
  if (!hoja || hoja.getLastRow() < 2) return '';
  const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 4).getValues();
  const fila = filas.find(f => f[0] === usuario);
  return fila ? fila[3] : '';
}

// Resuelve a qué profesional concreto (usuario) corresponde una reserva:
// si el rol solo tiene un profesional activo, se asigna automáticamente
// (comportamiento actual, sin cambios visibles); si tiene más de uno, el
// paciente debe haber elegido uno válido.
function resolverProfesionalReserva_(rol, profesionalSolicitado) {
  const activos = getProfesionalesPorRol_(rol);
  if (!activos.length) {
    throw new Error('No hay profesionales activos para este servicio. Contacta al centro.');
  }
  if (activos.length === 1) return activos[0].usuario;
  const encontrado = activos.find(p => p.usuario === profesionalSolicitado);
  if (!encontrado) throw new Error('Selecciona con qué profesional quieres la hora.');
  return encontrado.usuario;
}

// Lista de horas "de plantilla" del día (09:00, 09:45, 10:30...) según
// la duración de cita y el rango horaInicio–horaTermino de Config. Es
// solo la grilla de referencia que usa el dashboard para ofrecer horas
// al abrir agenda — NO implica que esas horas ya estén disponibles.
function getHorasBase_(cfg) {
  const slots = [];
  const [hi, mi] = cfg.horaInicio.split(':').map(Number);
  const [ht, mt] = cfg.horaTermino.split(':').map(Number);
  let cursor = hi * 60 + mi;
  const fin = ht * 60 + mt;
  while (cursor < fin) {
    const hh = Math.floor(cursor / 60), mm = cursor % 60;
    slots.push(('0' + hh).slice(-2) + ':' + ('0' + mm).slice(-2));
    cursor += cfg.duracionMin;
  }
  return slots;
}

// Horas que el profesional (o el admin a su nombre) abrió a mano para
// esa fecha/sucursal — ver hoja "Disponibilidad". "profesional" es el
// usuario de login (ej. 'rosi', 'catalina', 'kin'): cada profesional
// tiene su propia agenda, aunque compartan el mismo rol.
function getHorasAbiertas_(fechaStr, sucursal, profesional) {
  const sheet = getSpreadsheet_().getSheetByName(SHEET_DISPONIBILIDAD);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return data
    .filter(row => row[0] === fechaStr && row[2] === sucursal && row[3] === profesional)
    .map(row => row[1]);
}

// Acepta como tercer parámetro un nombre de servicio (lo normal, desde el
// formulario de pacientes) O directamente un usuario de profesional (lo
// que usa el dashboard al abrir/cerrar horarios, y reagendarReserva). Si
// no encuentra el servicio en la hoja "Servicios", asume que ya le
// pasaron el profesional tal cual. Cuando viene de un servicio real y ese
// rol tiene más de un profesional activo, se necesita el 4to parámetro
// (profesionalSolicitado) para saber de cuál de todos se quieren ver las
// horas.
//
// Modelo "cerrado por defecto": un horario solo aparece disponible si
// alguien lo abrió antes a mano (hoja "Disponibilidad") y todavía no
// tiene una reserva encima.
function getHorariosDisponibles(fechaStr, sucursal, servicioORol, profesionalSolicitado) {
  const rol = getRolServicio_(servicioORol);
  const profesional = rol ? resolverProfesionalReserva_(rol, profesionalSolicitado) : (servicioORol || '');
  const cfg = getConfigMap_();
  const fecha = new Date(fechaStr + 'T00:00:00');

  const hoy = new Date();
  const limite = new Date();
  limite.setDate(hoy.getDate() + cfg.diasAnticipacion);
  if (fecha < new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate()) || fecha > limite) return [];

  const abiertas = getHorasAbiertas_(fechaStr, sucursal, profesional);
  if (!abiertas.length) return [];

  const ocupadas = getHorasOcupadas_(fechaStr, sucursal, profesional);
  const ahora = new Date();
  const esHoy = fechaStr === Utilities.formatDate(ahora, cfg.zona, 'yyyy-MM-dd');

  return abiertas.filter(h => {
    if (ocupadas.indexOf(h) !== -1) return false;
    if (esHoy) {
      const [hh, mm] = h.split(':').map(Number);
      const slotTime = new Date(fecha); slotTime.setHours(hh, mm, 0, 0);
      if (slotTime <= ahora) return false;
    }
    return true;
  }).sort();
}

function getHorasOcupadas_(fechaStr, sucursal, profesional) {
  const sheet = getSpreadsheet_().getSheetByName(SHEET_RESERVAS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();
  return data
    .filter(row =>
      row[COL.FECHA - 1] === fechaStr &&
      row[COL.ESTADO - 1] !== 'Cancelada' &&
      // Reservas guardadas antes de tener sucursal (celda vacía) se
      // consideran ocupadas para cualquier sucursal, por seguridad.
      (!row[COL.SUCURSAL - 1] || row[COL.SUCURSAL - 1] === sucursal) &&
      // Lo mismo para "profesional": si no se sabe a quién pertenece
      // (reservas antiguas), se considera ocupada para cualquier profesional.
      (!row[COL.PROFESIONAL - 1] || row[COL.PROFESIONAL - 1] === profesional)
    )
    .map(row => row[COL.HORA - 1]);
}

// ============================================================
// 5. CREAR RESERVA
// ============================================================

function crearReserva(datos) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const cfg = getConfigMap_();

    if (!datos.nombre || !datos.email || !datos.servicio || !datos.sucursal || !datos.fecha || !datos.hora) {
      throw new Error('Faltan datos obligatorios.');
    }

    const rol = getRolServicio_(datos.servicio);
    const profesional = rol ? resolverProfesionalReserva_(rol, datos.profesional) : '';

    const disponibles = getHorariosDisponibles(datos.fecha, datos.sucursal, datos.servicio, datos.profesional);
    if (disponibles.indexOf(datos.hora) === -1) {
      throw new Error('Ese horario ya no está disponible. Por favor elige otro.');
    }

    const tipoAtencion = rol === 'nutricionista'
      ? (datos.tipoAtencion === 'Online' ? 'Online' : 'Presencial')
      : '';

    const sheet = getSpreadsheet_().getSheetByName(SHEET_RESERVAS);
    sheet.appendRow([
      new Date(), datos.nombre, datos.email, datos.telefono || '',
      datos.servicio, datos.fecha, datos.hora, 'Confirmada', 'No', datos.sucursal,
      profesional, '', tipoAtencion
    ]);

    datos = Object.assign({}, datos, { profesionalNombre: nombrePorUsuario_(profesional), tipoAtencion: tipoAtencion });
    enviarEmailConfirmacion_(datos, cfg);

    return { ok: true, mensaje: 'Tu hora quedó confirmada. Te enviamos un correo con los detalles.' };
  } catch (e) {
    return { ok: false, mensaje: e.message };
  } finally {
    lock.releaseLock();
  }
}

function enviarEmailConfirmacion_(datos, cfg) {
  const fechaLegible = Utilities.formatDate(new Date(datos.fecha + 'T00:00:00'), cfg.zona, "EEEE d 'de' MMMM 'de' yyyy");
  const cuerpo =
    `Hola ${datos.nombre},\n\n` +
    `Tu hora en ${cfg.nombreCentro} quedó confirmada:\n\n` +
    `  Servicio: ${datos.servicio}\n` +
    (datos.profesionalNombre ? `  Profesional: ${datos.profesionalNombre}\n` : '') +
    (datos.tipoAtencion ? `  Tipo de atención: ${datos.tipoAtencion}\n` : '') +
    `  Sucursal: ${datos.sucursal}\n` +
    `  Fecha: ${fechaLegible}\n` +
    `  Hora: ${datos.hora}\n\n` +
    `Si necesitas reprogramar o cancelar, responde directamente este correo.\n\n` +
    `${cfg.nombreCentro}`;

  MailApp.sendEmail({
    to: datos.email,
    subject: cfg.asuntoConfirmacion,
    body: cuerpo,
    ...(cfg.emailContacto ? { replyTo: cfg.emailContacto } : {})
  });
}

// ============================================================
// 6. RECORDATORIOS AUTOMÁTICOS (se ejecuta cada hora via trigger)
// ============================================================

function enviarRecordatorios() {
  const cfg = getConfigMap_();
  const sheet = getSpreadsheet_().getSheetByName(SHEET_RESERVAS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();
  const ahora = new Date();

  data.forEach((row, i) => {
    const [ , nombre, email, telefono, servicio, fecha, hora, estado, recordatorioEnviado, sucursal] = row;
    if (estado !== 'Confirmada' || recordatorioEnviado === 'Sí') return;
    if (!fecha || !hora || !email) return;

    const [hh, mm] = String(hora).split(':').map(Number);
    const fechaHoraCita = new Date(fecha);
    fechaHoraCita.setHours(hh, mm, 0, 0);

    const horasFaltantes = (fechaHoraCita - ahora) / 3600000;

    if (horasFaltantes > 0 && horasFaltantes <= cfg.horasRecordatorio) {
      const fechaLegible = Utilities.formatDate(fechaHoraCita, cfg.zona, "EEEE d 'de' MMMM");
      const cuerpo =
        `Hola ${nombre},\n\n` +
        `Te recordamos tu hora en ${cfg.nombreCentro}:\n\n` +
        `  Servicio: ${servicio}\n` +
        (sucursal ? `  Sucursal: ${sucursal}\n` : '') +
        `  Fecha: ${fechaLegible}\n` +
        `  Hora: ${hora}\n\n` +
        `Si no puedes asistir, responde este correo para reprogramar.\n\n` +
        `${cfg.nombreCentro}`;

      MailApp.sendEmail({
        to: email,
        subject: cfg.asuntoRecordatorio,
        body: cuerpo,
        ...(cfg.emailContacto ? { replyTo: cfg.emailContacto } : {})
      });

      sheet.getRange(i + 2, COL.RECORDATORIO).setValue('Sí');
    }
  });
}

// ============================================================
// 7. DASHBOARD DE PROFESIONALES
// ============================================================

// Devuelve las horas del profesional que inició sesión (o de todos, si
// es admin), desde ayer en adelante, ordenadas por fecha y hora.
function getMisReservas(token) {
  const sesion = validarToken_(token);
  if (!sesion) throw new Error('Tu sesión expiró. Vuelve a iniciar sesión.');

  const sheet = getSpreadsheet_().getSheetByName(SHEET_RESERVAS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const cfg = getConfigMap_();
  const ayer = new Date();
  ayer.setDate(ayer.getDate() - 1);
  const ayerStr = Utilities.formatDate(ayer, cfg.zona, 'yyyy-MM-dd');

  const data = sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();

  const filas = data
    .map((row, i) => ({
      fila: i + 2, // número real de la fila en la hoja, para poder editarla después
      nombre: row[COL.NOMBRE - 1],
      email: row[COL.EMAIL - 1],
      telefono: row[COL.TELEFONO - 1],
      servicio: row[COL.SERVICIO - 1],
      fecha: row[COL.FECHA - 1],
      hora: row[COL.HORA - 1],
      estado: row[COL.ESTADO - 1],
      sucursal: row[COL.SUCURSAL - 1],
      profesional: row[COL.PROFESIONAL - 1],
      profesionalNombre: nombrePorUsuario_(row[COL.PROFESIONAL - 1]),
      tipoAtencion: row[COL.TIPO_ATENCION - 1] || '',
      nota: row[COL.NOTA - 1] || '',
    }))
    .filter(r => r.fecha && r.fecha >= ayerStr)
    .filter(r => sesion.rol === 'admin' || r.profesional === sesion.usuario);

  filas.sort((a, b) => (a.fecha + a.hora).localeCompare(b.fecha + b.hora));

  return { filas, rol: sesion.rol, nombre: sesion.nombre };
}

function puedeEditarFila_(sesion, fila) {
  if (sesion.rol === 'admin') return true;
  const sheet = getSpreadsheet_().getSheetByName(SHEET_RESERVAS);
  const profesional = sheet.getRange(fila, COL.PROFESIONAL).getValue();
  return profesional === sesion.usuario;
}

function actualizarEstado(token, fila, nuevoEstado) {
  const sesion = validarToken_(token);
  if (!sesion) throw new Error('Tu sesión expiró. Vuelve a iniciar sesión.');
  if (!puedeEditarFila_(sesion, fila)) throw new Error('No tienes permiso para editar esta hora.');

  const estadosValidos = ['Confirmada', 'Atendida', 'Cancelada'];
  if (estadosValidos.indexOf(nuevoEstado) === -1) throw new Error('Estado inválido.');

  getSpreadsheet_().getSheetByName(SHEET_RESERVAS).getRange(fila, COL.ESTADO).setValue(nuevoEstado);
  return { ok: true };
}

function guardarNota(token, fila, nota) {
  const sesion = validarToken_(token);
  if (!sesion) throw new Error('Tu sesión expiró. Vuelve a iniciar sesión.');
  if (!puedeEditarFila_(sesion, fila)) throw new Error('No tienes permiso para editar esta hora.');

  getSpreadsheet_().getSheetByName(SHEET_RESERVAS).getRange(fila, COL.NOTA).setValue(String(nota || ''));
  return { ok: true };
}

// Mueve una reserva existente a otra fecha/hora, verificando que el
// nuevo horario esté disponible. Pensada sobre todo para el panel de
// administración ("Agenda general"), pero cualquiera que pueda editar
// la fila puede usarla.
function reagendarReserva(token, fila, nuevaFecha, nuevaHora) {
  const sesion = validarToken_(token);
  if (!sesion) throw new Error('Tu sesión expiró. Vuelve a iniciar sesión.');
  if (!puedeEditarFila_(sesion, fila)) throw new Error('No tienes permiso para editar esta hora.');
  if (!nuevaFecha || !nuevaHora) throw new Error('Falta la nueva fecha u hora.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSpreadsheet_().getSheetByName(SHEET_RESERVAS);
    const filaActual = sheet.getRange(fila, 1, 1, NUM_COLS).getValues()[0];
    if (filaActual[COL.ESTADO - 1] === 'Cancelada') {
      throw new Error('Esta hora está cancelada; no se puede reagendar.');
    }

    const sucursal = filaActual[COL.SUCURSAL - 1];
    const profesional = filaActual[COL.PROFESIONAL - 1];
    const mismoHorario = nuevaFecha === filaActual[COL.FECHA - 1] && nuevaHora === filaActual[COL.HORA - 1];

    if (!mismoHorario) {
      const disponibles = getHorariosDisponibles(nuevaFecha, sucursal, profesional);
      if (disponibles.indexOf(nuevaHora) === -1) {
        throw new Error('Ese nuevo horario no está disponible (revisa que esté abierto y libre).');
      }
    }

    sheet.getRange(fila, COL.FECHA).setValue(nuevaFecha);
    sheet.getRange(fila, COL.HORA).setValue(nuevaHora);
    return { ok: true, mensaje: 'Hora reagendada.' };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 8. ABRIR / CERRAR AGENDA (modelo "cerrado por defecto")
// ============================================================

function rolProfesionalValido_(rol) {
  return rol === 'nutricionista' || rol === 'kinesiologo';
}

// Cada profesional (usuario) gestiona su propia agenda, aunque comparta
// rol con otro (ej. dos nutricionistas). El admin puede gestionar la de
// cualquiera, pero debe indicar explícitamente de cuál (datos.profesional
// = usuario de esa persona).
function resolverProfesionalObjetivo_(sesion, datos) {
  if (sesion.rol === 'admin') {
    const usuario = datos && datos.profesional;
    if (!usuario || !usuarioProfesionalValido_(usuario)) {
      throw new Error('Indica de qué profesional quieres gestionar la agenda.');
    }
    return usuario;
  }
  if (!rolProfesionalValido_(sesion.rol)) {
    throw new Error('Esta cuenta no gestiona una agenda de profesional.');
  }
  return sesion.usuario;
}

function usuarioProfesionalValido_(usuario) {
  const hoja = getSpreadsheet_().getSheetByName(SHEET_USUARIOS);
  if (!hoja || hoja.getLastRow() < 2) return false;
  const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 5).getValues();
  return filas.some(f => f[0] === usuario && rolProfesionalValido_(f[2]) && f[4] !== 'No');
}

function validarFechaGestionable_(cfg, fechaStr) {
  const fecha = new Date(fechaStr + 'T00:00:00');
  const hoy = new Date();
  const limite = new Date();
  limite.setDate(hoy.getDate() + cfg.diasAnticipacion);
  if (fecha < new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate())) {
    throw new Error('No puedes abrir o cerrar una fecha pasada.');
  }
  if (fecha > limite) {
    throw new Error('Esa fecha está más allá de los días de anticipación permitidos (revisa Config).');
  }
}

// Datos para pintar el dashboard: sucursales, y si es admin, la lista de
// profesionales (para elegir de quién gestionar la agenda).
function getInfoDashboard(token) {
  const sesion = validarToken_(token);
  if (!sesion) throw new Error('Tu sesión expiró. Vuelve a iniciar sesión.');
  const cfg = getConfigMap_();

  const info = { rol: sesion.rol, nombre: sesion.nombre, sucursales: cfg.sucursales };

  if (sesion.rol === 'admin') {
    const hoja = getSpreadsheet_().getSheetByName(SHEET_USUARIOS);
    const lastRow = hoja.getLastRow();
    const profesionales = [];
    if (lastRow >= 2) {
      hoja.getRange(2, 1, lastRow - 1, 5).getValues().forEach(row => {
        const [usuario, , rol, nombre, activo] = row;
        if (rolProfesionalValido_(rol) && activo !== 'No') {
          profesionales.push({ usuario, rol, nombre });
        }
      });
    }
    info.profesionales = profesionales;
  }

  return info;
}

// Estado de una fecha/sucursal/profesional para la grilla de "Abrir
// agenda": qué horas base hay, cuáles están abiertas y cuáles ya tienen
// una reserva encima.
function getEstadoDia(token, datos) {
  const sesion = validarToken_(token);
  if (!sesion) throw new Error('Tu sesión expiró. Vuelve a iniciar sesión.');
  const profesional = resolverProfesionalObjetivo_(sesion, datos);
  if (!datos || !datos.fecha || !datos.sucursal) throw new Error('Faltan datos.');

  const cfg = getConfigMap_();
  return {
    horasBase: getHorasBase_(cfg),
    abiertas: getHorasAbiertas_(datos.fecha, datos.sucursal, profesional),
    ocupadas: getHorasOcupadas_(datos.fecha, datos.sucursal, profesional)
  };
}

// Abre un horario puntual para que los pacientes puedan reservarlo.
function abrirHorario(token, datos) {
  const sesion = validarToken_(token);
  if (!sesion) throw new Error('Tu sesión expiró. Vuelve a iniciar sesión.');
  const profesional = resolverProfesionalObjetivo_(sesion, datos);

  if (!datos.fecha || !datos.hora || !datos.sucursal) {
    throw new Error('Faltan datos para abrir el horario.');
  }

  const cfg = getConfigMap_();
  validarFechaGestionable_(cfg, datos.fecha);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const yaAbiertas = getHorasAbiertas_(datos.fecha, datos.sucursal, profesional);
    if (yaAbiertas.indexOf(datos.hora) !== -1) {
      return { ok: true, mensaje: 'Ese horario ya estaba abierto.' };
    }

    crearHojaDisponibilidad_(getSpreadsheet_()).appendRow([
      datos.fecha, datos.hora, datos.sucursal, profesional, sesion.usuario, new Date()
    ]);

    return { ok: true, mensaje: 'Horario abierto.' };
  } finally {
    lock.releaseLock();
  }
}

// Abre de una vez todas las horas base del día (útil para no tener que
// abrir hora por hora en un día normal de atención).
function abrirDiaCompleto(token, datos) {
  const sesion = validarToken_(token);
  if (!sesion) throw new Error('Tu sesión expiró. Vuelve a iniciar sesión.');
  const profesional = resolverProfesionalObjetivo_(sesion, datos);

  if (!datos.fecha || !datos.sucursal) throw new Error('Faltan datos.');

  const cfg = getConfigMap_();
  validarFechaGestionable_(cfg, datos.fecha);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const base = getHorasBase_(cfg);
    const yaAbiertas = getHorasAbiertas_(datos.fecha, datos.sucursal, profesional);
    const nuevas = base.filter(h => yaAbiertas.indexOf(h) === -1);

    if (!nuevas.length) {
      return { ok: true, mensaje: 'Ese día ya estaba completamente abierto.', abiertos: 0 };
    }

    const sheet = crearHojaDisponibilidad_(getSpreadsheet_());
    const filas = nuevas.map(h => [datos.fecha, h, datos.sucursal, profesional, sesion.usuario, new Date()]);
    sheet.getRange(sheet.getLastRow() + 1, 1, filas.length, 6).setValues(filas);

    return { ok: true, mensaje: 'Se abrieron ' + nuevas.length + ' horarios.', abiertos: nuevas.length };
  } finally {
    lock.releaseLock();
  }
}

// Cierra un horario que estaba abierto (si nadie lo reservó todavía).
function cerrarHorario(token, datos) {
  const sesion = validarToken_(token);
  if (!sesion) throw new Error('Tu sesión expiró. Vuelve a iniciar sesión.');
  const profesional = resolverProfesionalObjetivo_(sesion, datos);

  if (!datos.fecha || !datos.hora || !datos.sucursal) {
    throw new Error('Faltan datos para cerrar el horario.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ocupadas = getHorasOcupadas_(datos.fecha, datos.sucursal, profesional);
    if (ocupadas.indexOf(datos.hora) !== -1) {
      throw new Error('Ese horario ya tiene una reserva — cancélala primero desde "Mis horas" o "Agenda general".');
    }

    const sheet = getSpreadsheet_().getSheetByName(SHEET_DISPONIBILIDAD);
    const lastRow = sheet ? sheet.getLastRow() : 0;
    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
      for (let i = data.length - 1; i >= 0; i--) {
        if (data[i][0] === datos.fecha && data[i][1] === datos.hora &&
            data[i][2] === datos.sucursal && data[i][3] === profesional) {
          sheet.deleteRow(i + 2);
          return { ok: true, mensaje: 'Horario cerrado.' };
        }
      }
    }

    return { ok: true, mensaje: 'Ese horario ya estaba cerrado.' };
  } finally {
    lock.releaseLock();
  }
}
