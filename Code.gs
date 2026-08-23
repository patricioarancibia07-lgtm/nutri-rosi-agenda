/**
 * NUTRI-ROSI · Agenda online + recordatorios automáticos
 * ---------------------------------------------------------------
 * Backend en Google Apps Script, usando esta misma Hoja de cálculo
 * como base de datos (pestañas "Reservas" y "Config").
 *
 * PRIMEROS PASOS (hazlo en este orden):
 *   1. Ejecuta la función  configurarHojas  una vez (▶ en el editor).
 *      La primera vez pedirá autorización: acéptala.
 *   2. Ve a la pestaña "Config" en la hoja de cálculo y completa los
 *      datos de tu centro (servicios, horario, etc).
 *   3. Implementar > Nueva implementación > Tipo: Aplicación web.
 *      - Ejecutar como: Yo
 *      - Quién tiene acceso: Cualquier usuario
 *      Copia la URL que te entrega: esa es la página que subes a tu
 *      web (o compartes / embebes en un <iframe>).
 *   4. Ejecuta la función  crearTriggerRecordatorios  una vez, para
 *      activar el envío automático de recordatorios cada hora.
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
  FECHA: 6, HORA: 7, ESTADO: 8, RECORDATORIO: 9, SUCURSAL: 10
};
const NUM_COLS = 10;

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
    'Fecha', 'Hora', 'Estado', 'Recordatorio enviado', 'Sucursal'
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

  avisar_(
    'Listo. Ahora completa la pestaña "Config" con los datos de tu centro, ' +
    'y luego implementa este proyecto como Aplicación web (menú Implementar).'
  );
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

// Expuesto al formulario web (sin datos internos sensibles)
function getInfoPublica() {
  const cfg = getConfigMap_();
  return {
    nombreCentro: cfg.nombreCentro,
    servicios: cfg.servicios,
    sucursales: cfg.sucursales,
    diasAnticipacion: cfg.diasAnticipacion,
    zona: cfg.zona,
  };
}

// ============================================================
// 3. APLICACIÓN WEB
// ============================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Reservar hora — Nutri-Rosi')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// 4. HORARIOS DISPONIBLES
// ============================================================

function getHorariosDisponibles(fechaStr, sucursal) {
  const cfg = getConfigMap_();
  const fecha = new Date(fechaStr + 'T00:00:00');
  const diaSemana = fecha.getDay();

  if (cfg.diasAtencion.indexOf(diaSemana) === -1) return [];

  const hoy = new Date();
  const limite = new Date();
  limite.setDate(hoy.getDate() + cfg.diasAnticipacion);
  if (fecha < new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate()) || fecha > limite) return [];

  const slots = [];
  const [hi, mi] = cfg.horaInicio.split(':').map(Number);
  const [ht, mt] = cfg.horaTermino.split(':').map(Number);
  let cursor = new Date(fecha); cursor.setHours(hi, mi, 0, 0);
  const fin = new Date(fecha); fin.setHours(ht, mt, 0, 0);

  while (cursor < fin) {
    slots.push(Utilities.formatDate(cursor, cfg.zona, 'HH:mm'));
    cursor = new Date(cursor.getTime() + cfg.duracionMin * 60000);
  }

  const ocupadas = getHorasOcupadas_(fechaStr, sucursal);
  const ahora = new Date();
  const esHoy = fechaStr === Utilities.formatDate(ahora, cfg.zona, 'yyyy-MM-dd');

  return slots.filter(h => {
    if (ocupadas.indexOf(h) !== -1) return false;
    if (esHoy) {
      const [hh, mm] = h.split(':').map(Number);
      const slotTime = new Date(fecha); slotTime.setHours(hh, mm, 0, 0);
      if (slotTime <= ahora) return false;
    }
    return true;
  });
}

function getHorasOcupadas_(fechaStr, sucursal) {
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
      (!row[COL.SUCURSAL - 1] || row[COL.SUCURSAL - 1] === sucursal)
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

    const disponibles = getHorariosDisponibles(datos.fecha, datos.sucursal);
    if (disponibles.indexOf(datos.hora) === -1) {
      throw new Error('Ese horario ya no está disponible. Por favor elige otro.');
    }

    const sheet = getSpreadsheet_().getSheetByName(SHEET_RESERVAS);
    sheet.appendRow([
      new Date(), datos.nombre, datos.email, datos.telefono || '',
      datos.servicio, datos.fecha, datos.hora, 'Confirmada', 'No', datos.sucursal
    ]);

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
    if (estado === 'Cancelada' || recordatorioEnviado === 'Sí') return;
    if (!fecha || !hora) return;

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
