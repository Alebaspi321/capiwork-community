/**
 * CAPIMEL / Capiwork
 * Versión 2 - Dashboard Personal consumiendo directamente DB_PER_CONSUMO
 *
 * IMPORTANTE:
 * - Si este Apps Script está vinculado al Google Sheet que contiene DB_PER_CONSUMO,
 *   deje SOURCE_SPREADSHEET_ID vacío.
 * - Si es un proyecto independiente, coloque en SOURCE_SPREADSHEET_ID el ID
 *   del Google Sheet que contiene la pestaña DB_PER_CONSUMO.
 */

const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

function getScriptSetting_(key, fallback) {
  const value = SCRIPT_PROPERTIES.getProperty(key);
  return value === null || value === '' ? fallback : value;
}

function getEmailListSetting_(key, fallback) {
  return getScriptSetting_(key, fallback.join(','))
    .split(',')
    .map(value => normalizeEmail_(value))
    .filter(Boolean);
}

const CONFIG = {
  APP_NAME: 'Capiwork',
  APP_SUBTITLE: 'Monitoreo, Evaluación y Aprendizaje',
  ADMIN_EMAIL: getScriptSetting_('ADMIN_EMAIL', 'admin@capiwork.example'),
  ADMIN_NAME: getScriptSetting_('ADMIN_NAME', 'Administrador de demostración'),
  ADMIN_TEMP_PASSWORD: getScriptSetting_('ADMIN_TEMP_PASSWORD', 'CambiarEstaClave2026!'),
  SUPER_ADMIN_EMAILS: getEmailListSetting_('SUPER_ADMIN_EMAILS', ['admin@capiwork.example']),

  // Si el proyecto está VINCULADO a la hoja de cálculo, deje esto vacío.
  SOURCE_SPREADSHEET_ID: getScriptSetting_('SOURCE_SPREADSHEET_ID', ''),
  PERSONAL_DB_SHEET: getScriptSetting_('PERSONAL_DB_SHEET', 'Personal Demo'),

  // Logos de Google Drive: pegue SOLO el ID del archivo.
  LOGO_LOGIN_FILE_ID: getScriptSetting_('LOGO_LOGIN_FILE_ID', ''),
  LOGO_ICON_FILE_ID: getScriptSetting_('LOGO_ICON_FILE_ID', ''),
  LOGO_CENTROS_FILE_ID: getScriptSetting_('LOGO_CENTROS_FILE_ID', ''),

  SESSION_HOURS: 8,
  AUTH_DB_NAME: 'CAPIMEL - Seguridad',
  TZ: 'America/La_Paz'
};

const AUTH_SHEETS = {
  USERS: 'Usuarios',
  SESSIONS: 'Sesiones',
  LOG: 'Log',
  CONFIG: 'Config'
};

const RESEARCH_CONFIG = {
  SPREADSHEET_ID: getScriptSetting_('RESEARCH_SPREADSHEET_ID', ''),
  SHEET_NAME: 'Investigaciones',
  SHEET_URL: getScriptSetting_('RESEARCH_SHEET_URL', '#'),
  FOLDER_URL: getScriptSetting_('RESEARCH_FOLDER_URL', '#')
};

const UNIFIED_SHEETS = {
  PROJECTS: 'Proyectos 1S-2026',
  PUBLICATIONS: 'Publicaciones 1S-2026',
  FELLOWS: 'Becarios 1S-2026',
  PERSONNEL: 'Personal 1S-2026',
  EVENTS: 'Eventos'
};

const PROGRAM_COUNTRIES = ['Bolivia', 'Guatemala', 'Colombia'];
const FELLOWS_AGREEMENTS_SHEET = 'Acuerdos becarios';

function isProgramCountry_(value) {
  const key = normalizeHeader_(value);
  return PROGRAM_COUNTRIES.some(country => normalizeHeader_(country) === key);
}

const PROJECT_BANK = {
  SPREADSHEET_ID: getScriptSetting_('PROJECT_SPREADSHEET_ID', ''),
  URL: getScriptSetting_('PROJECT_SHEET_URL', '#'),
  MASTER_SHEET: 'Iniciatives & projects',
  PROGRESS_SHEET: 'Ongoing',
  DISSEMINATION_SHEET: 'Disemination'
};

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.APP_NAME = CONFIG.APP_NAME;
  t.APP_SUBTITLE = CONFIG.APP_SUBTITLE;
  t.LOGO_LOGIN = getAssetDataUrl_(CONFIG.LOGO_LOGIN_FILE_ID, 'login');
  t.LOGO_ICON = getAssetDataUrl_(CONFIG.LOGO_ICON_FILE_ID, 'icon');
  t.LOGO_CENTROS = getAssetDataUrl_(CONFIG.LOGO_CENTROS_FILE_ID, 'centros');

  return t.evaluate()
    .setTitle(CONFIG.APP_NAME + ' | MEL')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* =========================================================
   CONFIGURACIÓN INICIAL
   ========================================================= */

function configurarSistema() {
  // 1. Verifica primero que exista la pestaña consolidada de Personal.
  const source = getSourceSpreadsheet_();
  const personalSheet = source.getSheetByName(CONFIG.PERSONAL_DB_SHEET);
  if (!personalSheet) {
    throw new Error('No se encontró la hoja "' + CONFIG.PERSONAL_DB_SHEET + '".');
  }

  // 2. Crea una base separada SOLO para usuarios/sesiones.
  const props = PropertiesService.getScriptProperties();
  let authId = props.getProperty('AUTH_DB_ID');
  let authSs;

  if (!authId) {
    authSs = SpreadsheetApp.create(CONFIG.AUTH_DB_NAME);
    authId = authSs.getId();
    props.setProperty('AUTH_DB_ID', authId);
  } else {
    authSs = SpreadsheetApp.openById(authId);
  }

  setupSheet_(authSs, AUTH_SHEETS.USERS, [
    'ID', 'Email', 'Nombre', 'Rol', 'Salt', 'PasswordHash',
    'Estado', 'DebeCambiar', 'Creado', 'Actualizado', 'UltimoAcceso'
  ]);

  setupSheet_(authSs, AUTH_SHEETS.SESSIONS, [
    'TokenHash', 'UserID', 'Email', 'Expira', 'Creado', 'Activo'
  ]);

  setupSheet_(authSs, AUTH_SHEETS.LOG, [
    'FechaHora', 'Usuario', 'Accion', 'Modulo', 'Detalle'
  ]);

  setupSheet_(authSs, AUTH_SHEETS.CONFIG, ['Clave', 'Valor']);

  const existing = findUserByEmail_(CONFIG.ADMIN_EMAIL);
  if (!existing) {
    const salt = Utilities.getUuid();
    authSs.getSheetByName(AUTH_SHEETS.USERS).appendRow([
      Utilities.getUuid(),
      normalizeEmail_(CONFIG.ADMIN_EMAIL),
      CONFIG.ADMIN_NAME,
      'ADMIN',
      salt,
      hashPassword_(CONFIG.ADMIN_TEMP_PASSWORD, salt),
      'ACTIVO',
      true,
      new Date(),
      new Date(),
      ''
    ]);
  }

  setConfigValue_('SOURCE_SPREADSHEET_ID', source.getId());
  setConfigValue_('PERSONAL_DB_SHEET', CONFIG.PERSONAL_DB_SHEET);
  setConfigValue_('CREATED_AT', new Date().toISOString());

  return {
    ok: true,
    message: 'CAPIMEL configurado correctamente.',
    sourceUrl: source.getUrl(),
    authUrl: authSs.getUrl(),
    personalSheet: CONFIG.PERSONAL_DB_SHEET
  };
}

function autorizarServiciosCapiwork() {
  getSourceSpreadsheet_().getName();
  SpreadsheetApp.openById(RESEARCH_CONFIG.SPREADSHEET_ID).getName();
  const quota = MailApp.getRemainingDailyQuota();
  return 'Servicios autorizados. Cuota diaria disponible: ' + quota;
}

function setupSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#0B2349')
      .setFontColor('#FFFFFF');
  }
}

function getSourceSpreadsheet_() {
  if (CONFIG.SOURCE_SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  const savedId = PropertiesService.getScriptProperties().getProperty('SOURCE_DB_ID');
  if (savedId) return SpreadsheetApp.openById(savedId);

  throw new Error(
    'No se pudo identificar la hoja consolidada. Revise SOURCE_SPREADSHEET_ID.'
  );
}

function getAuthDb_() {
  const id = PropertiesService.getScriptProperties().getProperty('AUTH_DB_ID');
  if (!id) throw new Error('Primero debe ejecutar configurarSistema().');
  return SpreadsheetApp.openById(id);
}

function setConfigValue_(key, value) {
  const sh = getAuthDb_().getSheetByName(AUTH_SHEETS.CONFIG);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

/* =========================================================
   LOGIN / SESIÓN
   ========================================================= */

function login(email, password) {
  email = normalizeEmail_(email);
  password = String(password || '');

  if (!email || !password) {
    return { ok: false, message: 'Ingrese correo electrónico y contraseña.' };
  }

  const user = findUserByEmail_(email);
  if (!user || user.Estado !== 'ACTIVO') {
    log_('NO_AUTENTICADO', 'LOGIN_FALLIDO', 'Acceso', 'Correo: ' + email);
    return { ok: false, message: 'Correo o contraseña incorrectos.' };
  }

  if (hashPassword_(password, user.Salt) !== user.PasswordHash) {
    log_('NO_AUTENTICADO', 'LOGIN_FALLIDO', 'Acceso', 'Correo: ' + email);
    return { ok: false, message: 'Correo o contraseña incorrectos.' };
  }

  const token = createSession_(user);
  updateUserLastAccess_(user.ID);
  log_(email, 'LOGIN', 'Acceso', 'Inicio de sesión correcto');

  return {
    ok: true,
    token: token,
    user: sanitizeUser_(user),
    mustChangePassword: bool_(user.DebeCambiar)
  };
}

function logout(token) {
  if (!token) return { ok: true };
  const sh = getAuthDb_().getSheetByName(AUTH_SHEETS.SESSIONS);
  const data = sh.getDataRange().getValues();
  const tokenHash = hashToken_(token);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === tokenHash) {
      sh.getRange(i + 1, 6).setValue(false);
      break;
    }
  }
  return { ok: true };
}

function validateSession(token) {
  try {
    const user = requireSession_(token);
    return { ok: true, user: sanitizeUser_(user) };
  } catch (err) {
    return { ok: false, message: 'La sesión venció. Inicie sesión nuevamente.' };
  }
}

function createSession_(user) {
  cleanExpiredSessions_();
  const rawToken = Utilities.getUuid() + '-' + Utilities.getUuid();
  const expires = new Date(Date.now() + CONFIG.SESSION_HOURS * 60 * 60 * 1000);

  getAuthDb_().getSheetByName(AUTH_SHEETS.SESSIONS).appendRow([
    hashToken_(rawToken), user.ID, user.Email, expires, new Date(), true
  ]);

  return rawToken;
}

function requireSession_(token, roles) {
  if (!token) throw new Error('Sesión no válida.');

  const sh = getAuthDb_().getSheetByName(AUTH_SHEETS.SESSIONS);
  const data = sh.getDataRange().getValues();
  const tokenHash = hashToken_(token);
  const now = new Date();
  let session = null;

  for (let i = 1; i < data.length; i++) {
    const expiry = new Date(data[i][3]);
    if (
      String(data[i][0]) === tokenHash &&
      bool_(data[i][5]) &&
      expiry > now
    ) {
      session = { UserID: String(data[i][1]), Email: String(data[i][2]) };
      break;
    }
  }

  if (!session) throw new Error('Sesión vencida o inválida.');

  const user = findUserById_(session.UserID);
  if (!user || user.Estado !== 'ACTIVO') throw new Error('Usuario sin acceso.');

  if (roles && roles.length && roles.indexOf(user.Rol) === -1) {
    throw new Error('No tiene permisos para realizar esta acción.');
  }
  return user;
}

function cleanExpiredSessions_() {
  const sh = getAuthDb_().getSheetByName(AUTH_SHEETS.SESSIONS);
  if (sh.getLastRow() <= 1) return;
  const data = sh.getDataRange().getValues();
  const now = new Date();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const expiry = new Date(data[i][3]);
    if (!bool_(data[i][5]) || expiry <= now) rows.push(i + 1);
  }
  rows.reverse().forEach(r => sh.deleteRow(r));
}

function changeMyPassword(token, currentPassword, newPassword) {
  const user = requireSession_(token);

  if (hashPassword_(String(currentPassword || ''), user.Salt) !== user.PasswordHash) {
    return { ok: false, message: 'La contraseña actual no es correcta.' };
  }

  const validation = validatePassword_(String(newPassword || ''));
  if (!validation.ok) return validation;

  const sh = getAuthDb_().getSheetByName(AUTH_SHEETS.USERS);
  const salt = Utilities.getUuid();

  sh.getRange(user._row, 5).setValue(salt);
  sh.getRange(user._row, 6).setValue(hashPassword_(newPassword, salt));
  sh.getRange(user._row, 8).setValue(false);
  sh.getRange(user._row, 10).setValue(new Date());

  log_(user.Email, 'CAMBIO_PASSWORD', 'Cuenta', 'Contraseña actualizada');
  return { ok: true, message: 'Contraseña actualizada correctamente.' };
}

function requestPasswordReset(email) {
  email = normalizeEmail_(email);
  const generic = { ok:true, message:'Si el correo está registrado, recibirá un código de recuperación.' };
  if (!email) return generic;
  const cache = CacheService.getScriptCache();
  const rateKey = 'pwd-request:' + hashToken_(email);
  if (cache.get(rateKey)) return generic;
  const user = findUserByEmail_(email);
  cache.put(rateKey, '1', 60);
  if (!user || user.Estado !== 'ACTIVO') return generic;
  const code = String(Math.floor(100000 + Math.random() * 900000));
  cache.put('pwd-code:' + hashToken_(email), hashToken_(code), 600);
  MailApp.sendEmail({
    to: email,
    name: CONFIG.APP_NAME,
    subject: 'Código para recuperar su contraseña de Capiwork',
    htmlBody: '<div style="font-family:Arial;color:#172033;max-width:520px"><h2 style="color:#092548">Recuperación de contraseña</h2><p>Su código de Capiwork es:</p><p style="font-size:30px;font-weight:800;letter-spacing:7px;color:#1769e0">' + code + '</p><p>Vence en 10 minutos. Si no solicitó este cambio, ignore el mensaje.</p></div>'
  });
  log_(email, 'SOLICITAR_RESET', 'Cuenta', 'Código de recuperación enviado');
  return generic;
}

function resetForgottenPassword(email, code, newPassword) {
  email = normalizeEmail_(email);
  code = String(code || '').trim();
  const user = findUserByEmail_(email);
  if (!user || user.Estado !== 'ACTIVO') return {ok:false, message:'El código no es válido o venció.'};
  const stored = CacheService.getScriptCache().get('pwd-code:' + hashToken_(email));
  if (!/^\d{6}$/.test(code) || !stored || stored !== hashToken_(code)) {
    return {ok:false, message:'El código no es válido o venció.'};
  }
  const validation = validatePassword_(String(newPassword || ''));
  if (!validation.ok) return validation;
  const salt = Utilities.getUuid();
  const usersSheet = getAuthDb_().getSheetByName(AUTH_SHEETS.USERS);
  usersSheet.getRange(user._row, 5).setValue(salt);
  usersSheet.getRange(user._row, 6).setValue(hashPassword_(newPassword, salt));
  usersSheet.getRange(user._row, 8).setValue(false);
  usersSheet.getRange(user._row, 10).setValue(new Date());
  invalidateUserSessions_(user.ID);
  CacheService.getScriptCache().remove('pwd-code:' + hashToken_(email));
  log_(email, 'RESET_PASSWORD_AUTOGESTION', 'Cuenta', 'Contraseña recuperada');
  return {ok:true, message:'Contraseña actualizada. Ya puede iniciar sesión.'};
}

function invalidateUserSessions_(userId) {
  const sh = getAuthDb_().getSheetByName(AUTH_SHEETS.SESSIONS);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]) === String(userId)) sh.getRange(i + 1, 6).setValue(false);
  }
}

/* =========================================================
   DASHBOARD PERSONAL: DB_PER_CONSUMO
   ========================================================= */

function getAppData(token) {
  const user = requireSession_(token);
  const personal = buildPersonalDashboard_();
  const unified = buildUnifiedOverview_();
  const projectBank = buildProjectBank_();
  unified.fellowWorks = linkFellowWorks_(unified.fellows, projectBank.projects);
  const dataQuality = buildDataQuality_(personal.rows, unified, projectBank.projects);

  return {
    ok: true,
    user: sanitizeUser_(user),
    personal: personal.rows,
    summary: personal.summary,
    charts: personal.charts,
    filters: personal.filters,
    unified: unified,
    projectBank: projectBank,
    dataQuality: dataQuality,
    fellowAgreements: user.Rol === 'ADMIN' ? getFellowAgreements_() : [],
    users: user.Rol === 'ADMIN' ? getUsers_() : []
  };
}

function buildDataQuality_(personal, unified, projectBank) {
  personal = personal || [];
  unified = unified || {};
  projectBank = projectBank || [];
  const publications = unified.publications || [];
  const projects = unified.projects || [];
  const fellows = unified.fellows || [];
  const events = unified.events || [];
  const makeItems = (rows, labelFn, checks) => rows.map(row => {
    const fields = checks.filter(check => check.missing(row)).map(check => ({key:check.key,label:check.label,type:check.type||'text',options:check.options||[]}));
    return fields.length ? {label:String(labelFn(row) || 'Registro sin nombre'), rowNumber:Number(row.rowNumber || row._rowNumber || 0), fields:fields} : null;
  }).filter(Boolean);
  const personalItems = makeItems(personal, x => x.nombreCompleto, [
    {key:'correo_contacto',label:'Correo',type:'email',missing:x=>!x.correo},{key:'pais_residencia',label:'País',type:'select',options:PROGRAM_COUNTRIES,missing:x=>!x.paisResidencia},{key:'institucion',label:'Institución',missing:x=>!x.institucion},{key:'flag_retirado',label:'Estado',type:'select',options:['Vigente','Retirado'],missing:x=>!x.estado}
  ]);
  const projectItems = makeItems(projects, x => x.proyecto || x.titulo, [
    {key:'equipo',label:'Equipo',missing:x=>!x.equipo},{key:'difusiones_n',label:'Cantidad de difusiones',type:'number',missing:x=>!x.difusiones_n || x.difusiones_n === '—'}
  ]);
  const publicationItems = makeItems(publications, x => x.titulo, [
    {key:'fecha',label:'Fecha',type:'date',missing:x=>!x.fecha},{key:'pais',label:'País',type:'select',options:PROGRAM_COUNTRIES,missing:x=>!x.pais},{key:'clasificacion',label:'Clasificación final',missing:x=>/revisar/i.test(x.clasificacion || '')},{key:'avance',label:'Avance de publicación',type:'range',missing:x=>!/^\d{1,3}%$/.test(String(x.avance||''))}
  ]);
  const fellowItems = makeItems(fellows, x => x.becario, [
    {key:'de_avance',label:'Porcentaje de avance',type:'range',missing:x=>x.de_avance === '' || typeof x.de_avance === 'undefined'}
  ]);
  const eventItems = makeItems(events, x => x.evento, [
    {key:'inicio',label:'Fecha de inicio',type:'date',missing:x=>!x.inicio},{key:'pais',label:'País',type:'select',options:PROGRAM_COUNTRIES,missing:x=>!x.pais},{key:'participantes',label:'Participantes',type:'number',missing:x=>!x.participantes}
  ]);
  const bankItems = makeItems(projectBank, x => x.name, [
    {key:'progress',label:'Avance válido (%)',type:'range',missing:x=>!/^\d{1,3}%$/.test(String(x.progress || '').trim())},{key:'status',label:'Estado',type:'select',options:['Logistics','Ethical approval process','Collecting data','Analysing','Disseminating','Publishing','Completed'],missing:x=>!x.status}
  ]);
  const issues = [
    {module:'Personal', total:personal.length, pending:personalItems.length, detail:'Correo, país, institución o estado sin completar',items:personalItems},
    {module:'Proyectos', total:projects.length, pending:projectItems.length, detail:'Equipo o cantidad de difusiones pendiente',items:projectItems},
    {module:'Publicaciones', total:publications.length, pending:publicationItems.length, detail:'Fecha, país, clasificación o avance pendiente',items:publicationItems},
    {module:'Becarios', total:fellows.length, pending:fellowItems.length, detail:'Porcentaje de avance pendiente',items:fellowItems},
    {module:'Eventos', total:events.length, pending:eventItems.length, detail:'Fecha, país o participantes pendiente',items:eventItems},
    {module:'Banco de proyectos', total:projectBank.length, pending:bankItems.length, errors:bankItems.length, detail:'Avance inválido o estado pendiente',items:bankItems}
  ];
  return {
    totalRecords: issues.reduce((n,x) => n + x.total, 0),
    pending: issues.reduce((n,x) => n + x.pending, 0),
    errors: issues.reduce((n,x) => n + (x.errors || 0), 0),
    modules: issues,
    validatedAt: formatDate_(new Date())
  };
}

function validateDataQuality(token) {
  requireSession_(token);
  const personal = buildPersonalDashboard_();
  const unified = buildUnifiedOverview_();
  const projectBank = buildProjectBank_();
  return {ok:true, dataQuality:buildDataQuality_(personal.rows, unified, projectBank.projects)};
}

function updateMissingData(token, payload) {
  const user = requireSession_(token, ['ADMIN']);
  payload = payload || {};
  const moduleName = String(payload.module || '');
  const rowNumber = Number(payload.rowNumber);
  const values = payload.values || {};
  if (!rowNumber || rowNumber < 2) throw new Error('Registro no válido.');
  const configs = {
    'Personal':{sheet:CONFIG.PERSONAL_DB_SHEET,allowed:['correo_contacto','pais_residencia','institucion','flag_retirado'],personal:true},
    'Proyectos':{sheet:UNIFIED_SHEETS.PROJECTS,allowed:['equipo','difusiones_n']},
    'Publicaciones':{sheet:UNIFIED_SHEETS.PUBLICATIONS,allowed:['fecha','pais','clasificacion','avance']},
    'Becarios':{sheet:UNIFIED_SHEETS.FELLOWS,allowed:['de_avance']},
    'Eventos':{sheet:UNIFIED_SHEETS.EVENTS,allowed:['inicio','pais','participantes']}
  };
  const lock = LockService.getScriptLock();lock.waitLock(30000);
  try {
    if (moduleName === 'Banco de proyectos') {
      const sh = SpreadsheetApp.openById(PROJECT_BANK.SPREADSHEET_ID).getSheetByName(PROJECT_BANK.PROGRESS_SHEET);
      if (!sh || rowNumber > sh.getLastRow()) throw new Error('Proyecto no encontrado.');
      if (Object.prototype.hasOwnProperty.call(values,'progress')) sh.getRange(rowNumber,8).setValue(normalizePercent_(values.progress));
      if (Object.prototype.hasOwnProperty.call(values,'status')) sh.getRange(rowNumber,9).setValue(sanitizeModuleValue_(values.status));
    } else {
      const cfg = configs[moduleName];if(!cfg) throw new Error('Módulo no válido.');
      const sh = getSourceSpreadsheet_().getSheetByName(cfg.sheet);if(!sh || rowNumber > sh.getLastRow()) throw new Error('Registro no encontrado.');
      const headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
      const keys = headers.map(h => cfg.personal ? canonicalPersonalHeader_(h) : normalizeHeader_(h));
      cfg.allowed.forEach(key => {
        if (!Object.prototype.hasOwnProperty.call(values,key)) return;
        const col = keys.indexOf(key);if(col < 0) throw new Error('No se encontró la columna ' + key + '.');
        let value = sanitizeModuleValue_(values[key]);if(key === 'de_avance'||key==='avance') value = normalizePercent_(value);
        sh.getRange(rowNumber,col+1).setValue(value);
      });
    }
    SpreadsheetApp.flush();
    log_(user.Email || user.email, 'COMPLETAR_DATO', moduleName, 'Fila ' + rowNumber);
    return {ok:true,message:'Datos guardados y validados correctamente.'};
  } finally { lock.releaseLock(); }
}

function updateRecordData(token, payload) {
  const user = requireSession_(token, ['ADMIN']);payload=payload||{};
  const kind=String(payload.kind||''),rowNumber=Number(payload.rowNumber),values=payload.values||{};
  if(!rowNumber||rowNumber<2)throw new Error('Registro no válido.');
  const configs={
    project:{sheet:UNIFIED_SHEETS.PROJECTS,allowed:['proyecto','titulo','responsable','equipo','estado_iniciativa','estado_actual','avance','difusiones_n','nuevo_vs_2s_2025']},
    publication:{sheet:UNIFIED_SHEETS.PUBLICATIONS,allowed:['titulo','clasificacion','subtipo','estado','fecha','pais','origen','avance']},
    fellow:{sheet:UNIFIED_SHEETS.FELLOWS,allowed:['becario','estado','pais','nivel','programa','institucion','fecha_admision','semestres_cursados','de_avance','lectura_del_avance']},
    event:{sheet:UNIFIED_SHEETS.EVENTS,allowed:['evento','tipo','inicio','fin','pais','ciudad','modalidad','participantes','organiza','nuevo']},
    personal:{sheet:CONFIG.PERSONAL_DB_SHEET,allowed:['nombre_completo','correo_contacto','cargo_actual','institucion','flag_retirado','pais_residencia','ciudad_residencia','genero','profesion','tipo_vinculacion','orcid'],personal:true},
    research:{sheet:RESEARCH_CONFIG.SHEET_NAME,allowed:['investigador_principal','investigadores_participantes','correo_electronico','whatsapp','paises','titulo','descripcion','area_de_salud','poblaciones','clasificacion_de_investigacion','enfoques_transversales'],research:true}
  };
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    if(kind==='bank'){
      const sh=SpreadsheetApp.openById(PROJECT_BANK.SPREADSHEET_ID).getSheetByName(PROJECT_BANK.PROGRESS_SHEET);if(!sh||rowNumber>sh.getLastRow())throw new Error('Proyecto no encontrado.');
      const columns={name:2,title:3,objectives:4,start:5,expectedEnd:6,finalDate:7,progress:8,status:9,duration:10,coordinator:11,protocolLead:12,authors:13,collaborators:17,context:18,funding:19,budget:20,publications:21,drafts:22,comments:23};
      Object.keys(values).forEach(key=>{if(!columns[key])return;let value=values[key];if(key==='progress')value=normalizePercent_(value);sh.getRange(rowNumber,columns[key]).setValue(sanitizeModuleValue_(value));});
      sh.getRange(rowNumber,24).setValue(new Date());sh.getRange(rowNumber,25).setValue(user.Email||user.email);
    }else{
      const cfg=configs[kind];if(!cfg)throw new Error('Tipo de ficha no válido.');
      const ss=cfg.research?SpreadsheetApp.openById(RESEARCH_CONFIG.SPREADSHEET_ID):getSourceSpreadsheet_();const sh=ss.getSheetByName(cfg.sheet);if(!sh||rowNumber>sh.getLastRow())throw new Error('Registro no encontrado.');
      const headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0],keys=headers.map(h=>cfg.personal?canonicalPersonalHeader_(h):normalizeHeader_(h));
      cfg.allowed.forEach(key=>{if(!Object.prototype.hasOwnProperty.call(values,key))return;const col=keys.indexOf(key);if(col<0)return;let value=values[key];if(key==='de_avance'||key==='avance')value=normalizePercent_(value);sh.getRange(rowNumber,col+1).setValue(sanitizeModuleValue_(value));});
    }
    SpreadsheetApp.flush();log_(user.Email||user.email,'EDITAR_FICHA',kind,'Fila '+rowNumber);return{ok:true,message:'Ficha actualizada correctamente.'};
  }finally{lock.releaseLock();}
}

function getFellowAgreementsSheet_() {
  const ss=getSourceSpreadsheet_();let sh=ss.getSheetByName(FELLOWS_AGREEMENTS_SHEET);
  const headers=['ID','Becario','Correo','Trabajo acordado','Fecha de entrega','Avisar con días','Estado','Último aviso','Observaciones','Creado por','Fecha de creación'];
  if(!sh){sh=ss.insertSheet(FELLOWS_AGREEMENTS_SHEET);sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1769E0').setFontColor('#FFFFFF');sh.setFrozenRows(1);}
  return sh;
}

function getFellowAgreements_(){
  const sh=getFellowAgreementsSheet_(),values=sh.getDataRange().getValues();if(values.length<2)return[];
  return values.slice(1).map((r,i)=>({rowNumber:i+2,id:String(r[0]||''),becario:String(r[1]||''),correo:String(r[2]||''),trabajo:String(r[3]||''),fechaEntrega:formatAgreementDate_(r[4]),diasAviso:Number(r[5])||7,estado:String(r[6]||'PENDIENTE'),ultimoAviso:String(r[7]||''),observaciones:String(r[8]||''),creadoPor:String(r[9]||'')})).filter(x=>x.becario);
}

function createFellowAgreement(token,payload){
  const user=requireSession_(token,['ADMIN']);payload=payload||{};const becario=clean_(payload.becario),trabajo=clean_(payload.trabajo),fecha=clean_(payload.fechaEntrega),dias=Math.max(1,Math.min(60,Number(payload.diasAviso)||7));
  if(!becario||!trabajo||!/^\d{4}-\d{2}-\d{2}$/.test(fecha))throw new Error('Complete becario, trabajo y fecha de entrega.');
  const person=buildPersonalDashboard_().rows.find(x=>normalizePersonMatch_(x.nombreCompleto)===normalizePersonMatch_(becario));if(!person||!person.correo)throw new Error('El becario no tiene correo registrado.');
  const sh=getFellowAgreementsSheet_();sh.appendRow([Utilities.getUuid(),becario,person.correo,sanitizeModuleValue_(trabajo),fecha,dias,'PENDIENTE','',sanitizeModuleValue_(payload.observaciones),user.Email||user.email,new Date()]);
  log_(user.Email||user.email,'CREAR_ACUERDO','Becarios',becario+' · '+fecha);return{ok:true,message:'Acuerdo programado correctamente.',agreements:getFellowAgreements_()};
}

function completeFellowAgreement(token,rowNumber){
  const user=requireSession_(token,['ADMIN']);const sh=getFellowAgreementsSheet_(),row=Number(rowNumber);if(row<2||row>sh.getLastRow())throw new Error('Acuerdo no encontrado.');sh.getRange(row,7).setValue('COMPLETADO');log_(user.Email||user.email,'COMPLETAR_ACUERDO','Becarios','Fila '+row);return{ok:true,message:'Acuerdo marcado como completado.',agreements:getFellowAgreements_()};
}

function installFellowReminderTrigger(token){
  requireSession_(token,['ADMIN']);ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='processFellowReminders').forEach(t=>ScriptApp.deleteTrigger(t));ScriptApp.newTrigger('processFellowReminders').timeBased().everyDays(1).atHour(8).create();return{ok:true,message:'Recordatorios automáticos activados para las 08:00 (hora de Bolivia).'};
}

function processFellowReminders(){
  const sh=getFellowAgreementsSheet_(),values=sh.getDataRange().getValues(),todayKey=Utilities.formatDate(new Date(),CONFIG.TZ,'yyyy-MM-dd'),today=new Date(todayKey+'T12:00:00');
  for(let i=1;i<values.length;i++){const r=values[i],estado=String(r[6]||'PENDIENTE').toUpperCase();if(estado!=='PENDIENTE')continue;const dueKey=formatAgreementDate_(r[4]);if(!dueKey)continue;const due=new Date(dueKey+'T12:00:00'),remaining=Math.round((due-today)/86400000),lead=Number(r[5])||7,noticeKey=todayKey+'|'+remaining;if(remaining>lead||remaining<0||String(r[7]||'')===noticeKey)continue;
    const becario=String(r[1]||''),correo=String(r[2]||''),trabajo=String(r[3]||''),subject=remaining===0?'Entrega acordada para hoy: '+trabajo:'Recordatorio de entrega: faltan '+remaining+' día(s)';
    const message='Hola '+becario+',\n\nTe recordamos que el trabajo acordado "'+trabajo+'" tiene fecha de entrega '+dueKey+'. '+(remaining===0?'La entrega corresponde a hoy.':'Faltan '+remaining+' día(s).')+'\n\nEste es un recordatorio automático de Capiwork MEL.';
    MailApp.sendEmail({to:correo,subject:subject,body:message,name:'Capiwork MEL'});MailApp.sendEmail({to:CONFIG.ADMIN_EMAIL,subject:'[Capiwork] '+subject,body:'Becario: '+becario+'\nCorreo: '+correo+'\nTrabajo: '+trabajo+'\nFecha: '+dueKey+'\nDías restantes: '+remaining,name:'Capiwork MEL'});sh.getRange(i+1,8).setValue(noticeKey);
  }
}

function formatAgreementDate_(value){if(!value)return'';if(value instanceof Date&&!isNaN(value))return Utilities.formatDate(value,CONFIG.TZ,'yyyy-MM-dd');const text=String(value).trim();const m=text.match(/^(\d{4})-(\d{2})-(\d{2})/);return m?m[0]:'';}

function getProjectBank(token) {
  requireSession_(token);
  return {ok:true, projectBank:buildProjectBank_(), refreshedAt:formatDate_(new Date())};
}

function updateProjectProgress(token, payload) {
  const user = requireSession_(token);
  if (String(user.Rol || '').toUpperCase() !== 'ADMIN') throw new Error('Solo administradores pueden actualizar avances.');
  payload = payload || {};
  const rowNumber = Number(payload.rowNumber);
  const progress = String(payload.progress || '').trim();
  const status = clean_(payload.status);
  const comments = clean_(payload.comments);
  if (!rowNumber || rowNumber < 2) throw new Error('Proyecto no válido.');
  if (!/^\d{1,3}%$/.test(progress)) throw new Error('Ingrese un avance entre 0% y 100%.');
  const numericProgress = Number(progress.replace('%',''));
  if (numericProgress < 0 || numericProgress > 100) throw new Error('El avance debe estar entre 0% y 100%.');
  if (!status) throw new Error('Seleccione el estado actual del proyecto.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = SpreadsheetApp.openById(PROJECT_BANK.SPREADSHEET_ID).getSheetByName(PROJECT_BANK.PROGRESS_SHEET);
    if (!sh || rowNumber > sh.getLastRow()) throw new Error('No se encontró el proyecto en Ongoing.');
    const projectName = String(sh.getRange(rowNumber, 2).getDisplayValue() || '').trim();
    if (!projectName) throw new Error('La fila seleccionada no contiene un proyecto.');
    sh.getRange(rowNumber, 8).setValue(progress);
    sh.getRange(rowNumber, 9).setValue(status);
    if (comments) sh.getRange(rowNumber, 23).setValue(comments);
    sh.getRange(rowNumber, 24).setValue(new Date());
    sh.getRange(rowNumber, 25).setValue(user.Email || user.email || 'Usuario Capiwork');
    SpreadsheetApp.flush();
    log_(user.Email || user.email, 'ACTUALIZAR_AVANCE', 'Banco de proyectos', projectName + ' · ' + progress + ' · ' + status);
    return {ok:true, message:'Avance actualizado correctamente.', projectBank:buildProjectBank_()};
  } finally {
    lock.releaseLock();
  }
}

function buildProjectBank_() {
  const ss = SpreadsheetApp.openById(PROJECT_BANK.SPREADSHEET_ID);
  const sh = ss.getSheetByName(PROJECT_BANK.PROGRESS_SHEET);
  if (!sh) return {projects:[],summary:{total:0},sourceUrl:PROJECT_BANK.URL};
  const values = sh.getDataRange().getDisplayValues();
  const projects = values.slice(1).map((r,i)=>({
    rowNumber:i+2, id:String(r[0]||''), name:String(r[1]||'').trim(), title:String(r[2]||'').trim(), objectives:String(r[3]||'').trim(),
    start:cleanProjectCell_(r[4]), expectedEnd:cleanProjectCell_(r[5]), finalDate:cleanProjectCell_(r[6]), progress:cleanProjectCell_(r[7]), status:cleanProjectCell_(r[8]), duration:cleanProjectCell_(r[9]),
    coordinator:String(r[10]||'').trim(), protocolLead:String(r[11]||'').trim(), authors:String(r[12]||'').trim(), settings:[r[13],r[14],r[15]].map(String).map(x=>x.trim()).filter(Boolean), collaborators:String(r[16]||'').trim(), context:String(r[17]||'').trim(), funding:String(r[18]||'').trim(), budget:String(r[19]||'').trim(), publications:String(r[20]||'').trim(), drafts:String(r[21]||'').trim(), comments:String(r[22]||'').trim(), updatedAt:String(r[23]||'').trim(), updatedBy:String(r[24]||'').trim()
  })).filter(x=>x.name);
  const validProgress = projects.map(x=>Number(String(x.progress).replace('%',''))).filter(n=>!isNaN(n));
  return {
    projects:projects,
    summary:{
      total:projects.length,
      completed:projects.filter(x=>/completed|publishing/i.test(x.status)).length,
      active:projects.filter(x=>!/^completed$/i.test(x.status)).length,
      average:validProgress.length ? Math.round(validProgress.reduce((a,b)=>a+b,0)/validProgress.length) : 0,
      publications:projects.reduce((n,x)=>n+(Number(x.publications)||0),0)
    },
    statuses:uniqueSorted_(projects.map(x=>x.status)),
    sourceUrl:PROJECT_BANK.URL
  };
}

function cleanProjectCell_(value) {
  const text = String(value || '').trim();
  return /^#(ERROR!|REF!|N\/A|VALUE!|DIV\/0!|NAME\?)$/i.test(text) ? '' : text;
}

function refreshPersonalDashboard(token) {
  requireSession_(token);
  const personal = buildPersonalDashboard_();
  return {
    ok: true,
    personal: personal.rows,
    summary: personal.summary,
    charts: personal.charts,
    filters: personal.filters,
    refreshedAt: formatDate_(new Date())
  };
}

function getResearchData(token) {
  requireSession_(token);
  const ss = SpreadsheetApp.openById(RESEARCH_CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(RESEARCH_CONFIG.SHEET_NAME);
  if (!sh) throw new Error('No se encontró la pestaña Investigaciones.');
  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return {ok:true, records:[], summary:{total:0,paises:0,areas:0,colaboraciones:0}, filters:{paises:[],areas:[],clasificaciones:[]}, sheetUrl:RESEARCH_CONFIG.SHEET_URL, folderUrl:RESEARCH_CONFIG.FOLDER_URL};
  const headers = values[0].map(normalizeHeader_);
  const index = {}; headers.forEach((h,i) => index[h]=i);
  const get = (row, name) => typeof index[name] === 'undefined' ? '' : String(row[index[name]] || '').trim();
  const records = values.slice(1).filter(row=>get(row,'investigador_principal')).map((row,pos)=>({
    id:String(pos+1), rowNumber:pos+2, codigo:'INV-'+String(pos+1).padStart(3,'0'),
    principal:get(row,'investigador_principal'),
    participantes:splitResearchList_(get(row,'investigadores_participantes')),
    correo:get(row,'correo_electronico'), whatsapp:get(row,'whatsapp'),
    paises:splitResearchList_(get(row,'paises')).filter(isProgramCountry_),
    titulo:get(row,'titulo'), descripcion:get(row,'descripcion'),
    areas:splitResearchList_(get(row,'area_de_salud')),
    poblaciones:splitResearchList_(get(row,'poblaciones')),
    clasificaciones:splitResearchList_(get(row,'clasificacion_de_investigacion')),
    enfoques:splitResearchList_(get(row,'enfoques_transversales')),
    estado:'Investigación registrada', avance:'', difusiones:'', nuevo:'',
    foto:get(row,'foto_principal'), fotosParticipantes:get(row,'fotos_participantes'),
    fechaRegistro:get(row,'fecha_registro'), orcid:get(row,'orcid_investigador_principal')
  })).filter(record=>record.paises.length);
  const allCountries=[].concat.apply([],records.map(r=>r.paises));
  const allAreas=[].concat.apply([],records.map(r=>r.areas));
  const allClass=[].concat.apply([],records.map(r=>r.clasificaciones));
  const allPopulations=[].concat.apply([],records.map(r=>r.poblaciones));
  const allApproaches=[].concat.apply([],records.map(r=>r.enfoques));
  return {
    ok:true,
    records:records,
    summary:{
      total:records.length,
      paises:uniqueCount_(allCountries),
      areas:uniqueCount_(allAreas),
      lineas:uniqueCount_(allAreas),
      colaboraciones:records.filter(r=>r.paises.length>1).length
    },
    filters:{paises:uniqueSorted_(allCountries),areas:uniqueSorted_(allAreas),poblaciones:uniqueSorted_(allPopulations),clasificaciones:uniqueSorted_(allClass),enfoques:uniqueSorted_(allApproaches)},
    sheetUrl:RESEARCH_CONFIG.SHEET_URL,
    folderUrl:RESEARCH_CONFIG.FOLDER_URL
  };
}

function splitResearchList_(value) {
  return String(value || '').split(/\s*\|\s*|\s*;\s*/).map(x=>x.trim()).filter(Boolean);
}

function linkFellowWorks_(fellows, projects) {
  return (fellows || []).map(fellow => {
    const name = normalizePersonMatch_(fellow.becario);
    const tokens = name.split(' ').filter(x=>x.length>2);
    return (projects || []).filter(project => {
      const people = normalizePersonMatch_([project.coordinator, project.protocolLead, project.authors].join(' '));
      return name && (people.indexOf(name) >= 0 || (tokens.length >= 2 && tokens.every(t=>people.indexOf(t)>=0)));
    });
  });
}

function normalizePersonMatch_(value) {
  return String(value || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]+/g,' ').trim();
}

function buildUnifiedOverview_() {
  const ss = getSourceSpreadsheet_();
  const projects = readUnifiedRows_(ss, UNIFIED_SHEETS.PROJECTS, row => /^P\d+$/i.test(row.id));
  const publications = readUnifiedRows_(ss, UNIFIED_SHEETS.PUBLICATIONS, row => /^\d+$/.test(row._key));
  const fellows = readUnifiedRows_(ss, UNIFIED_SHEETS.FELLOWS, row => /^\d+$/.test(row._key) && isProgramCountry_(row.pais));
  const events = readUnifiedRows_(ss, UNIFIED_SHEETS.EVENTS, row => /^\d+$/.test(row.id) && isProgramCountry_(row.pais));
  return {
    sourceUrl: RESEARCH_CONFIG.SHEET_URL,
    projects: projects,
    publications: publications,
    fellows: fellows,
    events: events,
    summary: {
      projects: projects.length,
      publications: publications.length,
      fellows: fellows.length,
      events: events.length,
      countries: uniqueCount_([].concat(fellows.map(r=>r.pais), events.map(r=>r.pais)).filter(Boolean))
    }
  };
}

function readUnifiedRows_(ss, sheetName, predicate) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  const headers = values[0].map(normalizeHeader_);
  return values.slice(1).map((row,rowIndex) => {
    const item = {_key:String(row[0] || '').trim(), _rowNumber:rowIndex+2};
    headers.forEach((h,i) => item[h === '' ? '_' : h] = String(row[i] || '').trim());
    return item;
  }).filter(predicate || (item => Object.keys(item).some(k => item[k])));
}

function submitModuleRecord(token, moduleType, payload) {
  const user = requireSession_(token);
  const role = String(user.Rol || user.rol || '').toUpperCase();
  if (role !== 'ADMIN') throw new Error('Solo administradores pueden registrar información.');
  const configs = {
    project:{sheet:UNIFIED_SHEETS.PROJECTS,idHeader:'id',prefix:'P'},
    publication:{sheet:UNIFIED_SHEETS.PUBLICATIONS,idHeader:'',prefix:''},
    fellow:{sheet:UNIFIED_SHEETS.FELLOWS,idHeader:'',prefix:''},
    event:{sheet:UNIFIED_SHEETS.EVENTS,idHeader:'id',prefix:''}
  };
  const cfg = configs[String(moduleType || '')];
  if (!cfg) throw new Error('Tipo de formulario no válido.');
  payload = payload || {};
  const required = {
    project:['proyecto','titulo','responsable'],
    publication:['titulo','clasificacion','estado'],
    fellow:['becario','estado','pais','nivel','programa','institucion'],
    event:['evento','tipo','inicio','pais','modalidad','organiza']
  }[moduleType];
  required.forEach(k => { if (!String(payload[k] || '').trim()) throw new Error('Complete todos los campos obligatorios.'); });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = getSourceSpreadsheet_().getSheetByName(cfg.sheet);
    if (!sh) throw new Error('No se encontró la pestaña ' + cfg.sheet + '.');
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
    const keys = headers.map(normalizeHeader_);
    const firstKey = keys[0];
    const nextNumber = nextModuleNumber_(sh, cfg.prefix);
    const data = Object.assign({}, payload);
    data[firstKey] = cfg.prefix + nextNumber;
    if (moduleType === 'event') {
      data.id = String(nextNumber);
      data.duracion = calculateDurationDays_(data.inicio, data.fin);
    }
    if (moduleType === 'fellow') data.de_avance = normalizePercent_(data.de_avance);
    if (moduleType === 'project') data.ongoing_estado_avance = [data.estado_avance, normalizePercent_(data.avance)].filter(Boolean).join(' · ');
    const row = keys.map(k => sanitizeModuleValue_(data[k]));
    sh.getRange(sh.getLastRow()+1,1,1,row.length).setValues([row]);
    log_(user.Email || user.email, 'CREAR_REGISTRO', cfg.sheet, cfg.prefix + nextNumber);
    return {ok:true,id:cfg.prefix + nextNumber,message:'Registro guardado correctamente en ' + cfg.sheet + '.'};
  } finally { lock.releaseLock(); }
}

function nextModuleNumber_(sheet, prefix) {
  const last = sheet.getLastRow();
  if (last < 2) return 1;
  const values = sheet.getRange(2,1,last-1,1).getDisplayValues().flat();
  return values.reduce((max,v) => Math.max(max, Number(String(v).replace(/[^0-9]/g,'')) || 0), 0) + 1;
}

function calculateDurationDays_(start, end) {
  if (!start || !end) return '';
  const a = new Date(start), b = new Date(end);
  if (isNaN(a.getTime()) || isNaN(b.getTime()) || b < a) return '';
  return Math.floor((b-a)/86400000)+1;
}

function normalizePercent_(value) {
  if (value === '' || value === null || typeof value === 'undefined') return '';
  const n = Math.max(0,Math.min(100,Number(String(value).replace('%','')) || 0));
  return n + '%';
}

function sanitizeModuleValue_(value) {
  const text = Array.isArray(value) ? value.join(', ') : String(value == null ? '' : value).trim();
  if (/^[=+@]/.test(text)) return "'" + text;
  return text.slice(0,2000);
}

function buildPersonalDashboard_() {
  const source = getSourceSpreadsheet_();
  const sh = source.getSheetByName(CONFIG.PERSONAL_DB_SHEET);
  if (!sh) throw new Error('No se encontró la hoja Personal 1S-2026.');

  const range = sh.getDataRange();
  const values = range.getValues();
  if (!values || values.length < 2) {
    return emptyPersonalDashboard_();
  }

  const headers = values[0].map(h => canonicalPersonalHeader_(h));
  const required = [
    'id','nombres','apellidos','correo_contacto','pais_residencia',
    'genero','profesion','institucion','cargo_actual','tipo_vinculacion',
    'fecha_inicio_programa','flag_retirado','anos_antiguedad'
  ];

  // No detiene el sistema si alguna columna opcional falta.
  const index = {};
  headers.forEach((h, i) => index[h] = i);

  const rawRows = values.slice(1).filter(r =>
    r.some(v => v !== '' && v !== null && typeof v !== 'undefined')
  );

  const rows = rawRows.map((r, pos) => {
    const nombre = val_(r, index, 'nombres');
    const apellidos = val_(r, index, 'apellidos');
    const nombreCompletoFuente = val_(r, index, 'nombre_completo');
    const statusRaw = val_(r, index, 'estado') || val_(r, index, 'flag_retirado');
    const status = normalizeStatus_(statusRaw);

    return {
      id: String(val_(r, index, '#') || val_(r, index, 'id') || (pos + 1)),
      rowNumber: pos + 2,
      nombres: String(nombre || ''),
      apellidos: String(apellidos || ''),
      nombreCompleto: String(nombreCompletoFuente || [nombre, apellidos].filter(Boolean).join(' ')).trim(),
      correo: String(val_(r, index, 'correo_contacto') || val_(r, index, 'correo_electronico') || ''),
      carnetIdentidad: String(val_(r, index, 'carnet_identidad') || val_(r, index, 'numero_carnet') || val_(r, index, 'cedula_identidad') || val_(r, index, 'documento_identidad') || val_(r, index, 'numero_documento_identidad') || val_(r, index, 'ci') || ''),
      pasaporte: String(val_(r, index, 'numero_pasaporte') || val_(r, index, 'pasaporte') || val_(r, index, 'passport') || ''),
      paisNacimiento: String(val_(r, index, 'pais_nacimiento') || ''),
      paisResidencia: String(val_(r, index, 'pais') || val_(r, index, 'pais_residencia') || ''),
      ciudadResidencia: String(val_(r, index, 'ciudad') || val_(r, index, 'ciudad_residencia') || ''),
      genero: String(val_(r, index, 'genero') || 'Sin dato'),
      grupoEtnico: String(val_(r, index, 'grupo_etnico') || ''),
      grupoIndigena: String(val_(r, index, 'grupo_indigena') || ''),
      profesion: String(val_(r,index,'profesion') || 'Sin dato'),
      nivelEducativo: String(val_(r, index, 'nivel_educativo') || ''),
      tipoFormacion: String(val_(r, index, 'tipo_formacion_na') || ''),
      institucion: normalizeInstitution_(val_(r, index, 'institucion')),
      cargo: String(val_(r,index,'cargo/rol') || val_(r,index,'cargo_rol') || buildCargo_(r, index)),
      vinculacion: String(val_(r,index,'vinculacion') || buildVinculacion_(r, index)),
      orcid: String(val_(r, index, 'orcid') || ''),
      telefono: String(val_(r, index, 'telefono') || val_(r, index, 'numero_telefonico') || ''),
      estado: status,
      fechaInicio: formatSheetDate_(val_(r, index, 'fecha_inicio_programa')),
      fechaRetiro: formatSheetDate_(val_(r, index, 'fecha_retiro')),
      anosAntiguedad: safeNumber_(val_(r, index, 'anos_antiguedad')),
      mesesAntiguedad: safeNumber_(val_(r, index, 'meses_antiguedad')),
      diasAntiguedad: safeNumber_(val_(r, index, 'dias_antiguedad')),
      tipoMovimiento: String(val_(r, index, 'tipo_movimiento') || '')
    };
  });

  // Una persona puede tener varios IDs por movimientos históricos.
  // Se consolida primero por correo y, si no existe, por nombre completo.
  // Como se recorre en orden de la hoja, prevalece el registro más reciente.
  const uniqueMap = {};
  rows.forEach(r => {
    const email = String(r.correo || '').trim().toLowerCase();
    const normalizedName = String(r.nombreCompleto || '').trim().toLowerCase();
    const hasUsableEmail = email && email !== 'anonymous' && email !== 'anónimo' && email !== 'anonimo';
    const key = hasUsableEmail ? 'MAIL:' + email : (normalizedName ? 'NAME:' + normalizedName : 'ID:' + r.id);
    uniqueMap[key] = r;
  });
  const uniqueRows = Object.keys(uniqueMap).map(k => uniqueMap[k]);

  // Las filas históricas originales se conservan tal como fueron cargadas.
  // Los movimientos nuevos de actualización o retiro reemplazan la versión
  // anterior para los indicadores, evitando aumentar el número de personas.
  const currentRows = [];
  const latestIndex = {};
  rows.forEach(r => {
    const email = String(r.correo || '').trim().toLowerCase();
    const key = email && email !== 'anonymous' && email !== 'anónimo' && email !== 'anonimo'
      ? 'MAIL:' + email
      : '';
    const movement = normalizeHeader_(r.tipoMovimiento);
    const replacesPrevious = movement === 'actualizacion' || movement === 'retiro';
    if (replacesPrevious && key && typeof latestIndex[key] !== 'undefined') {
      currentRows[latestIndex[key]] = r;
    } else {
      currentRows.push(r);
      if (key) latestIndex[key] = currentRows.length - 1;
    }
  });

  const programRows = currentRows.filter(r => isProgramCountry_(r.paisResidencia));

  const rowsWithInstitution = programRows.filter(r => isValidInstitution_(r.institucion));
  const summary = {
    total: programRows.length,
    responses: rows.length,
    vigentes: programRows.filter(r => r.estado !== 'Retirado').length,
    retirados: programRows.filter(r => r.estado === 'Retirado').length,
    paises: uniqueCount_(programRows.map(r => r.paisResidencia)),
    instituciones: uniqueCount_(rowsWithInstitution.map(r => r.institucion))
  };

  const charts = {
    porPais: countTop_(programRows, 'paisResidencia', 10),
    porGenero: countTop_(programRows, 'genero', 10),
    porInstitucion: countTop_(rowsWithInstitution, 'institucion', 10),
    porCargo: countTop_(programRows, 'cargo', 10),
    porEstado: countTop_(programRows, 'estado', 10)
  };

  const filters = {
    paises: uniqueSorted_(programRows.map(r => r.paisResidencia)),
    instituciones: uniqueSorted_(rowsWithInstitution.map(r => r.institucion)),
    generos: uniqueSorted_(programRows.map(r => r.genero)),
    estados: uniqueSorted_(programRows.map(r => r.estado))
  };

  return {
    rows: programRows.sort((a,b) => a.nombreCompleto.localeCompare(b.nombreCompleto)),
    summary: summary,
    charts: charts,
    filters: filters
  };
}

function emptyPersonalDashboard_() {
  return {
    rows: [],
    summary: { total:0, vigentes:0, retirados:0, paises:0, instituciones:0 },
    charts: { porPais:[], porGenero:[], porInstitucion:[], porCargo:[], porEstado:[] },
    filters: { paises:[], instituciones:[], generos:[], estados:[] }
  };
}

function canonicalPersonalHeader_(value) {
  const key = normalizeHeader_(value).replace(/[()]/g,'').replace(/\//g,'_').replace(/_+/g,'_');
  const aliases = {
    id_maestro:'id', correo_de_contacto:'correo_contacto', pais_de_nacimiento:'pais_nacimiento',
    fecha_de_nacimiento:'fecha_nacimiento', pais_de_residencia:'pais_residencia',
    ciudad_de_residencia:'ciudad_residencia', pueblo_indigena:'grupo_indigena',
    profesion_reportada:'profesion', profesion_normalizada:'profesion_normalizada',
    formacion_otorgada:'tipo_formacion_na', fecha_de_inicio_en_nihr:'fecha_inicio_programa',
    antiguedad_anos:'anos_antiguedad', cargo_rol_actual:'cargo_actual', otro_cargo_rol:'otro_cargo',
    tipo_de_vinculacion:'tipo_vinculacion', estado_en_el_programa:'flag_retirado',
    fecha_de_retiro:'fecha_retiro', tipo_de_movimiento:'tipo_movimiento',
    motivo_de_retiro:'motivo_retiro', observaciones_de_movimiento:'observaciones_movimiento',
    fecha_actualizacion_movimiento:'fecha_actualizacion', consentimiento_de_datos:'autorizacion_datos',
    numero_telefonico:'telefono', telefono_celular:'telefono', carnet_de_identidad:'carnet_identidad',
    numero_de_carnet:'numero_carnet', numero_de_carnet_de_identidad:'numero_carnet', cedula:'cedula_identidad',
    cedula_de_identidad:'cedula_identidad', documento_de_identidad:'documento_identidad',
    numero_de_documento_de_identidad:'numero_documento_identidad', numero_de_pasaporte:'numero_pasaporte'
  };
  return aliases[key] || key;
}

function normalizeHeader_(header) {
  return String(header || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function val_(row, index, name) {
  if (typeof index[name] === 'undefined') return '';
  const v = row[index[name]];
  return (v === null || typeof v === 'undefined') ? '' : v;
}

function normalizeStatus_(value) {
  const v = String(value || '').trim().toLowerCase();
  if (!v) return 'Sin dato';
  if (v.indexOf('retir') >= 0) return 'Retirado';
  if (v.indexOf('vigent') >= 0 || v.indexOf('activ') >= 0) return 'Vigente';
  return titleCase_(String(value));
}

function buildProfession_(row, index) {
  const principal = String(val_(row, index, 'profesion') || '');
  const otra = String(val_(row, index, 'otra_profesion') || '');
  if (/otro/i.test(principal) && otra) return otra;
  return principal || otra || 'Sin dato';
}

function buildCargo_(row, index) {
  const principal = String(val_(row, index, 'cargo_actual') || '');
  const otro = String(val_(row, index, 'otro_cargo') || '');
  if (/otro/i.test(principal) && otro) return otro;
  return principal || otro || 'Sin dato';
}

function buildVinculacion_(row, index) {
  const principal = String(val_(row, index, 'tipo_vinculacion') || '');
  const otra = String(val_(row, index, 'otra_vinculacion') || '');
  if (/otro/i.test(principal) && otra) return otra;
  return principal || otra || 'Sin dato';
}

function normalizeInstitution_(value) {
  const institution = String(value || '').trim();
  return isValidInstitution_(institution) ? institution : '';
}

function isValidInstitution_(value) {
  const normalized = normalizeHeader_(value);
  return Boolean(normalized) &&
    normalized !== 'si' &&
    normalized !== 'no' &&
    normalized !== 'sin_dato' &&
    normalized !== 'n_a' &&
    normalized !== 'na';
}

function countTop_(rows, field, limit) {
  const counts = {};
  rows.forEach(r => {
    const key = String(r[field] || 'Sin dato').trim() || 'Sin dato';
    counts[key] = (counts[key] || 0) + 1;
  });
  return Object.keys(counts)
    .map(k => ({ label:k, value:counts[k] }))
    .sort((a,b) => b.value - a.value || a.label.localeCompare(b.label))
    .slice(0, limit || 10);
}

function uniqueSorted_(arr) {
  return Array.from(new Set(
    arr.map(x => String(x || '').trim()).filter(Boolean)
  )).sort((a,b) => a.localeCompare(b));
}

function uniqueCount_(arr) {
  return uniqueSorted_(arr).length;
}

function safeNumber_(value) {
  if (value === '' || value === null || typeof value === 'undefined') return null;
  if (value instanceof Date) return null;
  const n = Number(value);
  return isNaN(n) ? null : n;
}

function formatSheetDate_(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, CONFIG.TZ, 'dd/MM/yyyy');
  }

  // Si viene como número serial de Excel / Sheets.
  if (typeof value === 'number' && value > 1000) {
    const date = new Date(Date.UTC(1899, 11, 30) + value * 86400000);
    return Utilities.formatDate(date, CONFIG.TZ, 'dd/MM/yyyy');
  }

  const parsed = new Date(value);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, CONFIG.TZ, 'dd/MM/yyyy');
  }
  return String(value);
}

/* =========================================================
   USUARIOS Y PERMISOS
   ========================================================= */

function isSuperAdmin_(user){return Boolean(user)&&CONFIG.SUPER_ADMIN_EMAILS.indexOf(normalizeEmail_(user.Email||user.email))>=0;}
function requireSuperAdmin_(token){const user=requireSession_(token,['ADMIN']);if(!isSuperAdmin_(user))throw new Error('Solo los superadministradores pueden administrar roles y accesos.');return user;}

function createUser(token, payload) {
  const admin = requireSuperAdmin_(token);
  payload = payload || {};

  const email = normalizeEmail_(payload.email);
  const nombre = clean_(payload.nombre);
  const rol = String(payload.rol || 'VIEWER').toUpperCase();
  const password = String(payload.password || '');

  if (!email || !nombre) return { ok:false, message:'Ingrese nombre y correo.' };
  if (['ADMIN','EDITOR','VIEWER'].indexOf(rol) === -1) {
    return { ok:false, message:'Rol no válido.' };
  }
  if (findUserByEmail_(email)) {
    return { ok:false, message:'Ese correo ya tiene una cuenta.' };
  }

  const validation = validatePassword_(password);
  if (!validation.ok) return validation;

  const salt = Utilities.getUuid();
  getAuthDb_().getSheetByName(AUTH_SHEETS.USERS).appendRow([
    Utilities.getUuid(), email, nombre, rol, salt,
    hashPassword_(password, salt),
    'ACTIVO', true, new Date(), new Date(), ''
  ]);

  log_(admin.Email, 'CREAR_USUARIO', 'Usuarios', email);
  return { ok:true, message:'Usuario creado.', users:getUsers_() };
}

function toggleUserStatus(token, userId) {
  const admin = requireSuperAdmin_(token);
  const target = findUserById_(userId);

  if (!target) return { ok:false, message:'Usuario no encontrado.' };
  if (isSuperAdmin_(target)) return {ok:false,message:'Las cuentas principales de superadministración no pueden desactivarse.'};
  if (target.ID === admin.ID) {
    return { ok:false, message:'No puede desactivar su propia cuenta.' };
  }

  const sh = getAuthDb_().getSheetByName(AUTH_SHEETS.USERS);
  const newStatus = target.Estado === 'ACTIVO' ? 'INACTIVO' : 'ACTIVO';
  sh.getRange(target._row, 7).setValue(newStatus);
  sh.getRange(target._row, 10).setValue(new Date());

  log_(admin.Email, 'CAMBIAR_ESTADO', 'Usuarios', target.Email + ' -> ' + newStatus);
  return { ok:true, message:'Estado actualizado.', users:getUsers_() };
}

function resetUserPassword(token, userId, newPassword) {
  const admin = requireSuperAdmin_(token);
  const target = findUserById_(userId);
  if (!target) return { ok:false, message:'Usuario no encontrado.' };

  const validation = validatePassword_(String(newPassword || ''));
  if (!validation.ok) return validation;

  const salt = Utilities.getUuid();
  const sh = getAuthDb_().getSheetByName(AUTH_SHEETS.USERS);
  sh.getRange(target._row, 5).setValue(salt);
  sh.getRange(target._row, 6).setValue(hashPassword_(newPassword, salt));
  sh.getRange(target._row, 8).setValue(true);
  sh.getRange(target._row, 10).setValue(new Date());

  log_(admin.Email, 'RESET_PASSWORD', 'Usuarios', target.Email);
  return { ok:true, message:'Contraseña temporal actualizada.' };
}

function updateUserRole(token,userId,newRole){
  const admin=requireSuperAdmin_(token),target=findUserById_(userId),role=String(newRole||'').toUpperCase();if(!target)return{ok:false,message:'Usuario no encontrado.'};if(['ADMIN','EDITOR','VIEWER'].indexOf(role)<0)return{ok:false,message:'Rol no válido.'};if(isSuperAdmin_(target)&&role!=='ADMIN')return{ok:false,message:'Las cuentas de superadministración deben conservar el rol de administrador principal.'};const sh=getAuthDb_().getSheetByName(AUTH_SHEETS.USERS);sh.getRange(target._row,4).setValue(role);sh.getRange(target._row,10).setValue(new Date());log_(admin.Email,'CAMBIAR_ROL','Usuarios',target.Email+' -> '+role);return{ok:true,message:'Rol actualizado correctamente.',users:getUsers_()};
}

function getUsers_() {
  const sh = getAuthDb_().getSheetByName(AUTH_SHEETS.USERS);
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];

  return values.slice(1).map(r => ({
    id: String(r[0] || ''),
    email: String(r[1] || ''),
    nombre: String(r[2] || ''),
    rol: String(r[3] || ''),
    estado: String(r[6] || ''),
    debeCambiar: bool_(r[7]),
    creado: formatDate_(r[8]),
    ultimoAcceso: formatDate_(r[10])
  })).sort((a,b) => a.nombre.localeCompare(b.nombre));
}

/* =========================================================
   HELPERS USUARIOS / LOG
   ========================================================= */

function findUserByEmail_(email) {
  email = normalizeEmail_(email);
  const sh = getAuthDb_().getSheetByName(AUTH_SHEETS.USERS);
  const data = sh.getDataRange().getValues();

  for (let i=1; i<data.length; i++) {
    if (normalizeEmail_(data[i][1]) === email) return mapUser_(data[i], i+1);
  }
  return null;
}

function findUserById_(id) {
  const sh = getAuthDb_().getSheetByName(AUTH_SHEETS.USERS);
  const data = sh.getDataRange().getValues();

  for (let i=1; i<data.length; i++) {
    if (String(data[i][0]) === String(id)) return mapUser_(data[i], i+1);
  }
  return null;
}

function mapUser_(r, row) {
  return {
    _row: row,
    ID: String(r[0] || ''),
    Email: String(r[1] || ''),
    Nombre: String(r[2] || ''),
    Rol: String(r[3] || ''),
    Salt: String(r[4] || ''),
    PasswordHash: String(r[5] || ''),
    Estado: String(r[6] || ''),
    DebeCambiar: r[7],
    Creado: r[8],
    Actualizado: r[9],
    UltimoAcceso: r[10]
  };
}

function sanitizeUser_(u) {
  return {
    id:u.ID, email:u.Email, nombre:u.Nombre,
    rol:u.Rol, estado:u.Estado, debeCambiar:bool_(u.DebeCambiar), superAdmin:isSuperAdmin_(u)
  };
}

function updateUserLastAccess_(id) {
  const u = findUserById_(id);
  if (!u) return;
  getAuthDb_().getSheetByName(AUTH_SHEETS.USERS)
    .getRange(u._row, 11).setValue(new Date());
}

function log_(usuario, accion, modulo, detalle) {
  try {
    getAuthDb_().getSheetByName(AUTH_SHEETS.LOG)
      .appendRow([new Date(), usuario, accion, modulo, detalle]);
  } catch (err) {
    console.log(err);
  }
}

/* =========================================================
   UTILIDADES
   ========================================================= */

function normalizeEmail_(v) {
  return String(v || '').trim().toLowerCase();
}

function clean_(v) {
  return String(v || '').trim();
}

function bool_(v) {
  return v === true || String(v).toLowerCase() === 'true';
}

function validatePassword_(password) {
  if (password.length < 8) {
    return { ok:false, message:'La nueva contraseña debe tener al menos 8 caracteres.' };
  }
  if (!/[A-Za-z]/.test(password) || !/\d/.test(password)) {
    return { ok:false, message:'Use al menos una letra y un número.' };
  }
  return { ok:true };
}

function hashPassword_(password, salt) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + ':' + String(password),
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(bytes);
}

function hashToken_(token) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(token),
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(bytes);
}

function formatDate_(value) {
  if (!value) return '';
  try {
    return Utilities.formatDate(new Date(value), CONFIG.TZ, 'dd/MM/yyyy HH:mm');
  } catch (e) {
    return String(value);
  }
}

function titleCase_(text) {
  return String(text || '').toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
}

function getAssetDataUrl_(fileId, type) {
  // Los logos oficiales están incorporados en Assets.gs para evitar
  // imágenes rotas por permisos o sesiones de Google Drive.
  if (typeof EMBEDDED_ASSETS !== 'undefined' && EMBEDDED_ASSETS[type]) {
    return EMBEDDED_ASSETS[type];
  }
  if (fileId) {
    try {
      const blob = DriveApp.getFileById(fileId).getBlob();
      const mime = blob.getContentType() || 'image/svg+xml';
      return 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes());
    } catch (err) {
      console.log(err);
    }
  }
  return fallbackLogo_(type);
}

function fallbackLogo_(type) {
  const iconOnly = type === 'icon';
  const width = iconOnly ? 120 : 520;
  const height = iconOnly ? 120 : 160;
  const text = iconOnly ? '' :
    '<text x="155" y="92" font-family="Arial,sans-serif" font-size="46" font-weight="700" fill="#0B2349">Capi</text>' +
    '<text x="260" y="92" font-family="Arial,sans-serif" font-size="46" font-weight="700" fill="#3B82F6">work</text>';

  const svg =
    '<svg xmlns="http://www.w3.org/2000/svg" width="'+width+'" height="'+height+'" viewBox="0 0 '+width+' '+height+'">' +
    '<rect width="100%" height="100%" fill="white"/>' +
    '<g transform="translate(18,18)">' +
    '<rect x="12" y="16" rx="30" width="108" height="92" fill="#D6A66C"/>' +
    '<circle cx="44" cy="23" r="18" fill="#C88D52"/>' +
    '<circle cx="92" cy="23" r="18" fill="#C88D52"/>' +
    '<circle cx="52" cy="55" r="5" fill="#0B2349"/>' +
    '<ellipse cx="88" cy="72" rx="22" ry="15" fill="#E7C18E"/>' +
    '<circle cx="95" cy="70" r="5" fill="#0B2349"/>' +
    '<rect x="44" y="88" width="95" height="48" rx="7" fill="#2F6FED"/>' +
    '</g>' + text + '</svg>';

  return 'data:image/svg+xml;base64,' + Utilities.base64Encode(svg);
}
