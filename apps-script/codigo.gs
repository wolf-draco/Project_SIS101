const NOMBRE_HOJA_PRODUCTOS = 'Productos';
const NOMBRE_HOJA_COMBOS = 'Combos';
const NOMBRE_HOJA_CATEGORIAS = 'Categorias';
const NOMBRE_HOJA_VENTAS = 'Ventas';
const NOMBRE_HOJA_USUARIOS = 'Usuarios SIS101';
const ROL_ADMINISTRADOR = 'ADMINISTRADOR';
const ROL_EMPLEADO = 'EMPLEADO';
const ITERACIONES_CLAVE = 15000;
const ITERACIONES_CLAVE_ANTERIOR = 120000;
const PREFIJO_HASH_CLAVE = 'v2$';
const DURACION_SESION_SEGUNDOS = 21600;
const MAX_INTENTOS_LOGIN = 5;
const BLOQUEO_LOGIN_SEGUNDOS = 900;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('WebApp')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('Almacén | SIS101');
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Almacén')
    .addItem('📊 Ir al Dashboard', 'irAlDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('Productos')
      .addItem('Agregar Producto', 'abrirDialogoAgregarProducto')
      .addItem('Modificar Producto', 'abrirDialogoModificarProducto')
      .addItem('Eliminar Producto', 'abrirDialogoEliminarProducto')
      .addItem('Ver Productos', 'menuVerProductos'))
    .addSubMenu(ui.createMenu('Combos')
      .addItem('Agregar Combo', 'abrirDialogoAgregarCombo')
      .addItem('Modificar Combo', 'abrirDialogoModificarCombo')
      .addItem('Ver Combos', 'menuVerCombos'))
    .addSubMenu(ui.createMenu('Categorías')
      .addItem('Agregar Categoría', 'menuAgregarCategoria')
      .addItem('Eliminar Categoría', 'abrirDialogoEliminarCategoria')
      .addItem('Ver Categorías', 'menuVerCategorias')
      .addSeparator()
      .addItem('🛡️ Bloquear Edición Manual', 'configurarProteccionCategorias'))
    .addSubMenu(ui.createMenu('Cuentas de la aplicación')
      .addItem('Configurar administrador inicial', 'configurarAdministradorInicial')
      .addItem('Restablecer clave de administrador', 'restablecerClaveAdministrador'))
    .addSeparator()
    .addItem('🔍 Verificar Integridad', 'menuVerificarIntegridad')
    .addSeparator()
    .addItem('⚙️ Instalar/Reinstalar Dashboard', 'instalarDashboard')
    .addItem('⚠️ Reiniciar Datos de Prueba', 'menuReiniciarDatos')
    .addToUi();
}

function configurarAdministradorInicial() {
  const ui = SpreadsheetApp.getUi();
  const correoActivo = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const correoEfectivo = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if (!correoActivo || correoActivo !== correoEfectivo) {
    throw new Error('Configura el administrador desde la hoja de cálculo con la cuenta propietaria.');
  }

  const sheet = obtenerHojaUsuarios_(true);
  if (sheet.getLastRow() > 1) {
    ui.alert('Ya existen cuentas. Inicia sesión como administrador o usa la opción para restablecer su clave.');
    return;
  }

  const usernameResponse = ui.prompt('Administrador inicial', 'Nombre de usuario (3 a 32 caracteres):', ui.ButtonSet.OK_CANCEL);
  if (usernameResponse.getSelectedButton() !== ui.Button.OK) return;
  const username = normalizarUsuario_(usernameResponse.getResponseText());
  if (!esUsuarioValido_(username)) {
    ui.alert('El usuario debe tener 3 a 32 caracteres y usar solo letras, números, punto, guion o guion bajo.');
    return;
  }

  const passwordResponse = ui.prompt('Administrador inicial', 'Contraseña (mínimo 8 caracteres):', ui.ButtonSet.OK_CANCEL);
  if (passwordResponse.getSelectedButton() !== ui.Button.OK) return;
  const password = passwordResponse.getResponseText();
  const confirmResponse = ui.prompt('Administrador inicial', 'Confirma la contraseña:', ui.ButtonSet.OK_CANCEL);
  if (confirmResponse.getSelectedButton() !== ui.Button.OK) return;
  if (password !== confirmResponse.getResponseText() || !esClaveValida_(password)) {
    ui.alert('Las contraseñas no coinciden o no cumplen el mínimo de 8 caracteres.');
    return;
  }

  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty('SIS101_AUTH_PEPPER')) {
    properties.setProperty('SIS101_AUTH_PEPPER', Utilities.getUuid() + Utilities.getUuid());
  }
  const credentials = crearCredencial_(password);
  conLock_(() => {
    if (obtenerHojaUsuarios_(false).getLastRow() > 1) throw new Error('Ya se crearon cuentas mientras se configuraba el administrador.');
    guardarUsuario_(sheet, username, ROL_ADMINISTRADOR, credentials, true);
    properties.setProperty('SIS101_ADMIN_SHEET_EMAIL', correoActivo);
    sheet.hideSheet();
  });
  ui.alert('Administrador configurado. Abre la aplicación web e inicia sesión con estas credenciales.');
}

function restablecerClaveAdministrador() {
  requireAdministradorDeHoja_();
  const ui = SpreadsheetApp.getUi();
  const passwordResponse = ui.prompt('Restablecer administrador', 'Nueva contraseña (mínimo 8 caracteres):', ui.ButtonSet.OK_CANCEL);
  if (passwordResponse.getSelectedButton() !== ui.Button.OK) return;
  const password = passwordResponse.getResponseText();
  if (!esClaveValida_(password)) {
    ui.alert('La contraseña debe tener entre 8 y 128 caracteres.');
    return;
  }

  const confirmResponse = ui.prompt('Restablecer administrador', 'Confirma la nueva contraseña:', ui.ButtonSet.OK_CANCEL);
  if (confirmResponse.getSelectedButton() !== ui.Button.OK) return;
  if (password !== confirmResponse.getResponseText()) {
    ui.alert('Las contraseñas no coinciden.');
    return;
  }

  const sheet = obtenerHojaUsuarios_(false);
  const admin = obtenerUsuarioPorRol_(sheet, ROL_ADMINISTRADOR);
  if (!admin) throw new Error('No existe una cuenta de administrador para restablecer.');
  actualizarClaveUsuario_(sheet, admin.row, crearCredencial_(password));
  ui.alert('La contraseña del administrador se restableció.');
}

function obtenerHojaUsuarios_(crear) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(NOMBRE_HOJA_USUARIOS);
  if (!sheet && crear) {
    sheet = ss.insertSheet(NOMBRE_HOJA_USUARIOS);
    sheet.appendRow(['Usuario', 'Rol', 'Sal', 'Hash de clave', 'Activo', 'Creado']);
  }
  if (!sheet) throw new Error('No se encontró la hoja de usuarios.');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Usuario', 'Rol', 'Sal', 'Hash de clave', 'Activo', 'Creado']);
  } else {
    const encabezados = sheet.getRange(1, 1, 1, 6).getDisplayValues()[0];
    if (encabezados.join('|') !== 'Usuario|Rol|Sal|Hash de clave|Activo|Creado') {
      throw new Error(`La hoja "${NOMBRE_HOJA_USUARIOS}" tiene encabezados inesperados.`);
    }
  }
  return sheet;
}

function normalizarUsuario_(username) {
  return String(username || '').trim().toLowerCase();
}

function esUsuarioValido_(username) {
  return /^[a-z0-9._-]{3,32}$/.test(username);
}

function esClaveValida_(password) {
  return typeof password === 'string' && password.length >= 8 && password.length <= 128;
}

function obtenerUsuarios_(sheet) {
  if (sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues()
    .map((row, index) => ({
      usuario: normalizarUsuario_(row[0]),
      rol: String(row[1] || ''),
      sal: String(row[2] || ''),
      hash: String(row[3] || ''),
      activo: row[4] === true || String(row[4]).toLowerCase() === 'true',
      creado: String(row[5] || ''),
      row: index + 2
    }));
}

function obtenerUsuarioPorRol_(sheet, role) {
  return obtenerUsuarios_(sheet).find(user => user.rol === role) || null;
}

function obtenerUsuario_(sheet, username) {
  const normalizado = normalizarUsuario_(username);
  return obtenerUsuarios_(sheet).find(user => user.usuario === normalizado) || null;
}

function crearCredencial_(password) {
  const salt = Utilities.getUuid() + Utilities.getUuid();
  const hash = derivarHashClave_(password, salt, ITERACIONES_CLAVE);
  return { salt, hash: `${PREFIJO_HASH_CLAVE}${ITERACIONES_CLAVE}$${hash}` };
}

function guardarUsuario_(sheet, username, role, credentials, active) {
  sheet.appendRow([username, role, credentials.salt, credentials.hash, Boolean(active), new Date().toISOString()]);
  return username;
}

function actualizarClaveUsuario_(sheet, row, credentials) {
  sheet.getRange(row, 3, 1, 2).setValues([[credentials.salt, credentials.hash]]);
}

function leerHashClave_(storedHash) {
  const hashVersionado = String(storedHash || '').match(/^v2\$(\d+)\$([a-f0-9]{64})$/);
  if (hashVersionado) {
    const iterations = Number(hashVersionado[1]);
    if (Number.isSafeInteger(iterations) && iterations > 0 && iterations <= ITERACIONES_CLAVE_ANTERIOR) {
      return { iterations, hash: hashVersionado[2], versionado: true };
    }
  }
  if (/^[a-f0-9]{64}$/.test(String(storedHash || ''))) {
    return { iterations: ITERACIONES_CLAVE_ANTERIOR, hash: String(storedHash), versionado: false };
  }
  return { iterations: ITERACIONES_CLAVE, hash: '', versionado: false };
}

function derivarHashClave_(password, salt, iterations) {
  const pepper = PropertiesService.getScriptProperties().getProperty('SIS101_AUTH_PEPPER');
  if (!pepper) throw new Error('Falta la configuración de autenticación. Inicializa el acceso desde Sheets.');
  const key = Utilities.newBlob(String(password) + pepper).getBytes();
  const saltBytes = Utilities.newBlob(String(salt)).getBytes().concat([0, 0, 0, 1]);
  let u = Utilities.computeHmacSha256Signature(saltBytes, key);
  const result = u.slice();
  for (let iteration = 1; iteration < iterations; iteration += 1) {
    u = Utilities.computeHmacSha256Signature(u, key);
    for (let index = 0; index < result.length; index += 1) result[index] ^= u[index];
  }
  return result.map(byte => ((byte + 256) % 256).toString(16).padStart(2, '0')).join('');
}

function migrarHashUsuario_(sheet, user, password) {
  return conLock_(() => {
    const latestUser = obtenerUsuario_(sheet, user.usuario);
    if (!latestUser || !latestUser.activo || latestUser.rol !== user.rol) return null;
    if (latestUser.hash !== user.hash) {
      const latestHash = leerHashClave_(latestUser.hash);
      if (latestHash.versionado
        && latestHash.iterations === ITERACIONES_CLAVE
        && hashesIguales_(derivarHashClave_(password, latestUser.sal, latestHash.iterations), latestHash.hash)) {
        return latestUser.hash;
      }
      return null;
    }
    const credentials = crearCredencial_(password);
    actualizarClaveUsuario_(sheet, latestUser.row, credentials);
    return credentials.hash;
  });
}

function hashesIguales_(left, right) {
  if (left.length !== right.length) return false;
  let difference = 0;
  for (let index = 0; index < left.length; index += 1) difference |= left.charCodeAt(index) ^ right.charCodeAt(index);
  return difference === 0;
}

function claveDeSesion_(token) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token), Utilities.Charset.UTF_8);
  return 'sis101-session-' + Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '');
}

function requireSesion_(token, roles) {
  const cleanToken = String(token || '');
  if (!cleanToken) throw new Error('Inicia sesión para continuar.');
  const cache = CacheService.getScriptCache();
  const sessionJson = cache.get(claveDeSesion_(cleanToken));
  if (!sessionJson) throw new Error('La sesión venció. Inicia sesión nuevamente.');
  const session = JSON.parse(sessionJson);
  const sheet = obtenerHojaUsuarios_(false);
  const user = obtenerUsuario_(sheet, session.usuario);
  if (!user || !user.activo || user.rol !== session.rol || user.hash !== session.hash) {
    cache.remove(claveDeSesion_(cleanToken));
    throw new Error('La cuenta está desactivada. Inicia sesión nuevamente.');
  }
  if (roles && !roles.includes(user.rol)) throw new Error('No tienes permiso para esta operación.');
  return { usuario: user.usuario, rol: user.rol };
}

function requireAdministradorDeHoja_() {
  const activeEmail = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const adminEmail = String(PropertiesService.getScriptProperties().getProperty('SIS101_ADMIN_SHEET_EMAIL') || '').trim().toLowerCase();
  if (!activeEmail || !adminEmail || activeEmail !== adminEmail) {
    throw new Error('Esta operación solo está disponible para el administrador desde la hoja de cálculo.');
  }
}

function apiIniciarSesion(username, password) {
  const cleanUsername = normalizarUsuario_(username);
  if (!esUsuarioValido_(cleanUsername) || typeof password !== 'string' || password.length > 128) {
    return { ok: false, mensaje: 'Usuario o contraseña incorrectos.' };
  }
  if (!PropertiesService.getScriptProperties().getProperty('SIS101_AUTH_PEPPER')
    || !SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_USUARIOS)) {
    return { ok: false, mensaje: 'El acceso aún no está configurado. Pide al administrador que lo inicialice desde Sheets.' };
  }
  const cache = CacheService.getScriptCache();
  const failureKey = claveDeSesion_('login:' + cleanUsername);
  const failures = Number(cache.get(failureKey) || 0);
  if (failures >= MAX_INTENTOS_LOGIN) {
    return { ok: false, mensaje: 'Demasiados intentos. Espera 15 minutos antes de volver a probar.' };
  }

  const sheet = obtenerHojaUsuarios_(false);
  const user = obtenerUsuario_(sheet, cleanUsername);
  const salt = user ? user.sal : 'sis101-login-salt';
  const storedHash = leerHashClave_(user ? user.hash : '');
  const candidate = derivarHashClave_(password, salt, storedHash.iterations);
  if (!user || !user.activo || !hashesIguales_(candidate, storedHash.hash)) {
    cache.put(failureKey, String(failures + 1), BLOQUEO_LOGIN_SEGUNDOS);
    return { ok: false, mensaje: 'Usuario o contraseña incorrectos.' };
  }

  let sessionHash = user.hash;
  if (!storedHash.versionado || storedHash.iterations !== ITERACIONES_CLAVE) {
    sessionHash = migrarHashUsuario_(sheet, user, password);
    if (!sessionHash) return { ok: false, mensaje: 'La cuenta cambió. Inicia sesión nuevamente.' };
  }

  cache.remove(failureKey);
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put(claveDeSesion_(token), JSON.stringify({ usuario: user.usuario, rol: user.rol, hash: sessionHash }), DURACION_SESION_SEGUNDOS);
  return { ok: true, token, usuario: user.usuario, rol: user.rol };
}

function apiCerrarSesion(token) {
  if (token) CacheService.getScriptCache().remove(claveDeSesion_(token));
  return { ok: true };
}

function apiListarCuentas(token) {
  requireSesion_(token, [ROL_ADMINISTRADOR]);
  const sheet = obtenerHojaUsuarios_(false);
  return obtenerUsuarios_(sheet).map(user => ({
    usuario: user.usuario,
    rol: user.rol,
    activo: user.activo,
    creado: user.creado
  }));
}

function apiGuardarEnlacePago(token, url) {
  requireSesion_(token, [ROL_ADMINISTRADOR]);
  const paymentUrl = String(url || '').trim();
  if (paymentUrl.length > 2048 || (paymentUrl && !/^https?:\/\/\S+$/i.test(paymentUrl))) {
    return { ok: false, mensaje: 'La URL de destino del QR debe ser http o https y no superar 2048 caracteres.' };
  }
  PropertiesService.getScriptProperties().setProperty('SIS101_PAYMENT_QR_URL', paymentUrl);
  return { ok: true, enlacePagoUrl: paymentUrl, qrPagoUrl: paymentUrl };
}

function apiGuardarQrPago(token, url) {
  return apiGuardarEnlacePago(token, url);
}

function apiCrearCuentaEmpleado(token, username, password) {
  requireSesion_(token, [ROL_ADMINISTRADOR]);
  const cleanUsername = normalizarUsuario_(username);
  if (!esUsuarioValido_(cleanUsername)) return { ok: false, mensaje: 'El usuario debe tener 3 a 32 caracteres y usar letras, números, punto, guion o guion bajo.' };
  if (!esClaveValida_(password)) return { ok: false, mensaje: 'La contraseña debe tener entre 8 y 128 caracteres.' };
  const credentials = crearCredencial_(password);
  return conLock_(function () {
    const sheet = obtenerHojaUsuarios_(false);
    const users = obtenerUsuarios_(sheet);
    if (users.some(user => user.usuario === cleanUsername)) return { ok: false, mensaje: 'Ese usuario ya existe.' };
    if (users.filter(user => user.rol === ROL_EMPLEADO && user.activo).length >= 3) {
      return { ok: false, mensaje: 'Ya hay tres cuentas de empleados activas. Desactiva una antes de crear otra.' };
    }
    guardarUsuario_(sheet, cleanUsername, ROL_EMPLEADO, credentials, true);
    return { ok: true, cuentas: apiListarCuentas(token) };
  });
}

function apiCambiarEstadoCuentaEmpleado(token, username, active) {
  requireSesion_(token, [ROL_ADMINISTRADOR]);
  const cleanUsername = normalizarUsuario_(username);
  return conLock_(function () {
    const sheet = obtenerHojaUsuarios_(false);
    const user = obtenerUsuario_(sheet, cleanUsername);
    if (!user || user.rol !== ROL_EMPLEADO) return { ok: false, mensaje: 'No se encontró una cuenta de empleado.' };
    if (active && obtenerUsuarios_(sheet).filter(item => item.rol === ROL_EMPLEADO && item.activo).length >= 3 && !user.activo) {
      return { ok: false, mensaje: 'Ya hay tres cuentas de empleados activas.' };
    }
    sheet.getRange(user.row, 5).setValue(Boolean(active));
    return { ok: true, cuentas: apiListarCuentas(token) };
  });
}

function apiEliminarCuentaEmpleado(token, username) {
  requireSesion_(token, [ROL_ADMINISTRADOR]);
  const cleanUsername = normalizarUsuario_(username);
  return conLock_(function () {
    const sheet = obtenerHojaUsuarios_(false);
    const user = obtenerUsuario_(sheet, cleanUsername);
    if (!user || user.rol !== ROL_EMPLEADO) return { ok: false, mensaje: 'No se encontró una cuenta de empleado.' };
    sheet.deleteRow(user.row);
    return { ok: true, cuentas: apiListarCuentas(token) };
  });
}

function apiRestablecerClaveEmpleado(token, username, password) {
  requireSesion_(token, [ROL_ADMINISTRADOR]);
  const cleanUsername = normalizarUsuario_(username);
  if (!esClaveValida_(password)) return { ok: false, mensaje: 'La contraseña debe tener entre 8 y 128 caracteres.' };
  const credentials = crearCredencial_(password);
  return conLock_(function () {
    const sheet = obtenerHojaUsuarios_(false);
    const user = obtenerUsuario_(sheet, cleanUsername);
    if (!user || user.rol !== ROL_EMPLEADO) return { ok: false, mensaje: 'No se encontró una cuenta de empleado.' };
    actualizarClaveUsuario_(sheet, user.row, credentials);
    return { ok: true, mensaje: 'Contraseña actualizada.' };
  });
}

function apiCambiarClavePropia(token, currentPassword, newPassword) {
  const session = requireSesion_(token);
  if (!esClaveValida_(newPassword)) return { ok: false, mensaje: 'La nueva contraseña debe tener entre 8 y 128 caracteres.' };
  const sheet = obtenerHojaUsuarios_(false);
  const user = obtenerUsuario_(sheet, session.usuario);
  const storedHash = leerHashClave_(user ? user.hash : '');
  if (!user || !hashesIguales_(derivarHashClave_(String(currentPassword || ''), user.sal, storedHash.iterations), storedHash.hash)) {
    return { ok: false, mensaje: 'La contraseña actual no es correcta.' };
  }
  const credentials = crearCredencial_(newPassword);
  return conLock_(function () {
    const currentUser = obtenerUsuario_(sheet, session.usuario);
    if (!currentUser || currentUser.hash !== user.hash || !currentUser.activo) {
      return { ok: false, mensaje: 'La cuenta cambió. Inicia sesión de nuevo y vuelve a intentar.' };
    }
    actualizarClaveUsuario_(sheet, currentUser.row, credentials);
    CacheService.getScriptCache().remove(claveDeSesion_(token));
    return { ok: true, mensaje: 'Contraseña actualizada. Inicia sesión con la nueva clave.' };
  });
}

function cargarAlmacenDesdeSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const almacen = new Almacen();

  const hojaCategorias = ss.getSheetByName(NOMBRE_HOJA_CATEGORIAS);
  if (hojaCategorias && hojaCategorias.getLastRow() > 1) {
    const datos = hojaCategorias.getRange(2, 1, hojaCategorias.getLastRow() - 1, 3).getValues();
    almacen.categorias = datos.map(fila => ({
      nombre: fila[0],
      letra: fila[1],
      ultimo: Number(fila[2])
    }));
  } else {
    const backupJson = PropertiesService.getScriptProperties().getProperty('categorias_backup');
    if (backupJson) {
      try {
        const parsed = JSON.parse(backupJson);
        if (Array.isArray(parsed) && parsed.length > 0) almacen.categorias = parsed;
      } catch (e) { }
    }
  }

  const hojaProductos = ss.getSheetByName(NOMBRE_HOJA_PRODUCTOS);
  if (hojaProductos && hojaProductos.getLastRow() > 1) {
    const datos = hojaProductos.getRange(2, 1, hojaProductos.getLastRow() - 1, 9).getValues();
    almacen.productos = datos.map(fila => {
      const p = new Producto(fila[1], fila[2], fila[3], fila[4], fila[5], fila[6]);
      p.id = fila[0];
      p.vasosSueltos = Number(fila[7]);
      p.imagen = String(fila[8] || '');
      return p;
    });
  }

  const hojaCombos = ss.getSheetByName(NOMBRE_HOJA_COMBOS);
  if (hojaCombos && hojaCombos.getLastRow() > 1) {
    const datos = hojaCombos.getRange(2, 1, hojaCombos.getLastRow() - 1, 4).getValues();
    almacen.combos = datos.map(fila => {
      const nombre = fila[1];
      const precio = fila[2];
      const texto = fila[3];
      const componentes = [];
      const rotos = [];
      texto.split(' + ').forEach(parte => {
        const partes = parte.split(' x ');
        const cantStr = partes[0];
        const nombreProd = partes.slice(1).join(' x ');
        const producto = almacen.productos.find(p => p.nombre === nombreProd);
        if (producto) componentes.push({ producto, cantidad: Number(cantStr) });
        else rotos.push(nombreProd);
      });
      const combo = new Combo(nombre, precio, componentes);
      combo.id = fila[0];
      combo.componentesRotos = rotos;
      combo._textoOriginal = texto;
      return combo;
    });
  }

  if (hojaProductos && hojaProductos.getLastRow() > 1) {
    const ids = hojaProductos.getRange(2, 1, hojaProductos.getLastRow() - 1, 1).getValues();
    ids.forEach(fila => {
      const id = String(fila[0] || '');
      const m = id.match(/^([A-Z]+)(\d+)$/);
      if (m) {
        const cat = almacen.categorias.find(c => c.letra === m[1]);
        if (cat && parseInt(m[2]) > cat.ultimo) cat.ultimo = parseInt(m[2]);
      }
    });
  }

  if (hojaCombos && hojaCombos.getLastRow() > 1) {
    const ids = hojaCombos.getRange(2, 1, hojaCombos.getLastRow() - 1, 1).getValues();
    ids.forEach(fila => {
      const m = String(fila[0] || '').match(/^K(\d+)$/);
      if (m && parseInt(m[1]) > almacen.ultimoCombo) almacen.ultimoCombo = parseInt(m[1]);
    });
  }

  return almacen;
}

function guardarAlmacenEnSheet_(almacen) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let hojaCategorias = ss.getSheetByName(NOMBRE_HOJA_CATEGORIAS);
  if (!hojaCategorias) hojaCategorias = ss.insertSheet(NOMBRE_HOJA_CATEGORIAS);
  hojaCategorias.clear();
  hojaCategorias.appendRow(['Nombre', 'Letra', 'Último ID']);
  const filasCat = almacen.categorias.map(c => [c.nombre, c.letra, c.ultimo]);
  if (filasCat.length > 0) hojaCategorias.getRange(2, 1, filasCat.length, 3).setValues(filasCat);
  PropertiesService.getScriptProperties().setProperty(
    'categorias_backup',
    JSON.stringify(almacen.categorias.map(c => ({ nombre: c.nombre, letra: c.letra, ultimo: c.ultimo })))
  );

  let hojaProductos = ss.getSheetByName(NOMBRE_HOJA_PRODUCTOS);
  if (!hojaProductos) hojaProductos = ss.insertSheet(NOMBRE_HOJA_PRODUCTOS);
  hojaProductos.clear();
  hojaProductos.appendRow(['ID', 'Nombre', 'Categoría', 'Precio', 'Stock', 'Porciones', 'Precio Vaso', 'Vasos Sueltos', 'URL Imagen']);
  const filasProd = almacen.aFilasProductos();
  if (filasProd.length > 0) hojaProductos.getRange(2, 1, filasProd.length, 9).setValues(filasProd);
  if (filasProd.length > 0) {
    hojaProductos.getRange(2, 4, filasProd.length, 1).setNumberFormat('#,##0.00" Bs"');
    hojaProductos.getRange(2, 7, filasProd.length, 1).setNumberFormat('#,##0.00" Bs"');
    hojaProductos.getRange(2, 5, filasProd.length, 1).setHorizontalAlignment('center');
  }

  let hojaCombos = ss.getSheetByName(NOMBRE_HOJA_COMBOS);
  if (!hojaCombos) hojaCombos = ss.insertSheet(NOMBRE_HOJA_COMBOS);
  hojaCombos.clear();
  hojaCombos.appendRow(['ID', 'Nombre', 'Precio', 'Componentes']);
  const filasCombos = almacen.aFilasCombos();
  if (filasCombos.length > 0) hojaCombos.getRange(2, 1, filasCombos.length, 4).setValues(filasCombos);
  if (filasCombos.length > 0) hojaCombos.getRange(2, 3, filasCombos.length, 1).setNumberFormat('#,##0.00" Bs"');

  estilizarHojasDatos_([NOMBRE_HOJA_CATEGORIAS, NOMBRE_HOJA_PRODUCTOS, NOMBRE_HOJA_COMBOS]);
}

function estilizarTabla_(hoja, columnas, colorPestana, anchos) {
  const filas = Math.max(1, hoja.getLastRow());
  const tabla = hoja.getRange(1, 1, filas, columnas);
  hoja.setHiddenGridlines(true);
  hoja.setFrozenRows(1);
  hoja.setTabColor(colorPestana);
  hoja.getBandings().forEach(banda => banda.remove());

  tabla.setBackground('#FFFFFF');
  tabla.setFontFamily('Arial').setFontSize(10).setFontColor('#263244').setVerticalAlignment('middle');
  hoja.getRange(1, 1, 1, columnas)
    .setBackground('#1E5EFF')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('left')
    .setWrap(true);
  hoja.setRowHeight(1, 34);
  if (filas > 1) {
    hoja.setRowHeights(2, filas - 1, 30);
    for (let inicio = 2; inicio <= filas; inicio += 500) {
      const cantidad = Math.min(500, filas - inicio + 1);
      const fondos = Array.from({ length: cantidad }, (_, indice) => {
        const color = (inicio + indice) % 2 === 0 ? '#FFFFFF' : '#F3F6FB';
        return Array(columnas).fill(color);
      });
      hoja.getRange(inicio, 1, cantidad, columnas).setBackgrounds(fondos);
    }
    hoja.getRange(1, 1, filas, columnas)
      .setBorder(true, true, true, true, true, true, '#E2E8F0', SpreadsheetApp.BorderStyle.SOLID);
  }
  anchos.forEach((ancho, indice) => hoja.setColumnWidth(indice + 1, ancho));
  const filtro = hoja.getFilter();
  if (filtro) filtro.remove();
  if (filas > 1) hoja.getRange(1, 1, filas, columnas).createFilter();
}

function estilizarFilaNueva_(hoja, fila, columnas, colorPestana) {
  const rango = hoja.getRange(fila, 1, 1, columnas);
  rango.setBackground(fila % 2 === 0 ? '#FFFFFF' : '#F3F6FB')
    .setFontFamily('Arial')
    .setFontSize(10)
    .setFontColor('#263244')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#E2E8F0', SpreadsheetApp.BorderStyle.SOLID);
  hoja.setRowHeight(fila, 30);
  hoja.setTabColor(colorPestana);

  const filtro = hoja.getFilter();
  if (filtro) filtro.remove();
  hoja.getRange(1, 1, fila, columnas).createFilter();
}

function ajustarAlineacionHoja_(hoja) {
  const filasDatos = hoja.getLastRow() - 1;
  if (filasDatos < 1) return;

  const nombre = hoja.getName();
  if (nombre === NOMBRE_HOJA_PRODUCTOS) {
    hoja.getRange(2, 1, filasDatos, 1).setFontColor('#1746B0').setFontWeight('bold').setHorizontalAlignment('center');
    hoja.getRange(2, 4, filasDatos, 1).setHorizontalAlignment('right');
    hoja.getRange(2, 5, filasDatos, 4).setHorizontalAlignment('center');
  } else if (nombre === NOMBRE_HOJA_COMBOS) {
    hoja.getRange(2, 1, filasDatos, 1).setFontColor('#1746B0').setFontWeight('bold').setHorizontalAlignment('center');
    hoja.getRange(2, 3, filasDatos, 1).setHorizontalAlignment('right');
    hoja.getRange(2, 4, filasDatos, 1).setWrap(true);
  } else if (nombre === NOMBRE_HOJA_CATEGORIAS) {
    hoja.getRange(2, 2, filasDatos, 2).setHorizontalAlignment('center');
    hoja.getRange(2, 1, filasDatos, 1).setFontWeight('bold');
  } else if (nombre === NOMBRE_HOJA_VENTAS) {
    hoja.getRange(2, 1, filasDatos, 4).setHorizontalAlignment('center');
    hoja.getRange(2, 6, filasDatos, 1).setHorizontalAlignment('center');
    hoja.getRange(2, 7, filasDatos, 2).setHorizontalAlignment('right');
    hoja.getRange(2, 9, filasDatos, 2).setHorizontalAlignment('left');
  } else if (nombre === 'Eliminados') {
    hoja.getRange(2, 1, filasDatos, 4).setHorizontalAlignment('center');
    hoja.getRange(2, 7, filasDatos, 2).setHorizontalAlignment('right');
    hoja.getRange(2, 9, filasDatos, 1).setWrap(true);
  }
}

function estilizarHojasDatos_(nombresHojas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configuraciones = [
    [NOMBRE_HOJA_PRODUCTOS, 9, '#1E5EFF', [75, 200, 145, 105, 85, 100, 110, 110, 240]],
    [NOMBRE_HOJA_COMBOS, 4, '#60708F', [75, 210, 115, 420]],
    [NOMBRE_HOJA_CATEGORIAS, 3, '#0F3D2E', [190, 90, 110]],
    [NOMBRE_HOJA_VENTAS, 10, '#2E9B72', [120, 100, 100, 75, 220, 90, 130, 130, 120, 140]],
    ['Eliminados', 9, '#4D5666', [120, 100, 110, 75, 200, 140, 110, 120, 300]],
  ];

  configuraciones.forEach(([nombre, columnas, color, anchos]) => {
    if (nombresHojas && !nombresHojas.includes(nombre)) return;
    const hoja = ss.getSheetByName(nombre);
    if (!hoja) return;
    if (nombre === NOMBRE_HOJA_VENTAS && hoja.getLastRow() > 0 && !hoja.getRange(1, 10).getValue()) {
      hoja.getRange(1, 10).setValue('Método de pago');
    }
    estilizarTabla_(hoja, columnas, color, anchos);
    ajustarAlineacionHoja_(hoja);
    const filas = hoja.getLastRow() - 1;
    if (filas < 1) return;
    if (nombre === NOMBRE_HOJA_PRODUCTOS) {
      hoja.getRange(2, 4, filas, 1).setNumberFormat('#,##0.00" Bs"');
      hoja.getRange(2, 7, filas, 1).setNumberFormat('#,##0.00" Bs"');
    } else if (nombre === NOMBRE_HOJA_COMBOS) {
      hoja.getRange(2, 3, filas, 1).setNumberFormat('#,##0.00" Bs"');
    } else if (nombre === NOMBRE_HOJA_VENTAS) {
      hoja.getRange(2, 7, filas, 2).setNumberFormat('#,##0.00" Bs"');
    } else if (nombre === 'Eliminados') {
      hoja.getRange(2, 8, filas, 1).setNumberFormat('#,##0.00" Bs"');
    }
  });
}

function obtenerRevisionWeb_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tablas = [NOMBRE_HOJA_CATEGORIAS, NOMBRE_HOJA_PRODUCTOS, NOMBRE_HOJA_COMBOS].map(nombre => {
    const hoja = ss.getSheetByName(nombre);
    return hoja ? hoja.getDataRange().getValues() : [];
  });
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify(tablas),
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(digest);
}

function apiObtenerDatosWeb(token) {
  const session = requireSesion_(token, [ROL_ADMINISTRADOR, ROL_EMPLEADO]);
  const almacen = cargarAlmacenDesdeSheet_();
  const hojaVentas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_VENTAS);
  const zona = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const ventas = hojaVentas && hojaVentas.getLastRow() > 1
    ? hojaVentas.getRange(2, 1, hojaVentas.getLastRow() - 1, 10).getValues().map(fila => {
      const fecha = fila[0] instanceof Date ? Utilities.formatDate(fila[0], zona, 'yyyy-MM-dd') : String(fila[0] || '');
      const hora = fila[1] instanceof Date ? Utilities.formatDate(fila[1], zona, 'HH:mm:ss') : String(fila[1] || '');
      return {
        fecha: `${fecha} ${hora}`.trim(),
        tipo: String(fila[2] || ''),
        detalle: String(fila[4] || ''),
        cantidad: Number(fila[5]) || 0,
        total: Number(fila[7]) || 0,
        usuario: String(fila[8] || ''),
        metodoPago: String(fila[9] || '')
      };
    }).reverse()
    : [];

  return {
    version: obtenerRevisionWeb_(),
    categorias: almacen.categorias.map(c => ({ nombre: c.nombre, letra: c.letra, ultimo: Number(c.ultimo) || 0 })),
    productos: almacen.productos.map(p => ({
      id: p.id,
      nombre: p.nombre,
      categoria: p.categoria,
      precio: Number(p.precio) || 0,
      stock: Number(p.stock) || 0,
      porciones: Number(p.porciones) || 0,
      precioVaso: Number(p.precioVaso) || 0,
      vasosSueltos: Number(p.vasosSueltos) || 0,
      imagen: String(p.imagen || '')
    })),
    combos: almacen.combos.map(c => ({
      id: c.id,
      nombre: c.nombre,
      precio: Number(c.precio) || 0,
      componentes: c.componentes.map(parte => ({ id: parte.producto.id, cantidad: Number(parte.cantidad) || 0 })),
      componentesRotos: c.componentesRotos || [],
      textoOriginal: c._textoOriginal || null
    })),
    ultimoCombo: Number(almacen.ultimoCombo) || 0,
    enlacePagoUrl: String(PropertiesService.getScriptProperties().getProperty('SIS101_PAYMENT_QR_URL') || ''),
    qrPagoUrl: String(PropertiesService.getScriptProperties().getProperty('SIS101_PAYMENT_QR_URL') || ''),
    ventas: session.rol === ROL_ADMINISTRADOR
      ? ventas
      : ventas.filter(venta => venta.usuario === session.usuario)
  };
}

function apiGuardarInventarioWeb(token, datos, versionEsperada) {
  requireSesion_(token, [ROL_ADMINISTRADOR]);
  return conLock_(() => {
    const versionActual = obtenerRevisionWeb_();
    if (String(versionEsperada) !== versionActual) {
      return { ok: false, conflicto: true, mensaje: 'El inventario cambió en otra sesión. Se cargaron los datos más recientes.', datos: apiObtenerDatosWeb(token) };
    }

    try {
      if (!datos || !Array.isArray(datos.categorias) || !Array.isArray(datos.productos) || !Array.isArray(datos.combos)) {
        return { ok: false, mensaje: 'Los datos recibidos no tienen un formato válido.' };
      }

      const almacen = new Almacen();
      almacen.categorias = datos.categorias.map(c => ({ nombre: String(c.nombre || '').trim(), letra: String(c.letra || '').trim().toUpperCase(), ultimo: Number(c.ultimo) || 0 }));
      const nombresCategorias = almacen.categorias.map(c => c.nombre.toLowerCase());
      const letrasCategorias = almacen.categorias.map(c => c.letra);
      if (nombresCategorias.some((nombre, i) => !nombre || nombresCategorias.indexOf(nombre) !== i)
        || letrasCategorias.some((letra, i) => !/^[A-Z]$/.test(letra) || letra === 'K' || letrasCategorias.indexOf(letra) !== i)) {
        return { ok: false, mensaje: 'Hay nombres o prefijos de categoría inválidos o repetidos.' };
      }

      almacen.productos = datos.productos.map(d => {
        const producto = new Producto(String(d.nombre || '').trim(), String(d.categoria || ''), Number(d.precio), Number(d.stock), Number(d.porciones), Number(d.precioVaso));
        producto.id = String(d.id || '');
        producto.vasosSueltos = Number(d.vasosSueltos) || 0;
        producto.imagen = String(d.imagen || '').trim();
        return producto;
      });
      const idsProductos = almacen.productos.map(p => p.id);
      const nombresProductos = almacen.productos.map(p => p.nombre.toLowerCase());
      if (almacen.productos.some(p => !p.id || !p.nombre || !almacen.categorias.some(c => c.nombre === p.categoria)
        || [p.precio, p.stock, p.porciones, p.precioVaso, p.vasosSueltos].some(n => !Number.isFinite(n) || n < 0))
        || almacen.productos.some(p => !Number.isInteger(p.stock) || !Number.isInteger(p.porciones) || !Number.isInteger(p.vasosSueltos)
          || ((p.porciones > 0) !== (p.precioVaso > 0)))
        || idsProductos.some((id, i) => idsProductos.indexOf(id) !== i)
        || nombresProductos.some((nombre, i) => nombresProductos.indexOf(nombre) !== i)) {
        return { ok: false, mensaje: 'Hay productos inválidos, repetidos o con categorías inexistentes.' };
      }
      if (almacen.productos.some(p => p.imagen && !/^https?:\/\/\S+$/i.test(p.imagen))) {
        return { ok: false, mensaje: 'Cada imagen debe tener una URL válida que comience con http:// o https://.' };
      }

      almacen.combos = datos.combos.map(d => {
        const componentes = (Array.isArray(d.componentes) ? d.componentes : []).map(parte => ({
          producto: almacen.buscarProducto(String(parte.id || '')),
          cantidad: Number(parte.cantidad)
        }));
        if (componentes.some(parte => !parte.producto || !Number.isInteger(parte.cantidad) || parte.cantidad <= 0)) {
          throw new Error('Un combo contiene componentes inválidos.');
        }
        const idsComponentes = componentes.map(parte => parte.producto.id);
        if (idsComponentes.some((id, i) => idsComponentes.indexOf(id) !== i)) throw new Error('Un combo tiene productos repetidos.');
        const combo = new Combo(String(d.nombre || '').trim(), Number(d.precio), componentes);
        combo.id = String(d.id || '');
        combo.componentesRotos = Array.isArray(d.componentesRotos) ? d.componentesRotos.map(nombre => String(nombre)) : [];
        combo._textoOriginal = combo.componentesRotos.length ? String(d.textoOriginal || '') : null;
        return combo;
      });
      const idsCombos = almacen.combos.map(c => c.id);
      const nombresCombos = almacen.combos.map(c => c.nombre.toLowerCase());
      if (almacen.combos.some(c => !c.id || !c.nombre || !Number.isFinite(c.precio) || c.precio < 0
        || (!c.componentes.length && !c.componentesRotos.length)
        || (c.componentesRotos.length && !c._textoOriginal))
        || idsCombos.some((id, i) => idsCombos.indexOf(id) !== i)
        || nombresCombos.some((nombre, i) => nombresCombos.indexOf(nombre) !== i)) {
        return { ok: false, mensaje: 'Hay combos inválidos, repetidos o sin componentes.' };
      }
      almacen.ultimoCombo = Number(datos.ultimoCombo) || 0;

      const almacenAnterior = cargarAlmacenDesdeSheet_();
      guardarAlmacenEnSheet_(almacen);
      const idsProductosNuevos = almacen.productos.map(p => p.id);
      const idsCombosNuevos = almacen.combos.map(c => c.id);
      const nombresCategoriasNuevas = almacen.categorias.map(c => c.nombre);
      almacenAnterior.productos.filter(p => !idsProductosNuevos.includes(p.id)).forEach(p => {
        registrarEliminacion_('Producto', p.id, p.nombre, p.categoria, p.stock, p.precio, 'Desde la aplicación web');
      });
      almacenAnterior.combos.filter(c => !idsCombosNuevos.includes(c.id)).forEach(c => {
        registrarEliminacion_('Combo', c.id, c.nombre, '-', 0, c.precio, 'Desde la aplicación web');
      });
      almacenAnterior.categorias.filter(c => !nombresCategoriasNuevas.includes(c.nombre)).forEach(c => {
        registrarEliminacion_('Categoría', '-', c.nombre, '-', 0, 0, 'Desde la aplicación web');
      });
      return { ok: true, datos: apiObtenerDatosWeb(token) };
    } catch (e) {
      return { ok: false, mensaje: 'No se pudo guardar el inventario: ' + e.message };
    }
  });
}

function validarMetodoPagoVenta_(metodoPago, pagoConfirmado, enlaceMostrado) {
  if (metodoPago !== 'Efectivo' && metodoPago !== 'Enlace' && metodoPago !== 'QR') {
    throw new Error('Selecciona un método de pago válido.');
  }
  const hojaVentas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_VENTAS);
  if (hojaVentas && hojaVentas.getLastRow() > 0) {
    const encabezadoPago = String(hojaVentas.getRange(1, 10).getValue() || '');
    if (encabezadoPago && encabezadoPago !== 'Método de pago') {
      throw new Error('La columna J de la hoja Ventas ya está ocupada por otro dato.');
    }
  }
  if (metodoPago === 'Enlace' || metodoPago === 'QR') {
    const paymentUrl = String(PropertiesService.getScriptProperties().getProperty('SIS101_PAYMENT_QR_URL') || '').trim();
    if (!paymentUrl) throw new Error('El administrador todavía no configuró el enlace de pago.');
    if (enlaceMostrado !== paymentUrl) throw new Error('El enlace de pago cambió. Actualiza la página antes de registrar la venta.');
    if (pagoConfirmado !== true) throw new Error('Confirma que verificaste la recepción del pago.');
  }
  return metodoPago;
}

function apiVenderProductoWeb(token, idProducto, cantidad, metodoPago, pagoConfirmado, enlaceMostrado) {
  const session = requireSesion_(token, [ROL_EMPLEADO]);
  return conLock_(() => {
    try {
      const metodoPagoVenta = validarMetodoPagoVenta_(metodoPago, pagoConfirmado, enlaceMostrado);
      const almacen = cargarAlmacenDesdeSheet_();
      const producto = almacen.buscarProducto(String(idProducto || ''));
      const cantidadVenta = Number(cantidad);
      if (!producto) return { ok: false, mensaje: 'No existe el producto seleccionado.' };
      if (!Number.isInteger(cantidadVenta) || cantidadVenta <= 0 || !producto.reducirStock(cantidadVenta)) {
        return { ok: false, mensaje: 'No se pudo vender el producto. Revisa la cantidad y el stock disponible.' };
      }
      const precio = Number(producto.precio);
      const nombre = producto.nombre;
      guardarAlmacenEnSheet_(almacen);
      registrarVenta_('Producto', producto.id, nombre, cantidadVenta, precio, session.usuario, metodoPagoVenta);
      return { ok: true, mensaje: `Venta registrada: ${cantidadVenta} ${nombre}.`, datos: apiObtenerDatosWeb(token) };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

function apiVenderVasoWeb(token, idProducto, idVaso, cantidad, metodoPago, pagoConfirmado, enlaceMostrado) {
  const session = requireSesion_(token, [ROL_EMPLEADO]);
  return conLock_(() => {
    try {
      const metodoPagoVenta = validarMetodoPagoVenta_(metodoPago, pagoConfirmado, enlaceMostrado);
      const almacen = cargarAlmacenDesdeSheet_();
      const producto = almacen.buscarProducto(idProducto);
      if (!producto) return { ok: false, mensaje: `No existe el producto ${idProducto}.` };
      const precio = Number(producto.precioVaso);
      const nombre = producto.nombre;
      const cantidadVenta = Number(cantidad);
      if (!Number.isInteger(cantidadVenta) || !almacen.venderVaso(idProducto, cantidadVenta, idVaso)) {
        return { ok: false, mensaje: 'No se pudo realizar la venta. Revisa el stock disponible.' };
      }
      guardarAlmacenEnSheet_(almacen);
      registrarVenta_('Vaso', idProducto, nombre, cantidadVenta, precio, session.usuario, metodoPagoVenta);
      return { ok: true, mensaje: `Venta registrada: ${cantidadVenta} vaso(s) de ${nombre}.`, datos: apiObtenerDatosWeb(token) };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

function apiVenderComboWeb(token, idCombo, cantidad, metodoPago, pagoConfirmado, enlaceMostrado) {
  const session = requireSesion_(token, [ROL_EMPLEADO]);
  return conLock_(() => {
    try {
      const metodoPagoVenta = validarMetodoPagoVenta_(metodoPago, pagoConfirmado, enlaceMostrado);
      const almacen = cargarAlmacenDesdeSheet_();
      const combo = almacen.buscarCombo(idCombo);
      if (!combo) return { ok: false, mensaje: `No existe el combo ${idCombo}.` };
      const precio = Number(combo.precio);
      const nombre = combo.nombre;
      const cantidadVenta = Number(cantidad);
      if (!Number.isInteger(cantidadVenta) || !almacen.venderCombo(idCombo, cantidadVenta)) {
        return { ok: false, mensaje: 'No se pudo vender el combo. Revisa el stock de sus componentes.' };
      }
      guardarAlmacenEnSheet_(almacen);
      registrarVenta_('Combo', idCombo, nombre, cantidadVenta, precio, session.usuario, metodoPagoVenta);
      return { ok: true, mensaje: `Venta registrada: ${cantidadVenta} combo(s) ${nombre}.`, datos: apiObtenerDatosWeb(token) };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

class Producto {
  constructor(nombre, categoria, precio, stock, porciones, precioVaso, imagen = '') {
    this.id = "";
    this.nombre = nombre;
    this.categoria = categoria;
    this.precio = precio;
    this.stock = stock;
    this.porciones = porciones;
    this.precioVaso = precioVaso;
    this.vasosSueltos = 0;
    this.imagen = imagen;
  }

  ver() {
    Logger.log(`[${this.id}] ${this.nombre} | ${this.categoria} | ${this.precio} Bs | Stock: ${this.stock} | Vasos sueltos: ${this.vasosSueltos}`);
  }

  modificar(nombre, precio, porciones, precioVaso) {
    this.nombre = nombre;
    this.precio = precio;
    this.porciones = porciones;
    this.precioVaso = precioVaso;
  }

  aumentarStock(cantidad) {
    if (cantidad <= 0) { Logger.log("La cantidad debe ser mayor a 0"); return false; }
    this.stock = this.stock + cantidad;
    return true;
  }

  reducirStock(cantidad) {
    if (cantidad <= 0) { Logger.log("La cantidad debe ser mayor a 0"); return false; }
    if (cantidad > this.stock) {
      Logger.log(`Stock insuficiente de ${this.nombre} (hay ${this.stock})`);
      return false;
    }
    this.stock = this.stock - cantidad;
    return true;
  }

  hayStock() {
    return this.stock > 0 || this.vasosSueltos > 0;
  }

  vasosDisponibles() {
    return this.stock * this.porciones + this.vasosSueltos;
  }

  venderVaso(cantidad) {
    if (this.porciones === 0) { Logger.log(`${this.nombre} no se vende por vaso`); return false; }
    if (cantidad <= 0) { Logger.log("La cantidad debe ser mayor a 0"); return false; }
    if (cantidad > this.vasosDisponibles()) {
      Logger.log(`Vasos insuficientes de ${this.nombre} (quedan ${this.vasosDisponibles()})`);
      return false;
    }
    for (let i = 0; i < cantidad; i++) {
      if (this.vasosSueltos === 0) {
        this.stock = this.stock - 1;
        this.vasosSueltos = this.porciones;
      }
      this.vasosSueltos = this.vasosSueltos - 1;
    }
    return true;
  }

  aFila() {
    return [this.id, this.nombre, this.categoria, this.precio, this.stock, this.porciones, this.precioVaso, this.vasosSueltos, this.imagen];
  }
}

class Combo {
  constructor(nombre, precio, componentes) {
    this.id = "";
    this.nombre = nombre;
    this.precio = precio;
    this.componentes = componentes;
    this.componentesRotos = [];
  }

  estaRoto() {
    return this.componentesRotos && this.componentesRotos.length > 0;
  }

  ver() {
    Logger.log(`[${this.id}] ${this.nombre} | ${this.precio} Bs | Disponibles: ${this.estaRoto() ? 'ROTO' : this.disponibles()}`);
    this.componentes.forEach(c => Logger.log(`    - ${c.cantidad} x ${c.producto.nombre}`));
    if (this.estaRoto()) {
      this.componentesRotos.forEach(n => Logger.log(`    ⚠️ FALTANTE: ${n}`));
    }
  }

  modificar(nombre, precio, componentes) {
    this.nombre = nombre;
    this.precio = precio;
    this.componentes = componentes;
    this.componentesRotos = [];
    this._textoOriginal = null;
  }

  disponibles() {
    if (this.estaRoto()) return 0;
    if (this.componentes.length === 0) return 0;
    return this.componentes.reduce((menor, c) => {
      let resto = c.producto.stock, veces = 0;
      while (resto >= c.cantidad) { resto -= c.cantidad; veces++; }
      if (menor === -1 || veces < menor) return veces;
      return menor;
    }, -1);
  }

  vender(cantidad) {
    if (this.estaRoto()) { Logger.log(`El combo ${this.nombre} tiene componentes faltantes`); return false; }
    if (cantidad <= 0) { Logger.log("La cantidad debe ser mayor a 0"); return false; }
    if (cantidad > this.disponibles()) { Logger.log(`Combos insuficientes de ${this.nombre} (se pueden armar ${this.disponibles()})`); return false; }
    this.componentes.forEach(c => c.producto.reducirStock(c.cantidad * cantidad));
    return true;
  }

  aFila() {
    if (this.estaRoto() && this._textoOriginal) {
      return [this.id, this.nombre, this.precio, this._textoOriginal];
    }
    const texto = this.componentes.map(c => `${c.cantidad} x ${c.producto.nombre}`).join(' + ');
    return [this.id, this.nombre, this.precio, texto];
  }
}

class Almacen {
  constructor() {
    this.productos = [];
    this.combos = [];
    this.letraCombos = "K";
    this.ultimoCombo = 0;
    this.categorias = [
      { nombre: "Botellas", letra: "B", ultimo: 0 },
      { nombre: "Cervezas", letra: "C", ultimo: 0 },
      { nombre: "Mezcladores", letra: "M", ultimo: 0 },
      { nombre: "Insumos", letra: "I", ultimo: 0 },
    ];
  }

  agregarCategoria(nombre, letra) {
    if (this.categorias.some(c => c.nombre === nombre)) { Logger.log(`Ya existe la categoría ${nombre}`); return false; }
    if (letra.length !== 1) { Logger.log("La letra de la categoría debe ser un solo carácter"); return false; }
    if (letra === this.letraCombos) { Logger.log(`La letra ${letra} está reservada para los combos`); return false; }
    if (this.categorias.some(c => c.letra === letra)) { Logger.log(`La letra ${letra} ya la usa otra categoría`); return false; }
    this.categorias.push({ nombre: nombre, letra: letra, ultimo: 0 });
    return true;
  }

  verCategorias() {
    this.categorias.forEach(c => {
      const cantidad = this.productos.filter(p => p.categoria === c.nombre).length;
      Logger.log(`${c.letra} - ${c.nombre} | Productos: ${cantidad}`);
    });
    Logger.log(`${this.letraCombos} - Combos | Combos: ${this.combos.length}`);
  }

  agregarProducto(producto) {
    const categoria = this.categorias.find(c => c.nombre === producto.categoria);
    if (categoria === undefined) { Logger.log(`La categoría ${producto.categoria} no existe en el almacén`); return false; }
    if (this.productos.some(p => p === producto)) { Logger.log(`El producto ${producto.nombre} ya fue agregado`); return false; }
    categoria.ultimo = categoria.ultimo + 1;
    producto.id = categoria.letra + categoria.ultimo;
    this.productos.push(producto);
    return true;
  }

  agregarCombo(combo) {
    if (this.combos.some(c => c === combo)) { Logger.log(`El combo ${combo.nombre} ya fue agregado`); return false; }
    const hayNoRegistrados = combo.componentes.some(c => !this.productos.some(p => p === c.producto));
    if (hayNoRegistrados) { Logger.log(`El combo ${combo.nombre} usa productos que no están en el almacén`); return false; }
    this.ultimoCombo = this.ultimoCombo + 1;
    combo.id = this.letraCombos + this.ultimoCombo;
    this.combos.push(combo);
    return true;
  }

  buscarProducto(id) { return this.productos.find(p => p.id === id); }
  buscarCombo(id) { return this.combos.find(c => c.id === id); }

  eliminarProducto(id) {
    const posicion = this.productos.findIndex(p => p.id === id);
    if (posicion === -1) { Logger.log(`No existe un producto con el id ${id}`); return false; }
    const producto = this.productos[posicion];
    const comboQueLoUsa = this.combos.find(c => c.componentes.some(k => k.producto === producto));
    if (comboQueLoUsa !== undefined) { Logger.log(`No se puede eliminar: lo usa el combo ${comboQueLoUsa.nombre}`); return false; }
    this.productos.splice(posicion, 1);
    return true;
  }

  eliminarCombo(id) {
    const posicion = this.combos.findIndex(c => c.id === id);
    if (posicion === -1) { Logger.log(`No existe un combo con el id ${id}`); return false; }
    this.combos.splice(posicion, 1);
    return true;
  }

  verProductos() {
    if (this.productos.length === 0) Logger.log("No hay productos en el almacén");
    else this.productos.forEach(p => p.ver());
  }

  verCombos() {
    if (this.combos.length === 0) Logger.log("No hay combos registrados");
    else this.combos.forEach(c => c.ver());
  }

  filtrarPorCategoria(categoria) { return this.productos.filter(p => p.categoria === categoria); }
  stockBajo(limite) { return this.productos.filter(p => p.stock <= limite); }
  valorTotal() { return this.productos.reduce((total, p) => total + p.precio * p.stock, 0); }

  venderVaso(idProducto, cantidad, idVaso) {
    const licor = this.buscarProducto(idProducto);
    const vaso = this.buscarProducto(idVaso);
    if (licor === undefined || vaso === undefined) { Logger.log("El licor o el vaso no existen en el almacén"); return false; }
    if (cantidad <= 0) { Logger.log("La cantidad debe ser mayor a 0"); return false; }
    if (licor.porciones === 0) { Logger.log(`${licor.nombre} no se vende por vaso`); return false; }
    if (cantidad > licor.vasosDisponibles()) { Logger.log(`Vasos de licor insuficientes de ${licor.nombre} (quedan ${licor.vasosDisponibles()})`); return false; }
    if (cantidad > vaso.stock) { Logger.log(`Vasos descartables insuficientes (quedan ${vaso.stock})`); return false; }
    licor.venderVaso(cantidad);
    vaso.reducirStock(cantidad);
    return true;
  }

  venderCombo(idCombo, cantidad) {
    const combo = this.buscarCombo(idCombo);
    if (combo === undefined) { Logger.log(`No existe un combo con el id ${idCombo}`); return false; }
    return combo.vender(cantidad);
  }

  aFilasProductos() { return this.productos.map(p => p.aFila()); }
  aFilasCombos() { return this.combos.map(c => c.aFila()); }
}

function menuVerProductos() {
  requireAdministradorDeHoja_();
  const almacen = cargarAlmacenDesdeSheet_();
  almacen.verProductos();
  SpreadsheetApp.getUi().alert('Revisa los registros (Ver > Registros) para ver los productos.');
}

function menuVerCombos() {
  requireAdministradorDeHoja_();
  const almacen = cargarAlmacenDesdeSheet_();
  almacen.verCombos();
  SpreadsheetApp.getUi().alert('Revisa los registros (Ver > Registros) para ver los combos.');
}

function menuVerCategorias() {
  requireAdministradorDeHoja_();
  const almacen = cargarAlmacenDesdeSheet_();
  almacen.verCategorias();
  SpreadsheetApp.getUi().alert('Revisa los registros (Ver > Registros) para ver las categorías.');
}

function menuReiniciarDatos() {
  requireAdministradorDeHoja_();
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert('¿Estás seguro?', 'Esto borrará todos los datos actuales y creará un conjunto de prueba.', ui.ButtonSet.YES_NO);
  if (respuesta !== ui.Button.YES) return;

  return conLock_(() => {
    const almacen = new Almacen();
    almacen.agregarCategoria("Postres", "P");

    const singani = new Producto("Singani Casa Real", "Botellas", 180, 15, 12, 20);
    const fernet = new Producto("Fernet Branca", "Botellas", 220, 5, 12, 25);
    const cerveza = new Producto("Cerveza Paceña", "Cervezas", 25, 120, 0, 0);
    const coca = new Producto("Coca-Cola 2L", "Mezcladores", 15, 8, 0, 0);
    const hielo = new Producto("Bolsa de hielo", "Insumos", 10, 20, 0, 0);
    const balde = new Producto("Balde vacío", "Insumos", 30, 12, 0, 0);
    const vaso = new Producto("Vaso descartable", "Insumos", 1, 100, 0, 0);
    const postre = new Producto("Brownie", "Postres", 20, 10, 0, 0);

    almacen.agregarProducto(singani);
    almacen.agregarProducto(fernet);
    almacen.agregarProducto(cerveza);
    almacen.agregarProducto(coca);
    almacen.agregarProducto(hielo);
    almacen.agregarProducto(balde);
    almacen.agregarProducto(vaso);
    almacen.agregarProducto(postre);

    const comboSingani = new Combo("Combo Singani", 220, [
      { producto: singani, cantidad: 1 }, { producto: coca, cantidad: 2 },
      { producto: hielo, cantidad: 1 }, { producto: vaso, cantidad: 4 },
    ]);
    const comboFernet = new Combo("Combo Fernet", 260, [
      { producto: fernet, cantidad: 1 }, { producto: coca, cantidad: 2 },
      { producto: hielo, cantidad: 1 }, { producto: vaso, cantidad: 4 },
    ]);
    const comboCerveza = new Combo("Combo Cerveza", 90, [
      { producto: cerveza, cantidad: 6 }, { producto: hielo, cantidad: 1 },
      { producto: vaso, cantidad: 6 },
    ]);

    almacen.agregarCombo(comboSingani);
    almacen.agregarCombo(comboFernet);
    almacen.agregarCombo(comboCerveza);

    guardarAlmacenEnSheet_(almacen);
    ui.alert('Datos de prueba reiniciados. Revisa las hojas "Productos", "Combos" y "Categorias".');
  });
}

function registrarVenta_(tipo, id, nombre, cantidad, precioUnitario, usuario, metodoPago) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hojaVentas = ss.getSheetByName(NOMBRE_HOJA_VENTAS);
  if (!hojaVentas) hojaVentas = ss.insertSheet(NOMBRE_HOJA_VENTAS);
  if (hojaVentas.getLastRow() === 0) {
    hojaVentas.appendRow(['Fecha', 'Hora', 'Tipo', 'ID', 'Nombre', 'Cantidad', 'Precio Unitario', 'Total', 'Usuario', 'Método de pago']);
  } else if (!hojaVentas.getRange(1, 9).getValue()) {
    hojaVentas.getRange(1, 9).setValue('Usuario');
  }
  const encabezadoPago = String(hojaVentas.getRange(1, 10).getValue() || '');
  if (encabezadoPago && encabezadoPago !== 'Método de pago') {
    throw new Error('La columna J de la hoja Ventas ya está ocupada por otro dato.');
  }
  if (!encabezadoPago) hojaVentas.getRange(1, 10).setValue('Método de pago');
  const ahora = new Date();
  const zona = ss.getSpreadsheetTimeZone();
  const fecha = Utilities.formatDate(ahora, zona, 'yyyy-MM-dd');
  const hora = Utilities.formatDate(ahora, zona, 'HH:mm:ss');
  const fila = hojaVentas.getLastRow() + 1;
  hojaVentas.getRange(fila, 1, 1, 2).setNumberFormat('@');
  hojaVentas.getRange(fila, 1, 1, 10).setValues([[fecha, hora, tipo, id, nombre, cantidad, precioUnitario, cantidad * precioUnitario, usuario || 'Sheets', metodoPago]]);
  if (fila === 2) {
    estilizarTabla_(hojaVentas, 10, '#2E9B72', [120, 100, 100, 75, 220, 90, 130, 130, 120, 140]);
  } else {
    estilizarFilaNueva_(hojaVentas, fila, 10, '#2E9B72');
  }
  const filtro = hojaVentas.getFilter();
  if (filtro) {
    const rangoFiltro = filtro.getRange();
    if (rangoFiltro.getNumColumns() < 10 || rangoFiltro.getNumRows() < fila) {
      filtro.remove();
      hojaVentas.getRange(1, 1, fila, 10).createFilter();
    }
  }
  hojaVentas.getRange(fila, 1, 1, 4).setHorizontalAlignment('center');
  hojaVentas.getRange(fila, 6, 1, 1).setHorizontalAlignment('center');
  hojaVentas.getRange(fila, 7, 1, 2).setHorizontalAlignment('right').setNumberFormat('#,##0.00" Bs"');
  hojaVentas.getRange(fila, 9, 1, 2).setHorizontalAlignment('left');
  return true;
}

function include(nombreArchivo) {
  return HtmlService.createHtmlOutputFromFile(nombreArchivo).getContent();
}

function abrirDialogoVenderVaso() {
  requireAdministradorDeHoja_();
  const html = HtmlService.createTemplateFromFile('DialogoVenderVaso')
    .evaluate()
    .setWidth(440)
    .setHeight(430);
  SpreadsheetApp.getUi().showModalDialog(html, 'Vender Vaso');
}

function abrirDialogoVenderCombo() {
  requireAdministradorDeHoja_();
  const html = HtmlService.createTemplateFromFile('DialogoVenderCombo')
    .evaluate()
    .setWidth(440)
    .setHeight(390);
  SpreadsheetApp.getUi().showModalDialog(html, 'Vender Combo');
}

function getDatosParaVenderVaso() {
  requireAdministradorDeHoja_();
  const almacen = cargarAlmacenDesdeSheet_();
  const licores = almacen.productos
    .filter(p => p.porciones > 0 && p.vasosDisponibles() > 0)
    .map(p => ({ id: p.id, nombre: p.nombre, precioVaso: p.precioVaso, vasosDisponibles: p.vasosDisponibles() }));
  const vasos = almacen.productos
    .filter(p => p.stock > 0)
    .map(p => ({ id: p.id, nombre: p.nombre, stock: p.stock }));
  return { licores, vasos };
}

function getDatosParaVenderCombo() {
  requireAdministradorDeHoja_();
  const almacen = cargarAlmacenDesdeSheet_();
  const combos = almacen.combos
    .filter(c => c.disponibles() > 0)
    .map(c => ({ id: c.id, nombre: c.nombre, precio: c.precio, disponibles: c.disponibles() }));
  return { combos };
}

function ejecutarVentaVasoDesdeDialogo(idLicor, idVaso, cantidad) {
  requireAdministradorDeHoja_();
  return { ok: false, mensaje: 'Las ventas solo pueden registrarse desde la aplicación con una cuenta de empleado.' };
}

function ejecutarVentaComboDesdeDialogo(idCombo, cantidad) {
  requireAdministradorDeHoja_();
  return { ok: false, mensaje: 'Las ventas solo pueden registrarse desde la aplicación con una cuenta de empleado.' };
}

function abrirDialogoAgregarProducto() {
  requireAdministradorDeHoja_();
  const html = HtmlService.createTemplateFromFile('DialogoAgregarProducto')
    .evaluate()
    .setWidth(480)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Agregar Producto');
}

function getDatosParaAgregarProducto() {
  requireAdministradorDeHoja_();
  const almacen = cargarAlmacenDesdeSheet_();
  const categorias = almacen.categorias.map(c => ({
    nombre: c.nombre,
    letra: c.letra,
    ultimo: c.ultimo
  }));
  return { categorias };
}

function ejecutarAgregarProductoDesdeDialogo(nombre, categoria, precio, stock, porciones, precioVaso) {
  requireAdministradorDeHoja_();
  return conLock_(() => {
    try {
      if (!nombre || nombre.trim() === '') return { ok: false, mensaje: 'El nombre no puede estar vacío.' };
      if (porciones < 0 || precioVaso < 0 || precio < 0 || stock < 0) return { ok: false, mensaje: 'Los valores numéricos no pueden ser negativos.' };
      if ((porciones > 0) !== (precioVaso > 0)) return { ok: false, mensaje: 'Porciones y precio por vaso deben ser ambos > 0 o ambos = 0.' };
      const almacen = cargarAlmacenDesdeSheet_();
      if (almacen.productos.some(p => p.nombre.toLowerCase() === nombre.toLowerCase())) {
        return { ok: false, mensaje: `Ya existe un producto con el nombre "${nombre}".` };
      }
      const nuevoProd = new Producto(nombre, categoria, precio, stock, porciones, precioVaso);
      if (!almacen.agregarProducto(nuevoProd)) {
        return { ok: false, mensaje: 'No se pudo agregar. Revisa la categoría o si el producto ya existe.' };
      }
      guardarAlmacenEnSheet_(almacen);
      return { ok: true, mensaje: `Producto "${nombre}" agregado con ID ${nuevoProd.id}.` };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

function conLock_(fn) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); return fn(); }
  finally { lock.releaseLock(); }
}

function registrarEliminacion_(tipo, id, nombre, categoria, stockAlBorrar, precio, detalle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('Eliminados');
  if (!hoja) hoja = ss.insertSheet('Eliminados');
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(['Fecha', 'Hora', 'Tipo', 'ID', 'Nombre', 'Categoría', 'Stock al borrar', 'Precio', 'Detalle']);
  }
  const ahora = new Date();
  const zona = Session.getScriptTimeZone();
  const fecha = Utilities.formatDate(ahora, zona, 'dd/MM/yyyy');
  const hora = Utilities.formatDate(ahora, zona, 'HH:mm:ss');
  hoja.appendRow([fecha, hora, tipo, id, nombre, categoria, stockAlBorrar, precio, detalle]);
  estilizarTabla_(hoja, 9, '#4D5666', [120, 100, 110, 75, 200, 140, 110, 120, 300]);
  ajustarAlineacionHoja_(hoja);
  hoja.getRange(2, 8, Math.max(1, hoja.getLastRow() - 1), 1).setNumberFormat('#,##0.00" Bs"');
}

function menuAgregarCategoria() {
  requireAdministradorDeHoja_();
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Agregar Categoría', 'Nombre de la categoría:', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const nombre = r1.getResponseText().trim();
  if (!nombre) { ui.alert('El nombre no puede estar vacío.'); return; }
  const r2 = ui.prompt('Agregar Categoría', 'Letra (1 carácter mayúscula):', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const letra = r2.getResponseText().trim().toUpperCase();
  return conLock_(() => {
    const almacen = cargarAlmacenDesdeSheet_();
    if (!almacen.agregarCategoria(nombre, letra)) {
      ui.alert('No se pudo agregar. Revisa los registros (Ver → Registros en el editor).');
      return;
    }
    guardarAlmacenEnSheet_(almacen);
    ui.alert(`✅ Categoría "${nombre}" agregada con letra ${letra}.`);
  });
}

function configurarProteccionCategorias() {
  requireAdministradorDeHoja_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(NOMBRE_HOJA_CATEGORIAS);
  if (!hoja) { SpreadsheetApp.getUi().alert('No existe la pestaña Categorias.'); return; }
  hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  hoja.protect().setDescription('Solo gestionable desde el menú Almacén.').setWarningOnly(false);
  const triggers = ScriptApp.getProjectTriggers();
  if (!triggers.some(t => t.getHandlerFunction() === 'onEditRevertCategorias')) {
    ScriptApp.newTrigger('onEditRevertCategorias').forSpreadsheet(ss).onEdit().create();
  }
  const almacen = cargarAlmacenDesdeSheet_();
  PropertiesService.getScriptProperties().setProperty(
    'categorias_backup',
    JSON.stringify(almacen.categorias.map(c => ({ nombre: c.nombre, letra: c.letra, ultimo: c.ultimo })))
  );
  SpreadsheetApp.getUi().alert('✅ Protección activada. Usa el menú Almacén → Categorías para gestionarlas.');
}

function onEditRevertCategorias(e) {
  if (!e || !e.range) return;
  const hoja = e.range.getSheet();
  if (hoja.getName() !== NOMBRE_HOJA_CATEGORIAS) return;
  const backupJson = PropertiesService.getScriptProperties().getProperty('categorias_backup');
  if (!backupJson) return;
  let backup;
  try { backup = JSON.parse(backupJson); } catch (err) { return; }
  if (!Array.isArray(backup)) return;
  hoja.clearContents();
  hoja.appendRow(['Nombre', 'Letra', 'Último ID']);
  if (backup.length > 0) {
    const filas = backup.map(c => [c.nombre, c.letra, c.ultimo]);
    hoja.getRange(2, 1, filas.length, 3).setValues(filas);
  }
  try { hoja.toast('Las categorías se gestionan solo desde el menú Almacén. Tu cambio fue revertido.', 'Cambio bloqueado', 6); } catch (err) { }
}

function abrirDialogoEliminarProducto() {
  requireAdministradorDeHoja_();
  const html = HtmlService.createTemplateFromFile('DialogoEliminarProducto')
    .evaluate().setWidth(520).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Eliminar Producto');
}

function getDatosParaEliminarProducto() {
  requireAdministradorDeHoja_();
  const almacen = cargarAlmacenDesdeSheet_();
  const productos = almacen.productos.map(p => {
    const combosQueLoUsan = almacen.combos.filter(c =>
      c.componentes.some(k => k.producto.id === p.id)
    );
    return {
      id: p.id,
      nombre: p.nombre,
      categoria: p.categoria,
      precio: p.precio,
      stock: p.stock,
      vasosSueltos: p.vasosSueltos,
      valorInventario: p.precio * p.stock,
      combosQueLoUsan: combosQueLoUsan.map(c => c.nombre)
    };
  });
  return { productos };
}

function ejecutarEliminarProductoDesdeDialogo(idProducto, confirmoStock) {
  requireAdministradorDeHoja_();
  return conLock_(() => {
    try {
      const almacen = cargarAlmacenDesdeSheet_();
      const producto = almacen.buscarProducto(idProducto);
      if (!producto) return { ok: false, mensaje: `No existe el producto ${idProducto}.` };

      const combosQueLoUsan = almacen.combos.filter(c =>
        c.componentes.some(k => k.producto.id === idProducto)
      );
      if (combosQueLoUsan.length > 0) {
        const nombres = combosQueLoUsan.map(c => c.nombre).join(', ');
        return { ok: false, mensaje: `Bloqueado: se usa en ${combosQueLoUsan.length} combo(s): ${nombres}. Elimínalos primero.` };
      }

      if (producto.stock > 0 && !confirmoStock) {
        return { ok: false, mensaje: 'Debes confirmar que entiendes que se perderá el stock.' };
      }

      const stockAlBorrar = producto.stock;
      const precio = producto.precio;
      const nombre = producto.nombre;
      const categoria = producto.categoria;

      if (!almacen.eliminarProducto(idProducto)) {
        return { ok: false, mensaje: 'No se pudo eliminar. Revisa los registros.' };
      }
      guardarAlmacenEnSheet_(almacen);
      registrarEliminacion_('Producto', idProducto, nombre, categoria, stockAlBorrar, precio, 'Manual');
      return { ok: true, mensaje: `Producto ${idProducto} (${nombre}) eliminado.` };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

function abrirDialogoEliminarCategoria() {
  requireAdministradorDeHoja_();
  const html = HtmlService.createTemplateFromFile('DialogoEliminarCategoria')
    .evaluate().setWidth(560).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Eliminar Categoría');
}

function getDatosParaEliminarCategoria() {
  requireAdministradorDeHoja_();
  const almacen = cargarAlmacenDesdeSheet_();
  const categorias = almacen.categorias.map(c => {
    const productos = almacen.productos.filter(p => p.categoria === c.nombre);
    const idsProductos = productos.map(p => p.id);
    const combosAfectados = almacen.combos.filter(co =>
      co.componentes.some(k => idsProductos.includes(k.producto.id))
    );
    return {
      nombre: c.nombre,
      letra: c.letra,
      productos: productos.map(p => ({ id: p.id, nombre: p.nombre, stock: p.stock, precio: p.precio })),
      combos: combosAfectados.map(co => ({ id: co.id, nombre: co.nombre, precio: co.precio }))
    };
  });
  return { categorias };
}

function ejecutarEliminarCategoriaDesdeDialogo(nombreCategoria, nombreConfirmacion) {
  requireAdministradorDeHoja_();
  return conLock_(() => {
    try {
      const almacen = cargarAlmacenDesdeSheet_();
      const categoria = almacen.categorias.find(c => c.nombre === nombreCategoria);
      if (!categoria) return { ok: false, mensaje: `No existe la categoría ${nombreCategoria}.` };
      if (nombreConfirmacion.trim() !== nombreCategoria) {
        return { ok: false, mensaje: 'El nombre de confirmación no coincide exactamente.' };
      }

      const productosAEliminar = almacen.productos.filter(p => p.categoria === nombreCategoria);
      const idsProductos = productosAEliminar.map(p => p.id);
      const combosAEliminar = almacen.combos.filter(c =>
        c.componentes.some(k => idsProductos.includes(k.producto.id))
      );

      combosAEliminar.forEach(c => {
        const pos = almacen.combos.findIndex(x => x.id === c.id);
        if (pos !== -1) almacen.combos.splice(pos, 1);
      });
      productosAEliminar.forEach(p => {
        const pos = almacen.productos.findIndex(x => x.id === p.id);
        if (pos !== -1) almacen.productos.splice(pos, 1);
      });
      const posCat = almacen.categorias.findIndex(c => c.nombre === nombreCategoria);
      if (posCat !== -1) almacen.categorias.splice(posCat, 1);

      guardarAlmacenEnSheet_(almacen);

      combosAEliminar.forEach(c => {
        registrarEliminacion_('Combo', c.id, c.nombre, '-', 0, c.precio, `Cascada por categoría ${nombreCategoria}`);
      });
      productosAEliminar.forEach(p => {
        registrarEliminacion_('Producto', p.id, p.nombre, p.categoria, p.stock, p.precio, `Cascada por categoría ${nombreCategoria}`);
      });
      registrarEliminacion_('Categoría', '-', nombreCategoria, '-', 0, 0,
        `${productosAEliminar.length} productos y ${combosAEliminar.length} combos eliminados`);

      return { ok: true, mensaje: `Categoría "${nombreCategoria}" eliminada. Se borraron ${productosAEliminar.length} producto(s) y ${combosAEliminar.length} combo(s).` };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

function menuVerificarIntegridad() {
  requireAdministradorDeHoja_();
  const almacen = cargarAlmacenDesdeSheet_();
  const problemas = [];

  const idsVistos = {};
  almacen.productos.forEach(p => {
    if (idsVistos[p.id]) problemas.push(`ID duplicado en Productos: ${p.id} (${p.nombre} y ${idsVistos[p.id]})`);
    else idsVistos[p.id] = p.nombre;
  });

  const nombresVistos = {};
  almacen.productos.forEach(p => {
    const key = p.nombre.toLowerCase();
    if (nombresVistos[key]) problemas.push(`Nombre duplicado en Productos: "${p.nombre}" (IDs ${nombresVistos[key]} y ${p.id})`);
    else nombresVistos[key] = p.id;
  });

  const nombresCats = almacen.categorias.map(c => c.nombre);
  almacen.productos.forEach(p => {
    if (!nombresCats.includes(p.categoria)) problemas.push(`Producto ${p.id} (${p.nombre}) tiene categoría inexistente: "${p.categoria}"`);
  });

  almacen.productos.forEach(p => {
    const cat = almacen.categorias.find(c => c.nombre === p.categoria);
    if (cat) {
      const m = String(p.id).match(/^([A-Z]+)/);
      if (m && m[1] !== cat.letra) {
        problemas.push(`Producto ${p.id} (${p.nombre}) tiene prefijo ${m[1]} pero su categoría "${p.categoria}" usa ${cat.letra}`);
      }
    }
  });

  almacen.categorias.forEach(c => {
    const maxId = almacen.productos
      .map(p => String(p.id).match(/^([A-Z]+)(\d+)$/))
      .filter(m => m && m[1] === c.letra)
      .reduce((max, m) => Math.max(max, parseInt(m[2])), 0);
    if (maxId > c.ultimo) {
      problemas.push(`Categoría "${c.nombre}" tiene ultimo=${c.ultimo} pero el ID más alto es ${c.letra}${maxId}. Se autocorregirá al guardar.`);
    }
  });

  almacen.combos.forEach(c => {
    if (c.estaRoto()) problemas.push(`Combo ${c.id} (${c.nombre}) tiene componentes rotos: ${c.componentesRotos.join(', ')}`);
  });

  const idsCombos = {};
  almacen.combos.forEach(c => {
    if (idsCombos[c.id]) problemas.push(`ID duplicado en Combos: ${c.id} (${c.nombre} y ${idsCombos[c.id]})`);
    else idsCombos[c.id] = c.nombre;
  });

  const ui = SpreadsheetApp.getUi();
  if (problemas.length === 0) {
    ui.alert('✅ Verificación completada', 'No se encontraron problemas.', ui.ButtonSet.OK);
    return;
  }
  const texto = problemas.map((p, i) => `${i + 1}. ${p}`).join('\n');
  Logger.log('PROBLEMAS DE INTEGRIDAD:\n' + texto);
  const resumen = problemas.length > 15 ? problemas.slice(0, 15).map((p, i) => `${i + 1}. ${p}`).join('\n') + `\n\n... y ${problemas.length - 15} más.` : texto;
  ui.alert('⚠️ Se encontraron ' + problemas.length + ' problema(s)', resumen + '\n\nEl detalle completo está en Ver → Registros.', ui.ButtonSet.OK);
}

function irAlDashboard() {
  requireAdministradorDeHoja_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Proyecto SIS101');
  if (!hoja) {
    SpreadsheetApp.getUi().alert('No existe la pestaña "Proyecto SIS101". Ejecuta primero Almacén → Instalar/Reinstalar Dashboard.');
    return;
  }
  ss.setActiveSheet(hoja);
}

function instalarDashboard() {
  requireAdministradorDeHoja_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('Proyecto SIS101');
  if (!hoja) hoja = ss.insertSheet('Proyecto SIS101', 0);
  hoja.getRange('B2:I100').breakApart();
  hoja.clear();
  hoja.clearFormats();
  hoja.getBandings().forEach(banda => banda.remove());
  hoja.setHiddenGridlines(true);
  hoja.setTabColor('#1E5EFF');
  hoja.setColumnWidth(1, 24);
  [76, 185, 138, 82, 112, 28, 112, 126].forEach((ancho, indice) => hoja.setColumnWidth(indice + 2, ancho));
  hoja.getRange('B2:I55').setBackground('#F5F7FB').setFontFamily('Arial').setFontColor('#263244').setVerticalAlignment('middle');

  hoja.getRange('B2:I2').merge();
  hoja.getRange('B2').setValue('SIS101  /  CONTROL DE INVENTARIO');
  hoja.getRange('B2:I2').setFontSize(20).setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle').setBackground('#1E5EFF').setFontColor('#FFFFFF');
  hoja.setRowHeight(2, 58);

  hoja.getRange('B3:I3').merge();
  hoja.getRange('B3').setValue('INVENTARIO Y ACTIVIDAD  ·  DATOS ACTUALIZADOS DESDE LA HOJA DE CÁLCULO');
  hoja.getRange('B3:I3').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('left').setFontColor('#667085');
  hoja.setRowHeight(3, 32);

  hoja.getRange('B5:I5').merge();
  hoja.getRange('B5').setValue('RESUMEN DEL INVENTARIO');
  hoja.getRange('B5:I5').setFontSize(12).setFontWeight('bold').setFontColor('#1746B0').setHorizontalAlignment('left').setVerticalAlignment('middle');
  hoja.setRowHeight(5, 34);

  hoja.getRange('B7:C7').merge();
  hoja.getRange('B7').setValue('PRODUCTOS REGISTRADOS');
  hoja.getRange('B7').setFontSize(10).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e8f0fe');
  hoja.getRange('B8:C8').merge();
  hoja.getRange('B8').setFormula('=COUNTA(Productos!A2:A)');
  hoja.getRange('B8').setFontSize(28).setFontWeight('bold').setFontColor('#1a73e8').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e8f0fe');
  hoja.getRange('B7:C8').setBackground('#EAF1FF').setFontColor('#1746B0');

  hoja.getRange('D7:E7').merge();
  hoja.getRange('D7').setValue('COMBOS REGISTRADOS');
  hoja.getRange('D7').setFontSize(10).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#fef7e0');
  hoja.getRange('D8:E8').merge();
  hoja.getRange('D8').setFormula('=COUNTA(Combos!A2:A)');
  hoja.getRange('D8').setFontSize(28).setFontWeight('bold').setFontColor('#b06000').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#fef7e0');
  hoja.getRange('D7:E8').setBackground('#F1F3F7').setFontColor('#263244');
  hoja.getRange('D8:E8').setFontColor('#1746B0');

  hoja.getRange('F7:G7').merge();
  hoja.getRange('F7').setValue('VALOR INVENTARIO');
  hoja.getRange('F7').setFontSize(10).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e6f4ea');
  hoja.getRange('F8:G8').merge();
  hoja.getRange('F8').setFormula('=SUMPRODUCT(Productos!D2:D, Productos!E2:E)');
  hoja.getRange('F8').setFontSize(26).setFontWeight('bold').setFontColor('#137333').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e6f4ea').setNumberFormat('#,##0" Bs"');
  hoja.getRange('F7:G8').setBackground('#EAF5EF').setFontColor('#184D37');

  hoja.getRange('H7:I7').merge();
  hoja.getRange('H7').setValue('VENTAS HOY');
  hoja.getRange('H7').setFontSize(10).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#fce8e6');
  hoja.getRange('H8:I8').merge();
  hoja.getRange('H8').setFormula('=SUMIFS(Ventas!H2:H, Ventas!A2:A, TEXT(TODAY(), "yyyy-MM-dd"))');
  hoja.getRange('H8').setFontSize(26).setFontWeight('bold').setFontColor('#c5221f').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#fce8e6').setNumberFormat('#,##0" Bs"');
  hoja.getRange('H7:I8').setBackground('#F1F3F7').setFontColor('#263244');
  hoja.getRange('H8:I8').setFontColor('#1746B0');
  hoja.setRowHeight(7, 25);
  hoja.setRowHeight(8, 55);

  hoja.getRange('B11:F11').merge();
  hoja.getRange('B11').setValue('INVENTARIO');
  hoja.getRange('B11:F11').setFontSize(12).setFontWeight('bold').setFontColor('#1746B0').setHorizontalAlignment('left').setVerticalAlignment('middle');
  hoja.setRowHeight(11, 34);

  ['ID', 'Nombre', 'Categoría', 'Stock', 'Precio'].forEach((h, i) => {
    const celda = String.fromCharCode(66 + i) + '13';
    hoja.getRange(celda).setValue(h);
    hoja.getRange(celda).setFontSize(11).setFontWeight('bold').setFontColor('#ffffff').setBackground('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  hoja.setRowHeight(13, 28);
  hoja.getRange('B14').setFormula('=IFERROR(QUERY(Productos!A2:H, "select A, B, C, E, D where A is not null order by B limit 15", 0), "Sin productos registrados")');
  for (let f = 14; f <= 29; f++) hoja.setRowHeight(f, 24);

  hoja.getRange('B31:H31').merge();
  hoja.getRange('B31').setValue('ACTIVIDAD RECIENTE');
  hoja.getRange('B31:H31').setFontSize(12).setFontWeight('bold').setFontColor('#1746B0').setHorizontalAlignment('left').setVerticalAlignment('middle');
  hoja.setRowHeight(31, 34);

  ['Fecha', 'Hora', 'Tipo', 'ID', 'Nombre', 'Cant.', 'Total'].forEach((h, i) => {
    const celda = String.fromCharCode(66 + i) + '33';
    hoja.getRange(celda).setValue(h);
    hoja.getRange(celda).setFontSize(11).setFontWeight('bold').setFontColor('#ffffff').setBackground('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  hoja.setRowHeight(33, 28);
  hoja.getRange('B34').setFormula('=IFERROR(QUERY(Ventas!A2:H, "select A, B, C, D, E, F, H order by A desc, B desc limit 10", 0), "Sin ventas registradas")');
  for (let f = 34; f <= 44; f++) hoja.setRowHeight(f, 24);

  hoja.getRange('B46:C46').merge();
  hoja.getRange('B46').setValue('TOP 5 MÁS VENDIDOS');
  hoja.getRange('B46:C46').setFontSize(12).setFontWeight('bold').setFontColor('#1746B0').setHorizontalAlignment('left').setVerticalAlignment('middle');
  hoja.setRowHeight(46, 34);

  hoja.getRange('B48').setValue('Producto / Combo');
  hoja.getRange('C48').setValue('Cantidad vendida');
  ['B48', 'C48'].forEach(c => hoja.getRange(c).setFontSize(11).setFontWeight('bold').setFontColor('#ffffff').setBackground('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle'));
  hoja.setRowHeight(48, 28);
  hoja.getRange('B49').setFormula('=IFERROR(QUERY(Ventas!A2:H, "select E, sum(F) where E is not null group by E order by sum(F) desc limit 5", 0), "Sin ventas registradas")');
  for (let f = 49; f <= 54; f++) hoja.setRowHeight(f, 24);

  hoja.getRange('B56:I100').breakApart().clearContent().clearFormat();
  hoja.getRange('B2:I55').setFontFamily('Arial').setFontColor('#263244').setVerticalAlignment('middle');
  hoja.getRange('B2:I2').setBackground('#1E5EFF').setFontColor('#FFFFFF');
  hoja.getRange('B3:I3').setFontColor('#667085');
  ['B5:I5', 'B11:F11', 'B31:H31', 'B46:C46'].forEach(rango =>
    hoja.getRange(rango).setFontColor('#1746B0').setFontWeight('bold')
  );
  hoja.getRange('B7:C8').setBackground('#EAF1FF');
  hoja.getRange('B7:C7').setFontColor('#52627A');
  hoja.getRange('B8:C8').setFontColor('#1746B0');
  hoja.getRange('D7:E8').setBackground('#F1F3F7');
  hoja.getRange('D7:E7').setFontColor('#52627A');
  hoja.getRange('D8:E8').setFontColor('#1746B0');
  hoja.getRange('F7:G8').setBackground('#EAF5EF');
  hoja.getRange('F7:G7').setFontColor('#52627A');
  hoja.getRange('F8:G8').setFontColor('#184D37');
  hoja.getRange('H7:I8').setBackground('#F1F3F7');
  hoja.getRange('H7:I7').setFontColor('#52627A');
  hoja.getRange('H8:I8').setFontColor('#1746B0');
  hoja.getRange('B13:F13').setBackground('#1E5EFF').setFontColor('#FFFFFF').setFontWeight('bold');
  hoja.getRange('B33:H33').setBackground('#1E5EFF').setFontColor('#FFFFFF').setFontWeight('bold');
  hoja.getRange('B48:C48').setBackground('#1E5EFF').setFontColor('#FFFFFF').setFontWeight('bold');
  [
    ['B14:F28', 15, 5],
    ['B34:H44', 11, 7],
    ['B49:C54', 6, 2],
  ].forEach(([rango, filas, columnas]) => {
    const colores = Array.from({ length: filas }, (_, fila) =>
      Array.from({ length: columnas }, () => fila % 2 === 0 ? '#FFFFFF' : '#F0F4FA')
    );
    hoja.getRange(rango).setBackgrounds(colores).setFontColor('#263244');
  });
  hoja.getRange('F14:F28').setNumberFormat('#,##0.00" Bs"');
  hoja.getRange('H34:H44').setNumberFormat('#,##0.00" Bs"');
  ['B7:C8', 'D7:E8', 'F7:G8', 'H7:I8', 'B13:F28', 'B33:H44', 'B48:C54'].forEach(rango =>
    hoja.getRange(rango).setBorder(true, true, true, true, false, false, '#D7DEE9', SpreadsheetApp.BorderStyle.SOLID)
  );

  estilizarHojasDatos_();
  SpreadsheetApp.flush();
  hoja.insertRowsBefore(1, 1);
  ss.setActiveSheet(hoja);
  SpreadsheetApp.getUi().alert('Dashboard instalado con sus títulos y contenido desplazados una fila hacia abajo.');
}

function abrirDialogoAgregarCombo() {
  requireAdministradorDeHoja_();
  const html = HtmlService.createTemplateFromFile('DialogoAgregarCombo')
    .evaluate().setWidth(520).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Agregar Combo');
}

function abrirDialogoModificarCombo() {
  requireAdministradorDeHoja_();
  const html = HtmlService.createTemplateFromFile('DialogoModificarCombo')
    .evaluate().setWidth(520).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Modificar Combo');
}

function getDatosParaAgregarCombo() {
  requireAdministradorDeHoja_();
  const almacen = cargarAlmacenDesdeSheet_();
  return datosDialogoCombo_(almacen);
}

function getDatosParaModificarCombo() {
  requireAdministradorDeHoja_();
  const almacen = cargarAlmacenDesdeSheet_();
  return datosDialogoCombo_(almacen);
}

function datosDialogoCombo_(almacen) {
  const productos = almacen.productos.map(p => ({
    id: p.id,
    nombre: p.nombre,
    stock: p.stock,
    categoria: p.categoria
  }));
  const combos = almacen.combos.map(c => ({
    id: c.id,
    nombre: c.nombre,
    precio: Number(c.precio),
    componentes: c.componentes.map(parte => ({ id: parte.producto.id, cantidad: Number(parte.cantidad) })),
    componentesRotos: c.componentesRotos || []
  }));
  const nombresExistentes = combos.map(c => ({ id: c.id, nombre: c.nombre }));
  return { productos, combos, nombresExistentes };
}

function ejecutarAgregarComboDesdeDialogo(nombre, precio, componentes) {
  requireAdministradorDeHoja_();
  return conLock_(() => {
    try {
      if (!nombre || nombre.trim() === '') {
        return { ok: false, mensaje: 'El nombre no puede estar vacío.' };
      }
      if (precio < 0) {
        return { ok: false, mensaje: 'El precio no puede ser negativo.' };
      }
      if (!componentes || componentes.length === 0) {
        return { ok: false, mensaje: 'Debe tener al menos un componente.' };
      }

      const ids = componentes.map(c => c.id);
      if (ids.length !== new Set(ids).size) {
        return { ok: false, mensaje: 'Hay productos repetidos en los componentes.' };
      }

      for (const c of componentes) {
        if (!c.id || !c.cantidad || c.cantidad <= 0) {
          return { ok: false, mensaje: 'Todos los componentes deben tener producto y cantidad mayor a 0.' };
        }
      }

      const almacen = cargarAlmacenDesdeSheet_();

      if (almacen.combos.some(x => x.nombre.toLowerCase() === nombre.toLowerCase())) {
        return { ok: false, mensaje: `Ya existe un combo llamado "${nombre}".` };
      }

      const componentesObjetos = [];
      for (const c of componentes) {
        const producto = almacen.buscarProducto(c.id);
        if (!producto) {
          return { ok: false, mensaje: `No existe el producto ${c.id}.` };
        }
        componentesObjetos.push({ producto, cantidad: c.cantidad });
      }

      const nuevoCombo = new Combo(nombre, precio, componentesObjetos);
      if (!almacen.agregarCombo(nuevoCombo)) {
        return { ok: false, mensaje: 'No se pudo agregar el combo. Revisa los registros.' };
      }

      guardarAlmacenEnSheet_(almacen);
      return { ok: true, mensaje: `Combo "${nombre}" agregado con ID ${nuevoCombo.id}.` };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

function ejecutarModificarComboDesdeDialogo(idCombo, nombre, precio, componentes) {
  requireAdministradorDeHoja_();
  return conLock_(() => {
    try {
      const id = String(idCombo || '');
      const nombreLimpio = String(nombre || '').trim();
      const precioNumero = Number(precio);
      if (!id) return { ok: false, mensaje: 'Selecciona un combo para modificar.' };
      if (!nombreLimpio) return { ok: false, mensaje: 'El nombre no puede estar vacío.' };
      if (!Number.isFinite(precioNumero) || precioNumero < 0) {
        return { ok: false, mensaje: 'El precio debe ser un número válido y no negativo.' };
      }
      if (!Array.isArray(componentes) || componentes.length === 0) {
        return { ok: false, mensaje: 'El combo debe tener al menos un componente.' };
      }

      const ids = componentes.map(c => String(c.id || ''));
      if (ids.some(idProducto => !idProducto) || ids.length !== new Set(ids).size) {
        return { ok: false, mensaje: 'Cada componente debe tener un producto único.' };
      }
      if (componentes.some(c => !Number.isInteger(Number(c.cantidad)) || Number(c.cantidad) <= 0)) {
        return { ok: false, mensaje: 'Las cantidades deben ser números enteros mayores a 0.' };
      }

      const almacen = cargarAlmacenDesdeSheet_();
      const combo = almacen.buscarCombo(id);
      if (!combo) return { ok: false, mensaje: `No existe el combo ${id}.` };
      if (almacen.combos.some(c => c.id !== id && c.nombre.toLowerCase() === nombreLimpio.toLowerCase())) {
        return { ok: false, mensaje: `Ya existe otro combo llamado "${nombreLimpio}".` };
      }

      const componentesObjetos = [];
      for (const componente of componentes) {
        const producto = almacen.buscarProducto(String(componente.id));
        if (!producto) return { ok: false, mensaje: `No existe el producto ${componente.id}.` };
        componentesObjetos.push({ producto, cantidad: Number(componente.cantidad) });
      }

      combo.modificar(nombreLimpio, precioNumero, componentesObjetos);
      guardarAlmacenEnSheet_(almacen);
      return { ok: true, mensaje: `Combo "${nombreLimpio}" (${id}) modificado.` };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

function abrirDialogoModificarProducto() {
  requireAdministradorDeHoja_();
  const html = HtmlService.createTemplateFromFile('DialogoModificarProducto')
    .evaluate().setWidth(520).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Modificar Producto');
}

function getDatosParaModificarProducto() {
  requireAdministradorDeHoja_();
  const almacen = cargarAlmacenDesdeSheet_();
  const productos = almacen.productos.map(p => {
    const combosQueLoUsan = almacen.combos
      .filter(c => c.componentes.some(k => k.producto.id === p.id))
      .map(c => c.nombre);
    return {
      id: p.id,
      nombre: p.nombre,
      categoria: p.categoria,
      precio: p.precio,
      stock: p.stock,
      porciones: p.porciones,
      precioVaso: p.precioVaso,
      combosQueLoUsan
    };
  });
  return { productos };
}

function ejecutarModificarProductoDesdeDialogo(idProducto, nombre, precio, stock, porciones, precioVaso) {
  requireAdministradorDeHoja_();
  return conLock_(() => {
    try {
      if (!nombre || nombre.trim() === '') {
        return { ok: false, mensaje: 'El nombre no puede estar vacío.' };
      }
      if (precio < 0 || stock < 0 || porciones < 0 || precioVaso < 0) {
        return { ok: false, mensaje: 'Los valores numéricos no pueden ser negativos.' };
      }
      if ((porciones > 0) !== (precioVaso > 0)) {
        return { ok: false, mensaje: 'Porciones y precio por vaso deben ser ambos > 0 o ambos = 0.' };
      }

      const almacen = cargarAlmacenDesdeSheet_();
      const producto = almacen.buscarProducto(idProducto);
      if (!producto) {
        return { ok: false, mensaje: `No existe el producto ${idProducto}.` };
      }

      const otroConMismoNombre = almacen.productos.find(
        p => p.id !== idProducto && p.nombre.toLowerCase() === nombre.toLowerCase()
      );
      if (otroConMismoNombre) {
        return { ok: false, mensaje: `Ya existe otro producto con el nombre "${nombre}" (${otroConMismoNombre.id}).` };
      }

      const nombreCambio = producto.nombre !== nombre;
      producto.nombre = nombre;
      producto.precio = precio;
      producto.stock = stock;
      producto.porciones = porciones;
      producto.precioVaso = precioVaso;

      guardarAlmacenEnSheet_(almacen);

      let detalle = `ID ${idProducto}`;
      if (nombreCambio) {
        const combosAfectados = almacen.combos.filter(c =>
          c.componentes.some(k => k.producto.id === idProducto)
        );
        if (combosAfectados.length > 0) {
          detalle += ` (nombre cambiado; ${combosAfectados.length} combo(s) actualizados: ${combosAfectados.map(c => c.nombre).join(', ')})`;
        }
      }

      return { ok: true, mensaje: `Producto "${nombre}" modificado. ${detalle}` };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}
