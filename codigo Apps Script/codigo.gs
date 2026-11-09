const NOMBRE_HOJA_PRODUCTOS = 'Productos';
const NOMBRE_HOJA_COMBOS = 'Combos';
const NOMBRE_HOJA_CATEGORIAS = 'Categorias';
const NOMBRE_HOJA_VENTAS = 'Ventas';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('WebApp')
    .setTitle('Almacén | SIS101');
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Almacén')
    .addItem('📊 Ir al Dashboard', 'irAlDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('Ventas')
      .addItem('Vender Vaso', 'abrirDialogoVenderVaso')
      .addItem('Vender Combo', 'abrirDialogoVenderCombo'))
    .addSubMenu(ui.createMenu('Productos')
      .addItem('Agregar Producto', 'abrirDialogoAgregarProducto')
      .addItem('Modificar Producto', 'abrirDialogoModificarProducto')
      .addItem('Eliminar Producto', 'abrirDialogoEliminarProducto')
      .addItem('Ver Productos', 'menuVerProductos'))
    .addSubMenu(ui.createMenu('Combos')
      .addItem('Agregar Combo', 'abrirDialogoAgregarCombo')
      .addItem('Ver Combos', 'menuVerCombos'))
    .addSubMenu(ui.createMenu('Categorías')
      .addItem('Agregar Categoría', 'menuAgregarCategoria')
      .addItem('Eliminar Categoría', 'abrirDialogoEliminarCategoria')
      .addItem('Ver Categorías', 'menuVerCategorias')
      .addSeparator()
      .addItem('🛡️ Bloquear Edición Manual', 'configurarProteccionCategorias'))
    .addSeparator()
    .addItem('🔍 Verificar Integridad', 'menuVerificarIntegridad')
    .addSeparator()
    .addItem('⚙️ Instalar/Reinstalar Dashboard', 'instalarDashboard')
    .addItem('⚠️ Reiniciar Datos de Prueba', 'menuReiniciarDatos')
    .addToUi();
}

function cargarAlmacenDesdeSheet() {
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
    const datos = hojaProductos.getRange(2, 1, hojaProductos.getLastRow() - 1, 8).getValues();
    almacen.productos = datos.map(fila => {
      const p = new Producto(fila[1], fila[2], fila[3], fila[4], fila[5], fila[6]);
      p.id = fila[0];
      p.vasosSueltos = Number(fila[7]);
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

function guardarAlmacenEnSheet(almacen) {
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
  hojaProductos.appendRow(['ID', 'Nombre', 'Categoría', 'Precio', 'Stock', 'Porciones', 'Precio Vaso', 'Vasos Sueltos']);
  const filasProd = almacen.aFilasProductos();
  if (filasProd.length > 0) hojaProductos.getRange(2, 1, filasProd.length, 8).setValues(filasProd);

  let hojaCombos = ss.getSheetByName(NOMBRE_HOJA_COMBOS);
  if (!hojaCombos) hojaCombos = ss.insertSheet(NOMBRE_HOJA_COMBOS);
  hojaCombos.clear();
  hojaCombos.appendRow(['ID', 'Nombre', 'Precio', 'Componentes']);
  const filasCombos = almacen.aFilasCombos();
  if (filasCombos.length > 0) hojaCombos.getRange(2, 1, filasCombos.length, 4).setValues(filasCombos);

  try { compactarTablasDashboard(); } catch (e) { }
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

function apiObtenerDatosWeb() {
  const almacen = cargarAlmacenDesdeSheet();
  const hojaVentas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_VENTAS);
  const zona = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const ventas = hojaVentas && hojaVentas.getLastRow() > 1
    ? hojaVentas.getRange(2, 1, hojaVentas.getLastRow() - 1, 8).getValues().map(fila => {
      const fecha = fila[0] instanceof Date ? Utilities.formatDate(fila[0], zona, 'yyyy-MM-dd') : String(fila[0] || '');
      const hora = fila[1] instanceof Date ? Utilities.formatDate(fila[1], zona, 'HH:mm:ss') : String(fila[1] || '');
      return {
        fecha: `${fecha} ${hora}`.trim(),
        tipo: String(fila[2] || ''),
        detalle: String(fila[4] || ''),
        cantidad: Number(fila[5]) || 0,
        total: Number(fila[7]) || 0
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
      vasosSueltos: Number(p.vasosSueltos) || 0
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
    ventas
  };
}

function apiGuardarInventarioWeb(datos, versionEsperada) {
  return conLock(() => {
    const versionActual = obtenerRevisionWeb_();
    if (String(versionEsperada) !== versionActual) {
      return { ok: false, conflicto: true, mensaje: 'El inventario cambió en otra sesión. Se cargaron los datos más recientes.', datos: apiObtenerDatosWeb() };
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

      const almacenAnterior = cargarAlmacenDesdeSheet();
      guardarAlmacenEnSheet(almacen);
      const idsProductosNuevos = almacen.productos.map(p => p.id);
      const idsCombosNuevos = almacen.combos.map(c => c.id);
      const nombresCategoriasNuevas = almacen.categorias.map(c => c.nombre);
      almacenAnterior.productos.filter(p => !idsProductosNuevos.includes(p.id)).forEach(p => {
        registrarEliminacion('Producto', p.id, p.nombre, p.categoria, p.stock, p.precio, 'Desde la aplicación web');
      });
      almacenAnterior.combos.filter(c => !idsCombosNuevos.includes(c.id)).forEach(c => {
        registrarEliminacion('Combo', c.id, c.nombre, '-', 0, c.precio, 'Desde la aplicación web');
      });
      almacenAnterior.categorias.filter(c => !nombresCategoriasNuevas.includes(c.nombre)).forEach(c => {
        registrarEliminacion('Categoría', '-', c.nombre, '-', 0, 0, 'Desde la aplicación web');
      });
      return { ok: true, datos: apiObtenerDatosWeb() };
    } catch (e) {
      return { ok: false, mensaje: 'No se pudo guardar el inventario: ' + e.message };
    }
  });
}

function apiVenderVasoWeb(idProducto, idVaso, cantidad) {
  return conLock(() => {
    try {
      const almacen = cargarAlmacenDesdeSheet();
      const producto = almacen.buscarProducto(idProducto);
      if (!producto) return { ok: false, mensaje: `No existe el producto ${idProducto}.` };
      const precio = Number(producto.precioVaso);
      const nombre = producto.nombre;
      const cantidadVenta = Number(cantidad);
      if (!Number.isInteger(cantidadVenta) || !almacen.venderVaso(idProducto, cantidadVenta, idVaso)) {
        return { ok: false, mensaje: 'No se pudo realizar la venta. Revisa el stock disponible.' };
      }
      guardarAlmacenEnSheet(almacen);
      registrarVenta('Vaso', idProducto, nombre, cantidadVenta, precio);
      return { ok: true, mensaje: `Venta registrada: ${cantidadVenta} vaso(s) de ${nombre}.`, datos: apiObtenerDatosWeb() };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

function apiVenderComboWeb(idCombo, cantidad) {
  return conLock(() => {
    try {
      const almacen = cargarAlmacenDesdeSheet();
      const combo = almacen.buscarCombo(idCombo);
      if (!combo) return { ok: false, mensaje: `No existe el combo ${idCombo}.` };
      const precio = Number(combo.precio);
      const nombre = combo.nombre;
      const cantidadVenta = Number(cantidad);
      if (!Number.isInteger(cantidadVenta) || !almacen.venderCombo(idCombo, cantidadVenta)) {
        return { ok: false, mensaje: 'No se pudo vender el combo. Revisa el stock de sus componentes.' };
      }
      guardarAlmacenEnSheet(almacen);
      registrarVenta('Combo', idCombo, nombre, cantidadVenta, precio);
      return { ok: true, mensaje: `Venta registrada: ${cantidadVenta} combo(s) ${nombre}.`, datos: apiObtenerDatosWeb() };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

class Producto {
  constructor(nombre, categoria, precio, stock, porciones, precioVaso) {
    this.id = "";
    this.nombre = nombre;
    this.categoria = categoria;
    this.precio = precio;
    this.stock = stock;
    this.porciones = porciones;
    this.precioVaso = precioVaso;
    this.vasosSueltos = 0;
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
    return [this.id, this.nombre, this.categoria, this.precio, this.stock, this.porciones, this.precioVaso, this.vasosSueltos];
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
  const almacen = cargarAlmacenDesdeSheet();
  almacen.verProductos();
  SpreadsheetApp.getUi().alert('Revisa los registros (Ver > Registros) para ver los productos.');
}

function menuVerCombos() {
  const almacen = cargarAlmacenDesdeSheet();
  almacen.verCombos();
  SpreadsheetApp.getUi().alert('Revisa los registros (Ver > Registros) para ver los combos.');
}

function menuVerCategorias() {
  const almacen = cargarAlmacenDesdeSheet();
  almacen.verCategorias();
  SpreadsheetApp.getUi().alert('Revisa los registros (Ver > Registros) para ver las categorías.');
}

function menuAgregarProductoEjemplo() {
  const almacen = cargarAlmacenDesdeSheet();
  const nuevoProd = new Producto("Whisky de Prueba", "Botellas", 150, 10, 15, 25);
  if (almacen.agregarProducto(nuevoProd)) {
    guardarAlmacenEnSheet(almacen);
    SpreadsheetApp.getUi().alert(`Producto "${nuevoProd.nombre}" agregado con ID: ${nuevoProd.id}`);
  } else {
    SpreadsheetApp.getUi().alert('No se pudo agregar el producto. Revisa los registros.');
  }
}

function menuReiniciarDatos() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert('¿Estás seguro?', 'Esto borrará todos los datos actuales y creará un conjunto de prueba.', ui.ButtonSet.YES_NO);
  if (respuesta !== ui.Button.YES) return;

  return conLock(() => {
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

    guardarAlmacenEnSheet(almacen);
    ui.alert('Datos de prueba reiniciados. Revisa las hojas "Productos", "Combos" y "Categorias".');
  });
}

function registrarVenta(tipo, id, nombre, cantidad, precioUnitario) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hojaVentas = ss.getSheetByName(NOMBRE_HOJA_VENTAS);
  if (!hojaVentas) hojaVentas = ss.insertSheet(NOMBRE_HOJA_VENTAS);
  if (hojaVentas.getLastRow() === 0) {
    hojaVentas.appendRow(['Fecha', 'Hora', 'Tipo', 'ID', 'Nombre', 'Cantidad', 'Precio Unitario', 'Total']);
  }
  const ahora = new Date();
  const zona = ss.getSpreadsheetTimeZone();
  const fecha = Utilities.formatDate(ahora, zona, 'yyyy-MM-dd');
  const hora = Utilities.formatDate(ahora, zona, 'HH:mm:ss');
  const fila = hojaVentas.getLastRow() + 1;
  hojaVentas.getRange(fila, 1, 1, 2).setNumberFormat('@');
  hojaVentas.getRange(fila, 1, 1, 8).setValues([[fecha, hora, tipo, id, nombre, cantidad, precioUnitario, cantidad * precioUnitario]]);
  return true;
}

function verificarVentas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(NOMBRE_HOJA_VENTAS);
  if (!hoja || hoja.getLastRow() <= 1) {
    Logger.log("No hay ventas registradas todavía.");
    return;
  }
  const datos = hoja.getDataRange().getValues();
  Logger.log(`Total de ventas registradas: ${datos.length - 1}`);
  datos.slice(1).forEach(fila => {
    Logger.log(`${fila[0]} ${fila[1]} | ${fila[2]} ${fila[3]} | ${fila[5]} x ${fila[4]} = ${fila[7]} Bs`);
  });
}

function include(nombreArchivo) {
  return HtmlService.createHtmlOutputFromFile(nombreArchivo).getContent();
}

function abrirDialogoVenderVaso() {
  const html = HtmlService.createTemplateFromFile('DialogoVenderVaso')
    .evaluate()
    .setWidth(440)
    .setHeight(430);
  SpreadsheetApp.getUi().showModalDialog(html, 'Vender Vaso');
}

function abrirDialogoVenderCombo() {
  const html = HtmlService.createTemplateFromFile('DialogoVenderCombo')
    .evaluate()
    .setWidth(440)
    .setHeight(390);
  SpreadsheetApp.getUi().showModalDialog(html, 'Vender Combo');
}

function getDatosParaVenderVaso() {
  const almacen = cargarAlmacenDesdeSheet();
  const licores = almacen.productos
    .filter(p => p.porciones > 0 && p.vasosDisponibles() > 0)
    .map(p => ({ id: p.id, nombre: p.nombre, precioVaso: p.precioVaso, vasosDisponibles: p.vasosDisponibles() }));
  const vasos = almacen.productos
    .filter(p => p.stock > 0)
    .map(p => ({ id: p.id, nombre: p.nombre, stock: p.stock }));
  return { licores, vasos };
}

function getDatosParaVenderCombo() {
  const almacen = cargarAlmacenDesdeSheet();
  const combos = almacen.combos
    .filter(c => c.disponibles() > 0)
    .map(c => ({ id: c.id, nombre: c.nombre, precio: c.precio, disponibles: c.disponibles() }));
  return { combos };
}

function ejecutarVentaVasoDesdeDialogo(idLicor, idVaso, cantidad) {
  return conLock(() => {
    try {
      const almacen = cargarAlmacenDesdeSheet();
      const licor = almacen.buscarProducto(idLicor);
      if (!licor) return { ok: false, mensaje: `No existe el producto ${idLicor}.` };
      const precioUnitario = licor.precioVaso;
      const nombre = licor.nombre;
      if (!almacen.venderVaso(idLicor, cantidad, idVaso)) {
        return { ok: false, mensaje: 'No se pudo realizar la venta. Revisa el stock disponible.' };
      }
      guardarAlmacenEnSheet(almacen);
      registrarVenta('Vaso', idLicor, nombre, cantidad, precioUnitario);
      return { ok: true, mensaje: `Venta registrada: ${cantidad} vaso(s) de ${nombre} por ${cantidad * precioUnitario} Bs.` };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

function ejecutarVentaComboDesdeDialogo(idCombo, cantidad) {
  return conLock(() => {
    try {
      const almacen = cargarAlmacenDesdeSheet();
      const combo = almacen.buscarCombo(idCombo);
      if (!combo) return { ok: false, mensaje: `No existe el combo ${idCombo}.` };
      const precioUnitario = combo.precio;
      const nombre = combo.nombre;
      if (!almacen.venderCombo(idCombo, cantidad)) {
        return { ok: false, mensaje: 'No se pudo vender el combo. Revisa el stock de sus componentes.' };
      }
      guardarAlmacenEnSheet(almacen);
      registrarVenta('Combo', idCombo, nombre, cantidad, precioUnitario);
      return { ok: true, mensaje: `Venta registrada: ${cantidad} combo(s) ${nombre} por ${cantidad * precioUnitario} Bs.` };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

function abrirDialogoAgregarProducto() {
  const html = HtmlService.createTemplateFromFile('DialogoAgregarProducto')
    .evaluate()
    .setWidth(480)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Agregar Producto');
}

function getDatosParaAgregarProducto() {
  const almacen = cargarAlmacenDesdeSheet();
  const categorias = almacen.categorias.map(c => ({
    nombre: c.nombre,
    letra: c.letra,
    ultimo: c.ultimo
  }));
  return { categorias };
}

function ejecutarAgregarProductoDesdeDialogo(nombre, categoria, precio, stock, porciones, precioVaso) {
  return conLock(() => {
    try {
      if (!nombre || nombre.trim() === '') return { ok: false, mensaje: 'El nombre no puede estar vacío.' };
      if (porciones < 0 || precioVaso < 0 || precio < 0 || stock < 0) return { ok: false, mensaje: 'Los valores numéricos no pueden ser negativos.' };
      if ((porciones > 0) !== (precioVaso > 0)) return { ok: false, mensaje: 'Porciones y precio por vaso deben ser ambos > 0 o ambos = 0.' };
      const almacen = cargarAlmacenDesdeSheet();
      if (almacen.productos.some(p => p.nombre.toLowerCase() === nombre.toLowerCase())) {
        return { ok: false, mensaje: `Ya existe un producto con el nombre "${nombre}".` };
      }
      const nuevoProd = new Producto(nombre, categoria, precio, stock, porciones, precioVaso);
      if (!almacen.agregarProducto(nuevoProd)) {
        return { ok: false, mensaje: 'No se pudo agregar. Revisa la categoría o si el producto ya existe.' };
      }
      guardarAlmacenEnSheet(almacen);
      return { ok: true, mensaje: `Producto "${nombre}" agregado con ID ${nuevoProd.id}.` };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

function conLock(fn) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); return fn(); }
  finally { lock.releaseLock(); }
}

function registrarEliminacion(tipo, id, nombre, categoria, stockAlBorrar, precio, detalle) {
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
}

function menuAgregarCategoria() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Agregar Categoría', 'Nombre de la categoría:', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const nombre = r1.getResponseText().trim();
  if (!nombre) { ui.alert('El nombre no puede estar vacío.'); return; }
  const r2 = ui.prompt('Agregar Categoría', 'Letra (1 carácter mayúscula):', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const letra = r2.getResponseText().trim().toUpperCase();
  return conLock(() => {
    const almacen = cargarAlmacenDesdeSheet();
    if (!almacen.agregarCategoria(nombre, letra)) {
      ui.alert('No se pudo agregar. Revisa los registros (Ver → Registros en el editor).');
      return;
    }
    guardarAlmacenEnSheet(almacen);
    ui.alert(`✅ Categoría "${nombre}" agregada con letra ${letra}.`);
  });
}

function configurarProteccionCategorias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(NOMBRE_HOJA_CATEGORIAS);
  if (!hoja) { SpreadsheetApp.getUi().alert('No existe la pestaña Categorias.'); return; }
  hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  hoja.protect().setDescription('Solo gestionable desde el menú Almacén.').setWarningOnly(false);
  const triggers = ScriptApp.getProjectTriggers();
  if (!triggers.some(t => t.getHandlerFunction() === 'onEditRevertCategorias')) {
    ScriptApp.newTrigger('onEditRevertCategorias').forSpreadsheet(ss).onEdit().create();
  }
  const almacen = cargarAlmacenDesdeSheet();
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
  const html = HtmlService.createTemplateFromFile('DialogoEliminarProducto')
    .evaluate().setWidth(520).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Eliminar Producto');
}

function getDatosParaEliminarProducto() {
  const almacen = cargarAlmacenDesdeSheet();
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
  return conLock(() => {
    try {
      const almacen = cargarAlmacenDesdeSheet();
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
      guardarAlmacenEnSheet(almacen);
      registrarEliminacion('Producto', idProducto, nombre, categoria, stockAlBorrar, precio, 'Manual');
      return { ok: true, mensaje: `Producto ${idProducto} (${nombre}) eliminado.` };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

function abrirDialogoEliminarCategoria() {
  const html = HtmlService.createTemplateFromFile('DialogoEliminarCategoria')
    .evaluate().setWidth(560).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Eliminar Categoría');
}

function getDatosParaEliminarCategoria() {
  const almacen = cargarAlmacenDesdeSheet();
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
  return conLock(() => {
    try {
      const almacen = cargarAlmacenDesdeSheet();
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

      guardarAlmacenEnSheet(almacen);

      combosAEliminar.forEach(c => {
        registrarEliminacion('Combo', c.id, c.nombre, '-', 0, c.precio, `Cascada por categoría ${nombreCategoria}`);
      });
      productosAEliminar.forEach(p => {
        registrarEliminacion('Producto', p.id, p.nombre, p.categoria, p.stock, p.precio, `Cascada por categoría ${nombreCategoria}`);
      });
      registrarEliminacion('Categoría', '-', nombreCategoria, '-', 0, 0,
        `${productosAEliminar.length} productos y ${combosAEliminar.length} combos eliminados`);

      return { ok: true, mensaje: `Categoría "${nombreCategoria}" eliminada. Se borraron ${productosAEliminar.length} producto(s) y ${combosAEliminar.length} combo(s).` };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

function menuVerificarIntegridad() {
  const almacen = cargarAlmacenDesdeSheet();
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Proyecto SIS101');
  if (!hoja) {
    SpreadsheetApp.getUi().alert('No existe la pestaña "Proyecto SIS101". Ejecuta primero Almacén → Instalar/Reinstalar Dashboard.');
    return;
  }
  ss.setActiveSheet(hoja);
}

function instalarDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('Proyecto SIS101');
  if (!hoja) hoja = ss.insertSheet('Proyecto SIS101', 0);
  hoja.clear();
  hoja.clearFormats();
  hoja.setHiddenGridlines(true);
  hoja.setColumnWidth(1, 25);
  for (let c = 2; c <= 9; c++) hoja.setColumnWidth(c, 115);

  hoja.getRange('B2:I2').merge();
  hoja.getRange('B2').setValue('PROYECTO SIS101  ·  SISTEMA DE ALMACÉN');
  hoja.getRange('B2').setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#1a73e8').setFontColor('#ffffff');
  hoja.setRowHeight(2, 60);

  hoja.getRange('B3:I3').merge();
  hoja.getRange('B3').setValue('Usa el menú  Almacén  para operar el sistema  ·  Los datos se actualizan automáticamente');
  hoja.getRange('B3').setFontSize(11).setFontStyle('italic').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontColor('#5f6368');
  hoja.setRowHeight(3, 30);

  hoja.getRange('B5:I5').merge();
  hoja.getRange('B5').setValue('RESUMEN DEL INVENTARIO');
  hoja.getRange('B5').setFontSize(13).setFontWeight('bold').setFontColor('#1a73e8').setHorizontalAlignment('left').setVerticalAlignment('middle');
  hoja.setRowHeight(5, 32);

  hoja.getRange('B7:C7').merge();
  hoja.getRange('B7').setValue('PRODUCTOS REGISTRADOS');
  hoja.getRange('B7').setFontSize(10).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e8f0fe');
  hoja.getRange('B8:C8').merge();
  hoja.getRange('B8').setFormula('=COUNTA(Productos!A2:A)');
  hoja.getRange('B8').setFontSize(28).setFontWeight('bold').setFontColor('#1a73e8').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e8f0fe');

  hoja.getRange('D7:E7').merge();
  hoja.getRange('D7').setValue('COMBOS REGISTRADOS');
  hoja.getRange('D7').setFontSize(10).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#fef7e0');
  hoja.getRange('D8:E8').merge();
  hoja.getRange('D8').setFormula('=COUNTA(Combos!A2:A)');
  hoja.getRange('D8').setFontSize(28).setFontWeight('bold').setFontColor('#b06000').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#fef7e0');

  hoja.getRange('F7:G7').merge();
  hoja.getRange('F7').setValue('VALOR INVENTARIO');
  hoja.getRange('F7').setFontSize(10).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e6f4ea');
  hoja.getRange('F8:G8').merge();
  hoja.getRange('F8').setFormula('=SUMPRODUCT(Productos!D2:D, Productos!E2:E)');
  hoja.getRange('F8').setFontSize(26).setFontWeight('bold').setFontColor('#137333').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e6f4ea').setNumberFormat('#,##0" Bs"');

  hoja.getRange('H7:I7').merge();
  hoja.getRange('H7').setValue('VENTAS HOY');
  hoja.getRange('H7').setFontSize(10).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#fce8e6');
  hoja.getRange('H8:I8').merge();
  hoja.getRange('H8').setFormula('=SUMIFS(Ventas!H2:H, Ventas!A2:A, TEXT(TODAY(), "yyyy-MM-dd"))');
  hoja.getRange('H8').setFontSize(26).setFontWeight('bold').setFontColor('#c5221f').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#fce8e6').setNumberFormat('#,##0" Bs"');
  hoja.setRowHeight(7, 25);
  hoja.setRowHeight(8, 55);

  hoja.getRange('B11:F11').merge();
  hoja.getRange('B11').setValue('PRODUCTOS POR AGOTARSE  (stock ≤ 10)');
  hoja.getRange('B11').setFontSize(13).setFontWeight('bold').setFontColor('#b06000').setHorizontalAlignment('left').setVerticalAlignment('middle');
  hoja.setRowHeight(11, 32);

  ['ID', 'Nombre', 'Categoría', 'Stock', 'Vasos Sueltos'].forEach((h, i) => {
    const celda = String.fromCharCode(66 + i) + '13';
    hoja.getRange(celda).setValue(h);
    hoja.getRange(celda).setFontSize(11).setFontWeight('bold').setFontColor('#ffffff').setBackground('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  hoja.setRowHeight(13, 28);
  hoja.getRange('B14').setFormula('=IFERROR(QUERY(Productos!A2:H, "select A, B, C, E, H where A is not null and E is not null and E <= 10 order by E asc limit 15", 0), "Sin productos por agotarse")');
  for (let f = 14; f <= 29; f++) hoja.setRowHeight(f, 24);

  hoja.getRange('B31:H31').merge();
  hoja.getRange('B31').setValue('ÚLTIMAS 10 VENTAS');
  hoja.getRange('B31').setFontSize(13).setFontWeight('bold').setFontColor('#1a73e8').setHorizontalAlignment('left').setVerticalAlignment('middle');
  hoja.setRowHeight(31, 32);

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
  hoja.getRange('B46').setFontSize(13).setFontWeight('bold').setFontColor('#137333').setHorizontalAlignment('left').setVerticalAlignment('middle');
  hoja.setRowHeight(46, 32);

  hoja.getRange('B48').setValue('Producto / Combo');
  hoja.getRange('C48').setValue('Cantidad vendida');
  ['B48', 'C48'].forEach(c => hoja.getRange(c).setFontSize(11).setFontWeight('bold').setFontColor('#ffffff').setBackground('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle'));
  hoja.setRowHeight(48, 28);
  hoja.getRange('B49').setFormula('=IFERROR(QUERY(Ventas!A2:H, "select E, sum(F) where E is not null group by E order by sum(F) desc limit 5", 0), "Sin ventas registradas")');
  for (let f = 49; f <= 54; f++) hoja.setRowHeight(f, 24);

  hoja.getRange('B56:I56').merge();
  hoja.getRange('B56').setValue('MANUAL DE USO RÁPIDO');
  hoja.getRange('B56').setFontSize(13).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1a73e8').setHorizontalAlignment('center').setVerticalAlignment('middle');
  hoja.setRowHeight(56, 34);

  const tareas = [
    ['🛒', 'Vender un vaso', 'Almacén → Ventas → Vender Vaso'],
    ['🍾', 'Vender un combo', 'Almacén → Ventas → Vender Combo'],
    ['➕', 'Agregar producto', 'Almacén → Productos → Agregar Producto'],
    ['✏️', 'Modificar producto', 'Almacén → Productos → Modificar Producto (precio, stock, nombre)'],
    ['🗑️', 'Eliminar producto', 'Almacén → Productos → Eliminar Producto'],
    ['🎁', 'Agregar combo', 'Almacén → Combos → Agregar Combo'],
    ['📁', 'Agregar categoría', 'Almacén → Categorías → Agregar Categoría'],
    ['🗂️', 'Eliminar categoría', 'Almacén → Categorías → Eliminar Categoría (borra productos y combos en cascada)'],
    ['🔍', 'Verificar integridad', 'Almacén → Verificar Integridad (revisa IDs, referencias y combos rotos)']
  ];

  tareas.forEach((t, i) => {
    const fila = 57 + i;
    hoja.getRange(`B${fila}:C${fila}`).merge();
    hoja.getRange(`B${fila}`).setValue(`${t[0]}  ${t[1]}`);
    hoja.getRange(`B${fila}`).setFontSize(11).setFontWeight('bold').setFontColor('#1a73e8').setBackground('#e8f0fe').setHorizontalAlignment('left').setVerticalAlignment('middle');
    hoja.getRange(`D${fila}:I${fila}`).merge();
    hoja.getRange(`D${fila}`).setValue(t[2]);
    hoja.getRange(`D${fila}`).setFontSize(11).setFontColor('#3c4043').setBackground('#ffffff').setHorizontalAlignment('left').setVerticalAlignment('middle');
    hoja.setRowHeight(fila, 26);
  });

  const filaAdv = 57 + tareas.length + 1;

  hoja.getRange(`B${filaAdv}:I${filaAdv}`).merge();
  hoja.getRange(`B${filaAdv}`).setValue('ADVERTENCIAS IMPORTANTES');
  hoja.getRange(`B${filaAdv}`).setFontSize(13).setFontWeight('bold').setFontColor('#ffffff').setBackground('#b06000').setHorizontalAlignment('center').setVerticalAlignment('middle');
  hoja.setRowHeight(filaAdv, 34);

  const advertencias = [
    '⚠️ Las categorías solo se modifican desde el menú Almacén. No las edites directamente en la hoja.',
    '⚠️ No cambies manualmente la columna ID en Productos ni en Combos. Los IDs no se reutilizan.',
    '⚠️ No reordenes filas ni columnas en Productos, Combos, Categorias ni Ventas.',
    '⚠️ Para cambiar precio, stock o nombre de un producto usa "Modificar Producto".',
    '⚠️ Las ventas quedan registradas en la pestaña Ventas como histórico. No las borres si quieres conservar el reporte.',
    '⚠️ Si algo sale mal, recupera el archivo con Google Sheets → Archivo → Historial de versiones.'
  ];

  advertencias.forEach((a, i) => {
    const fila = filaAdv + 2 + i;
    hoja.getRange(`B${fila}:I${fila}`).merge();
    hoja.getRange(`B${fila}`).setValue(a);
    hoja.getRange(`B${fila}`).setFontSize(11).setFontColor('#8a4b00').setBackground('#fef7e0').setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);
    hoja.setRowHeight(fila, 30);
  });

  SpreadsheetApp.flush();
  ss.setActiveSheet(hoja);
  SpreadsheetApp.getUi().alert('✅ Dashboard y manual instalados en la pestaña "Proyecto SIS101".');
}

function fixCompletoDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('Proyecto SIS101');
  if (!hoja) hoja = ss.insertSheet('Proyecto SIS101', 0);
  hoja.clear();
  hoja.clearFormats();
  hoja.setHiddenGridlines(true);
  hoja.setColumnWidth(1, 25);
  for (let c = 2; c <= 9; c++) hoja.setColumnWidth(c, 115);

  hoja.getRange('B2:I2').merge();
  hoja.getRange('B2').setValue('PROYECTO SIS101  ·  SISTEMA DE ALMACÉN');
  hoja.getRange('B2').setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#1a73e8').setFontColor('#ffffff');
  hoja.setRowHeight(2, 60);

  hoja.getRange('B3:I3').merge();
  hoja.getRange('B3').setValue('Usa el menú  Almacén  para operar el sistema  ·  Los datos se actualizan automáticamente');
  hoja.getRange('B3').setFontSize(11).setFontStyle('italic').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontColor('#5f6368');
  hoja.setRowHeight(3, 30);

  hoja.getRange('B5:I5').merge();
  hoja.getRange('B5').setValue('RESUMEN DEL INVENTARIO');
  hoja.getRange('B5').setFontSize(13).setFontWeight('bold').setFontColor('#1a73e8').setHorizontalAlignment('left').setVerticalAlignment('middle');
  hoja.setRowHeight(5, 32);

  hoja.getRange('B7:C7').merge();
  hoja.getRange('B7').setValue('PRODUCTOS REGISTRADOS');
  hoja.getRange('B7').setFontSize(10).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e8f0fe');
  hoja.getRange('B8:C8').merge();
  hoja.getRange('B8').setFormula('=COUNTA(Productos!A2:A)');
  hoja.getRange('B8').setFontSize(28).setFontWeight('bold').setFontColor('#1a73e8').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e8f0fe');

  hoja.getRange('D7:E7').merge();
  hoja.getRange('D7').setValue('COMBOS REGISTRADOS');
  hoja.getRange('D7').setFontSize(10).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#fef7e0');
  hoja.getRange('D8:E8').merge();
  hoja.getRange('D8').setFormula('=COUNTA(Combos!A2:A)');
  hoja.getRange('D8').setFontSize(28).setFontWeight('bold').setFontColor('#b06000').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#fef7e0');

  hoja.getRange('F7:G7').merge();
  hoja.getRange('F7').setValue('VALOR INVENTARIO');
  hoja.getRange('F7').setFontSize(10).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e6f4ea');
  hoja.getRange('F8:G8').merge();
  hoja.getRange('F8').setFormula('=SUMPRODUCT(Productos!D2:D, Productos!E2:E)');
  hoja.getRange('F8').setFontSize(26).setFontWeight('bold').setFontColor('#137333').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e6f4ea').setNumberFormat('#,##0" Bs"');

  hoja.getRange('H7:I7').merge();
  hoja.getRange('H7').setValue('VENTAS HOY');
  hoja.getRange('H7').setFontSize(10).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#fce8e6');
  hoja.getRange('H8:I8').merge();
  hoja.getRange('H8').setFormula('=SUMIFS(Ventas!H2:H, Ventas!A2:A, TEXT(TODAY(), "dd/MM/yyyy"))');
  hoja.getRange('H8').setFontSize(26).setFontWeight('bold').setFontColor('#c5221f').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#fce8e6').setNumberFormat('#,##0" Bs"');
  hoja.setRowHeight(7, 25);
  hoja.setRowHeight(8, 55);

  hoja.getRange('B11:F11').merge();
  hoja.getRange('B11').setValue('PRODUCTOS POR AGOTARSE  (stock ≤ 10)');
  hoja.getRange('B11').setFontSize(13).setFontWeight('bold').setFontColor('#b06000').setHorizontalAlignment('left').setVerticalAlignment('middle');
  hoja.setRowHeight(11, 32);

  ['ID', 'Nombre', 'Categoría', 'Stock', 'Vasos Sueltos'].forEach((h, i) => {
    const celda = String.fromCharCode(66 + i) + '13';
    hoja.getRange(celda).setValue(h);
    hoja.getRange(celda).setFontSize(11).setFontWeight('bold').setFontColor('#ffffff').setBackground('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  hoja.setRowHeight(13, 28);

  hoja.getRange('B14').setFormula(
    '=IFERROR(QUERY(Productos!A2:H, "select A, B, C, E, H where A is not null and E is not null and E <= 10 order by E asc limit 15", 0), "Sin productos por agotarse")'
  );

  for (let f = 14; f <= 29; f++) hoja.setRowHeight(f, 24);

  hoja.getRange('B31:H31').merge();
  hoja.getRange('B31').setValue('ÚLTIMAS 10 VENTAS');
  hoja.getRange('B31').setFontSize(13).setFontWeight('bold').setFontColor('#1a73e8').setHorizontalAlignment('left').setVerticalAlignment('middle');
  hoja.setRowHeight(31, 32);

  ['Fecha', 'Hora', 'Tipo', 'ID', 'Nombre', 'Cant.', 'Total'].forEach((h, i) => {
    const celda = String.fromCharCode(66 + i) + '33';
    hoja.getRange(celda).setValue(h);
    hoja.getRange(celda).setFontSize(11).setFontWeight('bold').setFontColor('#ffffff').setBackground('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  hoja.setRowHeight(33, 28);

  hoja.getRange('B34').setFormula(
    '=IFERROR(QUERY(Ventas!A2:H, "select A, B, C, D, E, F, H order by A desc, B desc limit 10", 0), "Sin ventas registradas")'
  );
  for (let f = 34; f <= 44; f++) hoja.setRowHeight(f, 24);

  hoja.getRange('B46:C46').merge();
  hoja.getRange('B46').setValue('TOP 5 MÁS VENDIDOS');
  hoja.getRange('B46').setFontSize(13).setFontWeight('bold').setFontColor('#137333').setHorizontalAlignment('left').setVerticalAlignment('middle');
  hoja.setRowHeight(46, 32);

  hoja.getRange('B48').setValue('Producto / Combo');
  hoja.getRange('C48').setValue('Cantidad vendida');
  ['B48', 'C48'].forEach(c => hoja.getRange(c).setFontSize(11).setFontWeight('bold').setFontColor('#ffffff').setBackground('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle'));
  hoja.setRowHeight(48, 28);

  hoja.getRange('B49').setFormula(
    '=IFERROR(QUERY(Ventas!A2:H, "select E, sum(F) where E is not null group by E order by sum(F) desc limit 5 label sum(F) \'Cantidad\'", 0), "Sin ventas registradas")'
  );
  for (let f = 49; f <= 54; f++) hoja.setRowHeight(f, 24);

  SpreadsheetApp.flush();
  Logger.log('Dashboard reacomodado. B14: ' + hoja.getRange('B14').getDisplayValue());
  ss.setActiveSheet(hoja);
}

function abrirDialogoAgregarCombo() {
  const html = HtmlService.createTemplateFromFile('DialogoAgregarCombo')
    .evaluate().setWidth(520).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Agregar Combo');
}

function getDatosParaAgregarCombo() {
  const almacen = cargarAlmacenDesdeSheet();
  const productos = almacen.productos.map(p => ({
    id: p.id,
    nombre: p.nombre,
    stock: p.stock,
    categoria: p.categoria
  }));
  const nombresExistentes = almacen.combos.map(c => c.nombre);
  return { productos, nombresExistentes };
}

function ejecutarAgregarComboDesdeDialogo(nombre, precio, componentes) {
  return conLock(() => {
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

      const almacen = cargarAlmacenDesdeSheet();

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

      guardarAlmacenEnSheet(almacen);
      return { ok: true, mensaje: `Combo "${nombre}" agregado con ID ${nuevoCombo.id}.` };
    } catch (e) {
      return { ok: false, mensaje: 'Error: ' + e.message };
    }
  });
}

function abrirDialogoModificarProducto() {
  const html = HtmlService.createTemplateFromFile('DialogoModificarProducto')
    .evaluate().setWidth(520).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Modificar Producto');
}

function getDatosParaModificarProducto() {
  const almacen = cargarAlmacenDesdeSheet();
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
  return conLock(() => {
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

      const almacen = cargarAlmacenDesdeSheet();
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

      guardarAlmacenEnSheet(almacen);

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

function compactarTablasDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Proyecto SIS101');
  if (!hoja) return;
  if (hoja.getLastRow() < 56) return;
  const b11 = String(hoja.getRange('B11').getValue() || '');
  if (b11.indexOf('PRODUCTOS POR AGOTARSE') === -1) return;

  hoja.showRows(1, hoja.getMaxRows());

  let lastProd = 13;
  const prodData = hoja.getRange('B14:F28').getValues();
  for (let i = 0; i < prodData.length; i++) {
    const tieneContenido = prodData[i].some(v => v !== '' && v !== null);
    if (tieneContenido) lastProd = 14 + i;
    else break;
  }
  if (lastProd < 28) hoja.hideRows(lastProd + 1, 30 - lastProd);

  let lastVentas = 33;
  const ventasData = hoja.getRange('B34:H43').getValues();
  for (let i = 0; i < ventasData.length; i++) {
    const tieneContenido = ventasData[i].some(v => v !== '' && v !== null);
    if (tieneContenido) lastVentas = 34 + i;
    else break;
  }
  if (lastVentas < 43) hoja.hideRows(lastVentas + 1, 45 - lastVentas);

  let lastTop = 48;
  const topData = hoja.getRange('B49:C53').getValues();
  for (let i = 0; i < topData.length; i++) {
    const tieneContenido = topData[i].some(v => v !== '' && v !== null);
    if (tieneContenido) lastTop = 49 + i;
    else break;
  }
  if (lastTop < 53) hoja.hideRows(lastTop + 1, 55 - lastTop);

  SpreadsheetApp.flush();
}

function fixTopVentas() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Proyecto SIS101');
  if (!hoja) { Logger.log('❌ No existe dashboard'); return; }
  
  hoja.getRange('B49').setFormula(
    '=IFERROR(QUERY(Ventas!A2:H, "select E, sum(F) where E is not null group by E order by sum(F) desc limit 5", 0), "Sin ventas registradas")'
  );
  SpreadsheetApp.flush();
  
  Logger.log('B49: "' + hoja.getRange('B49').getDisplayValue() + '"');
  Logger.log('C49: "' + hoja.getRange('C49').getDisplayValue() + '"');
  Logger.log('B50: "' + hoja.getRange('B50').getDisplayValue() + '"');
  Logger.log('C50: "' + hoja.getRange('C50').getDisplayValue() + '"');
  
  compactarTablasDashboard();
}
