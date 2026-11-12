const STORAGE_KEY = "sis101-almacen-v1";
const LOW_STOCK_LIMIT = 5;
const APPS_SCRIPT_MODE = typeof google !== "undefined" && Boolean(google.script?.run);
const sectionNames = {
  dashboard: "Resumen",
  productos: "Productos",
  combos: "Combos",
  ventas: "Ventas",
  historial: "Historial de ventas",
  categorias: "Categorías",
  usuarios: "Cuentas",
};
const currency = new Intl.NumberFormat("es-BO", { style: "currency", currency: "BOB", minimumFractionDigits: 2 });
let almacen = new Almacen();
let salesHistory = [];
let paymentQrUrl = "";
let paymentQrImageData = "";
let paymentQrError = "";
let currentView = "dashboard";
let toastTimer;
let sheetVersion = 0;
let sessionToken = "";
let currentUser = null;
let userAccounts = [];

function isAdministrator() {
  return !APPS_SCRIPT_MODE || currentUser?.rol === "ADMINISTRADOR";
}

function showLogin(message = "") {
  sessionToken = "";
  currentUser = null;
  document.getElementById("app-shell").hidden = true;
  document.getElementById("login-screen").hidden = false;
  document.getElementById("modal-backdrop").hidden = true;
  document.getElementById("login-password").value = "";
  const error = document.getElementById("login-error");
  error.textContent = message;
  error.hidden = !message;
  document.getElementById("login-username").focus();
}

function handleSessionError(message) {
  if (!APPS_SCRIPT_MODE) return false;
  const text = String(message || "");
  if (!/sesión|inicia sesión|cuenta está desactivada/i.test(text)) return false;
  showLogin(text);
  return true;
}

function showAuthenticatedApp(user) {
  currentUser = user;
  document.getElementById("login-screen").hidden = true;
  document.getElementById("app-shell").hidden = false;
  document.getElementById("user-name").textContent = user.usuario;
  document.getElementById("user-role").textContent = user.rol === "ADMINISTRADOR" ? "Administrador" : "Empleado";
  document.getElementById("user-avatar").textContent = initials(user.usuario);
  document.querySelectorAll(".admin-only").forEach((element) => {
    element.hidden = !isAdministrator() || (!APPS_SCRIPT_MODE && element.dataset.view === "usuarios");
  });
  document.querySelectorAll(".employee-only").forEach((element) => {
    element.hidden = APPS_SCRIPT_MODE && isAdministrator();
  });
  currentView = isAdministrator() ? "dashboard" : "ventas";
}

function login(username, password) {
  const submit = document.querySelector("#login-form button[type=submit]");
  const error = document.getElementById("login-error");
  submit.disabled = true;
  error.hidden = true;
  google.script.run
    .withSuccessHandler((result) => {
      submit.disabled = false;
      if (!result.ok) {
        error.textContent = result.mensaje || "No se pudo iniciar sesión.";
        error.hidden = false;
        return;
      }
      sessionToken = result.token;
      showAuthenticatedApp({ usuario: result.usuario, rol: result.rol });
      render();
      loadData();
    })
    .withFailureHandler((failure) => {
      submit.disabled = false;
      error.textContent = failure.message || "No se pudo iniciar sesión.";
      error.hidden = false;
    })
    .apiIniciarSesion(username, password);
}

function logout() {
  const token = sessionToken;
  sessionToken = "";
  currentUser = null;
  if (APPS_SCRIPT_MODE && token) google.script.run.apiCerrarSesion(token);
  showLogin();
}

function escapeHTML(value) {
  return String(value ?? "").replace(/[&<>"']/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  })[character]);
}

function normalizeImageUrl(value) {
  const imageUrl = String(value ?? "").trim();
  if (!imageUrl) return "";
  try {
    const url = new URL(imageUrl);
    return url.protocol === "http:" || url.protocol === "https:" ? url.href : "";
  } catch {
    return "";
  }
}

function configurePaymentQr(value) {
  paymentQrUrl = normalizeImageUrl(value);
  paymentQrImageData = "";
  paymentQrError = "";
  if (!paymentQrUrl) return;
  if (typeof qrcode !== "function") {
    paymentQrError = "No se cargó el generador de códigos QR.";
    return;
  }
  try {
    const qr = qrcode(0, "M");
    qr.addData(paymentQrUrl, "Byte");
    qr.make();
    paymentQrImageData = qr.createDataURL(5, 4);
  } catch {
    paymentQrError = "No se pudo generar el QR. Revisa que el enlace no sea demasiado largo.";
  }
}

function productAvatar(product) {
  const fallback = escapeHTML(initials(product.nombre));
  const imageUrl = normalizeImageUrl(product.imagen);
  if (!imageUrl) return `<span class="item-avatar">${fallback}</span>`;
  return `<span class="item-avatar product-avatar"><img class="product-thumbnail" src="${escapeHTML(imageUrl)}" data-fallback="${fallback}" alt="" loading="lazy"></span>`;
}

function bindProductImageFallbacks(root) {
  root.querySelectorAll(".product-thumbnail").forEach((image) => {
    image.addEventListener("error", () => {
      image.parentElement.textContent = image.dataset.fallback || "";
      image.parentElement.classList.remove("product-avatar");
    }, { once: true });
  });
}

function loadData() {
  if (APPS_SCRIPT_MODE) {
    document.getElementById("storage-label").textContent = "Conectando con Google Sheets...";
    google.script.run
      .withSuccessHandler((data) => {
        applySheetData(data);
        document.getElementById("storage-label").textContent = "Conectado a Google Sheets";
        render();
      })
      .withFailureHandler((error) => {
        document.getElementById("storage-label").textContent = "Sin conexión con Google Sheets";
        if (handleSessionError(error.message)) return;
        showToast(`No se pudo cargar la hoja: ${error.message}`, true);
      })
      .apiObtenerDatosWeb(sessionToken);
    return;
  }
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "null");
    if (!saved) return;
    almacen.categorias = Array.isArray(saved.categorias) && saved.categorias.length
      ? saved.categorias
      : almacen.categorias;
    almacen.productos = (saved.productos || []).map((data) => {
      const product = new Producto(data.nombre, data.categoria, Number(data.precio), Number(data.stock), Number(data.porciones), Number(data.precioVaso));
      product.id = data.id;
      product.vasosSueltos = Number(data.vasosSueltos || 0);
      product.imagen = String(data.imagen || "");
      return product;
    });
    almacen.ultimoCombo = Number(saved.ultimoCombo || 0);
    almacen.categorias.forEach((category) => { category.ultimo = Number(category.ultimo || 0); });
    almacen.combos = (saved.combos || []).map((data) => {
      const components = data.componentes
        .map((part) => ({ producto: almacen.buscarProducto(part.id), cantidad: Number(part.cantidad) }))
        .filter((part) => part.producto);
      const combo = new Combo(data.nombre, Number(data.precio), components);
      combo.id = data.id;
      combo.componentesRotos = data.componentesRotos || [];
      combo._textoOriginal = data.textoOriginal || null;
      return combo;
    });
    salesHistory = Array.isArray(saved.ventas) ? saved.ventas : [];
  } catch (error) {
    localStorage.removeItem(STORAGE_KEY);
    showToast("No se pudo leer el respaldo local. Se inició un almacén vacío.", true);
  }
}

function saveData() {
  const data = {
    categorias: almacen.categorias,
    productos: almacen.productos.map((product) => ({ ...product })),
    combos: almacen.combos.map((combo) => ({
      id: combo.id,
      nombre: combo.nombre,
      precio: combo.precio,
      componentes: combo.componentes.map((part) => ({ id: part.producto.id, cantidad: part.cantidad })),
      componentesRotos: combo.componentesRotos || [],
      textoOriginal: combo._textoOriginal || null,
    })),
    ultimoCombo: almacen.ultimoCombo,
    ventas: salesHistory,
  };
  if (APPS_SCRIPT_MODE) {
    if (!isAdministrator()) {
      showToast("Tu cuenta solo puede registrar ventas.", true);
      return;
    }
    const inventory = {
      categorias: data.categorias,
      productos: data.productos,
      combos: data.combos,
      ultimoCombo: data.ultimoCombo,
    };
    google.script.run
      .withSuccessHandler((result) => {
        if (!result.ok) {
          if (result.datos) applySheetData(result.datos);
          if (handleSessionError(result.mensaje)) return;
          showToast(result.mensaje || "No se pudo guardar en Google Sheets.", true);
          render();
          return;
        }
        applySheetData(result.datos);
        document.getElementById("storage-label").textContent = "Conectado a Google Sheets";
        render();
      })
      .withFailureHandler((error) => {
        if (handleSessionError(error.message)) return;
        showToast(`Error al guardar en Google Sheets: ${error.message}`, true);
      })
      .apiGuardarInventarioWeb(sessionToken, inventory, sheetVersion);
    updateCounts();
    return;
  }
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  updateCounts();
}

function applySheetData(data) {
  almacen = new Almacen();
  almacen.categorias = data.categorias || [];
  almacen.productos = (data.productos || []).map((item) => {
    const product = new Producto(item.nombre, item.categoria, Number(item.precio), Number(item.stock), Number(item.porciones), Number(item.precioVaso));
    product.id = item.id;
    product.vasosSueltos = Number(item.vasosSueltos) || 0;
    product.imagen = String(item.imagen || "");
    return product;
  });
  almacen.combos = (data.combos || []).map((item) => {
    const components = item.componentes
      .map((part) => ({ producto: almacen.buscarProducto(part.id), cantidad: Number(part.cantidad) }))
      .filter((part) => part.producto);
    const combo = new Combo(item.nombre, Number(item.precio), components);
    combo.id = item.id;
    combo.componentesRotos = item.componentesRotos || [];
    combo._textoOriginal = item.textoOriginal || null;
    return combo;
  });
  almacen.ultimoCombo = Number(data.ultimoCombo) || 0;
  sheetVersion = String(data.version || "");
  configurePaymentQr(data.qrPagoUrl || data.enlacePagoUrl || "");
  salesHistory = data.ventas || [];
}

function updateCounts() {
  document.getElementById("nav-product-count").textContent = almacen.productos.length;
  document.getElementById("nav-combo-count").textContent = almacen.combos.length;
}

function money(value) {
  return currency.format(Number(value) || 0);
}

function showToast(message, isError = false) {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.classList.toggle("error", isError);
  toast.classList.add("visible");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove("visible"), 2800);
}

function initials(name) {
  return String(name).trim().split(/\s+/).slice(0, 2).map((part) => part[0] || "").join("").toUpperCase();
}

function render() {
  if (APPS_SCRIPT_MODE && !currentUser) return;
  if (APPS_SCRIPT_MODE && isAdministrator() && currentView === "ventas") currentView = "dashboard";
  if (!isAdministrator() && !["ventas", "productos", "combos", "historial"].includes(currentView)) currentView = "ventas";
  document.querySelectorAll(".admin-only").forEach((element) => {
    element.hidden = !isAdministrator() || (!APPS_SCRIPT_MODE && element.dataset.view === "usuarios");
  });
  document.querySelectorAll(".employee-only").forEach((element) => {
    element.hidden = APPS_SCRIPT_MODE && isAdministrator();
  });
  document.getElementById("current-section").textContent = sectionNames[currentView];
  document.querySelectorAll(".nav-item").forEach((item) => item.classList.toggle("active", item.dataset.view === currentView));
  const views = {
    dashboard: renderDashboard,
    productos: renderProducts,
    combos: renderCombos,
    ventas: renderSales,
    historial: renderSalesHistory,
    categorias: renderCategories,
    usuarios: renderUsers,
  };
  document.getElementById("page-content").innerHTML = views[currentView]();
  bindProductImageFallbacks(document.getElementById("page-content"));
  if (currentView === "ventas") bindSales();
  if (currentView === "usuarios") loadUsers();
  updateCounts();
}

function pageHeading(eyebrow, title, description, button = "") {
  return `<div class="page-heading"><div><p class="eyebrow">${eyebrow}</p><h1>${title}</h1><p class="page-subtitle">${description}</p></div>${button}</div>`;
}

function productRow(product, actions = false) {
  const stockClass = product.stock <= LOW_STOCK_LIMIT ? "stock-low" : "";
  const stockLabel = product.stock <= 0 ? '<span class="pill danger">Agotado</span>'
    : product.stock <= LOW_STOCK_LIMIT ? '<span class="pill warning">Bajo</span>' : '<span class="pill">Disponible</span>';
  return `<tr>
    <td><div class="item-cell">${productAvatar(product)}<span class="item-name"><strong>${escapeHTML(product.nombre)}</strong><small>${escapeHTML(product.categoria)}</small></span></div></td>
    <td class="mono">${escapeHTML(product.id)}</td>
    <td><span class="stock-value ${stockClass}">${product.stock}</span> <span class="mono">unid.</span></td>
    <td>${product.porciones > 0 ? `<span class="mono">${product.vasosDisponibles()} vasos</span>` : '<span class="mono">—</span>'}</td>
    <td>${money(product.precio)}</td><td>${product.porciones}</td><td>${product.porciones > 0 ? money(product.precioVaso) : '<span class="mono">—</span>'}</td><td>${stockLabel}</td>
    ${actions ? `<td><div class="row-actions"><button class="row-action" data-action="edit-product" data-id="${escapeHTML(product.id)}" title="Editar producto" aria-label="Editar ${escapeHTML(product.nombre)}">✎</button><button class="row-action" data-action="add-stock" data-id="${escapeHTML(product.id)}" title="Aumentar stock" aria-label="Aumentar stock de ${escapeHTML(product.nombre)}">＋</button><button class="row-action delete" data-action="delete-product" data-id="${escapeHTML(product.id)}" title="Eliminar producto" aria-label="Eliminar ${escapeHTML(product.nombre)}">×</button></div></td>` : ""}
  </tr>`;
}

function emptyTable(message, detail, action = "", columns = 8) {
  return `<tr><td colspan="${columns}"><div class="empty-state"><strong>${message}</strong><p>${detail}</p>${action}</div></td></tr>`;
}

function renderDashboard() {
  const inventoryUnits = almacen.productos.reduce((sum, product) => sum + product.stock, 0);
  const inventoryValue = almacen.valorTotal();
  const totalSales = salesHistory.reduce((sum, sale) => sum + sale.total, 0);
  const categoryRows = almacen.categorias.map((category) => {
    const count = almacen.filtrarPorCategoria(category.nombre).length;
    return `<div class="category-row"><div class="category-label"><span class="category-letter">${escapeHTML(category.letra)}</span>${escapeHTML(category.nombre)}</div><span class="category-count">${count} ${count === 1 ? "producto" : "productos"}</span></div>`;
  }).join("");
  const inventoryRows = almacen.productos.slice(-6).reverse().map((product) => `<tr>
    <td><div class="item-cell">${productAvatar(product)}<span class="item-name"><strong>${escapeHTML(product.nombre)}</strong><small>${escapeHTML(product.id)}</small></span></div></td>
    <td><span class="category-tag">${escapeHTML(product.categoria)}</span></td>
    <td><span class="stock-value">${product.stock}</span> <span class="mono">unid.</span></td>
    <td class="sale-price">${money(product.precio)}</td>
  </tr>`).join("");
  const productCountLabel = almacen.productos.length === 1 ? "producto registrado" : "productos registrados";
  return `${pageHeading("PANEL DE CONTROL", "Resumen del almacén", "Inventario, disponibilidad y actividad en un solo lugar.", '<button class="button button-primary" data-action="add-product"><span class="button-plus">+</span> Agregar producto</button>')}
    <section class="metric-grid" aria-label="Indicadores del almacén">
      <article class="metric-card metric-highlight"><div class="metric-label">Valor del inventario <span class="metric-icon">◈</span></div><div class="metric-value">${money(inventoryValue)}</div><div class="metric-note">${inventoryUnits} unidades registradas</div></article>
      <article class="metric-card"><div class="metric-label">Productos <span class="metric-icon">▤</span></div><div class="metric-value">${almacen.productos.length}</div><div class="metric-note">En ${almacen.categorias.length} categorías</div></article>
      <article class="metric-card"><div class="metric-label">Combos <span class="metric-icon">▦</span></div><div class="metric-value">${almacen.combos.length}</div><div class="metric-note">${almacen.combos.filter((combo) => !combo.componentesRotos?.length && combo.disponibles() > 0).length} con stock disponible</div></article>
      <article class="metric-card"><div class="metric-label">Ventas registradas <span class="metric-icon">↗</span></div><div class="metric-value">${money(totalSales)}</div><div class="metric-note">${salesHistory.length} operaciones guardadas</div></article>
    </section>
    <div class="dashboard-grid">
      <section class="panel dashboard-inventory-panel"><header class="panel-header"><div><h2>Inventario reciente</h2><p>${almacen.productos.length} ${productCountLabel}</p></div><button class="text-link" data-view="productos">Ver inventario <span aria-hidden="true">→</span></button></header>
        <div class="table-wrap"><table><thead><tr><th>Producto</th><th>Categoría</th><th>Stock</th><th>Precio</th></tr></thead><tbody>${inventoryRows || emptyTable("Inventario vacío", "Los productos aparecerán aquí.", "", 4)}</tbody></table></div>
      </section>
      <section class="panel"><header class="panel-header"><div><h2>Categorías</h2><p>Distribución del inventario</p></div><button class="text-link" data-view="categorias">Administrar <span aria-hidden="true">→</span></button></header><div class="category-list">${categoryRows || '<div class="empty-state">Aún no hay categorías.</div>'}</div></section>
    </div>`;
}

function renderProducts() {
  const canEdit = isAdministrator();
  const rows = almacen.productos.map((product) => productRow(product, canEdit)).join("");
  const options = ['<option value="">Todas las categorías</option>', ...almacen.categorias.map((category) => `<option value="${escapeHTML(category.nombre)}">${escapeHTML(category.nombre)}</option>`)].join("");
  const emptyAction = canEdit ? '<button class="button button-primary button-small" data-action="add-product">Agregar producto</button>' : "";
  const actionHeader = canEdit ? "<th></th>" : "";
  const description = canEdit
    ? "Administra existencias, precios y productos por vaso."
    : "Consulta productos, precios y existencias disponibles para la venta.";
  const exportButton = canEdit ? '<button class="button button-secondary button-small" data-action="export-csv">↓ Exportar CSV</button>' : "";
  return `${pageHeading("INVENTARIO", "Productos", description, canEdit ? '<button class="button button-primary" data-action="add-product"><span class="button-plus">+</span> Nuevo producto</button>' : "")}
    <div class="toolbar"><label class="search-box"><span>⌕</span><input class="field" id="product-search" type="search" placeholder="Buscar por nombre o ID..."></label><select class="select-field" id="category-filter" aria-label="Filtrar por categoría">${options}</select>${exportButton}</div>
    <section class="panel table-panel"><div class="table-wrap"><table><thead><tr><th>Producto</th><th>ID</th><th>Stock</th><th>Vasos disponibles</th><th>Precio unidad</th><th>Porciones/unidad</th><th>Precio vaso</th><th>Estado</th>${actionHeader}</tr></thead><tbody id="product-rows">${rows || emptyTable(canEdit ? "Tu inventario está vacío" : "No hay productos registrados", canEdit ? "Agrega productos para comenzar a controlar las existencias." : "Los productos aparecerán aquí cuando el administrador los registre.", emptyAction, canEdit ? 9 : 8)}</tbody></table></div><footer class="table-footer"><span id="product-result-count">${almacen.productos.length} productos</span><span>Stock bajo: ${almacen.stockBajo(LOW_STOCK_LIMIT).length}</span></footer></section>`;
}

function renderCombos() {
  const canEdit = isAdministrator();
  const rows = almacen.combos.map((combo) => {
    const broken = combo.componentesRotos?.length > 0;
    const available = broken ? 0 : combo.disponibles();
    const parts = combo.componentes.map((part) => `${part.cantidad} × ${escapeHTML(part.producto.nombre)}`).join(" · ");
    const missing = broken ? ` · Faltan: ${combo.componentesRotos.map(escapeHTML).join(", ")}` : "";
    const actions = canEdit ? `<td><div class="row-actions"><button class="row-action" data-action="edit-combo" data-id="${escapeHTML(combo.id)}" title="Editar combo" aria-label="Editar ${escapeHTML(combo.nombre)}">✎</button><button class="row-action delete" data-action="delete-combo" data-id="${escapeHTML(combo.id)}" title="Eliminar combo" aria-label="Eliminar ${escapeHTML(combo.nombre)}">×</button></div></td>` : "";
    return `<tr><td><div class="item-cell"><span class="item-avatar combo-avatar">${escapeHTML(initials(combo.nombre))}</span><span class="item-name"><strong>${escapeHTML(combo.nombre)}</strong><small>${parts}${escapeHTML(missing)}</small></span></div></td><td class="mono">${escapeHTML(combo.id)}</td><td>${money(combo.precio)}</td><td><span class="stock-value ${available === 0 ? "stock-low" : ""}">${available}</span></td><td>${broken ? '<span class="pill danger">Componentes faltantes</span>' : available > 0 ? '<span class="pill">Disponible</span>' : '<span class="pill danger">Sin stock</span>'}</td>${actions}</tr>`;
  }).join("");
  const actionHeader = canEdit ? "<th></th>" : "";
  const emptyAction = canEdit ? '<button class="button button-primary button-small" data-action="add-combo">Crear primer combo</button>' : "";
  const description = canEdit
    ? "Arma paquetes con productos del inventario y controla su disponibilidad."
    : "Consulta los combos, sus componentes, precios y disponibilidad.";
  return `${pageHeading("CATÁLOGO", "Combos", description, canEdit ? '<button class="button button-primary" data-action="add-combo"><span class="button-plus">+</span> Nuevo combo</button>' : "")}
    <section class="panel table-panel"><div class="table-wrap"><table><thead><tr><th>Combo y componentes</th><th>ID</th><th>Precio</th><th>Disponibles</th><th>Estado</th>${actionHeader}</tr></thead><tbody>${rows || emptyTable("Aún no hay combos", canEdit ? "Crea un combo a partir de productos registrados." : "Los combos aparecerán aquí cuando el administrador los registre.", emptyAction, canEdit ? 6 : 5)}</tbody></table></div><footer class="table-footer"><span>${almacen.combos.length} combos registrados</span><span>La disponibilidad se calcula con el stock actual</span></footer></section>`;
}

function renderSales() {
  const productOptions = almacen.productos.filter((product) => product.stock > 0).map((product) => `<option value="${escapeHTML(product.id)}">${escapeHTML(product.id)} · ${escapeHTML(product.nombre)} (${product.stock} disponibles)</option>`).join("");
  const drinkOptions = almacen.productos.filter((product) => product.porciones > 0).map((product) => `<option value="${escapeHTML(product.id)}">${escapeHTML(product.id)} · ${escapeHTML(product.nombre)} (${product.vasosDisponibles()} vasos)</option>`).join("");
  const cupOptions = almacen.productos.filter((product) => product.nombre.toLowerCase().includes("vaso")).map((product) => `<option value="${escapeHTML(product.id)}">${escapeHTML(product.id)} · ${escapeHTML(product.nombre)} (${product.stock})</option>`).join("");
  const comboOptions = almacen.combos.filter((combo) => !combo.componentesRotos?.length).map((combo) => `<option value="${escapeHTML(combo.id)}">${escapeHTML(combo.id)} · ${escapeHTML(combo.nombre)} (${combo.disponibles()} disponibles)</option>`).join("");
  return `${pageHeading("PUNTO DE VENTA", "Ventas", "Registra ventas y descuenta existencias automáticamente.")}
    <div class="sales-grid">
      <form class="panel sale-panel" id="product-sale-form"><h2>Venta de producto</h2><p>Vende unidades completas y descuenta el stock.</p>
        <div class="form-field"><label for="sale-product">Producto</label><select class="select-field" id="sale-product" required>${productOptions || '<option value="">No hay productos con stock</option>'}</select></div>
        <div class="form-field"><label for="sale-product-quantity">Cantidad</label><input class="field" id="sale-product-quantity" type="number" min="1" step="1" value="1" required></div>
        <p class="form-hint" id="product-sale-hint">Selecciona un producto para ver su precio y stock.</p><div class="sale-total"><span>Total estimado</span><strong id="product-sale-total">${money(0)}</strong></div>${paymentMethodFields("product")}<button class="button button-primary" type="submit">Registrar venta</button>
      </form>
      <form class="panel sale-panel" id="glass-sale-form"><h2>Venta por vaso</h2><p>Descuenta las porciones y los vasos descartables.</p>
        <div class="form-field"><label for="sale-drink">Producto por vaso</label><select class="select-field" id="sale-drink" required>${drinkOptions || '<option value="">No hay productos configurados por vaso</option>'}</select></div>
        <div class="form-field"><label for="sale-cup">Vaso descartable</label><select class="select-field" id="sale-cup" required>${cupOptions || '<option value="">No se encontró un producto de vasos</option>'}</select></div>
        <div class="form-field"><label for="sale-glass-quantity">Cantidad de vasos</label><input class="field" id="sale-glass-quantity" type="number" min="1" step="1" value="1" required></div>
        <p class="form-hint" id="glass-sale-hint">Selecciona un producto para ver su precio y disponibilidad.</p><div class="sale-total"><span>Total estimado</span><strong id="glass-sale-total">${money(0)}</strong></div>${paymentMethodFields("glass")}<button class="button button-primary" type="submit">Registrar venta</button>
      </form>
      <form class="panel sale-panel" id="combo-sale-form"><h2>Venta de combo</h2><p>El stock de cada componente se actualizará al confirmar.</p>
        <div class="form-field"><label for="sale-combo">Combo</label><select class="select-field" id="sale-combo" required>${comboOptions || '<option value="">No hay combos registrados</option>'}</select></div>
        <div class="form-field"><label for="sale-combo-quantity">Cantidad</label><input class="field" id="sale-combo-quantity" type="number" min="1" step="1" value="1" required></div>
        <p class="form-hint" id="combo-sale-hint">Selecciona un combo para ver disponibilidad.</p><div class="sale-total"><span>Total estimado</span><strong id="combo-sale-total">${money(0)}</strong></div>${paymentMethodFields("combo")}<button class="button button-primary" type="submit">Registrar venta</button>
      </form>
    </div>`;
}

function paymentMethodFields(suffix) {
  const qrContent = paymentQrImageData
    ? `<img class="payment-qr-image" data-qr-image src="${escapeHTML(paymentQrImageData)}" alt="Código QR de pago"><p class="form-hint">Escanea el QR para abrir el enlace configurado. Verifica que recibiste el pago antes de confirmar la venta.</p><label class="payment-confirmation"><input type="checkbox" data-qr-confirmation> Confirmo que verifiqué la recepción del pago.</label>`
    : paymentQrError
      ? `<p class="form-error">${escapeHTML(paymentQrError)}</p>`
      : '<p class="form-hint payment-qr-missing">El administrador aún no configuró la URL del QR. Podrás usar este método cuando la agregue en Cuentas.</p>';
  return `<div class="form-field"><label for="sale-payment-${suffix}">Método de pago</label><select class="select-field" id="sale-payment-${suffix}" data-payment-method><option value="Efectivo">Efectivo</option><option value="QR">QR</option></select></div><div class="payment-qr" data-payment-qr hidden>${qrContent}</div>`;
}

function bindPaymentFields(form) {
  const methodSelect = form.querySelector("[data-payment-method]");
  const qrPanel = form.querySelector("[data-payment-qr]");
  if (!methodSelect || !qrPanel) return;
  const confirmation = form.querySelector("[data-qr-confirmation]");
  const updateVisibility = () => {
    qrPanel.hidden = methodSelect.value !== "QR";
    if (methodSelect.value !== "QR" && confirmation) confirmation.checked = false;
  };
  const clearConfirmation = () => {
    if (confirmation) confirmation.checked = false;
  };
  methodSelect.addEventListener("change", updateVisibility);
  form.querySelectorAll('input[type="number"]').forEach((input) => input.addEventListener("input", clearConfirmation));
  form.querySelectorAll("select:not([data-payment-method])").forEach((select) => select.addEventListener("change", clearConfirmation));
  updateVisibility();
}

function getPaymentDetails(form) {
  const method = form.querySelector("[data-payment-method]").value;
  if (method === "QR") {
    if (!paymentQrUrl) {
      showToast("El administrador todavía no configuró la URL del QR de pago.", true);
      return null;
    }
    if (!paymentQrImageData) {
      showToast(paymentQrError || "No se pudo generar el QR.", true);
      return null;
    }
    const qrImage = form.querySelector("[data-qr-image]");
    if (!qrImage?.complete || qrImage.naturalWidth === 0) {
      showToast("No se pudo cargar el QR generado. Pide al administrador que revise la URL.", true);
      return null;
    }
    if (!form.querySelector("[data-qr-confirmation]").checked) {
      showToast("Confirma que verificaste la recepción del pago.", true);
      return null;
    }
  }
  return { metodoPago: method, qrConfirmado: method === "QR", qrPagoUrl: method === "QR" ? paymentQrUrl : "" };
}

function renderSalesHistory() {
  const showSeller = APPS_SCRIPT_MODE && isAdministrator();
  const historyRows = salesHistory.map((sale) => {
    const type = sale.tipo === "Vaso" || sale.tipo === "vaso" ? "Venta por vaso"
      : sale.tipo === "Producto" ? "Venta de producto" : "Venta de combo";
    return `<tr><td class="mono">${escapeHTML(sale.fecha)}</td><td>${type}</td><td>${escapeHTML(sale.detalle)}</td><td>${Number(sale.cantidad) || 0}</td><td>${escapeHTML(sale.metodoPago || "—")}</td><td class="sale-price">${money(sale.total)}</td>${showSeller ? `<td>${escapeHTML(sale.usuario || "Sheets")}</td>` : ""}</tr>`;
  }).join("");
  const history = historyRows || emptyTable("Sin ventas registradas", "Las ventas guardadas en la hoja Ventas aparecerán aquí.", "", showSeller ? 7 : 6);
  const title = showSeller ? "Historial de ventas" : "Tus ventas";
  return `${pageHeading("OPERACIONES", title, "Consulta las ventas registradas en la aplicación.")}
    <section class="panel history-panel"><header class="panel-header"><div><h2>${title}</h2><p>Operaciones almacenadas en la hoja Ventas</p></div><span class="pill">${salesHistory.length} ventas</span></header><div class="table-wrap"><table><thead><tr><th>Fecha</th><th>Tipo</th><th>Detalle</th><th>Cantidad</th><th>Método de pago</th><th>Total</th>${showSeller ? "<th>Usuario</th>" : ""}</tr></thead><tbody>${history}</tbody></table></div></section>`;
}

function renderCategories() {
  const cards = almacen.categorias.map((category) => {
    const count = almacen.filtrarPorCategoria(category.nombre).length;
    return `<article class="category-card"><div class="category-card-top"><span class="category-card-letter">${escapeHTML(category.letra)}</span><button class="row-action delete" data-action="delete-category" data-name="${escapeHTML(category.nombre)}" title="Eliminar categoría" aria-label="Eliminar ${escapeHTML(category.nombre)}">×</button></div><h3>${escapeHTML(category.nombre)}</h3><p>${count} ${count === 1 ? "producto" : "productos"} · Prefijo ${escapeHTML(category.letra)}</p></article>`;
  }).join("");
  return `${pageHeading("ORGANIZACIÓN", "Categorías", "Agrupa productos y conserva sus prefijos de identificación.", '<button class="button button-primary" data-action="add-category"><span class="button-plus">+</span> Nueva categoría</button>')}
    <div class="category-cards">${cards}</div>`;
}

function renderUsers() {
  if (!APPS_SCRIPT_MODE) return `${pageHeading("ADMINISTRACIÓN", "Cuentas", "La administración de cuentas solo está disponible en la aplicación conectada a Google Sheets.")}`;
  return `${pageHeading("ADMINISTRACIÓN", "Cuentas", "Crea y administra las cuentas de los tres empleados.")}
    <section class="panel account-create-panel payment-settings-panel"><header class="panel-header"><div><h2>Pagos con QR</h2><p>Configura el enlace que se codificará en el QR.</p></div></header>
      <form id="payment-qr-form"><div class="form-field"><label for="payment-qr-url">URL de destino del QR</label><input class="field" id="payment-qr-url" type="url" maxlength="2048" value="${escapeHTML(paymentQrUrl)}" placeholder="https://ejemplo.com/video"><p class="form-hint">Pega una URL http o https, como un enlace a un video. El QR se genera en este dispositivo; la aplicación no verifica pagos automáticamente.</p></div>
        ${paymentQrImageData ? `<div class="payment-qr-preview"><img class="payment-qr-image" src="${escapeHTML(paymentQrImageData)}" alt="Vista previa del QR"><p class="form-hint">El QR abrirá la URL configurada.</p></div>` : paymentQrError ? `<p class="form-error">${escapeHTML(paymentQrError)}</p>` : ""}
        <button class="button button-primary" type="submit">Guardar QR</button>
      </form>
    </section>
    <section class="panel account-create-panel">
      <form id="create-employee-form">
        <div class="form-row"><div class="form-field"><label for="employee-username">Usuario</label><input class="field" id="employee-username" name="usuario" minlength="3" maxlength="32" pattern="[A-Za-z0-9._-]{3,32}" autocomplete="off" required></div>
        <div class="form-field"><label for="employee-password">Contraseña inicial</label><input class="field" id="employee-password" name="password" type="password" minlength="8" maxlength="128" autocomplete="new-password" required></div></div>
        <p class="form-hint">La contraseña debe tener 8 caracteres o más. Se permiten hasta tres cuentas de empleados activas.</p>
        <button class="button button-primary" type="submit">Crear cuenta de empleado</button>
      </form>
    </section>
    <section class="panel table-panel account-list-panel"><header class="panel-header"><div><h2>Cuentas existentes</h2><p>Los empleados pueden registrar ventas, pero no editar el inventario.</p></div><span class="pill" id="active-employee-count">Cargando…</span></header>
      <div class="table-wrap" id="account-table"><div class="empty-state">Cargando cuentas…</div></div>
    </section>`;
}

function loadUsers() {
  google.script.run
    .withSuccessHandler((accounts) => {
      userAccounts = accounts;
      renderUserRows();
    })
    .withFailureHandler((error) => {
      if (handleSessionError(error.message)) return;
      showToast(`No se pudieron cargar las cuentas: ${error.message}`, true);
    })
    .apiListarCuentas(sessionToken);

  document.getElementById("payment-qr-form").addEventListener("submit", (event) => {
    event.preventDefault();
    const input = document.getElementById("payment-qr-url");
    const rawUrl = input.value.trim();
    const qrUrl = normalizeImageUrl(rawUrl);
    if (rawUrl && !qrUrl) return showToast("La URL de destino del QR debe usar HTTP o HTTPS.", true);
    const submit = event.currentTarget.querySelector('button[type="submit"]');
    submit.disabled = true;
    google.script.run
      .withSuccessHandler((result) => {
        submit.disabled = false;
        if (!result.ok) {
          if (handleSessionError(result.mensaje)) return;
          return showToast(result.mensaje || "No se pudo guardar el QR.", true);
        }
        configurePaymentQr(result.qrPagoUrl || result.enlacePagoUrl || "");
        render();
        showToast(paymentQrUrl ? "QR de pago guardado." : "QR de pago eliminado.");
      })
      .withFailureHandler((error) => {
        submit.disabled = false;
        if (handleSessionError(error.message)) return;
        showToast(`No se pudo guardar el QR: ${error.message}`, true);
      })
      .apiGuardarQrPago(sessionToken, qrUrl);
  }, { once: true });

  document.getElementById("create-employee-form").addEventListener("submit", (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const username = document.getElementById("employee-username").value.trim();
    const password = document.getElementById("employee-password").value;
    const submit = form.querySelector('button[type="submit"]');
    submit.disabled = true;
    google.script.run
      .withSuccessHandler((result) => {
        submit.disabled = false;
        if (!result.ok) {
          if (handleSessionError(result.mensaje)) return;
          return showToast(result.mensaje, true);
        }
        userAccounts = result.cuentas;
        form.reset();
        renderUserRows();
        showToast(`Cuenta ${username} creada.`);
      })
      .withFailureHandler((error) => {
        submit.disabled = false;
        if (handleSessionError(error.message)) return;
        showToast(`No se pudo crear la cuenta: ${error.message}`, true);
      })
      .apiCrearCuentaEmpleado(sessionToken, username, password);
  }, { once: true });

  document.getElementById("account-table").addEventListener("click", (event) => {
    const button = event.target.closest("[data-account-action]");
    if (!button) return;
    const username = button.dataset.username;
    if (button.dataset.accountAction === "toggle-account") {
      updateEmployeeStatus(username, button.dataset.active !== "true");
    } else if (button.dataset.accountAction === "reset-password") {
      resetEmployeePasswordForm(username);
    } else if (button.dataset.accountAction === "delete-account") {
      confirmDeleteEmployee(username);
    }
  });
}

function renderUserRows() {
  const activeCount = userAccounts.filter((user) => user.rol === "EMPLEADO" && user.activo).length;
  document.getElementById("active-employee-count").textContent = `${activeCount}/3 activos`;
  const rows = userAccounts.map((user) => {
    const isEmployee = user.rol === "EMPLEADO";
    const roleLabel = isEmployee ? "Empleado" : "Administrador";
    const actions = isEmployee
      ? `<div class="row-actions"><button class="row-action" data-account-action="reset-password" data-username="${escapeHTML(user.usuario)}" title="Restablecer contraseña" aria-label="Restablecer contraseña de ${escapeHTML(user.usuario)}">⚿</button><button class="button button-small ${user.activo ? "button-secondary" : "button-primary"}" data-account-action="toggle-account" data-username="${escapeHTML(user.usuario)}" data-active="${user.activo}" type="button">${user.activo ? "Desactivar" : "Activar"}</button><button class="row-action delete" data-account-action="delete-account" data-username="${escapeHTML(user.usuario)}" title="Eliminar cuenta" aria-label="Eliminar cuenta de ${escapeHTML(user.usuario)}">×</button></div>`
      : '<span class="mono">Cuenta principal</span>';
    return `<tr><td><strong>${escapeHTML(user.usuario)}</strong></td><td>${roleLabel}</td><td><span class="pill ${user.activo ? "" : "danger"}">${user.activo ? "Activa" : "Inactiva"}</span></td><td class="mono">${escapeHTML(user.creado || "—")}</td><td>${actions}</td></tr>`;
  }).join("");
  document.getElementById("account-table").innerHTML = `<table><thead><tr><th>Usuario</th><th>Rol</th><th>Estado</th><th>Creada</th><th></th></tr></thead><tbody>${rows || emptyTable("Sin cuentas", "No se encontraron cuentas.", "", 5)}</tbody></table>`;
}

function confirmDeleteEmployee(username) {
  openModal("Eliminar cuenta de empleado", `<p>¿Eliminar permanentemente la cuenta <strong>${escapeHTML(username)}</strong>? Ya no podrá iniciar sesión y sus sesiones activas dejarán de funcionar.</p>
    <p class="form-hint">El historial de ventas se conserva. Esta acción no se puede deshacer.</p>
    <div class="form-actions"><button class="button button-secondary" type="button" data-action="close-modal">Cancelar</button><button class="button button-danger" id="confirm-delete-employee" type="button">Eliminar cuenta</button></div>`);
  document.getElementById("confirm-delete-employee").addEventListener("click", (event) => {
    const button = event.currentTarget;
    button.disabled = true;
    google.script.run
      .withSuccessHandler((result) => {
        if (!result.ok) {
          button.disabled = false;
          if (handleSessionError(result.mensaje)) return;
          return showToast(result.mensaje, true);
        }
        userAccounts = result.cuentas;
        closeModal();
        renderUserRows();
        showToast(`Cuenta ${username} eliminada.`);
      })
      .withFailureHandler((error) => {
        button.disabled = false;
        if (handleSessionError(error.message)) return;
        showToast(`No se pudo eliminar la cuenta: ${error.message}`, true);
      })
      .apiEliminarCuentaEmpleado(sessionToken, username);
  });
}

function updateEmployeeStatus(username, active) {
  google.script.run
    .withSuccessHandler((result) => {
      if (!result.ok) {
        if (handleSessionError(result.mensaje)) return;
        return showToast(result.mensaje, true);
      }
      userAccounts = result.cuentas;
      renderUserRows();
      showToast(active ? `Cuenta ${username} activada.` : `Cuenta ${username} desactivada.`);
    })
    .withFailureHandler((error) => {
      if (handleSessionError(error.message)) return;
      showToast(`No se pudo actualizar la cuenta: ${error.message}`, true);
    })
    .apiCambiarEstadoCuentaEmpleado(sessionToken, username, active);
}

function resetEmployeePasswordForm(username) {
  openModal(`Restablecer contraseña`, `<form id="reset-employee-password-form">
    <p class="form-hint">Asigna una contraseña temporal de al menos 8 caracteres para <strong>${escapeHTML(username)}</strong>.</p>
    <div class="form-field"><label for="reset-employee-password">Nueva contraseña</label><input class="field" id="reset-employee-password" type="password" minlength="8" maxlength="128" autocomplete="new-password" required></div>
    <div class="form-actions"><button class="button button-secondary" type="button" data-action="close-modal">Cancelar</button><button class="button button-primary" type="submit">Guardar contraseña</button></div>
  </form>`);
  document.getElementById("reset-employee-password-form").addEventListener("submit", (event) => {
    event.preventDefault();
    const password = document.getElementById("reset-employee-password").value;
    const submit = event.currentTarget.querySelector('button[type="submit"]');
    submit.disabled = true;
    google.script.run
      .withSuccessHandler((result) => {
        submit.disabled = false;
        if (!result.ok) {
          if (handleSessionError(result.mensaje)) return;
          return showToast(result.mensaje, true);
        }
        closeModal();
        showToast(`Contraseña de ${username} restablecida.`);
      })
      .withFailureHandler((error) => {
        submit.disabled = false;
        if (handleSessionError(error.message)) return;
        showToast(`No se pudo restablecer la contraseña: ${error.message}`, true);
      })
      .apiRestablecerClaveEmpleado(sessionToken, username, password);
  });
}

function openModal(title, content, eyebrow = "ALMACÉN") {
  document.getElementById("modal-eyebrow").textContent = eyebrow;
  document.getElementById("modal-title").textContent = title;
  document.getElementById("modal-body").innerHTML = content;
  document.getElementById("modal-backdrop").hidden = false;
  document.body.style.overflow = "hidden";
  const firstInput = document.querySelector("#modal-body input, #modal-body select");
  if (firstInput) firstInput.focus();
}

function closeModal() {
  document.getElementById("modal-backdrop").hidden = true;
  document.body.style.overflow = "";
}

function productForm(product = null) {
  if (almacen.categorias.length === 0) {
    showToast("Primero crea una categoría para agregar productos.", true);
    return;
  }
  const editing = Boolean(product);
  const categoryOptions = almacen.categorias.map((category) => `<option value="${escapeHTML(category.nombre)}" ${product?.categoria === category.nombre ? "selected" : ""}>${escapeHTML(category.nombre)}</option>`).join("");
  openModal(editing ? "Editar producto" : "Nuevo producto", `<form id="product-form">
    <div class="form-field"><label for="product-name">Nombre</label><input class="field" id="product-name" name="nombre" maxlength="70" value="${escapeHTML(product?.nombre || "")}" placeholder="Ej. Singani Casa Real" required></div>
    <div class="form-field"><label for="product-image">URL de imagen (opcional)</label><input class="field" id="product-image" name="imagen" type="url" maxlength="2048" value="${escapeHTML(product?.imagen || "")}" placeholder="https://ejemplo.com/imagen.jpg"><p class="form-hint">Usa un enlace público directo a la imagen.</p><div class="product-image-preview" id="product-image-preview" hidden><img id="product-image-preview-img" alt="Vista previa del producto"></div><p class="form-error" id="product-image-error" hidden>Ingresa una URL http o https válida.</p></div>
    <div class="form-row"><div class="form-field"><label for="product-category">Categoría</label><select class="select-field" id="product-category" name="categoria" ${editing ? "disabled" : ""}>${categoryOptions}</select></div><div class="form-field"><label for="product-price">Precio (Bs)</label><input class="field" id="product-price" name="precio" type="number" min="0" step="0.01" value="${product?.precio ?? ""}" required></div></div>
    <div class="form-row"><div class="form-field"><label for="product-stock">Stock actual</label><input class="field" id="product-stock" name="stock" type="number" min="0" step="1" value="${product?.stock ?? 0}" required></div><div class="form-field"><label for="product-portions">Porciones por unidad</label><input class="field" id="product-portions" name="porciones" type="number" min="0" step="1" value="${product?.porciones ?? 0}" required></div></div>
    <div class="form-field"><label for="product-glass-price">Precio por vaso (Bs)</label><input class="field" id="product-glass-price" name="precioVaso" type="number" min="0" step="0.01" value="${product?.precioVaso ?? 0}" required></div>
    ${editing ? `<p class="form-hint">ID ${escapeHTML(product.id)} · La categoría y el ID no se modifican.</p>` : '<p class="form-hint">El ID se asigna automáticamente según el prefijo de categoría.</p>'}
    <div class="form-actions"><button class="button button-secondary" type="button" data-action="close-modal">Cancelar</button><button class="button button-primary" type="submit">${editing ? "Guardar cambios" : "Crear producto"}</button></div>
  </form>`);
  const imageInput = document.getElementById("product-image");
  const preview = document.getElementById("product-image-preview");
  const previewImage = document.getElementById("product-image-preview-img");
  const imageError = document.getElementById("product-image-error");
  const updateImagePreview = () => {
    const rawUrl = imageInput.value.trim();
    const imageUrl = normalizeImageUrl(rawUrl);
    imageError.hidden = !rawUrl || Boolean(imageUrl);
    preview.hidden = !imageUrl;
    previewImage.onerror = () => {
      preview.hidden = true;
      imageError.hidden = false;
      imageError.textContent = "No se pudo cargar la imagen desde esa URL.";
    };
    if (imageUrl) previewImage.src = imageUrl;
    else previewImage.removeAttribute("src");
  };
  imageInput.addEventListener("input", updateImagePreview);
  previewImage.addEventListener("load", () => {
    imageError.hidden = true;
    preview.hidden = false;
  });
  updateImagePreview();
  document.getElementById("product-form").addEventListener("submit", (event) => {
    event.preventDefault();
    const form = new FormData(event.currentTarget);
    const nombre = String(form.get("nombre")).trim();
    const precio = Number(form.get("precio"));
    const stock = Number(form.get("stock"));
    const porciones = Number(form.get("porciones"));
    const precioVaso = Number(form.get("precioVaso"));
    const imageUrlInput = String(form.get("imagen") || "").trim();
    const imagen = normalizeImageUrl(imageUrlInput);
    if (imageUrlInput && !imagen) {
      showToast("La URL debe ser válida y usar HTTP o HTTPS.", true);
      return;
    }
    if (!nombre || [precio, stock, porciones, precioVaso].some((value) => !Number.isFinite(value) || value < 0)) {
      showToast("Revisa los datos: deben ser valores válidos y no negativos.", true);
      return;
    }
    if (editing) {
      product.modificar(nombre, precio, porciones, precioVaso, imagen);
      product.stock = stock;
      product.vasosSueltos = Math.min(product.vasosSueltos, porciones || 0);
      showToast("Producto actualizado.");
    } else {
      const duplicate = almacen.productos.some((item) => item.nombre.toLowerCase() === nombre.toLowerCase());
      if (duplicate) return showToast("Ya existe un producto con ese nombre.", true);
      const created = new Producto(nombre, form.get("categoria"), precio, stock, porciones, precioVaso);
      created.imagen = imagen;
      if (!almacen.agregarProducto(created)) return showToast("No se pudo crear el producto.", true);
      showToast(`Producto ${created.id} agregado.`);
    }
    saveData(); closeModal(); render();
  });
}

function addStockForm(product) {
  openModal("Aumentar stock", `<form id="stock-form"><p class="form-hint">${escapeHTML(product.nombre)} · ${escapeHTML(product.id)} · Stock actual: ${product.stock}</p><div class="form-field"><label for="stock-quantity">Unidades a agregar</label><input class="field" id="stock-quantity" type="number" min="1" step="1" value="1" required></div><div class="form-actions"><button class="button button-secondary" type="button" data-action="close-modal">Cancelar</button><button class="button button-primary" type="submit">Actualizar stock</button></div></form>`);
  document.getElementById("stock-form").addEventListener("submit", (event) => {
    event.preventDefault();
    const quantity = Number(document.getElementById("stock-quantity").value);
    if (!product.aumentarStock(quantity)) return showToast("La cantidad debe ser mayor a cero.", true);
    saveData(); closeModal(); render(); showToast("Stock actualizado.");
  });
}

function comboForm(combo = null) {
  const editing = Boolean(combo);
  if (almacen.productos.length === 0) return showToast("Agrega productos antes de crear un combo.", true);
  const missingComponents = combo?.componentesRotos?.length
    ? `<p class="form-hint">Faltan productos del combo original: ${combo.componentesRotos.map(escapeHTML).join(", ")}. Al guardar, el combo se reemplazará por los componentes seleccionados.</p>`
    : "";
  openModal(editing ? "Editar combo" : "Nuevo combo", `<form id="combo-form"><div class="form-field"><label for="combo-name">Nombre del combo</label><input class="field" id="combo-name" maxlength="70" value="${escapeHTML(combo?.nombre || "")}" placeholder="Ej. Combo de la casa" required></div><div class="form-field"><label for="combo-price">Precio del combo (Bs)</label><input class="field" id="combo-price" type="number" min="0" step="0.01" value="${combo?.precio ?? 0}" required></div><div class="form-field"><label>Componentes</label><div class="component-list" id="component-list"></div><button class="component-add" type="button" id="add-component">+ Añadir componente</button></div>${missingComponents}<p class="form-hint">Cada producto solo puede aparecer una vez. Los combos disponibles dependen del stock actual.</p><div class="form-actions"><button class="button button-secondary" type="button" data-action="close-modal">Cancelar</button><button class="button button-primary" type="submit">${editing ? "Guardar cambios" : "Crear combo"}</button></div></form>`);
  const componentList = document.getElementById("component-list");
  const addComponent = (selectedId = "", quantity = 1) => {
    const usedIds = [...componentList.querySelectorAll("select")].map((select) => select.value);
    const available = almacen.productos.filter((product) => !usedIds.includes(product.id));
    if (!available.length) return showToast("Ya agregaste todos los productos disponibles.", true);
    const row = document.createElement("div");
    row.className = "component-row";
    row.innerHTML = `<select class="select-field" required>${available.map((product) => `<option value="${escapeHTML(product.id)}" ${product.id === selectedId ? "selected" : ""}>${escapeHTML(product.id)} · ${escapeHTML(product.nombre)}</option>`).join("")}</select><input class="field" type="number" min="1" step="1" value="${quantity}" aria-label="Cantidad" required><button class="row-action delete" type="button" aria-label="Quitar componente">×</button>`;
    row.querySelector("button").addEventListener("click", () => row.remove());
    componentList.appendChild(row);
  };
  document.getElementById("add-component").addEventListener("click", addComponent);
  if (editing) {
    combo.componentes.forEach((part) => addComponent(part.producto.id, part.cantidad));
    if (!combo.componentes.length && !combo.componentesRotos?.length) addComponent();
  } else {
    addComponent();
  }
  document.getElementById("combo-form").addEventListener("submit", (event) => {
    event.preventDefault();
    const nombre = document.getElementById("combo-name").value.trim();
    const precio = Number(document.getElementById("combo-price").value);
    const components = [...componentList.querySelectorAll(".component-row")].map((row) => ({ producto: almacen.buscarProducto(row.querySelector("select").value), cantidad: Number(row.querySelector("input").value) }));
    const componentIds = components.map((part) => part.producto?.id);
    if (!nombre || !Number.isFinite(precio) || precio < 0 || !components.length || components.some((part) => !part.producto || !Number.isInteger(part.cantidad) || part.cantidad < 1)) return showToast("Completa el nombre, precio y componentes válidos.", true);
    if (componentIds.length !== new Set(componentIds).size) return showToast("Cada producto solo puede aparecer una vez en el combo.", true);
    if (almacen.combos.some((item) => item.id !== combo?.id && item.nombre.toLowerCase() === nombre.toLowerCase())) return showToast("Ya existe un combo con ese nombre.", true);
    let comboId;
    if (editing) {
      if (almacen.buscarCombo(combo.id) !== combo) return showToast("El combo ya no está disponible. Actualiza la lista e inténtalo de nuevo.", true);
      combo.modificar(nombre, precio, components);
      combo.componentesRotos = [];
      combo._textoOriginal = null;
      comboId = combo.id;
    } else {
      const newCombo = new Combo(nombre, precio, components);
      if (!almacen.agregarCombo(newCombo)) return showToast("No se pudo crear el combo.", true);
      comboId = newCombo.id;
    }
    saveData(); closeModal(); render(); showToast(editing ? `Combo ${comboId} actualizado.` : `Combo ${comboId} creado.`);
  });
}

function categoryForm() {
  openModal("Nueva categoría", `<form id="category-form"><div class="form-field"><label for="category-name">Nombre</label><input class="field" id="category-name" maxlength="35" placeholder="Ej. Snacks" required></div><div class="form-field"><label for="category-letter">Prefijo del ID</label><input class="field" id="category-letter" maxlength="1" pattern="[A-Za-z]" placeholder="S" required></div><p class="form-hint">Usa una letra única. La K está reservada para los combos.</p><div class="form-actions"><button class="button button-secondary" type="button" data-action="close-modal">Cancelar</button><button class="button button-primary" type="submit">Crear categoría</button></div></form>`);
  document.getElementById("category-form").addEventListener("submit", (event) => {
    event.preventDefault();
    const name = document.getElementById("category-name").value.trim();
    const letter = document.getElementById("category-letter").value.trim().toUpperCase();
    if (!name || !/^[A-Z]$/.test(letter)) return showToast("Escribe un nombre y una sola letra para el prefijo.", true);
    if (!almacen.agregarCategoria(name, letter)) return showToast("El nombre o prefijo ya existe, o la letra K está reservada.", true);
    saveData(); closeModal(); render(); showToast("Categoría agregada.");
  });
}

function confirmAction(title, detail, confirmLabel, action, danger = false) {
  openModal(title, `<p class="form-hint">${detail}</p><div class="form-actions"><button class="button button-secondary" type="button" data-action="close-modal">Cancelar</button><button class="button ${danger ? "button-danger" : "button-primary"}" type="button" id="confirm-action">${confirmLabel}</button></div>`);
  document.getElementById("confirm-action").addEventListener("click", () => {
    const result = action();
    if (!result.ok) return showToast(result.message, true);
    closeModal(); saveData(); render(); showToast(result.message);
  });
}

function recordSale(type, detail, quantity, total, metodoPago) {
  salesHistory.unshift({ tipo: type, detalle: detail, cantidad: quantity, total, metodoPago, fecha: new Date().toLocaleString("es-BO", { dateStyle: "short", timeStyle: "short" }) });
  saveData(); render(); showToast("Venta registrada y stock actualizado.");
}

function executeSheetSale(method, submitButton, ...args) {
  if (submitButton) submitButton.disabled = true;
  google.script.run
    .withSuccessHandler((result) => {
      if (submitButton) submitButton.disabled = false;
      if (!result.ok) {
        if (handleSessionError(result.mensaje)) return;
        return showToast(result.mensaje, true);
      }
      applySheetData(result.datos);
      render();
      showToast(result.mensaje);
    })
    .withFailureHandler((error) => {
      if (submitButton) submitButton.disabled = false;
      if (handleSessionError(error.message)) return;
      showToast(`No se pudo registrar la venta: ${error.message}`, true);
    })[method](sessionToken, ...args);
}

function bindSales() {
  document.querySelectorAll(".sale-panel").forEach(bindPaymentFields);

  const productSelect = document.getElementById("sale-product");
  const productQuantity = document.getElementById("sale-product-quantity");
  const updateProductEstimate = () => {
    const product = almacen.buscarProducto(productSelect.value);
    const quantity = Number(productQuantity.value) || 0;
    document.getElementById("product-sale-total").textContent = money((product?.precio || 0) * quantity);
    document.getElementById("product-sale-hint").textContent = product
      ? `${product.stock} unidades disponibles · ${money(product.precio)} por unidad`
      : "Selecciona un producto para ver su precio y stock.";
  };
  [productSelect, productQuantity].forEach((element) => element.addEventListener("input", updateProductEstimate));
  updateProductEstimate();
  document.getElementById("product-sale-form").addEventListener("submit", (event) => {
    event.preventDefault();
    const payment = getPaymentDetails(event.currentTarget);
    if (!payment) return;
    const product = almacen.buscarProducto(productSelect.value);
    const quantity = Number(productQuantity.value);
    if (!product || !Number.isInteger(quantity) || quantity <= 0) return showToast("Selecciona un producto y una cantidad válida.", true);
    if (APPS_SCRIPT_MODE) return executeSheetSale("apiVenderProductoWeb", event.submitter, product.id, quantity, payment.metodoPago, payment.qrConfirmado, payment.qrPagoUrl);
    if (!product.reducirStock(quantity)) return showToast("No hay stock suficiente para completar la venta.", true);
    recordSale("Producto", product.nombre, quantity, product.precio * quantity, payment.metodoPago);
  });

  const drinkSelect = document.getElementById("sale-drink");
  const cupSelect = document.getElementById("sale-cup");
  const drinkQuantity = document.getElementById("sale-glass-quantity");
  const updateGlassEstimate = () => {
    const drink = almacen.buscarProducto(drinkSelect.value);
    const quantity = Number(drinkQuantity.value) || 0;
    document.getElementById("glass-sale-total").textContent = money((drink?.precioVaso || 0) * quantity);
    document.getElementById("glass-sale-hint").textContent = drink
      ? `${drink.vasosDisponibles()} vasos disponibles · ${money(drink.precioVaso)} por vaso · ${almacen.buscarProducto(cupSelect.value)?.stock || 0} vasos descartables`
      : "Selecciona un producto para ver su precio y disponibilidad.";
  };
  [drinkSelect, cupSelect, drinkQuantity].forEach((element) => element.addEventListener("input", updateGlassEstimate));
  updateGlassEstimate();
  document.getElementById("glass-sale-form").addEventListener("submit", (event) => {
    event.preventDefault();
    const payment = getPaymentDetails(event.currentTarget);
    if (!payment) return;
    const drink = almacen.buscarProducto(drinkSelect.value);
    const quantity = Number(drinkQuantity.value);
    const cup = almacen.buscarProducto(cupSelect.value);
    if (!drink || !cup || !Number.isInteger(quantity) || quantity <= 0) return showToast("Selecciona productos y una cantidad válida.", true);
    if (APPS_SCRIPT_MODE) return executeSheetSale("apiVenderVasoWeb", event.submitter, drink.id, cup.id, quantity, payment.metodoPago, payment.qrConfirmado, payment.qrPagoUrl);
    if (!almacen.venderVaso(drink.id, quantity, cup.id)) return showToast("No hay stock suficiente para completar la venta.", true);
    recordSale("vaso", `${drink.nombre} · ${cup.nombre}`, quantity, drink.precioVaso * quantity, payment.metodoPago);
  });

  const comboSelect = document.getElementById("sale-combo");
  const comboQuantity = document.getElementById("sale-combo-quantity");
  const updateComboEstimate = () => {
    const combo = almacen.buscarCombo(comboSelect.value);
    const quantity = Number(comboQuantity.value) || 0;
    document.getElementById("combo-sale-total").textContent = money((combo?.precio || 0) * quantity);
    document.getElementById("combo-sale-hint").textContent = combo
      ? `${combo.disponibles()} combos disponibles · ${money(combo.precio)} por combo`
      : "Selecciona un combo para ver disponibilidad.";
  };
  [comboSelect, comboQuantity].forEach((element) => element.addEventListener("input", updateComboEstimate));
  updateComboEstimate();
  document.getElementById("combo-sale-form").addEventListener("submit", (event) => {
    event.preventDefault();
    const payment = getPaymentDetails(event.currentTarget);
    if (!payment) return;
    const combo = almacen.buscarCombo(comboSelect.value);
    const quantity = Number(comboQuantity.value);
    if (!combo || !Number.isInteger(quantity) || quantity <= 0) return showToast("Selecciona un combo y una cantidad válida.", true);
    const total = combo.precio * quantity;
    if (APPS_SCRIPT_MODE) return executeSheetSale("apiVenderComboWeb", event.submitter, combo.id, quantity, payment.metodoPago, payment.qrConfirmado, payment.qrPagoUrl);
    if (!almacen.venderCombo(combo.id, quantity)) return showToast("No hay stock suficiente para completar la venta.", true);
    recordSale("combo", combo.nombre, quantity, total, payment.metodoPago);
  });
}

function exportBackup() {
  const data = JSON.stringify({
    categorias: almacen.categorias,
    productos: almacen.productos.map((product) => ({ ...product })),
    combos: almacen.combos.map((combo) => ({
      id: combo.id,
      nombre: combo.nombre,
      precio: combo.precio,
      componentes: combo.componentes.map((part) => ({ id: part.producto.id, cantidad: part.cantidad })),
    })),
    ventas: salesHistory,
  }, null, 2);
  downloadFile(`respaldo-almacen-${new Date().toISOString().slice(0, 10)}.json`, data, "application/json");
  showToast("Respaldo descargado.");
}

function downloadFile(filename, content, type) {
  const url = URL.createObjectURL(new Blob([content], { type }));
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function exportProductsCSV() {
  const rows = [["ID", "Nombre", "Categoría", "Precio", "Stock", "Porciones", "Precio vaso", "Vasos sueltos", "URL imagen"], ...almacen.aFilasProductos()];
  const csv = rows.map((row) => row.map((value) => `"${String(value ?? "").replace(/"/g, '""')}"`).join(";")).join("\r\n");
  downloadFile("productos-almacen.csv", `\uFEFF${csv}`, "text/csv;charset=utf-8");
  showToast("Inventario exportado a CSV.");
}

document.addEventListener("click", (event) => {
  const viewButton = event.target.closest("[data-view]");
  if (viewButton) {
    currentView = viewButton.dataset.view;
    render();
    document.getElementById("sidebar").classList.remove("open");
    return;
  }
  const actionButton = event.target.closest("[data-action]");
  if (!actionButton) return;
  if (!isAdministrator()) return;
  const { action, id, name } = actionButton.dataset;
  if (action === "add-product") productForm();
  if (action === "edit-product") productForm(almacen.buscarProducto(id));
  if (action === "add-stock") addStockForm(almacen.buscarProducto(id));
  if (action === "add-combo") comboForm();
  if (action === "edit-combo") {
    const combo = almacen.buscarCombo(id);
    if (combo) comboForm(combo);
  }
  if (action === "add-category") categoryForm();
  if (action === "close-modal") closeModal();
  if (action === "export-data") exportBackup();
  if (action === "export-csv") exportProductsCSV();
  if (action === "delete-product") {
    const product = almacen.buscarProducto(id);
    const combo = almacen.combos.find((item) => item.componentes.some((part) => part.producto.id === id));
    if (combo) return showToast(`No se puede eliminar: ${product.nombre} forma parte de ${combo.nombre}.`, true);
    confirmAction("Eliminar producto", `Se eliminará <strong>${escapeHTML(product.nombre)}</strong> (${escapeHTML(id)}). Esta acción no se puede deshacer.`, "Eliminar producto", () => ({ ok: almacen.eliminarProducto(id), message: "Producto eliminado." }), true);
  }
  if (action === "delete-combo") {
    const combo = almacen.buscarCombo(id);
    confirmAction("Eliminar combo", `Se eliminará <strong>${escapeHTML(combo.nombre)}</strong> (${escapeHTML(id)}).`, "Eliminar combo", () => ({ ok: almacen.eliminarCombo(id), message: "Combo eliminado." }), true);
  }
  if (action === "delete-category") {
    const category = almacen.categorias.find((item) => item.nombre === name);
    const used = almacen.productos.some((product) => product.categoria === name);
    if (used) return showToast("No se puede eliminar una categoría que contiene productos.", true);
    confirmAction("Eliminar categoría", `Se eliminará la categoría <strong>${escapeHTML(name)}</strong>.`, "Eliminar categoría", () => {
      const index = almacen.categorias.findIndex((item) => item.nombre === name);
      if (index < 0) return { ok: false, message: "La categoría ya no existe." };
      almacen.categorias.splice(index, 1);
      return { ok: true, message: "Categoría eliminada." };
    }, true);
  }
});

document.getElementById("page-content").addEventListener("input", (event) => {
  if (event.target.id !== "product-search" && event.target.id !== "category-filter") return;
  const search = (document.getElementById("product-search")?.value || "").trim().toLowerCase();
  const category = document.getElementById("category-filter")?.value || "";
  const filtered = almacen.productos.filter((product) =>
    (!category || product.categoria === category)
    && (!search || `${product.nombre} ${product.id} ${product.categoria}`.toLowerCase().includes(search)));
  const canEdit = isAdministrator();
  document.getElementById("product-rows").innerHTML = filtered.map((product) => productRow(product, canEdit)).join("") || emptyTable("Sin resultados", "Prueba con otro nombre, ID o categoría.", "", canEdit ? 9 : 8);
  document.getElementById("product-result-count").textContent = `${filtered.length} ${filtered.length === 1 ? "producto" : "productos"}`;
});

document.getElementById("page-content").addEventListener("change", (event) => {
  if (event.target.id === "category-filter") event.target.dispatchEvent(new Event("input", { bubbles: true }));
});

document.getElementById("modal-close").addEventListener("click", closeModal);
document.getElementById("modal-backdrop").addEventListener("click", (event) => {
  if (event.target.id === "modal-backdrop") closeModal();
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !document.getElementById("modal-backdrop").hidden) closeModal();
});
document.getElementById("mobile-menu").addEventListener("click", () => document.getElementById("sidebar").classList.toggle("open"));
document.getElementById("export-data").addEventListener("click", exportBackup);
document.getElementById("login-form").addEventListener("submit", (event) => {
  event.preventDefault();
  login(document.getElementById("login-username").value, document.getElementById("login-password").value);
});
document.getElementById("logout-button").addEventListener("click", logout);
document.getElementById("change-password").addEventListener("click", changeOwnPasswordForm);

function changeOwnPasswordForm() {
  openModal("Cambiar contraseña", `<form id="change-own-password-form">
    <div class="form-field"><label for="current-password">Contraseña actual</label><input class="field" id="current-password" type="password" autocomplete="current-password" required></div>
    <div class="form-field"><label for="new-password">Nueva contraseña</label><input class="field" id="new-password" type="password" minlength="8" maxlength="128" autocomplete="new-password" required></div>
    <p class="form-hint">La contraseña debe tener entre 8 y 128 caracteres.</p>
    <div class="form-actions"><button class="button button-secondary" type="button" data-action="close-modal">Cancelar</button><button class="button button-primary" type="submit">Actualizar contraseña</button></div>
  </form>`);
  document.getElementById("change-own-password-form").addEventListener("submit", (event) => {
    event.preventDefault();
    const currentPassword = document.getElementById("current-password").value;
    const newPassword = document.getElementById("new-password").value;
    const submit = event.currentTarget.querySelector('button[type="submit"]');
    submit.disabled = true;
    google.script.run
      .withSuccessHandler((result) => {
        submit.disabled = false;
        if (!result.ok) {
          if (handleSessionError(result.mensaje)) return;
          return showToast(result.mensaje, true);
        }
        closeModal();
        showToast(result.mensaje);
        logout();
      })
      .withFailureHandler((error) => {
        submit.disabled = false;
        if (handleSessionError(error.message)) return;
        showToast(`No se pudo cambiar la contraseña: ${error.message}`, true);
      })
      .apiCambiarClavePropia(sessionToken, currentPassword, newPassword);
  });
}

const now = new Date();
document.getElementById("today-label").textContent = new Intl.DateTimeFormat("es-BO", { weekday: "short", day: "numeric", month: "short", year: "numeric" }).format(now);
if (APPS_SCRIPT_MODE) {
  document.getElementById("storage-label").textContent = "Conectando con Google Sheets...";
  showLogin();
} else {
  document.getElementById("login-screen").hidden = true;
  document.getElementById("app-shell").hidden = false;
  document.getElementById("logout-button").hidden = true;
  document.getElementById("change-password").hidden = true;
  document.querySelector('[data-view="usuarios"]').hidden = true;
  loadData();
  saveData();
  render();
}