const STORAGE_KEY = "sis101-almacen-v1";
const LOW_STOCK_LIMIT = 5;
const APPS_SCRIPT_MODE = typeof google !== "undefined" && Boolean(google.script?.run);
const sectionNames = {
  dashboard: "Resumen",
  productos: "Productos",
  combos: "Combos",
  ventas: "Ventas",
  categorias: "Categorías",
};
const currency = new Intl.NumberFormat("es-BO", { style: "currency", currency: "BOB", minimumFractionDigits: 2 });
let almacen = new Almacen();
let salesHistory = [];
let currentView = "dashboard";
let toastTimer;
let sheetVersion = 0;

function escapeHTML(value) {
  return String(value ?? "").replace(/[&<>"']/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  })[character]);
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
        showToast(`No se pudo cargar la hoja: ${error.message}`, true);
      })
      .apiObtenerDatosWeb();
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
          showToast(result.mensaje || "No se pudo guardar en Google Sheets.", true);
          render();
          return;
        }
        applySheetData(result.datos);
        document.getElementById("storage-label").textContent = "Conectado a Google Sheets";
        render();
      })
      .withFailureHandler((error) => showToast(`Error al guardar en Google Sheets: ${error.message}`, true))
      .apiGuardarInventarioWeb(inventory, sheetVersion);
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
  document.getElementById("current-section").textContent = sectionNames[currentView];
  document.querySelectorAll(".nav-item").forEach((item) => item.classList.toggle("active", item.dataset.view === currentView));
  const views = {
    dashboard: renderDashboard,
    productos: renderProducts,
    combos: renderCombos,
    ventas: renderSales,
    categorias: renderCategories,
  };
  document.getElementById("page-content").innerHTML = views[currentView]();
  if (currentView === "ventas") bindSales();
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
    <td><div class="item-cell"><span class="item-avatar">${escapeHTML(initials(product.nombre))}</span><span class="item-name"><strong>${escapeHTML(product.nombre)}</strong><small>${escapeHTML(product.categoria)}</small></span></div></td>
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
    <td><div class="item-cell"><span class="item-avatar">${escapeHTML(initials(product.nombre))}</span><span class="item-name"><strong>${escapeHTML(product.nombre)}</strong><small>${escapeHTML(product.id)}</small></span></div></td>
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
  const rows = almacen.productos.map((product) => productRow(product, true)).join("");
  const options = ['<option value="">Todas las categorías</option>', ...almacen.categorias.map((category) => `<option value="${escapeHTML(category.nombre)}">${escapeHTML(category.nombre)}</option>`)].join("");
  return `${pageHeading("INVENTARIO", "Productos", "Administra existencias, precios y productos por vaso.", '<button class="button button-primary" data-action="add-product"><span class="button-plus">+</span> Nuevo producto</button>')}
    <div class="toolbar"><label class="search-box"><span>⌕</span><input class="field" id="product-search" type="search" placeholder="Buscar por nombre o ID..."></label><select class="select-field" id="category-filter" aria-label="Filtrar por categoría">${options}</select><button class="button button-secondary button-small" data-action="export-csv">↓ Exportar CSV</button></div>
    <section class="panel table-panel"><div class="table-wrap"><table><thead><tr><th>Producto</th><th>ID</th><th>Stock</th><th>Vasos disponibles</th><th>Precio unidad</th><th>Porciones/unidad</th><th>Precio vaso</th><th>Estado</th><th></th></tr></thead><tbody id="product-rows">${rows || emptyTable("Tu inventario está vacío", "Agrega productos para comenzar a controlar las existencias.", '<button class="button button-primary button-small" data-action="add-product">Agregar producto</button>', 9)}</tbody></table></div><footer class="table-footer"><span id="product-result-count">${almacen.productos.length} productos</span><span>Stock bajo: ${almacen.stockBajo(LOW_STOCK_LIMIT).length}</span></footer></section>`;
}

function renderCombos() {
  const rows = almacen.combos.map((combo) => {
    const broken = combo.componentesRotos?.length > 0;
    const available = broken ? 0 : combo.disponibles();
    const parts = combo.componentes.map((part) => `${part.cantidad} × ${escapeHTML(part.producto.nombre)}`).join(" · ");
    const missing = broken ? ` · Faltan: ${combo.componentesRotos.map(escapeHTML).join(", ")}` : "";
    return `<tr><td><div class="item-cell"><span class="item-avatar combo-avatar">${escapeHTML(initials(combo.nombre))}</span><span class="item-name"><strong>${escapeHTML(combo.nombre)}</strong><small>${parts}${escapeHTML(missing)}</small></span></div></td><td class="mono">${escapeHTML(combo.id)}</td><td>${money(combo.precio)}</td><td><span class="stock-value ${available === 0 ? "stock-low" : ""}">${available}</span></td><td>${broken ? '<span class="pill danger">Componentes faltantes</span>' : available > 0 ? '<span class="pill">Disponible</span>' : '<span class="pill danger">Sin stock</span>'}</td><td><div class="row-actions"><button class="row-action" data-action="edit-combo" data-id="${escapeHTML(combo.id)}" title="Editar combo" aria-label="Editar ${escapeHTML(combo.nombre)}">✎</button><button class="row-action delete" data-action="delete-combo" data-id="${escapeHTML(combo.id)}" title="Eliminar combo" aria-label="Eliminar ${escapeHTML(combo.nombre)}">×</button></div></td></tr>`;
  }).join("");
  return `${pageHeading("CATÁLOGO", "Combos", "Arma paquetes con productos del inventario y controla su disponibilidad.", '<button class="button button-primary" data-action="add-combo"><span class="button-plus">+</span> Nuevo combo</button>')}
    <section class="panel table-panel"><div class="table-wrap"><table><thead><tr><th>Combo y componentes</th><th>ID</th><th>Precio</th><th>Disponibles</th><th>Estado</th><th></th></tr></thead><tbody>${rows || emptyTable("Aún no hay combos", "Crea un combo a partir de productos registrados.", '<button class="button button-primary button-small" data-action="add-combo">Crear primer combo</button>', 6)}</tbody></table></div><footer class="table-footer"><span>${almacen.combos.length} combos registrados</span><span>La disponibilidad se calcula con el stock actual</span></footer></section>`;
}

function renderSales() {
  const drinkOptions = almacen.productos.filter((product) => product.porciones > 0).map((product) => `<option value="${escapeHTML(product.id)}">${escapeHTML(product.id)} · ${escapeHTML(product.nombre)} (${product.vasosDisponibles()} vasos)</option>`).join("");
  const cupOptions = almacen.productos.filter((product) => product.nombre.toLowerCase().includes("vaso")).map((product) => `<option value="${escapeHTML(product.id)}">${escapeHTML(product.id)} · ${escapeHTML(product.nombre)} (${product.stock})</option>`).join("");
  const comboOptions = almacen.combos.filter((combo) => !combo.componentesRotos?.length).map((combo) => `<option value="${escapeHTML(combo.id)}">${escapeHTML(combo.id)} · ${escapeHTML(combo.nombre)} (${combo.disponibles()} disponibles)</option>`).join("");
  const historyRows = salesHistory.map((sale) => `<tr><td class="mono">${escapeHTML(sale.fecha)}</td><td>${sale.tipo === "Vaso" || sale.tipo === "vaso" ? "Venta por vaso" : "Venta de combo"}</td><td>${escapeHTML(sale.detalle)}</td><td>${sale.cantidad}</td><td class="sale-price">${money(sale.total)}</td></tr>`).join("");
  const history = historyRows || emptyTable("Sin ventas registradas", "Las ventas guardadas en la hoja Ventas aparecerán aquí.", "", 5);
  return `${pageHeading("PUNTO DE VENTA", "Ventas", "Registra ventas y descuenta existencias automáticamente.")}
    <div class="sales-grid">
      <form class="panel sale-panel" id="glass-sale-form"><h2>Venta por vaso</h2><p>Descuenta las porciones y los vasos descartables.</p>
        <div class="form-field"><label for="sale-drink">Producto por vaso</label><select class="select-field" id="sale-drink" required>${drinkOptions || '<option value="">No hay productos configurados por vaso</option>'}</select></div>
        <div class="form-field"><label for="sale-cup">Vaso descartable</label><select class="select-field" id="sale-cup" required>${cupOptions || '<option value="">No se encontró un producto de vasos</option>'}</select></div>
        <div class="form-field"><label for="sale-glass-quantity">Cantidad de vasos</label><input class="field" id="sale-glass-quantity" type="number" min="1" step="1" value="1" required></div>
        <p class="form-hint" id="glass-sale-hint">Selecciona un producto para ver su precio y disponibilidad.</p><div class="sale-total"><span>Total estimado</span><strong id="glass-sale-total">${money(0)}</strong></div><button class="button button-primary" type="submit">Registrar venta</button>
      </form>
      <form class="panel sale-panel" id="combo-sale-form"><h2>Venta de combo</h2><p>El stock de cada componente se actualizará al confirmar.</p>
        <div class="form-field"><label for="sale-combo">Combo</label><select class="select-field" id="sale-combo" required>${comboOptions || '<option value="">No hay combos registrados</option>'}</select></div>
        <div class="form-field"><label for="sale-combo-quantity">Cantidad</label><input class="field" id="sale-combo-quantity" type="number" min="1" step="1" value="1" required></div>
        <p class="form-hint" id="combo-sale-hint">Selecciona un combo para ver disponibilidad.</p><div class="sale-total"><span>Total estimado</span><strong id="combo-sale-total">${money(0)}</strong></div><button class="button button-primary" type="submit">Registrar venta</button>
      </form>
    </div>
    <section class="panel history-panel"><header class="panel-header"><div><h2>Historial de ventas</h2><p>Operaciones almacenadas en la hoja Ventas</p></div><span class="pill">${salesHistory.length} ventas</span></header><div class="table-wrap"><table><thead><tr><th>Fecha</th><th>Tipo</th><th>Detalle</th><th>Cantidad</th><th>Total</th></tr></thead><tbody>${history}</tbody></table></div></section>`;
}

function renderCategories() {
  const cards = almacen.categorias.map((category) => {
    const count = almacen.filtrarPorCategoria(category.nombre).length;
    return `<article class="category-card"><div class="category-card-top"><span class="category-card-letter">${escapeHTML(category.letra)}</span><button class="row-action delete" data-action="delete-category" data-name="${escapeHTML(category.nombre)}" title="Eliminar categoría" aria-label="Eliminar ${escapeHTML(category.nombre)}">×</button></div><h3>${escapeHTML(category.nombre)}</h3><p>${count} ${count === 1 ? "producto" : "productos"} · Prefijo ${escapeHTML(category.letra)}</p></article>`;
  }).join("");
  return `${pageHeading("ORGANIZACIÓN", "Categorías", "Agrupa productos y conserva sus prefijos de identificación.", '<button class="button button-primary" data-action="add-category"><span class="button-plus">+</span> Nueva categoría</button>')}
    <div class="category-cards">${cards}</div>`;
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
    <div class="form-row"><div class="form-field"><label for="product-category">Categoría</label><select class="select-field" id="product-category" name="categoria" ${editing ? "disabled" : ""}>${categoryOptions}</select></div><div class="form-field"><label for="product-price">Precio (Bs)</label><input class="field" id="product-price" name="precio" type="number" min="0" step="0.01" value="${product?.precio ?? ""}" required></div></div>
    <div class="form-row"><div class="form-field"><label for="product-stock">Stock actual</label><input class="field" id="product-stock" name="stock" type="number" min="0" step="1" value="${product?.stock ?? 0}" required></div><div class="form-field"><label for="product-portions">Porciones por unidad</label><input class="field" id="product-portions" name="porciones" type="number" min="0" step="1" value="${product?.porciones ?? 0}" required></div></div>
    <div class="form-field"><label for="product-glass-price">Precio por vaso (Bs)</label><input class="field" id="product-glass-price" name="precioVaso" type="number" min="0" step="0.01" value="${product?.precioVaso ?? 0}" required></div>
    ${editing ? `<p class="form-hint">ID ${escapeHTML(product.id)} · La categoría y el ID no se modifican.</p>` : '<p class="form-hint">El ID se asigna automáticamente según el prefijo de categoría.</p>'}
    <div class="form-actions"><button class="button button-secondary" type="button" data-action="close-modal">Cancelar</button><button class="button button-primary" type="submit">${editing ? "Guardar cambios" : "Crear producto"}</button></div>
  </form>`);
  document.getElementById("product-form").addEventListener("submit", (event) => {
    event.preventDefault();
    const form = new FormData(event.currentTarget);
    const nombre = String(form.get("nombre")).trim();
    const precio = Number(form.get("precio"));
    const stock = Number(form.get("stock"));
    const porciones = Number(form.get("porciones"));
    const precioVaso = Number(form.get("precioVaso"));
    if (!nombre || [precio, stock, porciones, precioVaso].some((value) => !Number.isFinite(value) || value < 0)) {
      showToast("Revisa los datos: deben ser valores válidos y no negativos.", true);
      return;
    }
    if (editing) {
      product.modificar(nombre, precio, porciones, precioVaso);
      product.stock = stock;
      product.vasosSueltos = Math.min(product.vasosSueltos, porciones || 0);
      showToast("Producto actualizado.");
    } else {
      const duplicate = almacen.productos.some((item) => item.nombre.toLowerCase() === nombre.toLowerCase());
      if (duplicate) return showToast("Ya existe un producto con ese nombre.", true);
      const created = new Producto(nombre, form.get("categoria"), precio, stock, porciones, precioVaso);
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

function recordSale(type, detail, quantity, total) {
  salesHistory.unshift({ tipo: type, detalle: detail, cantidad: quantity, total, fecha: new Date().toLocaleString("es-BO", { dateStyle: "short", timeStyle: "short" }) });
  saveData(); render(); showToast("Venta registrada y stock actualizado.");
}

function executeSheetSale(method, submitButton, ...args) {
  submitButton.disabled = true;
  google.script.run
    .withSuccessHandler((result) => {
      submitButton.disabled = false;
      if (!result.ok) return showToast(result.mensaje, true);
      applySheetData(result.datos);
      render();
      showToast(result.mensaje);
    })
    .withFailureHandler((error) => {
      submitButton.disabled = false;
      showToast(`No se pudo registrar la venta: ${error.message}`, true);
    })[method](...args);
}

function bindSales() {
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
    const drink = almacen.buscarProducto(drinkSelect.value);
    const quantity = Number(drinkQuantity.value);
    const cup = almacen.buscarProducto(cupSelect.value);
    if (!drink || !cup || !Number.isInteger(quantity) || quantity <= 0) return showToast("Selecciona productos y una cantidad válida.", true);
    if (APPS_SCRIPT_MODE) return executeSheetSale("apiVenderVasoWeb", event.submitter, drink.id, cup.id, quantity);
    if (!almacen.venderVaso(drink.id, quantity, cup.id)) return showToast("No hay stock suficiente para completar la venta.", true);
    recordSale("vaso", `${drink.nombre} · ${cup.nombre}`, quantity, drink.precioVaso * quantity);
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
    const combo = almacen.buscarCombo(comboSelect.value);
    const quantity = Number(comboQuantity.value);
    if (!combo || !Number.isInteger(quantity) || quantity <= 0) return showToast("Selecciona un combo y una cantidad válida.", true);
    const total = combo.precio * quantity;
    if (APPS_SCRIPT_MODE) return executeSheetSale("apiVenderComboWeb", event.submitter, combo.id, quantity);
    if (!almacen.venderCombo(combo.id, quantity)) return showToast("No hay stock suficiente para completar la venta.", true);
    recordSale("combo", combo.nombre, quantity, total);
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
  const rows = [["ID", "Nombre", "Categoría", "Precio", "Stock", "Porciones", "Precio vaso", "Vasos sueltos"], ...almacen.aFilasProductos()];
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
  document.getElementById("product-rows").innerHTML = filtered.map((product) => productRow(product, true)).join("") || emptyTable("Sin resultados", "Prueba con otro nombre, ID o categoría.", "", 9);
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

const now = new Date();
document.getElementById("today-label").textContent = new Intl.DateTimeFormat("es-BO", { weekday: "short", day: "numeric", month: "short", year: "numeric" }).format(now);
if (APPS_SCRIPT_MODE) {
  document.getElementById("storage-label").textContent = "Conectando con Google Sheets...";
  loadData();
  render();
} else {
  loadData();
  saveData();
  render();
}