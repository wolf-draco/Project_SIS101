# 🔷 SIS101

### Control de inventario

> Una herramienta sencilla para administrar productos, combos y ventas, desde el
> navegador o desde Google Sheets.

**JavaScript** · **Node.js** · **Google Apps Script** · **Google Sheets**

---

## Contenido

- [Funciones principales](#funciones-principales)
- [Inicio rápido](#inicio-rápido)
- [Aplicación web](#aplicación-web)
- [Google Sheets y Apps Script](#google-sheets-y-apps-script)
- [Estructura del proyecto](#estructura-del-proyecto)
- [Modelo de datos](#modelo-de-datos)
- [Validación](#validación)
- [Licencia](#licencia)

## Funciones principales

| Área | Qué permite hacer |
| --- | --- |
| **Productos** | Registrar productos, categorías, precios y existencias. |
| **Combos** | Crear combos asociados a productos y consultar cuántos pueden prepararse. |
| **Ventas** | Registrar ventas por vaso o combo y consultar la actividad. |
| **Respaldo** | Exportar el inventario a CSV o descargar los datos en JSON. |

La aplicación web usa una interfaz oscura con acentos azules. El Dashboard de
Google Sheets utiliza una paleta clara para mantener la lectura cómoda en la
hoja de cálculo.

## Inicio rápido

Se requiere **Node.js 18 o posterior**. El proyecto no tiene dependencias
externas ni necesita `npm install`.

Abre [`index.html`](./index.html) en un navegador para usar la aplicación.
También puedes ejecutar la demostración del modelo desde la terminal:

```bash
node tests/Prueba.js
```

En el navegador, los datos se guardan en el almacenamiento local de ese
dispositivo. No se sincronizan automáticamente entre dispositivos y son
independientes de los datos de Google Sheets.

## Aplicación web

La interfaz permite administrar productos, categorías y combos, registrar
ventas por vaso o combo, y exportar el inventario. Los archivos principales son
[`index.html`](./index.html), [`app.css`](./src/css/app.css),
[`app.js`](./src/js/app.js) y [`Blu.js`](./src/js/Blu.js).

### Conectar con Google Sheets

La versión de Apps Script puede usar las pestañas `Productos`, `Combos`,
`Categorias` y `Ventas` como almacenamiento. Para generar la página integrada
con las fuentes web actuales:

```bash
node scripts/build-gas-web.js
```

El comando actualiza [`WebApp.html`](./apps-script/WebApp.html). Después,
actualiza los archivos HTML y `.gs` de [`apps-script/`](./apps-script) en el
proyecto de Apps Script vinculado a la hoja de cálculo y guarda los cambios.
La función `doGet` sirve la aplicación y `google.script.run` se comunica con
las pestañas de datos.

Para publicar la aplicación web, despliega el proyecto de Apps Script como una
aplicación web y limita el acceso a tu cuenta o a las personas que deban usar
el inventario. No habilites acceso anónimo para una hoja con información
comercial.

Para aplicar el formato de las hojas, copia también `codigo.gs` y
`Estilos.html` desde [`apps-script/`](./apps-script). Guarda y recarga la hoja
de cálculo; luego ejecuta **Almacén → Instalar/Reinstalar Dashboard**. El
instalador crea la pestaña `Proyecto SIS101` y aplica el formato a las tablas
de inventario.

> **Importante:** si editas `index.html` o algún archivo de `src/`, vuelve a
> generar `WebApp.html` y actualiza ese archivo en Apps Script. Los archivos
> del repositorio no se sincronizan automáticamente con Google Apps Script.

## Estructura del proyecto

```text
.
├── apps-script/          # Código de Apps Script y páginas HTML
├── scripts/
│   └── build-gas-web.js  # Genera la página web integrada
├── src/
│   ├── css/
│   │   └── app.css       # Estilos de la aplicación web
│   └── js/
│       ├── app.js        # Interfaz y lógica de la aplicación
│       └── Blu.js        # Modelo de productos, combos y almacén
├── tests/
│   └── Prueba.js         # Demostración ejecutable del modelo
├── index.html
├── LICENSE
└── README.md
```

## Modelo de datos

El módulo [`Blu.js`](./src/js/Blu.js) exporta las clases `Producto`, `Combo` y
`Almacen`; se puede usar directamente desde Node.js:

```js
const { Producto, Combo, Almacen } = require("./src/js/Blu.js");

const almacen = new Almacen();
const singani = new Producto(
  "Singani Casa Real",
  "Botellas",
  180,
  15,
  12,
  20
);

almacen.agregarProducto(singani);

const combo = new Combo("Combo Singani", 220, [
  { producto: singani, cantidad: 1 },
]);

almacen.agregarCombo(combo);
console.log(combo.id); // K1
```

Los IDs se asignan al registrar un elemento en el almacén; no se pasan al
constructor.

### `Producto`

```js
new Producto(nombre, categoria, precio, stock, porciones, precioVaso)
```

| Dato | Descripción |
| --- | --- |
| `nombre` | Nombre del producto. |
| `categoria` | Nombre de una categoría existente. |
| `precio` | Precio de una unidad o presentación, en bolivianos. |
| `stock` | Unidades o presentaciones disponibles. |
| `porciones` | Vasos que produce una unidad; usa `0` si no se vende por vaso. |
| `precioVaso` | Precio de venta por vaso. |

Métodos principales: `modificar`, `aumentarStock`, `reducirStock`,
`hayStock`, `vasosDisponibles`, `venderVaso`, `ver` y `aFila`.

### `Combo`

```js
new Combo(nombre, precio, componentes)
```

`componentes` es un arreglo de objetos con la forma
`{ producto, cantidad }`. Registra los productos en el almacén antes de
agregar el combo.

Métodos principales: `disponibles`, `vender`, `modificar`, `ver` y `aFila`.

### `Almacen`

Al crear un almacén se registran estas categorías:

| Categoría | Prefijo |
| --- | --- |
| Botellas | `B` |
| Cervezas | `C` |
| Mezcladores | `M` |
| Insumos | `I` |

La letra `K` está reservada para combos. Los IDs se generan secuencialmente
por categoría (`B1`, `B2`, `I1`) y para combos (`K1`, `K2`); no se reutilizan
después de eliminar un producto o combo.

Operaciones principales:

| Método | Descripción |
| --- | --- |
| `agregarCategoria(nombre, letra)` | Registra una categoría con un prefijo único. |
| `verCategorias()` | Muestra categorías y cantidades registradas. |
| `agregarProducto(producto)` | Valida la categoría, asigna ID y registra el producto. |
| `agregarCombo(combo)` | Valida los componentes, asigna ID y registra el combo. |
| `buscarProducto(id)` / `buscarCombo(id)` | Busca un producto o combo por ID. |
| `eliminarProducto(id)` | Elimina un producto que no esté asociado a un combo. |
| `eliminarCombo(id)` | Elimina un combo. |
| `verProductos()` / `verCombos()` | Imprime los registros actuales. |
| `filtrarPorCategoria(categoria)` | Devuelve los productos de una categoría. |
| `stockBajo(limite)` | Devuelve productos cuyo stock no supera el límite. |
| `valorTotal()` | Calcula el valor de las unidades almacenadas. |
| `venderVaso(idProducto, cantidad, idVaso)` | Vende vasos y descuenta los vasos descartables. |
| `venderCombo(idCombo, cantidad)` | Vende un combo por su ID. |
| `aFilasProductos()` / `aFilasCombos()` | Convierte registros en arreglos para listados o exportación. |

Las operaciones inválidas devuelven `false` e imprimen el motivo en consola.
Las búsquedas exitosas devuelven el objeto correspondiente; si no existe,
devuelven `undefined`.

## Validación

Ejecuta la demostración del modelo para comprobar las operaciones principales:

```bash
node tests/Prueba.js
```

El modelo no usa una base de datos: los datos de esta demostración viven en
memoria y se reinician al terminar el proceso.

## Licencia

Distribuido bajo la licencia [MIT](./LICENSE).
