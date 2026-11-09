# Project SIS101

Sistema de inventario para administrar productos, combos y ventas por unidad
en un almacén. El proyecto está implementado en JavaScript puro y funciona
desde Node.js sin dependencias externas.

## Requisitos

- Node.js 18 o una versión posterior.

No es necesario ejecutar `npm install`: el repositorio no utiliza paquetes
externos ni incluye un `package.json`.

## Ejecución

El archivo [`Prueba.js`](./Prueba.js) contiene una demostración completa de
las clases y sus operaciones:

```bash
node Prueba.js
```

La salida muestra la creación del almacén, validaciones, administración de
stock, combos, ventas, eliminaciones y exportación de datos a filas.

## Aplicación web

Abre [`index.html`](./index.html) en un navegador para usar la interfaz de
inventario. Permite administrar productos, categorías y combos, registrar
ventas por vaso o combo, y exportar el inventario a CSV o descargar un
respaldo JSON. Los datos se guardan en el almacenamiento local del navegador;
no se sincronizan entre dispositivos ni reemplazan la hoja de cálculo de
Google Apps Script.

### Conectar con Google Sheets

El proyecto de Apps Script incluye una versión web que usa las hojas
`Productos`, `Combos`, `Categorias` y `Ventas` como base de datos. Para
prepararla, ejecuta `node build-gas-web.js`; el comando genera
`codigo Apps Script/WebApp.html` con los archivos actuales de la página.
Agrega o reemplaza ese archivo en el mismo proyecto de Apps Script vinculado a
la hoja de cálculo y despliega el proyecto como una aplicación web. La función
`doGet` sirve la página y `google.script.run` lee y guarda los datos en las
hojas existentes.

Al desplegar, limita el acceso a tu cuenta o a las personas que deban usar el
inventario. No publiques una hoja con información comercial para acceso
anónimo. Después de cambiar `index.html`, `app.css` o `app.js`, vuelve a
ejecutar el empaquetador y actualiza `WebApp.html` en Apps Script.

## Uso como módulo

[`Blu.js`](./Blu.js) exporta las clases `Producto`, `Combo` y `Almacen`:

```js
const { Producto, Combo, Almacen } = require("./Blu.js");

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

## Modelo de datos

### `Producto`

Constructor:

```js
new Producto(nombre, categoria, precio, stock, porciones, precioVaso)
```

- `nombre`: nombre del producto.
- `categoria`: nombre de una categoría existente.
- `precio`: precio de una unidad o presentación, en bolivianos.
- `stock`: cantidad de unidades o presentaciones disponibles.
- `porciones`: vasos que produce una unidad. Debe ser `0` para productos que
  no se venden por vaso.
- `precioVaso`: precio de venta por vaso.

Métodos principales:

| Método | Descripción |
| --- | --- |
| `modificar(nombre, precio, porciones, precioVaso)` | Actualiza los datos comerciales del producto. |
| `aumentarStock(cantidad)` | Incrementa el stock si la cantidad es positiva. |
| `reducirStock(cantidad)` | Reduce el stock si hay unidades suficientes. |
| `hayStock()` | Indica si quedan unidades o vasos sueltos. |
| `vasosDisponibles()` | Calcula los vasos disponibles según unidades y porciones. |
| `venderVaso(cantidad)` | Descuenta vasos de las unidades abiertas y del stock. |
| `ver()` | Imprime el producto en consola. |
| `aFila()` | Devuelve los datos del producto como arreglo. |

### `Combo`

Constructor:

```js
new Combo(nombre, precio, componentes)
```

`componentes` es un arreglo de objetos con la forma
`{ producto, cantidad }`. Los productos deben estar registrados en el
almacén antes de agregar el combo.

Métodos principales:

- `disponibles()`: calcula cuántos combos pueden armarse con el stock actual.
- `vender(cantidad)`: descuenta de cada producto los componentes de la venta.
- `modificar(nombre, precio, componentes)`: actualiza el combo.
- `ver()`: imprime el combo y sus componentes.
- `aFila()`: devuelve sus datos como arreglo.

### `Almacen`

Al crear un almacén se registran estas categorías:

| Categoría | Prefijo |
| --- | --- |
| Botellas | `B` |
| Cervezas | `C` |
| Mezcladores | `M` |
| Insumos | `I` |

La letra `K` está reservada para combos. Los IDs se generan
secuencialmente por categoría (`B1`, `B2`, `I1`) y para combos (`K1`, `K2`).
Los IDs no se reutilizan después de eliminar un producto o combo.

Operaciones disponibles:

| Método | Descripción |
| --- | --- |
| `agregarCategoria(nombre, letra)` | Registra una categoría con un prefijo único. |
| `verCategorias()` | Muestra categorías y cantidades registradas. |
| `agregarProducto(producto)` | Valida la categoría, asigna ID y registra el producto. |
| `agregarCombo(combo)` | Valida sus componentes, asigna ID y registra el combo. |
| `buscarProducto(id)` / `buscarCombo(id)` | Busca por ID. |
| `eliminarProducto(id)` | Elimina un producto que no sea usado por un combo. |
| `eliminarCombo(id)` | Elimina un combo. |
| `verProductos()` / `verCombos()` | Imprime los registros actuales. |
| `filtrarPorCategoria(categoria)` | Devuelve los productos de una categoría. |
| `stockBajo(limite)` | Devuelve productos cuyo stock no supera el límite. |
| `valorTotal()` | Calcula el valor de las unidades almacenadas. |
| `venderVaso(idProducto, cantidad, idVaso)` | Vende vasos de un producto y descuenta vasos descartables. |
| `venderCombo(idCombo, cantidad)` | Vende un combo por su ID. |
| `aFilasProductos()` / `aFilasCombos()` | Convierte los registros a arreglos para listados o exportación. |

Las operaciones inválidas devuelven `false` y muestran el motivo en consola.
Las búsquedas exitosas devuelven el objeto correspondiente; si no existe,
devuelven `undefined`.

## Estructura

```text
.
├── Blu.js       # Implementación de Producto, Combo y Almacen
├── Prueba.js    # Ejemplo ejecutable de todas las funcionalidades
└── README.md    # Documentación del proyecto
```

## Validación rápida

Para comprobar que el ejemplo funciona:

```bash
node Prueba.js
```

El script no usa una base de datos: todos los datos viven en memoria y se
reinician cada vez que se ejecuta.

## Licencia

Este proyecto se distribuye bajo la licencia [MIT](./LICENSE).
