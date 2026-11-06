const { Producto, Combo, Almacen } = require("./Blu.js");

function titulo(texto) {
  console.log(`\n===== ${texto} =====`);
}

titulo("1. CREAR PRODUCTOS");
const singani = new Producto(1, "Singani Casa Real", "Botellas", 180, 15, 12, 20);
const fernet = new Producto(2, "Fernet Branca", "Botellas", 220, 5, 12, 25);
const coca = new Producto(3, "Coca-Cola 2L", "Mezcladores", 15, 8, 0, 0);
const hielo = new Producto(4, "Bolsa de hielo", "Insumos", 10, 20, 0, 0);
const balde = new Producto(5, "Balde vacío", "Insumos", 30, 12, 0, 0);
const vaso = new Producto(6, "Vaso descartable", "Insumos", 1, 100, 0, 0);
[singani, fernet, coca, hielo, balde, vaso].forEach((p) => p.ver());

titulo("2. METODOS DE PRODUCTO");

console.log("\n-- modificar --");
singani.modificar("Singani Casa Real Etiqueta Negra", "Botellas", 200, 12, 22);
singani.ver();

console.log("\n-- aumentarStock --");
console.log("aumentarStock(5):", singani.aumentarStock(5));
console.log("aumentarStock(-1):", singani.aumentarStock(-1));
singani.ver();

console.log("\n-- reducirStock --");
console.log("reducirStock(3):", singani.reducirStock(3));
console.log("reducirStock(100):", singani.reducirStock(100));
singani.ver();

console.log("\n-- hayStock --");
console.log("singani:", singani.hayStock());
const agotado = new Producto(99, "Producto agotado", "Prueba", 1, 0, 0, 0);
console.log("agotado:", agotado.hayStock());

console.log("\n-- vasosDisponibles --");
console.log("singani:", singani.vasosDisponibles());

console.log("\n-- venderVaso --");
console.log("venderVaso(5):", singani.venderVaso(5));
singani.ver();
console.log("venderVaso(7):", singani.venderVaso(7));
singani.ver();
console.log("venderVaso(1000):", singani.venderVaso(1000));
console.log("coca.venderVaso(1):", coca.venderVaso(1));
console.log("vasosDisponibles:", singani.vasosDisponibles());

console.log("\n-- aFila --");
console.log(singani.aFila());

titulo("3. CREAR COMBOS");
const comboSingani = new Combo(1, "Combo Singani", 220, [
  { producto: singani, cantidad: 1 },
  { producto: coca, cantidad: 2 },
  { producto: hielo, cantidad: 1 },
  { producto: vaso, cantidad: 4 },
]);
const comboFernet = new Combo(2, "Combo Fernet", 260, [
  { producto: fernet, cantidad: 1 },
  { producto: coca, cantidad: 2 },
  { producto: hielo, cantidad: 1 },
  { producto: vaso, cantidad: 4 },
]);
comboSingani.ver();
comboFernet.ver();

titulo("4. METODOS DE COMBO");

console.log("\n-- disponibles --");
console.log("comboSingani:", comboSingani.disponibles());
console.log("comboFernet:", comboFernet.disponibles());

console.log("\n-- vender --");
console.log("comboSingani.vender(2):", comboSingani.vender(2));
console.log("comboSingani.vender(100):", comboSingani.vender(100));
console.log("comboSingani.vender(0):", comboSingani.vender(0));
[singani, coca, hielo, vaso].forEach((p) => p.ver());
comboSingani.ver();

console.log("\n-- modificar --");
comboSingani.modificar("Combo Singani Premium", 250, [
  { producto: singani, cantidad: 1 },
  { producto: coca, cantidad: 3 },
  { producto: hielo, cantidad: 2 },
]);
comboSingani.ver();

console.log("\n-- aFila --");
console.log(comboSingani.aFila());

titulo("5. CREAR ALMACEN Y AGREGAR");
const almacen = new Almacen();
console.log("agregarProducto singani:", almacen.agregarProducto(singani));
console.log("agregarProducto fernet:", almacen.agregarProducto(fernet));
console.log("agregarProducto coca:", almacen.agregarProducto(coca));
console.log("agregarProducto hielo:", almacen.agregarProducto(hielo));
console.log("agregarProducto balde:", almacen.agregarProducto(balde));
console.log("agregarProducto vaso:", almacen.agregarProducto(vaso));
console.log(
  "agregarProducto repetido:",
  almacen.agregarProducto(new Producto(1, "Repetido", "Prueba", 1, 1, 0, 0))
);

console.log("agregarCombo singani:", almacen.agregarCombo(comboSingani));
console.log("agregarCombo fernet:", almacen.agregarCombo(comboFernet));
console.log(
  "agregarCombo repetido:",
  almacen.agregarCombo(new Combo(1, "Combo repetido", 100, []))
);
console.log(
  "agregarCombo con producto no registrado:",
  almacen.agregarCombo(
    new Combo(3, "Combo fantasma", 100, [{ producto: agotado, cantidad: 1 }])
  )
);

titulo("6. METODOS DE ALMACEN");

console.log("\n-- verProductos / verCombos --");
almacen.verProductos();
almacen.verCombos();

console.log("\n-- buscarProducto / buscarCombo --");
console.log("buscarProducto(2):", almacen.buscarProducto(2).nombre);
console.log("buscarProducto(99):", almacen.buscarProducto(99));
console.log("buscarCombo(2):", almacen.buscarCombo(2).nombre);
console.log("buscarCombo(9):", almacen.buscarCombo(9));

console.log("\n-- filtrarPorCategoria('Insumos') --");
almacen.filtrarPorCategoria("Insumos").forEach((p) => p.ver());

console.log("\n-- stockBajo(10) --");
almacen.stockBajo(10).forEach((p) => p.ver());

console.log("\n-- valorTotal --");
console.log(`${almacen.valorTotal()} Bs`);

console.log("\n-- venderCombo --");
console.log("venderCombo(2, 1):", almacen.venderCombo(2, 1));
console.log("venderCombo(2, 100):", almacen.venderCombo(2, 100));
console.log("venderCombo(9, 1):", almacen.venderCombo(9, 1));
almacen.verProductos();

console.log("\n-- venderVaso (licor + vaso descartable) --");
console.log("venderVaso(2, 4, 6):", almacen.venderVaso(2, 4, 6));
console.log("venderVaso(3, 1, 6):", almacen.venderVaso(3, 1, 6));
console.log("venderVaso(2, 1000, 6):", almacen.venderVaso(2, 1000, 6));
console.log("venderVaso(99, 1, 6):", almacen.venderVaso(99, 1, 6));
fernet.ver();
vaso.ver();

console.log("\n-- eliminarProducto / eliminarCombo --");
console.log("eliminarProducto(1) (lo usa un combo):", almacen.eliminarProducto(1));
console.log("eliminarProducto(5) (balde):", almacen.eliminarProducto(5));
console.log("eliminarProducto(77) (no existe):", almacen.eliminarProducto(77));
console.log("eliminarCombo(1):", almacen.eliminarCombo(1));
console.log("eliminarCombo(9) (no existe):", almacen.eliminarCombo(9));
console.log("eliminarProducto(1) (ya sin combo):", almacen.eliminarProducto(1));
almacen.verProductos();
almacen.verCombos();

console.log("\n-- aFilasProductos / aFilasCombos --");
console.log(almacen.aFilasProductos());
console.log(almacen.aFilasCombos());