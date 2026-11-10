const { Producto, Combo, Almacen } = require("../src/js/Blu.js");

function titulo(texto) {
  console.log(`\n===== ${texto} =====`);
}

titulo("1. CREAR ALMACEN Y CATEGORIAS");
const almacen = new Almacen();
almacen.verCategorias();

console.log("\n-- agregarCategoria --");
console.log("agregarCategoria('Postres', 'P'):", almacen.agregarCategoria("Postres", "P"));
console.log("agregarCategoria('Postres', 'Q') (nombre repetido):", almacen.agregarCategoria("Postres", "Q"));
console.log("agregarCategoria('Pasteles', 'P') (letra repetida):", almacen.agregarCategoria("Pasteles", "P"));
console.log("agregarCategoria('Snacks', 'SN') (letra inválida):", almacen.agregarCategoria("Snacks", "SN"));
console.log("agregarCategoria('Kiosco', 'K') (letra de combos):", almacen.agregarCategoria("Kiosco", "K"));

titulo("2. CREAR PRODUCTOS Y AGREGARLOS (el almacén asigna el id)");
const singani = new Producto("Singani Casa Real", "Botellas", 180, 15, 12, 20);
const fernet = new Producto("Fernet Branca", "Botellas", 220, 5, 12, 25);
const cerveza = new Producto("Cerveza Paceña", "Cervezas", 25, 120, 0, 0);
const coca = new Producto("Coca-Cola 2L", "Mezcladores", 15, 8, 0, 0);
const hielo = new Producto("Bolsa de hielo", "Insumos", 10, 20, 0, 0);
const balde = new Producto("Balde vacío", "Insumos", 30, 12, 0, 0);
const vaso = new Producto("Vaso descartable", "Insumos", 1, 100, 0, 0);

console.log("agregarProducto singani:", almacen.agregarProducto(singani));
console.log("agregarProducto fernet:", almacen.agregarProducto(fernet));
console.log("agregarProducto cerveza:", almacen.agregarProducto(cerveza));
console.log("agregarProducto coca:", almacen.agregarProducto(coca));
console.log("agregarProducto hielo:", almacen.agregarProducto(hielo));
console.log("agregarProducto balde:", almacen.agregarProducto(balde));
console.log("agregarProducto vaso:", almacen.agregarProducto(vaso));

console.log("\n-- errores al agregar --");
console.log("mismo producto otra vez:", almacen.agregarProducto(singani));
const postre = new Producto("Brownie", "Postres", 20, 10, 0, 0);
const sinCategoria = new Producto("Cigarrillos", "Tabaco", 15, 30, 0, 0);
console.log("categoría inexistente:", almacen.agregarProducto(sinCategoria));
console.log("categoría nueva (Postres):", almacen.agregarProducto(postre));

console.log("\n-- productos con sus id --");
almacen.verProductos();

titulo("3. METODOS DE PRODUCTO");

console.log("\n-- modificar --");
singani.modificar("Singani Casa Real Etiqueta Negra", 200, 12, 22);
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
const agotado = new Producto("Producto agotado", "Botellas", 1, 0, 0, 0);
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

titulo("4. CREAR COMBOS Y AGREGARLOS (el almacén asigna el id)");
const comboSingani = new Combo("Combo Singani", 220, [
  { producto: singani, cantidad: 1 },
  { producto: coca, cantidad: 2 },
  { producto: hielo, cantidad: 1 },
  { producto: vaso, cantidad: 4 },
]);
const comboFernet = new Combo("Combo Fernet", 260, [
  { producto: fernet, cantidad: 1 },
  { producto: coca, cantidad: 2 },
  { producto: hielo, cantidad: 1 },
  { producto: vaso, cantidad: 4 },
]);
const comboCerveza = new Combo("Combo Cerveza", 90, [
  { producto: cerveza, cantidad: 6 },
  { producto: hielo, cantidad: 1 },
  { producto: vaso, cantidad: 6 },
]);
console.log("agregarCombo singani:", almacen.agregarCombo(comboSingani));
console.log("agregarCombo fernet:", almacen.agregarCombo(comboFernet));
console.log("agregarCombo cerveza:", almacen.agregarCombo(comboCerveza));

console.log("\n-- errores al agregar --");
console.log("mismo combo otra vez:", almacen.agregarCombo(comboSingani));
console.log(
  "combo con producto no registrado:",
  almacen.agregarCombo(
    new Combo("Combo fantasma", 100, [{ producto: agotado, cantidad: 1 }])
  )
);

console.log("\n-- combos con sus id --");
almacen.verCombos();

titulo("5. METODOS DE COMBO");

console.log("\n-- disponibles --");
console.log("comboSingani:", comboSingani.disponibles());
console.log("comboFernet:", comboFernet.disponibles());
console.log("comboCerveza:", comboCerveza.disponibles());

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

titulo("6. METODOS DE ALMACEN");

console.log("\n-- categorías --");
almacen.verCategorias();

console.log("\n-- buscarProducto / buscarCombo --");
console.log("buscarProducto('B2'):", almacen.buscarProducto("B2").nombre);
console.log("buscarProducto('Z9'):", almacen.buscarProducto("Z9"));
console.log("buscarCombo('K2'):", almacen.buscarCombo("K2").nombre);
console.log("buscarCombo('K9'):", almacen.buscarCombo("K9"));

console.log("\n-- filtrarPorCategoria('Insumos') --");
almacen.filtrarPorCategoria("Insumos").forEach((p) => p.ver());

console.log("\n-- stockBajo(10) --");
almacen.stockBajo(10).forEach((p) => p.ver());

console.log("\n-- valorTotal --");
console.log(`${almacen.valorTotal()} Bs`);

console.log("\n-- venderCombo --");
console.log("venderCombo('K2', 1):", almacen.venderCombo("K2", 1));
console.log("venderCombo('K2', 100):", almacen.venderCombo("K2", 100));
console.log("venderCombo('K9', 1):", almacen.venderCombo("K9", 1));
almacen.verProductos();

console.log("\n-- venderVaso (licor + vaso descartable) --");
console.log("venderVaso('B2', 4, 'I3'):", almacen.venderVaso("B2", 4, "I3"));
console.log("venderVaso('M1', 1, 'I3'):", almacen.venderVaso("M1", 1, "I3"));
console.log("venderVaso('B2', 1000, 'I3'):", almacen.venderVaso("B2", 1000, "I3"));
console.log("venderVaso('Z9', 1, 'I3'):", almacen.venderVaso("Z9", 1, "I3"));
fernet.ver();
vaso.ver();

titulo("7. ELIMINAR (los id no cambian ni se reutilizan)");

console.log("\n-- antes de eliminar --");
almacen.verProductos();
almacen.verCombos();

console.log("\n-- eliminar con errores --");
console.log("eliminarProducto('B1') (lo usa un combo):", almacen.eliminarProducto("B1"));
console.log("eliminarProducto('Z7') (no existe):", almacen.eliminarProducto("Z7"));
console.log("eliminarCombo('K9') (no existe):", almacen.eliminarCombo("K9"));

console.log("\n-- eliminarCombo('K1'): los demás combos conservan su id --");
console.log("eliminarCombo('K1'):", almacen.eliminarCombo("K1"));
almacen.verCombos();

console.log("\n-- eliminarProducto('B1'): el Fernet sigue siendo B2 --");
console.log("eliminarProducto('B1'):", almacen.eliminarProducto("B1"));
almacen.verProductos();

console.log("\n-- el combo Fernet sigue funcionando con su mismo id --");
console.log("venderCombo('K2', 1):", almacen.venderCombo("K2", 1));
almacen.verCombos();

console.log("\n-- eliminarProducto('I2'): el vaso sigue siendo I3 --");
console.log("eliminarProducto('I2'):", almacen.eliminarProducto("I2"));
almacen.verProductos();

console.log("\n-- los id eliminados no se reutilizan --");
const servilletas = new Producto("Servilletas", "Insumos", 5, 50, 0, 0);
console.log("agregarProducto servilletas:", almacen.agregarProducto(servilletas));
console.log("id asignado a servilletas:", servilletas.id);
const comboMini = new Combo("Combo Cerveza Mini", 50, [
  { producto: cerveza, cantidad: 3 },
  { producto: vaso, cantidad: 3 },
]);
console.log("agregarCombo mini:", almacen.agregarCombo(comboMini));
console.log("id asignado al combo:", comboMini.id);

console.log("\n-- categorías --");
almacen.verCategorias();

console.log("\n-- aFilasProductos / aFilasCombos --");
console.log(almacen.aFilasProductos());
console.log(almacen.aFilasCombos());