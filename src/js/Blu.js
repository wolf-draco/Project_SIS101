class Producto {
  constructor(nombre, categoria, precio, stock, porciones, precioVaso, imagen = "") {
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
    console.log(
      `[${this.id}] ${this.nombre} | ${this.categoria} | ${this.precio} Bs | Stock: ${this.stock} | Vasos sueltos: ${this.vasosSueltos}`
    );
  }

  modificar(nombre, precio, porciones, precioVaso, imagen = this.imagen) {
    this.nombre = nombre;
    this.precio = precio;
    this.porciones = porciones;
    this.precioVaso = precioVaso;
    this.imagen = imagen;
  }

  aumentarStock(cantidad) {
    if (cantidad <= 0) {
      console.log("La cantidad debe ser mayor a 0");
      return false;
    }
    this.stock = this.stock + cantidad;
    return true;
  }

  reducirStock(cantidad) {
    if (cantidad <= 0) {
      console.log("La cantidad debe ser mayor a 0");
      return false;
    }
    if (cantidad > this.stock) {
      console.log(`Stock insuficiente de ${this.nombre} (hay ${this.stock})`);
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
    if (this.porciones === 0) {
      console.log(`${this.nombre} no se vende por vaso`);
      return false;
    }
    if (cantidad <= 0) {
      console.log("La cantidad debe ser mayor a 0");
      return false;
    }
    if (cantidad > this.vasosDisponibles()) {
      console.log(
        `Vasos insuficientes de ${this.nombre} (quedan ${this.vasosDisponibles()})`
      );
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
    return [
      this.id,
      this.nombre,
      this.categoria,
      this.precio,
      this.stock,
      this.porciones,
      this.precioVaso,
      this.vasosSueltos,
      this.imagen,
    ];
  }
}

class Combo {
  constructor(nombre, precio, componentes) {
    this.id = "";
    this.nombre = nombre;
    this.precio = precio;
    this.componentes = componentes;
  }

  ver() {
    console.log(
      `[${this.id}] ${this.nombre} | ${this.precio} Bs | Disponibles: ${this.disponibles()}`
    );
    this.componentes.forEach((c) => {
      console.log(`    - ${c.cantidad} x ${c.producto.nombre}`);
    });
  }

  modificar(nombre, precio, componentes) {
    this.nombre = nombre;
    this.precio = precio;
    this.componentes = componentes;
  }

  disponibles() {
    if (this.componentes.length === 0) {
      return 0;
    }
    return this.componentes.reduce((menor, c) => {
      let resto = c.producto.stock;
      let veces = 0;
      while (resto >= c.cantidad) {
        resto = resto - c.cantidad;
        veces = veces + 1;
      }
      if (menor === -1 || veces < menor) {
        return veces;
      }
      return menor;
    }, -1);
  }

  vender(cantidad) {
    if (cantidad <= 0) {
      console.log("La cantidad debe ser mayor a 0");
      return false;
    }
    if (cantidad > this.disponibles()) {
      console.log(
        `Combos insuficientes de ${this.nombre} (se pueden armar ${this.disponibles()})`
      );
      return false;
    }
    this.componentes.forEach((c) => {
      c.producto.reducirStock(c.cantidad * cantidad);
    });
    return true;
  }

  aFila() {
    const texto = this.componentes.reduce((acumulado, c) => {
      const parte = `${c.cantidad} x ${c.producto.nombre}`;
      if (acumulado === "") {
        return parte;
      }
      return acumulado + " + " + parte;
    }, "");
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
    if (this.categorias.some((c) => c.nombre === nombre)) {
      console.log(`Ya existe la categoría ${nombre}`);
      return false;
    }
    if (letra.length !== 1) {
      console.log("La letra de la categoría debe ser un solo carácter");
      return false;
    }
    if (letra === this.letraCombos) {
      console.log(`La letra ${letra} está reservada para los combos`);
      return false;
    }
    if (this.categorias.some((c) => c.letra === letra)) {
      console.log(`La letra ${letra} ya la usa otra categoría`);
      return false;
    }
    this.categorias.push({ nombre: nombre, letra: letra, ultimo: 0 });
    return true;
  }

  verCategorias() {
    this.categorias.forEach((c) => {
      const cantidad = this.productos.filter(
        (p) => p.categoria === c.nombre
      ).length;
      console.log(`${c.letra} - ${c.nombre} | Productos: ${cantidad}`);
    });
    console.log(`${this.letraCombos} - Combos | Combos: ${this.combos.length}`);
  }

  agregarProducto(producto) {
    const categoria = this.categorias.find((c) => c.nombre === producto.categoria);
    if (categoria === undefined) {
      console.log(`La categoría ${producto.categoria} no existe en el almacén`);
      return false;
    }
    if (this.productos.some((p) => p === producto)) {
      console.log(`El producto ${producto.nombre} ya fue agregado`);
      return false;
    }
    categoria.ultimo = categoria.ultimo + 1;
    producto.id = categoria.letra + categoria.ultimo;
    this.productos.push(producto);
    return true;
  }

  agregarCombo(combo) {
    if (this.combos.some((c) => c === combo)) {
      console.log(`El combo ${combo.nombre} ya fue agregado`);
      return false;
    }
    const hayNoRegistrados = combo.componentes.some(
      (c) => !this.productos.some((p) => p === c.producto)
    );
    if (hayNoRegistrados) {
      console.log(
        `El combo ${combo.nombre} usa productos que no están en el almacén`
      );
      return false;
    }
    this.ultimoCombo = this.ultimoCombo + 1;
    combo.id = this.letraCombos + this.ultimoCombo;
    this.combos.push(combo);
    return true;
  }

  buscarProducto(id) {
    return this.productos.find((p) => p.id === id);
  }

  buscarCombo(id) {
    return this.combos.find((c) => c.id === id);
  }

  eliminarProducto(id) {
    const posicion = this.productos.findIndex((p) => p.id === id);
    if (posicion === -1) {
      console.log(`No existe un producto con el id ${id}`);
      return false;
    }
    const producto = this.productos[posicion];
    const comboQueLoUsa = this.combos.find((c) =>
      c.componentes.some((k) => k.producto === producto)
    );
    if (comboQueLoUsa !== undefined) {
      console.log(`No se puede eliminar: lo usa el combo ${comboQueLoUsa.nombre}`);
      return false;
    }
    this.productos.splice(posicion, 1);
    return true;
  }

  eliminarCombo(id) {
    const posicion = this.combos.findIndex((c) => c.id === id);
    if (posicion === -1) {
      console.log(`No existe un combo con el id ${id}`);
      return false;
    }
    this.combos.splice(posicion, 1);
    return true;
  }

  verProductos() {
    if (this.productos.length === 0) {
      console.log("No hay productos en el almacén");
    } else {
      this.productos.forEach((p) => p.ver());
    }
  }

  verCombos() {
    if (this.combos.length === 0) {
      console.log("No hay combos registrados");
    } else {
      this.combos.forEach((c) => c.ver());
    }
  }

  filtrarPorCategoria(categoria) {
    return this.productos.filter((p) => p.categoria === categoria);
  }

  stockBajo(limite) {
    return this.productos.filter((p) => p.stock <= limite);
  }

  valorTotal() {
    return this.productos.reduce((total, p) => total + p.precio * p.stock, 0);
  }

  venderVaso(idProducto, cantidad, idVaso) {
    const licor = this.buscarProducto(idProducto);
    const vaso = this.buscarProducto(idVaso);
    if (licor === undefined || vaso === undefined) {
      console.log("El licor o el vaso no existen en el almacén");
      return false;
    }
    if (cantidad <= 0) {
      console.log("La cantidad debe ser mayor a 0");
      return false;
    }
    if (licor.porciones === 0) {
      console.log(`${licor.nombre} no se vende por vaso`);
      return false;
    }
    if (cantidad > licor.vasosDisponibles()) {
      console.log(
        `Vasos de licor insuficientes de ${licor.nombre} (quedan ${licor.vasosDisponibles()})`
      );
      return false;
    }
    if (cantidad > vaso.stock) {
      console.log(`Vasos descartables insuficientes (quedan ${vaso.stock})`);
      return false;
    }
    licor.venderVaso(cantidad);
    vaso.reducirStock(cantidad);
    return true;
  }

  venderCombo(idCombo, cantidad) {
    const combo = this.buscarCombo(idCombo);
    if (combo === undefined) {
      console.log(`No existe un combo con el id ${idCombo}`);
      return false;
    }
    return combo.vender(cantidad);
  }

  aFilasProductos() {
    return this.productos.map((p) => p.aFila());
  }

  aFilasCombos() {
    return this.combos.map((c) => c.aFila());
  }
}

if (typeof module !== "undefined") {
  module.exports = { Producto, Combo, Almacen };
}