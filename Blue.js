class producto{
    constructor(id, nombre, precio, categoria) {
        this.id = id;
        this.nombre = nombre;
        this.precio = precio;
        this.categoria = categoria;
        this.imagen = `./img/${this.id}.jpg`;
    }

    getProducto(solicitud) {
        console.log(producto.filter((l) => l.solicitud === solicitud));
    }
}

class listaProductos extends producto {
    constructor() {
        super();
        this.productos = [];
    }
}


const producto1 = new producto(1, "Camiseta", 20, "Ropa");
const producto2 = new producto(2, "Pantalón", 30, "Ropa");
const producto3 = new producto(3, "Zapatos", 50, "Calzado");
const producto4 = new producto(4, "Sombrero", 15, "Accesorios");
