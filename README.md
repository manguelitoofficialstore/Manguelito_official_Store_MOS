# MOS Store

Proyecto estático preparado para GitHub Pages.

## Archivos principales

- `index.html`: interfaz de la tienda.
- `css/styles.css`: diseño visual.
- `js/app.js`: lectura del Excel, catálogo, carrito y WhatsApp.
- `Productos.xlsx`: fuente de datos del catálogo.
- `Imagenes_Finales/`: carpeta de fotografías de los productos.
- `assets/logos/`: logos oficiales MOS y MosBasic.

## Cómo actualizar la tienda

El archivo fuente del catálogo es `Productos.xlsx`, ubicado en la raíz del proyecto.

Cuando cambies precios, inventarios, descripción, tallas, colores, categorías, referencias o nombres de fotos, reemplaza únicamente `Productos.xlsx` en GitHub y haz Commit. El código no necesita cambios.

La página vuelve a consultar el Excel cada 60 segundos y evita la caché del navegador.

## Fotografías

La carpeta se llama exactamente:

`Imagenes_Finales`

Sin tilde.

El nombre de cada archivo debe coincidir con la columna `Foto` del Excel. Si la celda no incluye extensión, la tienda prueba automáticamente:

- nombre exacto
- `.jpg`
- `.jpeg`
- `.png`
- `.webp`

Ejemplo:

Excel: `Camiseta_Manguelito_FS`

Archivo: `Imagenes_Finales/Camiseta_Manguelito_FS.jpg`

GitHub Pages diferencia mayúsculas y minúsculas, así que respeta exactamente los nombres.

## Estructura del Excel

La primera hoja se interpreta así:

- A: Referencia
- B: Colección
- C: Categoría
- D: Subcategoría
- E: Descripción
- F: Talla
- G: Género
- H: Color
- L: Precio de venta
- M: Inventario
- P: Foto

La descripción visible se construye en el orden solicitado:

`E + F + H + G`

El precio mostrado al cliente es exclusivamente la columna L.

## Orden del catálogo

1. Productos oficiales
   - Ropa
   - Accesorios
   - Items
2. Colección Basic
   - Ropa
   - Accesorios
   - Items

El código normaliza automáticamente `Accesorio` como `Accesorios` e `Item` como `Items`.

## WhatsApp

La segunda hoja del Excel puede contener todos los números de WhatsApp de ventas en la primera columna, desde A1. La tienda los muestra en el selector de finalización de compra.

Los números colombianos móviles de 10 dígitos se convierten automáticamente al formato con indicativo `57`.

## GitHub Pages

Sube el contenido de esta carpeta a la raíz del repositorio. En GitHub:

`Settings → Pages → Deploy from a branch → main → /(root)`
