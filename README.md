# Cocinería Sra. Jake · Pedidos

App web para recibir pedidos de almuerzo del día:

- **Público general** ve el menú con la disponibilidad en vivo y hace su pedido (paga al retirar).
- **Portal de empresas** (`#empresa`): la empresa entra con su código y cada trabajador se anota día a día en la lista de su empresa con su plato. Ven la lista del día, pueden quitar a alguien antes del cierre y revisan su consumo del mes.
- **Cocina y caja** ven cuántas porciones preparar de cada plato y agregado, avanzan los pedidos (pendiente → listo → entregado), ven la caja del día y la cuenta mensual de cada empresa.
- **Menú del día**: platos, porciones disponibles, agregados, precios, hora de cierre de reservas de empresa, empresas (código, contacto, email, RUT) y datos del negocio.
- **Balances** (`#balances`): ventas entre dos fechas (hoy, semana, mes, mes anterior, año o un rango a elección): total vendido, público cobrado y por cobrar, empresas a cuenta, desglose por empresa, por plato, por agregado y por día. Se descarga en Excel (CSV) o en PDF, y se puede enviar por correo.
- **Cotizaciones y detalles** (`#documentos`):
  - *Detalle de consumo* para una empresa en un período: colaciones por día y por persona, con el total a pagar.
  - *Cotizaciones* numeradas para clientes nuevos, con calculadora de plan mensual (personas × días), descuento, validez y condiciones.
  - Cada documento se ve en pantalla y se puede **descargar en PDF, imprimir, enviar por correo con el PDF adjunto** o mandar el resumen por **WhatsApp**.

Funciona con **Google Apps Script** como backend y una **planilla de Google Sheets** como base de datos (se crea sola).

## Reglas del negocio

| Regla | Valor inicial | Dónde se cambia |
|---|---|---|
| Colación normal (fondo + agregado + sopa, ensalada, pan y jugo) | $5.000 | Menú del día |
| Hipocalórico: sin agregado (arroz, tallarín o puré) | +$1.000 | Menú del día |
| Cierre de reservas de empresa | 10:30 | Menú del día |
| Porciones por plato | Pollo 40 · Cazuela 25 · Panitas 15 · Carne a la olla 30 | Menú del día |
| PIN de cocina | `1234` | Menú del día → "PIN nuevo" (**cámbialo al instalar**) |

- Los platos que no llevan agregado (la cazuela) no tienen opción hipocalórica.
- El público general paga al retirar. Las empresas cargan a su cuenta y se les cobra a fin de mes.
- Cada empresa entra al portal con su **código** (se genera al agregar la empresa; entrégaselo al encargado). También puedes mandarles el enlace directo `.../?empresa=CODIGO#empresa` (en GitHub Pages).
- Cada empresa tiene **una lista por día**: la cocina la ve como un solo pedido con el nombre de cada persona.
- La lista se cierra a la hora de cierre, o antes si la cocina la marca como "listo". Después, quien falte pide como público general según lo que quede.
- Un pedido anulado devuelve sus porciones al cupo del día.

## Archivos

- `index.html` — la app (cliente, cocina y menú). Se publica en GitHub Pages y también la sirve Apps Script.
- `Codigo.gs` — backend: lee y escribe la planilla, valida cupos con bloqueo para no vender de más.
- `appsscript.json` — manifest de Apps Script.
- `.clasp.json.example` — plantilla para [`clasp`](https://github.com/google/clasp).

## Puesta en marcha

1. En <https://script.google.com> crea un proyecto nuevo y copia su **ID de script** (Configuración del proyecto).
2. Sube el código:
   ```bash
   npm install -g @google/clasp
   clasp login
   cp .clasp.json.example .clasp.json   # pega el scriptId
   clasp push
   ```
   (O copia `Codigo.gs` e `index.html` a mano en el editor de Apps Script.)
3. En el editor ejecuta la función **`instalar`** una vez: crea la planilla "Cocinería Sra. Jake · Datos" en tu Drive y pide los permisos (planillas y **envío de correos** desde tu cuenta de Gmail, para cotizaciones y detalles). El registro muestra el enlace a la planilla.
4. **Implementar → Nueva implementación → Aplicación web** (Ejecutar como: *Yo*; Acceso: *Cualquier usuario*). Copia la URL `/exec`.
5. Pega esa URL en `API_URL` al inicio del `<script>` de `index.html` y haz push a `main`; GitHub Pages publica la página sola.

Puedes usar cualquiera de los dos enlaces:

- URL `/exec` de Apps Script (funciona apenas implementas).
- `https://<usuario>.github.io/cocineria-sra-jake/` (necesita `API_URL`).

Enlaces por sección: agrega `#empresa` al final para el portal de empresas, o `#cocina` para cocina (entra con el PIN).

Si cambias `Codigo.gs`, vuelve a hacer `clasp push` y crea una **nueva versión** de la implementación.

## Hojas de la planilla

`CONFIG`, `MENU`, `AGREGADOS`, `EMPRESAS`, `PEDIDOS` (un registro por pedido), `ITEMS` (un registro por plato pedido, útil para reportes) y `COTIZACIONES`. El detalle de columnas está al inicio de `Codigo.gs`.

## Correos

Los correos salen desde la cuenta de Google que implementó el script (con el nombre del negocio como remitente y el email de contacto como dirección de respuesta). Una cuenta Gmail gratuita puede enviar unos 100 correos al día.
