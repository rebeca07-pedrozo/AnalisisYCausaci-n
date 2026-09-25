README — Cómo correr la solución
## Paso 1: Subir los archivos a Drive

Crea una carpeta en Drive y sube los 5 .xlsx. Después, abre cada uno y conviértelo a Google Sheets: Archivo → Guardar como Hojas de cálculo de Google. El script no puede leer .xlsx directamente, necesita el formato nativo.

Te quedan 5 archivos nuevos. Esos son los que vas a usar.

## Paso 2: Sacar los IDs

Abre cada Sheet convertido y mira la URL:

https://docs.google.com/spreadsheets/d/1a2B3cD4eF5gH6iJ7kL8mN9oP0qR/edit#gid=0
                                      └──────── esto es el ID ────────┘

Copia lo que está entre /d/ y /edit. Anota los 5.

## Paso 3: Verificar los nombres de las hojas

Al convertir, Google a veces renombra las pestañas. Abre cada archivo y confirma el nombre exacto de la pestaña (abajo):

Archivo	Pestaña esperada
Programación de Pagos	sheet1
Histórico 2024	Histórico Pagos 2024
Histórico 2025	Histórico Pagos 2025
Radicación de Facturas	sheet1
Formas de pago	PSE, ach, Otras formas de pago

Si alguno cambió (por ejemplo a Hoja 1), lo corriges en el CONFIG.

## Paso 4: Crear el proyecto de Apps Script

Ve a script.google.com → Nuevo proyecto. Crea tres archivos con el botón + → Secuencia de comandos:

primerPunto.gs
segundoPunto.gs
tercerEjercicio.gs

Nota: Apps Script los guarda con extensión .gs, no .js. Es lo mismo.

## Paso 5: Llenar el CONFIG

En primerPunto.gs, el bloque CONFIG debe quedar así, reemplazando cada PEGA_AQUI_... por el ID correspondiente:

javascript
const CONFIG = {
  progId:  'aquí el ID del Sheet de Programación de Pagos',
  h24Id:   'aquí el ID del Sheet de Histórico 2024',
  h25Id:   'aquí el ID del Sheet de Histórico 2025',
  radId:   'aquí el ID del Sheet de Radicación de Facturas',
  formasId:'aquí el ID del Sheet de Formas de pago',

  progHoja: 'sheet1',
  h24Hoja:  'Histórico Pagos 2024',
  h25Hoja:  'Histórico Pagos 2025',
  radHoja:  'sheet1',
  formasPse:   'PSE',
  formasAch:   'ach',
  formasOtras: 'Otras formas de pago',

  diasVentana: 182
};

Los IDs van entre comillas simples, sin espacios. Guarda con Ctrl+S.

## Paso 6: Correr, en este orden

Selecciona la función en el desplegable de arriba y dale Ejecutar:

##	Función	Archivo	Qué escribe
1	ejercicio1	primerPunto	Columnas Alerta, Alerta principal
2	ejercicio2	segundoPunto	Fecha Radicación, Días..., Cumplimiento, Impacto económico
3	ejercicio2Historico	segundoPunto	Pestaña nueva Incumplimientos 2025
4	ejercicio3	tercerEjercicio	Forma de Pago Detectada, Formas de pago del proveedor, Fuente forma de pago

La primera vez te va a pedir permisos: Revisar permisos → tu cuenta → Configuración avanzada → Ir al proyecto (no seguro) → Permitir. Es normal, es tu propio script.

## Paso 7: Validar los resultados

Abre Ver → Registro de ejecución después de cada función. Deberías ver:

Función	Resultado esperado
ejercicio1	OK 1.929 · Sin histórico 60 · Cuenta nueva 28 · Beneficiario 5 · Ambas 6
ejercicio2	Cumple 1.135 · Incumple 0 · Sin radicación 893
ejercicio2Historico	8.501 evaluados · 166 incumplimientos · máx 52 días · ~$2.358.145
ejercicio3	Histórico 1.733 · Formas de pago 250 · Sin info 45

Si los números no coinciden, no sigas. Lo más probable es que el nombre de una pestaña esté mal o que un ID apunte al archivo equivocado. Mándame el log y lo revisamos.

Si algo falla
Error	Causa
No encontré la hoja "X"	Nombre de pestaña distinto → Paso 3
No encontré ninguna de estas columnas	El archivo convertido perdió un encabezado o apuntaste al ID equivocado
Exceeded maximum execution time	Muy raro acá, pero si pasa corre las funciones de a una, no seguidas
Columnas duplicadas	No pasa: el script reutiliza las columnas si ya existen, puedes correrlo varias veces

Cuando tengas los tres corriendo y los números cuadrados, seguimos con el tablero.