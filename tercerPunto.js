/*************************************************************
 * Bloque 3: Ejercicio 3 — Revisión forma de pago
 * Requiere primerPunto.js (CONFIG, limpiarNit, leerHoja,
 * mapaColumnas, col) y segundoPunto.js (columnaSalida,
 * escribirColumna, soloFecha).
 *************************************************************/

// ⚙️ AJUSTE 1 — Traducción de "Ccl Pagos" del histórico a forma de pago.
// Estos son los códigos internos que trae la base:
//   ABONOS -> abono a cuenta Davivienda (método EFT)
//   PAGACH / CONACH -> ACH hacia otros bancos
//   CH0360, CH7000, CH1660 -> cheque (cualquier código que empiece por CH)
const MAPA_CCL = {
  'ABONOS': 'EFT: Abono Davivienda',
  'PAGACH': 'ACH: Pagos a otros bancos',
  'CONACH': 'ACH: Pagos a otros bancos'
};

const SIN_INFO = 'Sin información';

function traducirCcl(codigo) {
  if (!codigo) return null;
  var c = String(codigo).trim().toUpperCase();
  if (MAPA_CCL[c]) return MAPA_CCL[c];
  if (c.indexOf('CH') === 0) return 'Cheque';
  return c;  // código nuevo que no conocemos: se deja tal cual, visible
}


/*************************************************************
 * Acumulador: cuenta cuántas veces cada NIT usó cada forma
 * y con qué fecha fue la última. Estructura:
 *   { nit: { forma: {veces: n, ultima: Date} } }
 *************************************************************/
function sumarForma(acumulador, nit, forma, fecha) {
  if (!nit || !forma) return;
  if (!acumulador[nit]) acumulador[nit] = {};
  if (!acumulador[nit][forma]) acumulador[nit][forma] = { veces: 0, ultima: null };

  var registro = acumulador[nit][forma];
  registro.veces++;
  if (fecha instanceof Date && (!registro.ultima || fecha > registro.ultima)) {
    registro.ultima = fecha;
  }
}

// De todas las formas de un proveedor, elige la principal:
// primero la más reciente; si ninguna tiene fecha, la más usada.
function formaPrincipal(formas) {
  var mejor = null;
  for (var nombre in formas) {
    var actual = formas[nombre];
    if (!mejor) { mejor = { nombre: nombre, dato: actual }; continue; }

    var hayFechas = actual.ultima && mejor.dato.ultima;
    if (hayFechas) {
      if (actual.ultima > mejor.dato.ultima) mejor = { nombre: nombre, dato: actual };
    } else if (actual.ultima && !mejor.dato.ultima) {
      mejor = { nombre: nombre, dato: actual };
    } else if (!mejor.dato.ultima && actual.veces > mejor.dato.veces) {
      mejor = { nombre: nombre, dato: actual };
    }
  }
  return mejor ? mejor.nombre : null;
}

// Texto para la columna de listado: "EFT: Abono Davivienda (12) | PSE (3)"
function listarFormas(formas) {
  var items = [];
  for (var nombre in formas) items.push({ n: nombre, v: formas[nombre].veces });
  items.sort(function (a, b) { return b.v - a.v; });
  return items.map(function (x) { return x.n + ' (' + x.v + ')'; }).join(' | ');
}


/*************************************************************
 * FUENTE 1 (PRIORITARIA) — archivo "Formas de pago"
 * Tres hojas con estructuras distintas:
 *   PSE   -> toda la hoja es PSE, no hay columna de forma
 *   ach   -> toda la hoja es ACH
 *   Otras -> sí trae columna "Forma de Pago" (PSE, CHK, SEBRA)
 *************************************************************/
function cargarFormasDePago() {
  var acumulador = {};

  // --- Hoja PSE ---
  var pse = leerHoja(CONFIG.formasId, CONFIG.formasPse);
  var mPse = mapaColumnas(pse.encabezados);
  pse.filas.forEach(function (fila) {
    sumarForma(acumulador, limpiarNit(fila[col(mPse, ['NIT'])]),
               'PSE', soloFecha(fila[col(mPse, ['Fecha'])]));
  });

  // --- Hoja ach ---
  // Su columna de fecha viene con tipos mezclados (texto y fecha);
  // soloFecha() devuelve null en los textos y simplemente no se usa
  // para el criterio de recencia. No rompe nada.
  var ach = leerHoja(CONFIG.formasId, CONFIG.formasAch);
  var mAch = mapaColumnas(ach.encabezados);
  ach.filas.forEach(function (fila) {
    sumarForma(acumulador, limpiarNit(fila[col(mAch, ['NIT'])]),
               'ACH: Pagos a otros bancos', soloFecha(fila[col(mAch, ['Fecha de abono'])]));
  });

  // --- Hoja Otras formas de pago ---
  var otras = leerHoja(CONFIG.formasId, CONFIG.formasOtras);
  var mOtras = mapaColumnas(otras.encabezados);
  otras.filas.forEach(function (fila) {
    var forma = String(fila[col(mOtras, ['Forma de Pago'])] || '').trim();
    if (!forma) return;  // 19 filas vienen sin forma: se ignoran
    sumarForma(acumulador, limpiarNit(fila[col(mOtras, ['NIT'])]),
               forma, soloFecha(fila[col(mOtras, ['Fecha'])]));
  });

  return acumulador;
}


/*************************************************************
 * FUENTE 2 (RESPALDO) — columna "Ccl Pagos" de los históricos
 *************************************************************/
function cargarFormasHistorico() {
  var acumulador = {};

  [[CONFIG.h24Id, CONFIG.h24Hoja], [CONFIG.h25Id, CONFIG.h25Hoja]].forEach(function (par) {
    var hoja = leerHoja(par[0], par[1]);
    var m = mapaColumnas(hoja.encabezados);
    var iNit   = col(m, ['NIT']);
    var iCcl   = col(m, ['Ccl Pagos']);
    var iFecha = col(m, ['Fecha Pago']);

    hoja.filas.forEach(function (fila) {
      var forma = traducirCcl(fila[iCcl]);
      if (!forma) return;  // 243 filas sin Ccl (pagos manuales)
      sumarForma(acumulador, limpiarNit(fila[iNit]), forma, soloFecha(fila[iFecha]));
    });
  });

  return acumulador;
}


/*************************************************************
 * EJERCICIO 3
 *************************************************************/
function ejercicio3() {
  var formas    = cargarFormasDePago();
  var historico = cargarFormasHistorico();

  var hoja  = SpreadsheetApp.openById(CONFIG.progId).getSheetByName(CONFIG.progHoja);
  var datos = hoja.getDataRange().getValues();
  var encabezados = datos[0];
  var iNit = col(mapaColumnas(encabezados), ['Proveedor']);  // recuerda: es el NIT

  var cDetectada = columnaSalida(hoja, encabezados, 'Forma de Pago Detectada');
  var cListado   = columnaSalida(hoja, encabezados, 'Formas de pago del proveedor');
  var cFuente    = columnaSalida(hoja, encabezados, 'Fuente forma de pago');

  var salida = [];
  var kpiFuente = {}, kpiForma = {};

  for (var f = 1; f < datos.length; f++) {
    var nit = limpiarNit(datos[f][iNit]);

    // Prioridad literal del enunciado: primero "Formas de pago",
    // y solo si el proveedor NO aparece ahí, se usa el histórico.
    var fuente = 'Formas de pago';
    var encontrado = formas[nit];
    if (!encontrado) { encontrado = historico[nit]; fuente = 'Histórico'; }
    if (!encontrado) { encontrado = null;           fuente = SIN_INFO; }

    var detectada = encontrado ? formaPrincipal(encontrado) : SIN_INFO;
    var listado   = encontrado ? listarFormas(encontrado)   : '';

    salida.push([detectada, listado, fuente]);
    kpiFuente[fuente] = (kpiFuente[fuente] || 0) + 1;
    kpiForma[detectada] = (kpiForma[detectada] || 0) + 1;
  }

  escribirColumna(hoja, cDetectada, salida.map(function (r) { return [r[0]]; }));
  escribirColumna(hoja, cListado,   salida.map(function (r) { return [r[1]]; }));
  escribirColumna(hoja, cFuente,    salida.map(function (r) { return [r[2]]; }));

  Logger.log('Fuente usada: %s', JSON.stringify(kpiFuente));
  Logger.log('Forma detectada: %s', JSON.stringify(kpiForma));
}
/ejercicio3