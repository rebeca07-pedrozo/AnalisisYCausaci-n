/*************************************************************
 * Bloque 2: Ejercicio 2 — Cumplimiento de política de pago
 * Requiere primerPunto.js (usa CONFIG y las utilidades).
 *************************************************************/

// ⚙️ AJUSTE 1 — Parámetros de la política.
// Ojo con la tasa: el enunciado dice 0,05% pero no dice si es diaria
// o mensual. La fórmula que dan (importe * tasa * días) solo tiene
// sentido si es DIARIA, así que la interpretamos así y lo documentamos.
const POLITICA = {
  diasLimite: 30,
  tasaDiaria: 0.0005   // 0,05%
};

const CUMPLIMIENTO = {
  cumple:      'Cumple',
  incumple:    'Incumple',
  noEncontrada:'Radicación no encontrada'
};


/*************************************************************
 * UTILIDADES
 *************************************************************/

// Deja la fecha a medianoche. Indispensable: la radicación trae hora
// (13:15:38) y la programación trae 02:00:00. Sin esto, restar fechas
// da decimales y un pago de 30 días exactos se contaría como 29.
function soloFecha(valor) {
  if (!(valor instanceof Date)) return null;
  return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
}

function diferenciaDias(desde, hasta) {
  return Math.round((hasta - desde) / (24 * 60 * 60 * 1000));
}

// Devuelve el índice de una columna de salida: si ya existe la reutiliza,
// si no la crea al final. Así puedes correr el script varias veces sin
// que se dupliquen las columnas.
function columnaSalida(hoja, encabezados, nombre) {
  var i = encabezados.indexOf(nombre);
  if (i !== -1) return i + 1;
  encabezados.push(nombre);
  var nueva = encabezados.length;
  hoja.getRange(1, nueva).setValue(nombre);
  return nueva;
}


/*************************************************************
 * ÍNDICE DE RADICACIÓN
 * Un mismo comprobante puede aparecer varias veces (4.691 casos)
 * y en 15 de ellos con fechas distintas. Nos quedamos con la MÁS
 * ANTIGUA: es la más exigente, el reloj arranca desde que la
 * factura entró por primera vez.
 *************************************************************/
function indiceRadicacion() {
  var hoja = leerHoja(CONFIG.radId, CONFIG.radHoja);
  var m = mapaColumnas(hoja.encabezados);
  var iComp  = col(m, ['Comprobante']);
  var iRadic = col(m, ['N° Radicación', 'No Radicación']);
  var iFecha = col(m, ['Fecha Radicación']);

  var porComprobante = {};
  var porRadicacion  = {};

  hoja.filas.forEach(function (fila) {
    var fecha = soloFecha(fila[iFecha]);
    if (!fecha) return;

    var comp = String(fila[iComp]).trim();
    if (comp && (!porComprobante[comp] || fecha < porComprobante[comp])) {
      porComprobante[comp] = fecha;
    }

    var rad = String(fila[iRadic]).trim();
    if (rad && (!porRadicacion[rad] || fecha < porRadicacion[rad])) {
      porRadicacion[rad] = fecha;
    }
  });

  return { porComprobante: porComprobante, porRadicacion: porRadicacion };
}


/*************************************************************
 * EJERCICIO 2A — Sobre la programación (lo que pide el enunciado)
 *************************************************************/
function ejercicio2() {
  var rad = indiceRadicacion();

  var hoja  = SpreadsheetApp.openById(CONFIG.progId).getSheetByName(CONFIG.progHoja);
  var datos = hoja.getDataRange().getValues();
  var encabezados = datos[0];
  var m = mapaColumnas(encabezados);
  var iComp  = col(m, ['Comprobante']);
  var iValor = col(m, ['Valor']);
  var iFPago = col(m, ['F Prog Pago']);

  // Se crean (o reutilizan) las 4 columnas de salida, en orden.
  var cFecha   = columnaSalida(hoja, encabezados, 'Fecha Radicación');
  var cDias    = columnaSalida(hoja, encabezados, 'Días entre radicación y pago');
  var cCumple  = columnaSalida(hoja, encabezados, 'Cumplimiento');
  var cImpacto = columnaSalida(hoja, encabezados, 'Impacto económico');

  var salida = [];
  var kpi = { cumple: 0, incumple: 0, noEncontrada: 0, impacto: 0, sumaDias: 0 };

  for (var f = 1; f < datos.length; f++) {
    var comp     = String(datos[f][iComp]).trim();
    var fechaRad = rad.porComprobante[comp] || null;
    var fechaPago= soloFecha(datos[f][iFPago]);
    var valor    = Number(datos[f][iValor]) || 0;

    if (!fechaRad || !fechaPago) {
      // No es un error del cruce: son pagos que no nacen de factura
      // radicada (ajustes, otras unidades). Se marcan, no se descartan.
      salida.push(['', '', CUMPLIMIENTO.noEncontrada, 0]);
      kpi.noEncontrada++;
      continue;
    }

    var dias = diferenciaDias(fechaRad, fechaPago);
    var estado, impacto = 0;

    if (dias <= POLITICA.diasLimite) {
      estado = CUMPLIMIENTO.cumple;
      kpi.cumple++;
    } else {
      estado = CUMPLIMIENTO.incumple;
      // Solo se cobran los días EN EXCESO, no los 30 de gracia.
      impacto = valor * POLITICA.tasaDiaria * (dias - POLITICA.diasLimite);
      kpi.incumple++;
      kpi.impacto += impacto;
    }

    kpi.sumaDias += dias;
    salida.push([fechaRad, dias, estado, impacto]);
  }

  // Escritura en bloque, columna por columna (no son contiguas si
  // el ejercicio 1 ya escribió antes).
  escribirColumna(hoja, cFecha,   salida.map(function (r) { return [r[0]]; }));
  escribirColumna(hoja, cDias,    salida.map(function (r) { return [r[1]]; }));
  escribirColumna(hoja, cCumple,  salida.map(function (r) { return [r[2]]; }));
  escribirColumna(hoja, cImpacto, salida.map(function (r) { return [r[3]]; }));

  hoja.getRange(2, cFecha, salida.length, 1).setNumberFormat('yyyy-mm-dd');
  hoja.getRange(2, cImpacto, salida.length, 1).setNumberFormat('$#,##0');

  var conCruce = kpi.cumple + kpi.incumple;
  Logger.log('PROGRAMACIÓN → Cumple: %s | Incumple: %s | Sin radicación: %s',
             kpi.cumple, kpi.incumple, kpi.noEncontrada);
  Logger.log('Promedio días (solo con cruce): %s | Impacto: $%s',
             conCruce ? (kpi.sumaDias / conCruce).toFixed(1) : 0,
             Math.round(kpi.impacto).toLocaleString());
}

function escribirColumna(hoja, columna, valores) {
  hoja.getRange(2, columna, valores.length, 1).setValues(valores);
}


/*************************************************************
 * EJERCICIO 2B — Análisis extra sobre el histórico 2025
 *
 * Por qué: la programación está armada para pegarle justo al
 * límite (máximo 30 días exactos), así que 2A da 0 incumplimientos
 * e impacto $0. El histórico sí tiene fechas de pago REALES y ahí
 * sí aparece el incumplimiento. Esto es lo que alimenta el tablero.
 *
 * El cruce acá es por N° Radicación, no por Comprobante, porque el
 * histórico sí trae esa llave.
 *************************************************************/
function ejercicio2Historico() {
  var rad = indiceRadicacion();

  var hoja = leerHoja(CONFIG.h25Id, CONFIG.h25Hoja);
  var m = mapaColumnas(hoja.encabezados);
  var iRadic = col(m, ['N° Radicación', 'No Radicación']);
  var iFecha = col(m, ['Fecha Pago']);
  var iValor = col(m, ['Valor']);
  var iNit   = col(m, ['NIT']);
  var iComp  = col(m, ['Comprobante']);

  var detalle = [['Comprobante', 'NIT', 'N° Radicación', 'Fecha Radicación',
                  'Fecha Pago', 'Días', 'Exceso', 'Valor', 'Impacto económico']];
  var kpi = { cumple: 0, incumple: 0, sinCruce: 0, impacto: 0, sumaDias: 0, maxDias: 0 };

  hoja.filas.forEach(function (fila) {
    var clave = String(fila[iRadic]).trim();
    var fechaRad = rad.porRadicacion[clave];
    var fechaPago = soloFecha(fila[iFecha]);

    if (!fechaRad || !fechaPago) { kpi.sinCruce++; return; }

    var dias  = diferenciaDias(fechaRad, fechaPago);
    var valor = Number(fila[iValor]) || 0;
    kpi.sumaDias += dias;
    if (dias > kpi.maxDias) kpi.maxDias = dias;

    if (dias <= POLITICA.diasLimite) {
      kpi.cumple++;
    } else {
      var exceso  = dias - POLITICA.diasLimite;
      var impacto = valor * POLITICA.tasaDiaria * exceso;
      kpi.incumple++;
      kpi.impacto += impacto;
      detalle.push([fila[iComp], fila[iNit], clave, fechaRad, fechaPago,
                    dias, exceso, valor, impacto]);
    }
  });

  // ⚙️ AJUSTE 2 — Si prefieres esta hoja en otro archivo, cambia el ID.
  var libro = SpreadsheetApp.openById(CONFIG.progId);
  var destino = libro.getSheetByName('Incumplimientos 2025');
  if (destino) destino.clear(); else destino = libro.insertSheet('Incumplimientos 2025');

  destino.getRange(1, 1, detalle.length, detalle[0].length).setValues(detalle);
  destino.getRange(1, 1, 1, detalle[0].length).setFontWeight('bold');
  if (detalle.length > 1) {
    destino.getRange(2, 4, detalle.length - 1, 2).setNumberFormat('yyyy-mm-dd');
    destino.getRange(2, 8, detalle.length - 1, 2).setNumberFormat('$#,##0');
  }

  var total = kpi.cumple + kpi.incumple;
  Logger.log('HISTÓRICO 2025 → Evaluados: %s | Cumple: %s | Incumple: %s | Sin radicación: %s',
             total, kpi.cumple, kpi.incumple, kpi.sinCruce);
  Logger.log('Promedio: %s días | Máximo: %s días | %% incumplimiento: %s%%',
             (kpi.sumaDias / total).toFixed(1), kpi.maxDias,
             ((kpi.incumple / total) * 100).toFixed(2));
  Logger.log('Impacto económico total: $%s', Math.round(kpi.impacto).toLocaleString());
}