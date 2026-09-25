/*************************************************************
 * CONCURSO CLAUDE + APPS SCRIPT
 * Bloque 1: Configuración, carga/normalización y Ejercicio 1
 *************************************************************/

// ⚙️ AJUSTE 1 — Pega acá los IDs de tus 5 Sheets.
// El ID es lo que va entre /d/ y /edit en la URL del archivo.
// Si mueves un archivo de carpeta el ID NO cambia; solo cambia si lo recreas.
const CONFIG = {
  progId:  'PEGA_AQUI_ID_PROGRAMACION',
  h24Id:   'PEGA_AQUI_ID_HISTORICO_2024',
  h25Id:   'PEGA_AQUI_ID_HISTORICO_2025',

  // ⚙️ AJUSTE 2 — Nombres de hoja. Si al subir a Drive el nombre cambia
  // (a veces Google renombra "sheet1" a "Hoja 1"), corrígelo acá.
  progHoja: 'sheet1',
  h24Hoja:  'Histórico Pagos 2024',
  h25Hoja:  'Histórico Pagos 2025',
  radId:   'PEGA_AQUI_ID_RADICACION',
  radHoja: 'sheet1',
  formasId:    'PEGA_AQUI_ID_FORMAS_DE_PAGO',
  formasPse:   'PSE',
  formasAch:   'ach',
  formasOtras: 'Otras formas de pago',

  // ⚙️ AJUSTE 3 — Ventana de "pagos recientes" en días.
  // Ojo: NO se cuenta desde hoy, se cuenta desde el último pago que
  // exista en el histórico (30-abr-2025). Si contáramos desde hoy,
  // la ventana quedaría vacía y no detectaríamos nada.
  diasVentana: 182  // ~6 meses
};

// Textos de las alertas, por si te piden cambiar la redacción exacta
const ALERTAS = {
  sinHistorico: 'Cuenta no registrada históricamente',
  beneficiario: 'Cambio de beneficiario',
  cuentaNueva:  'Cambio de cuenta bancaria reciente',
  ok:           'Validación OK'
};


/*************************************************************
 * UTILIDADES DE NORMALIZACIÓN
 * Sin esto los cruces fallan en silencio: las cuentas vienen
 * como texto en 2024 y como número en 2025.
 *************************************************************/

// Deja la cuenta como string de puros dígitos, sin ceros a la izquierda.
// Ej: "0470-1004-2408-9" y 470100424089 terminan igual.
function limpiarCuenta(valor) {
  if (valor === null || valor === '') return '';
  var texto = (typeof valor === 'number')
      ? valor.toFixed(0)        // evita que salga "4.701e+11"
      : String(valor);
  texto = texto.replace(/\D/g, '').replace(/^0+/, '');
  return texto;
}

// El NIT lo dejamos como string sin dígito de verificación ni guiones.
function limpiarNit(valor) {
  if (valor === null || valor === '') return '';
  var texto = (typeof valor === 'number') ? valor.toFixed(0) : String(valor);
  return texto.replace(/\D/g, '').replace(/^0+/, '');
}

// Lee una hoja completa y devuelve {encabezados, filas}
function leerHoja(id, nombreHoja) {
  var hoja = SpreadsheetApp.openById(id).getSheetByName(nombreHoja);
  if (!hoja) throw new Error('No encontré la hoja "' + nombreHoja + '". Revisa AJUSTE 2.');
  var datos = hoja.getDataRange().getValues();
  return { encabezados: datos[0], filas: datos.slice(1) };
}

// Ubica columnas por NOMBRE, no por posición.
// Así el script no se rompe si alguien inserta o mueve una columna.
function mapaColumnas(encabezados) {
  var mapa = {};
  encabezados.forEach(function (nombre, i) {
    mapa[String(nombre).trim()] = i;
  });
  return mapa;
}

// Busca una columna aceptando varios nombres posibles.
// Necesario porque 2024 la llama "Código Banco" y 2025 "Banco".
function col(mapa, posibles) {
  for (var i = 0; i < posibles.length; i++) {
    if (mapa[posibles[i]] !== undefined) return mapa[posibles[i]];
  }
  throw new Error('No encontré ninguna de estas columnas: ' + posibles.join(', '));
}


/*************************************************************
 * CARGA DEL HISTÓRICO (2024 + 2025 unificados)
 * Devuelve tres índices en memoria. Se arman en UNA sola pasada
 * por cada archivo: con 68.000 filas, cualquier búsqueda
 * anidada revienta el límite de 6 minutos de Apps Script.
 *************************************************************/
function cargarHistorico() {
  var nitsPagados   = {};  // NIT -> true  (¿alguna vez le pagamos?)
  var cuentasPorNit = {};  // NIT -> { cuenta: true }  (todas, cualquier fecha)
  var nitsPorCuenta = {};  // cuenta -> { NIT: true }  (¿de quién es esta cuenta?)
  var registros     = [];  // guardamos {nit, cuenta, fecha} para la ventana de 6 meses
  var fechaMax      = null;

  [[CONFIG.h24Id, CONFIG.h24Hoja], [CONFIG.h25Id, CONFIG.h25Hoja]].forEach(function (par) {
    var hoja = leerHoja(par[0], par[1]);
    var m = mapaColumnas(hoja.encabezados);
    var iNit    = col(m, ['NIT']);
    var iCuenta = col(m, ['Cuenta']);
    var iFecha  = col(m, ['Fecha Pago']);

    hoja.filas.forEach(function (fila) {
      var nit    = limpiarNit(fila[iNit]);
      var cuenta = limpiarCuenta(fila[iCuenta]);
      var fecha  = fila[iFecha];
      if (!nit) return;

      nitsPagados[nit] = true;

      if (cuenta) {
        if (!cuentasPorNit[nit]) cuentasPorNit[nit] = {};
        cuentasPorNit[nit][cuenta] = true;

        if (!nitsPorCuenta[cuenta]) nitsPorCuenta[cuenta] = {};
        nitsPorCuenta[cuenta][nit] = true;
      }

      if (fecha instanceof Date && cuenta) {
        registros.push({ nit: nit, cuenta: cuenta, fecha: fecha });
        if (!fechaMax || fecha > fechaMax) fechaMax = fecha;
      }
    });
  });

  // Ahora sí armamos la ventana, anclada al último pago real del histórico.
  var corte = new Date(fechaMax.getTime() - CONFIG.diasVentana * 24 * 60 * 60 * 1000);
  var cuentasRecientes = {};  // NIT -> { cuenta: true } solo dentro de la ventana
  registros.forEach(function (r) {
    if (r.fecha >= corte) {
      if (!cuentasRecientes[r.nit]) cuentasRecientes[r.nit] = {};
      cuentasRecientes[r.nit][r.cuenta] = true;
    }
  });

  Logger.log('Histórico cargado. Último pago: ' + fechaMax + ' | Corte 6m: ' + corte);
  return {
    nitsPagados: nitsPagados,
    cuentasPorNit: cuentasPorNit,
    nitsPorCuenta: nitsPorCuenta,
    cuentasRecientes: cuentasRecientes
  };
}


/*************************************************************
 * EJERCICIO 1 — Detección de inconsistencias
 *************************************************************/
function ejercicio1() {
  var hist = cargarHistorico();

  var libro = SpreadsheetApp.openById(CONFIG.progId);
  var hoja  = libro.getSheetByName(CONFIG.progHoja);
  var datos = hoja.getDataRange().getValues();
  var m = mapaColumnas(datos[0]);
  var iNit    = col(m, ['Proveedor']);           // ojo: "Proveedor" es el NIT
  var iCuenta = col(m, ['Nº Cuenta Bancaria', 'No Cuenta Bancaria']);

  var salida = [];  // [Alerta, Alerta principal]
  var conteo = {};

  for (var f = 1; f < datos.length; f++) {
    var nit    = limpiarNit(datos[f][iNit]);
    var cuenta = limpiarCuenta(datos[f][iCuenta]);
    var alertas = [];

    if (!hist.nitsPagados[nit]) {
      // Nunca le hemos pagado a este proveedor: no hay con qué comparar.
      alertas.push(ALERTAS.sinHistorico);
    } else {
      // ¿Esta cuenta le pertenece históricamente a OTRO NIT?
      var duenos = hist.nitsPorCuenta[cuenta];
      if (duenos && !duenos[nit]) alertas.push(ALERTAS.beneficiario);

      // ¿En los últimos 6 meses le pagamos a otra cuenta?
      // Si no tiene movimientos en la ventana, no alertamos: sería un
      // falso positivo por falta de datos, no por un cambio real.
      var recientes = hist.cuentasRecientes[nit];
      if (recientes && !recientes[cuenta]) alertas.push(ALERTAS.cuentaNueva);
    }

    if (alertas.length === 0) alertas.push(ALERTAS.ok);

    // ⚙️ AJUSTE 4 — Acá se decide el formato de salida.
    // "Alerta" concatena todo (6 filas tienen 2 alertas y son las más graves).
    // "Alerta principal" deja solo la más crítica, para agrupar en el tablero.
    // El orden de criticidad es el orden en que se agregaron arriba.
    var texto = alertas.join(' | ');
    salida.push([texto, alertas[0]]);
    conteo[texto] = (conteo[texto] || 0) + 1;
  }

  // Escribimos de un solo golpe al final de la hoja (nunca celda por celda).
  var colInicio = datos[0].length + 1;
  hoja.getRange(1, colInicio, 1, 2).setValues([['Alerta', 'Alerta principal']]);
  hoja.getRange(2, colInicio, salida.length, 2).setValues(salida);

  Logger.log('Resumen: ' + JSON.stringify(conteo));
}