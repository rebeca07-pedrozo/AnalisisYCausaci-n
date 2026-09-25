/*************************************************************
 * TABLERO DE CONTROL DE PAGOS — Davivienda
 * Servidor: lee las hojas ya procesadas y las sirve al HTML.
 * Requiere primerPunto.gs (CONFIG, leerHoja, mapaColumnas, col).
 *************************************************************/

// ⚙️ AJUSTE 1 — Pega aquí tu logo en base64.
// Si tu cadena YA empieza con "data:image/png;base64," pégala completa.
// Si es solo la cadena cruda (empieza con "iVBOR..." o "/9j/..."),
// pégala igual: el código le agrega el prefijo solo.
const LOGO_BASE64 = 'PEGA_AQUI_TU_BASE64';

// ⚙️ AJUSTE 2 — Tipo de imagen del logo. Cámbialo si tu logo es jpg o svg.
const LOGO_TIPO = 'image/png';

// Nombre de la pestaña que creó ejercicio2Historico
const HOJA_INCUMPLIMIENTOS = 'Incumplimientos 2025';


/*************************************************************
 * Punto de entrada del web app
 *************************************************************/
function doGet() {
  var plantilla = HtmlService.createTemplateFromFile('dashboard');
  plantilla.logo = construirLogo();
  return plantilla.evaluate()
    .setTitle('Tablero de Control de Pagos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function construirLogo() {
  var b64 = String(LOGO_BASE64).trim();
  if (!b64 || b64.indexOf('PEGA_AQUI') === 0) return '';
  if (b64.indexOf('data:') === 0) return b64;          // ya viene completo
  return 'data:' + LOGO_TIPO + ';base64,' + b64;       // le ponemos el prefijo
}


/*************************************************************
 * obtenerDatos() — lo que el HTML pide por google.script.run
 *
 * Devuelve las 2.028 filas de programación (pesan poco y permiten
 * filtrar en el navegador sin volver al servidor) más el detalle
 * de incumplimientos y los KPIs guardados del histórico.
 *************************************************************/
function obtenerDatos() {
  return {
    generado:  Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    programacion: leerProgramacion(),
    incumplimientos: leerIncumplimientos(),
    kpiHistorico: leerKpiHistorico()
  };
}

function leerProgramacion() {
  var hoja = leerHoja(CONFIG.progId, CONFIG.progHoja);
  var m = mapaColumnas(hoja.encabezados);

  var iUn      = col(m, ['Unidad Negocio']);
  var iComp    = col(m, ['Comprobante']);
  var iNit     = col(m, ['Proveedor']);
  var iNombre  = col(m, ['Nombre Proveedor']);
  var iValor   = col(m, ['Valor']);
  var iBanco   = col(m, ['Banco']);
  var iAlerta  = col(m, ['Alerta']);
  var iAlertaP = col(m, ['Alerta principal']);
  var iDias    = col(m, ['Días entre radicación y pago']);
  var iCumple  = col(m, ['Cumplimiento']);
  var iForma   = col(m, ['Forma de Pago Detectada']);
  var iFuente  = col(m, ['Fuente forma de pago']);

  return hoja.filas.map(function (f) {
    return {
      un:      String(f[iUn] || ''),
      comp:    String(f[iComp] || ''),
      nit:     String(f[iNit] || ''),
      nombre:  String(f[iNombre] || ''),
      valor:   Number(f[iValor]) || 0,
      banco:   String(f[iBanco] || 'Sin banco'),
      alerta:  String(f[iAlerta] || ''),
      alertaP: String(f[iAlertaP] || ''),
      dias:    f[iDias] === '' ? null : Number(f[iDias]),
      cumple:  String(f[iCumple] || ''),
      forma:   String(f[iForma] || ''),
      fuente:  String(f[iFuente] || '')
    };
  });
}

function leerIncumplimientos() {
  var libro = SpreadsheetApp.openById(CONFIG.progId);
  var hoja = libro.getSheetByName(HOJA_INCUMPLIMIENTOS);
  if (!hoja || hoja.getLastRow() < 2) return [];

  var datos = hoja.getDataRange().getValues();
  var m = mapaColumnas(datos[0]);
  var iNit     = col(m, ['NIT']);
  var iDias    = col(m, ['Días']);
  var iValor   = col(m, ['Valor']);
  var iImpacto = col(m, ['Impacto económico']);

  return datos.slice(1).map(function (f) {
    return {
      nit:     String(f[iNit] || ''),
      dias:    Number(f[iDias]) || 0,
      valor:   Number(f[iValor]) || 0,
      impacto: Number(f[iImpacto]) || 0
    };
  });
}

function leerKpiHistorico() {
  var guardado = PropertiesService.getScriptProperties().getProperty('kpiHistorico');
  // Si aún no has corrido ejercicio2Historico, el tablero no se rompe:
  // muestra ceros y un aviso en vez de fallar.
  return guardado ? JSON.parse(guardado)
                  : { evaluados: 0, cumple: 0, incumple: 0, impacto: 0, promDias: 0, maxDias: 0 };
}