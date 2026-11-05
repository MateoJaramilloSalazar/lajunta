/**
 * Tertulias · Encuesta MaxDiff -> Google Sheets
 *
 * Script vinculado a esta hoja de cálculo. Recibe las respuestas de la encuesta
 * (la página HTML alojada en GitHub Pages) y las guarda en la pestaña "Bloques".
 * Las pestañas "Resultados", "Ranking" y "Grupos" se calculan solas con fórmulas.
 *
 * Solo tiene permiso sobre ESTA hoja de cálculo.
 * @OnlyCurrentDoc
 */

var VERSION = 2;   // versión del formato de datos que envía la página
var K = 4;         // temas por bloque

var H_LEEME = 'Léeme', H_RANK = 'Ranking', H_GRUPOS = 'Grupos', H_RES = 'Resultados',
    H_TEMAS = 'Temas', H_BLOQUES = 'Bloques', H_ENVIOS = 'Envíos';

var GRUPOS = [
  'Emprendimiento y negocios',
  'Dinero y finanzas',
  'Proyecto de vida y propósito',
  'Filosofía y modelos mentales',
  'Sociedad y política',
  'Hábitos y productividad',
  'Liderazgo y carácter',
  'Emociones y salud mental',
  'Familia y relaciones'
];

// [índice de grupo, título, pregunta]. El ID de cada tema es su posición (1, 2, 3...).
// El orden debe ser idéntico al de la página HTML.
var TEMAS = [
  [0, 'Estrategia y experimentación', '¿Planear o adaptarse? Cómo equilibrar la estrategia planificada con la que surge al actuar, y cómo encontrar respuestas rápidas experimentando.'],
  [0, 'Mercadeo y experiencias memorables', '¿Cómo lograr que un producto importe al cliente y que recuerde la experiencia?'],
  [0, 'Riesgo y cisnes negros', '¿Cómo decidir cuando los datos históricos no bastan?'],
  [1, 'Inteligencia financiera básica', 'Ahorro, inversión y deudas. ¿Cómo se acumulan las ventajas con el interés compuesto?'],
  [1, 'Deuda e inversión', '¿Cuándo la deuda crea valor y cuándo amplifica errores? ¿CDT, inmueble o negocio? Cómo comparar alternativas.'],
  [1, 'Dinero, felicidad y suficiencia', '¿El dinero compra felicidad? ¿Cuánto es suficiente? Avaricia, despilfarro y administración sabia.'],
  [2, 'Legado personal', '¿Qué es lo que construimos a lo largo de nuestra vida?'],
  [2, 'Trabajo y propósito', '¿Trabajamos solo por dinero o por excelencia? ¿Qué harías si no necesitaras plata?'],
  [2, 'La muerte', 'Si te quedara un año, ¿qué cambiarías ya?'],
  [3, 'Pareto', '¿Cómo identificar el 20 % que genera más valor?'],
  [3, 'Teoría de sistemas', '¿Por qué solucionar un problema puede empeorarlo?'],
  [3, 'Aprendizaje', '¿Cómo aprender más rápido y dominar un conocimiento?'],
  [3, 'Libre albedrío', '¿Qué tanto de nuestra vida está escrito?'],
  [3, 'La búsqueda de la verdad', '¿Creemos lo que nos conviene o lo que es verdad?'],
  [4, 'Democracia', '¿Es el mejor sistema político posible?'],
  [4, 'Sociedades del futuro', '¿Cómo nos imaginamos una sociedad del futuro?'],
  [4, 'Religión y política en la sociedad', '¿Qué rol cumplen la religión y la política en nuestra sociedad?'],
  [4, 'Patriotismo y política sin pelear', '¿Cómo conversar sin dividirnos?'],
  [5, 'Productividad y disciplina', '¿Cómo administrar el tiempo y las prioridades, y construir hábitos que no dependan del ánimo?'],
  [5, 'IA y productividad', '¿Cómo usar inteligencia artificial para trabajar mejor?'],
  [5, 'Ocio que edifica vs. ocio que destruye', '¿En qué se diferencia el ocio que edifica del que destruye?'],
  [5, 'Minimalismo masculino', '¿Es un defecto? Ordenar casa, taller, carro y mente.'],
  [6, 'Liderazgo: de equipos y personal', '¿Qué hace que un equipo funcione bien? ¿Cómo liderarnos antes de liderar a otros?'],
  [6, 'Fortaleza ante el fracaso', '¿Qué hicimos con nuestro mayor fracaso?'],
  [6, 'La palabra empeñada', 'Volver a ser hombres que cumplen lo que dicen.'],
  [6, 'Mandar a la mierda con educación', 'Cómo poner límites y decir que no sin perder la compostura.'],
  [7, 'Ansiedad y rumiación', '¿Cómo dejar de darle vueltas a todo en la cabeza?'],
  [7, 'Silencio y soledad', '¿Cuándo fue la última vez que estuviste una hora sin estímulos?'],
  [7, 'Cómo pedir ayuda', 'La fortaleza de reconocer que no podemos solos.'],
  [7, 'Presión profesional e impostor', '¿Cómo gestionar la presión ante retos profesionales y evitar el síndrome del impostor?'],
  [8, 'Relaciones de pareja sanas', '¿Cómo construir una relación sana sin perder individualidad?'],
  [8, 'Ser cabeza de hogar hoy', '¿Qué significa hoy ser cabeza de hogar?'],
  [8, 'Amistad masculina', '¿Por qué nos cuesta tener amigos profundos después de los 30?'],
  [8, 'La paternidad', '¿La paternidad es cosa del pasado?']
];

/* ------------------------------------------------------------------ */
/* Menú                                                                */
/* ------------------------------------------------------------------ */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Tertulias')
    .addItem('Preparar o reiniciar pestañas de resultados', 'setup')
    .addItem('Importar un código de respuestas', 'importarCodigo')
    .addItem('Borrar respuestas de prueba', 'borrarPruebas')
    .addToUi();
}

/* ------------------------------------------------------------------ */
/* Recepción de respuestas (la página HTML hace POST a esta URL)       */
/* ------------------------------------------------------------------ */
function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) || '';
    var data = JSON.parse(raw);
    return json_(guardarConLock_(data));
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function doGet() {
  return json_({ ok: true, servicio: 'tertulias-maxdiff', version: VERSION });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function guardarConLock_(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return guardar_(data);
  } finally {
    lock.releaseLock();
  }
}

/**
 * data = { v:2, rid:'...', alias:'...', blocks:[ {i:[4 ids], m:id, l:id}, ... ] }
 * Los ids van de 1 a N (posición en la pestaña "Temas").
 * data.ping === true solo comprueba que el servicio responde, sin escribir nada.
 */
function guardar_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sT = ss.getSheetByName(H_TEMAS), sB = ss.getSheetByName(H_BLOQUES), sE = ss.getSheetByName(H_ENVIOS);
  if (!sT || !sB || !sE) throw new Error('La hoja no está preparada. Ejecuta Tertulias > Preparar pestañas.');
  var n = sT.getLastRow() - 1;
  var nb = Math.ceil(n * 2 / K);

  if (!data || data.v !== VERSION) throw new Error('Versión de datos no soportada.');
  if (data.ping === true) return { ok: true, ping: true, temas: n, bloques: nb };

  var rid = String(data.rid || '');
  if (!/^[A-Za-z0-9_-]{6,64}$/.test(rid)) throw new Error('Identificador inválido.');
  var alias = String(data.alias || '').replace(/[\r\n\t]+/g, ' ').replace(/^[=+\-@\s]+/, '').trim().slice(0, 40);

  var titulos = sT.getRange(2, 3, n, 1).getValues().map(function (r) { return String(r[0]); });
  var bl = data.blocks;
  if (!Array.isArray(bl) || bl.length !== nb) throw new Error('Se esperaban ' + nb + ' bloques.');

  var ahora = new Date();
  var filas = [];
  for (var b = 0; b < bl.length; b++) {
    var it = bl[b] && bl[b].i, m = bl[b] && bl[b].m, l = bl[b] && bl[b].l;
    if (!Array.isArray(it) || it.length !== K) throw new Error('Bloque ' + (b + 1) + ' inválido.');
    var vistos = {};
    for (var k = 0; k < K; k++) {
      var x = it[k];
      if (typeof x !== 'number' || x % 1 !== 0 || x < 1 || x > n || vistos[x]) throw new Error('Tema inválido en el bloque ' + (b + 1) + '.');
      vistos[x] = true;
    }
    if (typeof m !== 'number' || typeof l !== 'number' || !vistos[m] || !vistos[l] || m === l) {
      throw new Error('Elección inválida en el bloque ' + (b + 1) + '.');
    }
    filas.push([rid, ahora, alias, b + 1, it[0], it[1], it[2], it[3], m, l, titulos[m - 1], titulos[l - 1]]);
  }

  if (sB.getLastRow() > 30000) throw new Error('Se alcanzó el límite de filas de la hoja.');

  // Si la misma persona (mismo rid) vuelve a enviar, se reemplazan sus respuestas anteriores.
  var reemplazadas = 0;
  var ultima = sB.getLastRow();
  if (ultima > 1) {
    var ids = sB.getRange(2, 1, ultima - 1, 1).getValues();
    for (var r = ids.length - 1; r >= 0; r--) {
      if (String(ids[r][0]) === rid) { sB.deleteRow(r + 2); reemplazadas++; }
    }
  }

  var inicio = sB.getLastRow() + 1;
  var necesarias = inicio + filas.length - 1;
  if (sB.getMaxRows() < necesarias) sB.insertRowsAfter(sB.getMaxRows(), necesarias - sB.getMaxRows());
  sB.getRange(inicio, 1, filas.length, 12).setValues(filas);

  sE.appendRow([ahora, rid, alias, nb, reemplazadas ? 'reemplazó ' + reemplazadas + ' filas anteriores' : 'nuevo']);
  return { ok: true, bloques: nb, reemplazo: reemplazadas > 0 };
}

/* ------------------------------------------------------------------ */
/* Importar manualmente un código MD2|alias|... (respaldo)             */
/* ------------------------------------------------------------------ */
function importarCodigo() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Importar código', 'Pega el código que envió el participante (empieza con MD2|):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  try {
    var res = guardarConLock_(decodificar_(r.getResponseText()));
    ui.alert('Listo', 'Respuesta importada (' + res.bloques + ' bloques).', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('No se pudo importar', String((err && err.message) || err), ui.ButtonSet.OK);
  }
}

function decodificar_(txt) {
  txt = String(txt).trim();
  var p = txt.split('|');
  if (p.length !== 3 || p[0] !== 'MD2') throw new Error('El código no tiene el formato esperado.');
  var bloques = p[2].split(',').map(function (s) {
    var v = s.split('.').map(Number);
    return { i: v.slice(0, K), m: v[K], l: v[K + 1] };
  });
  var h = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, txt);
  var hex = h.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  return { v: VERSION, rid: 'cod-' + hex.slice(0, 16), alias: p[1], blocks: bloques };
}

/* ------------------------------------------------------------------ */
/* Limpieza de pruebas: borra las filas cuyo respondiente empieza por PRUEBA */
/* ------------------------------------------------------------------ */
function borrarPruebas() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('Borrar pruebas', 'Se borrarán las filas de Bloques cuyo respondiente empiece con PRUEBA. ¿Continuar?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  var sB = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H_BLOQUES);
  var ultima = sB.getLastRow(), borradas = 0;
  if (ultima > 1) {
    var ids = sB.getRange(2, 1, ultima - 1, 1).getValues();
    for (var r = ids.length - 1; r >= 0; r--) {
      if (String(ids[r][0]).indexOf('PRUEBA') === 0) { sB.deleteRow(r + 2); borradas++; }
    }
  }
  ui.alert('Listo', 'Filas borradas: ' + borradas, ui.ButtonSet.OK);
}

/* ------------------------------------------------------------------ */
/* Preparación de pestañas (se ejecuta una vez; no borra respuestas)   */
/* ------------------------------------------------------------------ */
function hoja_(ss, nombre) {
  return ss.getSheetByName(nombre) || ss.insertSheet(nombre);
}

function encabezado_(rango, color) {
  rango.setFontWeight('bold').setBackground(color || '#0E7C6B').setFontColor('#FFFFFF').setVerticalAlignment('middle');
}

function gradiente_(rango) {
  return SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#9A3B5B', SpreadsheetApp.InterpolationType.NUMBER, '-1')
    .setGradientMidpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMaxpointWithValue('#0E7C6B', SpreadsheetApp.InterpolationType.NUMBER, '1')
    .setRanges([rango]).build();
}

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var n = TEMAS.length, g = GRUPOS.length;

  // La hoja inicial pasa a ser "Léeme"
  var primera = ss.getSheets()[0];
  var sL = ss.getSheetByName(H_LEEME) || (['Hoja 1', 'Sheet1'].indexOf(primera.getName()) >= 0 ? primera.setName(H_LEEME) : ss.insertSheet(H_LEEME));
  var sK = hoja_(ss, H_RANK), sG = hoja_(ss, H_GRUPOS), sR = hoja_(ss, H_RES),
      sT = hoja_(ss, H_TEMAS), sB = hoja_(ss, H_BLOQUES), sE = hoja_(ss, H_ENVIOS);

  // --- Temas ---
  sT.clear();
  var filasT = [['ID', 'Grupo', 'Tema', 'Pregunta']];
  TEMAS.forEach(function (t, i) { filasT.push([i + 1, GRUPOS[t[0]], t[1], t[2]]); });
  sT.getRange(1, 1, filasT.length, 4).setValues(filasT);
  encabezado_(sT.getRange(1, 1, 1, 4));
  sT.setFrozenRows(1);
  sT.setColumnWidth(1, 50); sT.setColumnWidth(2, 220); sT.setColumnWidth(3, 280); sT.setColumnWidth(4, 560);
  sT.getRange(2, 4, n, 1).setWrap(true);

  // --- Bloques (no se borra si ya tiene datos) ---
  if (sB.getLastRow() === 0) {
    var hb = ['respondiente', 'fecha', 'alias', 'bloque', 'tema 1', 'tema 2', 'tema 3', 'tema 4', 'más (id)', 'menos (id)', 'más (tema)', 'menos (tema)'];
    sB.getRange(1, 1, 1, hb.length).setValues([hb]);
    encabezado_(sB.getRange(1, 1, 1, hb.length));
    sB.setFrozenRows(1);
    sB.getRange('A:A').setNumberFormat('@');
    sB.getRange('C:C').setNumberFormat('@');
    sB.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sB.setColumnWidth(1, 150); sB.setColumnWidth(2, 130); sB.setColumnWidth(3, 120);
    sB.setColumnWidth(11, 240); sB.setColumnWidth(12, 240);
  }

  // --- Envíos (registro) ---
  if (sE.getLastRow() === 0) {
    sE.getRange(1, 1, 1, 5).setValues([['fecha', 'respondiente', 'alias', 'bloques', 'nota']]);
    encabezado_(sE.getRange(1, 1, 1, 5));
    sE.setFrozenRows(1);
    sE.getRange('B:C').setNumberFormat('@');
    sE.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
    sE.setColumnWidth(1, 130); sE.setColumnWidth(2, 150); sE.setColumnWidth(5, 220);
  }

  // --- Resultados (fórmulas por tema) ---
  sR.clear();
  sR.getRange(1, 1, 1, 7).setValues([['ID', 'Grupo', 'Tema', 'Más', 'Menos', 'Mostrado', 'Puntaje']]);
  encabezado_(sR.getRange(1, 1, 1, 7));
  var f = [];
  for (var r = 2; r <= n + 1; r++) {
    f.push([
      '=Temas!A' + r, '=Temas!B' + r, '=Temas!C' + r,
      '=COUNTIF(Bloques!$I$2:$I,A' + r + ')',
      '=COUNTIF(Bloques!$J$2:$J,A' + r + ')',
      '=COUNTIF(Bloques!$E$2:$H,A' + r + ')',
      '=IF(F' + r + '=0,0,(D' + r + '-E' + r + ')/F' + r + ')'
    ]);
  }
  sR.getRange(2, 1, n, 7).setFormulas(f);
  sR.getRange(2, 7, n, 1).setNumberFormat('+0.00;-0.00;0.00');
  sR.setFrozenRows(1);
  sR.setColumnWidth(1, 40); sR.setColumnWidth(2, 220); sR.setColumnWidth(3, 280);
  sR.setConditionalFormatRules([gradiente_(sR.getRange(2, 7, n, 1))]);

  // --- Ranking (ordenado de mayor a menor puntaje) ---
  sK.clear();
  sK.getRange('A1').setValue('Personas que respondieron').setFontWeight('bold');
  sK.getRange('D1').setFormula('=IFERROR(COUNTA(UNIQUE(FILTER(Bloques!A2:A,Bloques!A2:A<>""))),0)').setFontWeight('bold').setHorizontalAlignment('left');
  sK.getRange(3, 1, 1, 7).setValues([['Posición', 'Grupo', 'Tema', 'Más', 'Menos', 'Mostrado', 'Puntaje']]);
  encabezado_(sK.getRange(3, 1, 1, 7));
  sK.getRange('A4').setFormula('=SEQUENCE(' + n + ')');
  sK.getRange('B4').setFormula('=SORT(Resultados!B2:G' + (n + 1) + ',6,FALSE,3,FALSE)');
  sK.getRange(4, 7, n, 1).setNumberFormat('+0.00;-0.00;0.00');
  sK.setFrozenRows(3);
  sK.setColumnWidth(1, 80); sK.setColumnWidth(2, 220); sK.setColumnWidth(3, 280);
  sK.setConditionalFormatRules([gradiente_(sK.getRange(4, 7, n, 1))]);

  // --- Grupos ---
  sG.clear();
  sG.getRange(1, 1, 1, 6).setValues([['Grupo', 'Temas', 'Puntaje promedio', 'Más', 'Menos', 'Posición']]);
  encabezado_(sG.getRange(1, 1, 1, 6));
  var nombres = [], fg = [];
  var rngB = 'Resultados!$B$2:$B$' + (n + 1);
  for (var i = 0; i < g; i++) {
    var rr = i + 2;
    nombres.push([GRUPOS[i]]);
    fg.push([
      '=COUNTIF(' + rngB + ',A' + rr + ')',
      '=IFERROR(AVERAGEIF(' + rngB + ',A' + rr + ',Resultados!$G$2:$G$' + (n + 1) + '),0)',
      '=SUMIF(' + rngB + ',A' + rr + ',Resultados!$D$2:$D$' + (n + 1) + ')',
      '=SUMIF(' + rngB + ',A' + rr + ',Resultados!$E$2:$E$' + (n + 1) + ')',
      '=RANK(C' + rr + ',$C$2:$C$' + (g + 1) + ',0)'
    ]);
  }
  sG.getRange(2, 1, g, 1).setValues(nombres);
  sG.getRange(2, 2, g, 5).setFormulas(fg);
  sG.getRange(2, 3, g, 1).setNumberFormat('+0.00;-0.00;0.00');
  sG.getRange('H1:J1').setValues([['Grupo (ordenado)', 'Temas', 'Puntaje promedio']]);
  encabezado_(sG.getRange('H1:J1'), '#546869');
  sG.getRange('H2').setFormula('=SORT(A2:C' + (g + 1) + ',3,FALSE)');
  sG.getRange(2, 10, g, 1).setNumberFormat('+0.00;-0.00;0.00');
  sG.setColumnWidth(1, 240); sG.setColumnWidth(3, 130); sG.setColumnWidth(8, 240); sG.setColumnWidth(10, 130);
  sG.setFrozenRows(1);
  sG.setConditionalFormatRules([gradiente_(sG.getRange(2, 3, g, 1)), gradiente_(sG.getRange(2, 10, g, 1))]);

  // --- Léeme ---
  sL.clear();
  var texto = [
    ['Tertulias · Priorización de temas (MaxDiff)'],
    [''],
    ['Cómo funciona'],
    ['Cada persona responde la encuesta (página HTML) y sus elecciones llegan solas a la pestaña "Bloques": una fila por bloque, con los 4 temas mostrados y cuál eligió como MÁS y como MENOS.'],
    ['Si la misma persona vuelve a enviar la encuesta desde el mismo navegador, sus respuestas anteriores se reemplazan (no se duplican).'],
    [''],
    ['Dónde mirar los resultados'],
    ['Ranking: los temas ordenados por puntaje. Puntaje = (veces MÁS − veces MENOS) / veces mostrado. Va de −1 a +1.'],
    ['Grupos: puntaje promedio por grupo de temas.'],
    ['Resultados: lo mismo que Ranking, pero en el orden original de la lista.'],
    [''],
    ['Cómo leerlos'],
    ['Con unas 12 personas, el error típico de un puntaje es de ±0,14. Dos temas solo se pueden distinguir con confianza si su puntaje difiere en unos 0,4 o más.'],
    ['Úsalo para decidir los temas de arriba y los de abajo con claridad. El orden fino del medio es ruido: ahí decide el criterio del grupo.'],
    [''],
    ['Importante'],
    ['No cambies el orden ni borres filas de la pestaña "Temas" después de recibir respuestas: cada respuesta guarda el número (ID) del tema, no su nombre. Corregir la redacción de un tema sí es seguro.'],
    ['No edites a mano las pestañas Resultados, Ranking ni Grupos: se calculan con fórmulas.'],
    ['Menú Tertulias: importar un código enviado por mensaje (respaldo) y borrar respuestas de prueba (las que tengan respondiente que empiece por PRUEBA).']
  ];
  sL.getRange(1, 1, texto.length, 1).setValues(texto);
  sL.setColumnWidth(1, 900);
  sL.getRange(1, 1, texto.length, 1).setWrap(true).setVerticalAlignment('top');
  sL.getRange('A1').setFontSize(16).setFontWeight('bold');
  ['A3', 'A7', 'A12', 'A16'].forEach(function (a) { sL.getRange(a).setFontWeight('bold').setFontColor('#0E7C6B'); });

  // Orden de las pestañas
  [H_LEEME, H_RANK, H_GRUPOS, H_RES, H_TEMAS, H_BLOQUES, H_ENVIOS].forEach(function (nombre, idx) {
    ss.setActiveSheet(ss.getSheetByName(nombre));
    ss.moveActiveSheet(idx + 1);
  });
  ss.setActiveSheet(ss.getSheetByName(H_RANK));
  ss.rename('Tertulias · Priorización de temas (MaxDiff)');
}
