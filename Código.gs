function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Optimizador')
    .addItem('Probar SKU de prueba', 'probarSkuDePrueba')
    .addItem('Optimizar todos', 'optimizarTodos')
    .addItem('Abrir app web local', 'abrirSidebarDemo')
    .addToUi();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Optimizador de cajas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function abrirSidebarDemo() {
  const html = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Optimizador de cajas');
  SpreadsheetApp.getUi().showSidebar(html);
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const hoja = e.range.getSheet().getName();
    const hojasQueRecalculan = ['Parametros', 'Rollos', 'Cajas'];

    if (!hojasQueRecalculan.includes(hoja)) return;

    optimizarTodosSilencioso_();
  } catch (error) {
    Logger.log('Error en onEdit: ' + error);
  }
}

function probarSkuDePrueba() {
  const codigo = '038025TD011P0000C1G1K0';
  const resultado = simularSku(codigo);

  escribirResultados([resultado.mejor], true);
  escribirResultadosDetalle(resultado.resultados, true);
  actualizarMarcaDeTiempo_();

  SpreadsheetApp.getUi().alert(
    'SKU: ' + codigo +
    '\nCaja recomendada: ' + (resultado.mejor?.tipo_caja || 'N/A') +
    '\nCantidad máxima: ' + (resultado.mejor?.cantidad_max || 0)
  );
}

function optimizarTodos() {
  const rollos = getSheetObjects_('Rollos');
  const cajas = getSheetObjects_('Cajas');

  const resultadosFinales = [];
  const resultadosDetalle = [];

  for (const rollo of rollos) {
    if (!rollo.codigo) continue;

    const matches = evaluarTodasLasCajas_(rollo);

    matches.forEach(m => resultadosDetalle.push(m));

    const mejor = matches[0] || null;

    resultadosFinales.push(formatearResultadoFinal_(rollo, mejor));
  }

  escribirResultados(resultadosFinales, true);
  escribirResultadosDetalle(resultadosDetalle, true);
  actualizarMarcaDeTiempo_();

  SpreadsheetApp.getUi().alert(
    'Proceso terminado.\n' +
    'Resultados: ' + resultadosFinales.length + ' SKUs\n' +
    'Detalle: ' + resultadosDetalle.length + ' combinaciones SKU × caja'
  );
}

function optimizarTodosSilencioso_() {
  const rollos = getSheetObjects_('Rollos');

  const resultadosFinales = [];
  const resultadosDetalle = [];

  for (const rollo of rollos) {
    if (!rollo.codigo) continue;

    const matches = evaluarTodasLasCajas_(rollo);

    matches.forEach(m => resultadosDetalle.push(m));

    const mejor = matches[0] || null;

    resultadosFinales.push(formatearResultadoFinal_(rollo, mejor));
  }

  escribirResultados(resultadosFinales, true);
  escribirResultadosDetalle(resultadosDetalle, true);
  actualizarMarcaDeTiempo_();
}

function simularSku(codigo) {
  const rollos = getSheetObjects_('Rollos');

  const rollo = rollos.find(r => String(r.codigo) === String(codigo));
  if (!rollo) {
    throw new Error('No se encontró el SKU: ' + codigo);
  }

  const resultados = evaluarTodasLasCajas_(rollo);

  return {
    rollo,
    resultados,
    mejor: resultados[0] || null
  };
}

// Función para la web app
function simularSkuWeb(codigo) {
  const data = simularSku(codigo);
  return {
    rollo: data.rollo,
    mejor: data.mejor,
    resultados: data.resultados,
    skus: getSheetObjects_('Rollos').map(r => ({
      codigo: r.codigo,
      descripcion: r.descripcion || ''
    }))
  };
}

function getSkusWeb() {
  return getSheetObjects_('Rollos').map(r => ({
    codigo: r.codigo,
    descripcion: r.descripcion || ''
  }));
}

function evaluarTodasLasCajas_(rollo) {
  const cajas = getSheetObjects_('Cajas');
  const parametros = getParametros_('Parametros');

  const matches = cajas.map(caja => calcularMatch_(rollo, caja, parametros));

  // Regla:
  // 1) mayor cantidad
  // 2) menor volumen de caja
  // 3) mayor aprovechamiento
  matches.sort((a, b) => {
    if (b.cantidad_max !== a.cantidad_max) {
      return b.cantidad_max - a.cantidad_max;
    }

    const volA = a.ancho_caja_mm * a.largo_caja_mm * a.alto_caja_mm;
    const volB = b.ancho_caja_mm * b.largo_caja_mm * b.alto_caja_mm;
    if (volA !== volB) {
      return volA - volB;
    }

    return b.aprovechamiento - a.aprovechamiento;
  });

  return matches;
}

function formatearResultadoFinal_(rollo, mejor) {
  return {
    codigo: rollo.codigo,
    descripcion: rollo.descripcion || '',
    caja_actual: rollo.caja_actual || '',
    caja_recomendada: mejor ? mejor.tipo_caja : '',
    orientacion: mejor ? mejor.orientacion : 'vertical_torre',
    ancho_rollo_mm: mejor ? mejor.ancho_rollo_mm : '',
    diametro_nominal_mm: mejor ? mejor.diametro_nominal_mm : '',
    diametro_min_mm: mejor ? mejor.diametro_min_mm : '',
    diametro_max_mm: mejor ? mejor.diametro_max_mm : '',
    diametro_calculo_mm: mejor ? mejor.diametro_calculo_mm : '',
    tolerancia_por_lado_mm: mejor ? mejor.tolerancia_por_lado_mm : '',
    ancho_efectivo_mm: mejor ? mejor.ancho_efectivo_mm : '',
    diametro_efectivo_mm: mejor ? mejor.diametro_efectivo_mm : '',
    ancho_caja_mm: mejor ? mejor.ancho_caja_mm : '',
    largo_caja_mm: mejor ? mejor.largo_caja_mm : '',
    alto_caja_mm: mejor ? mejor.alto_caja_mm : '',
    piezas_largo: mejor ? mejor.piezas_largo : 0,
    piezas_ancho: mejor ? mejor.piezas_ancho : 0,
    niveles: mejor ? mejor.niveles : 0,
    cantidad_max: mejor ? mejor.cantidad_max : 0,
    aprovechamiento: mejor ? mejor.aprovechamiento : 0,
    criterio: 'vertical_torre_max_cantidad',
    estatus: mejor ? mejor.estatus : 'SIN RESULTADO'
  };
}

function calcularMatch_(rollo, caja, parametros) {
  const toleranciaPorLado = toNumber_(parametros.tolerancia_por_lado_mm, 0.5);
  const usarDiametro = String(parametros.usar_diametro || 'maximo').toLowerCase();
  const criterio = String(parametros.criterio || 'garantizado').toLowerCase();

  const anchoRollo = toNumber_(rollo.ancho_mm, 0);
  const diametroNominal = toNumber_(rollo.diametro_nominal_mm, 0);
  const diametroMin = toNumber_(rollo.diametro_min_mm, 0);
  const diametroMax = toNumber_(rollo.diametro_max_mm, 0);

  let diametroCalculo = diametroNominal;
  if (usarDiametro === 'maximo') diametroCalculo = diametroMax || diametroNominal;
  if (usarDiametro === 'minimo') diametroCalculo = diametroMin || diametroNominal;
  if (usarDiametro === 'recomendado') diametroCalculo = Math.max(diametroNominal, diametroMax);

  const anchoCaja = toNumber_(caja.ancho_caja_mm, 0);
  const largoCaja = toNumber_(caja.largo_caja_mm, 0);
  const altoCaja = toNumber_(caja.alto_caja_mm, 0);

  // Vertical en torre:
  // base = diametro x diametro
  // altura = ancho del rollo
  const anchoEfectivo = anchoRollo + (2 * toleranciaPorLado);
  const diametroEfectivo = diametroCalculo + (2 * toleranciaPorLado);

  let piezasLargo = 0;
  let piezasAncho = 0;
  let niveles = 0;
  let cantidadMax = 0;
  let aprovechamiento = 0;
  let estatus = 'NO CABE';

  if (
    anchoEfectivo > 0 &&
    diametroEfectivo > 0 &&
    anchoCaja > 0 &&
    largoCaja > 0 &&
    altoCaja > 0
  ) {
    piezasLargo = Math.floor(largoCaja / diametroEfectivo);
    piezasAncho = Math.floor(anchoCaja / diametroEfectivo);
    niveles = Math.floor(altoCaja / anchoEfectivo);

    cantidadMax = piezasLargo * piezasAncho * niveles;

    const volumenRollo = Math.PI * Math.pow(diametroCalculo / 2, 2) * anchoRollo;
    const volumenCaja = anchoCaja * largoCaja * altoCaja;
    aprovechamiento = volumenCaja > 0 ? (cantidadMax * volumenRollo) / volumenCaja : 0;

    estatus = cantidadMax > 0 ? 'OK' : 'NO CABE';
  }

  return {
    codigo: rollo.codigo,
    descripcion: rollo.descripcion || '',
    caja_actual: rollo.caja_actual || '',
    tipo_caja: caja.tipo_caja,
    orientacion: 'vertical_torre',
    ancho_rollo_mm: anchoRollo,
    diametro_nominal_mm: diametroNominal,
    diametro_min_mm: diametroMin,
    diametro_max_mm: diametroMax,
    diametro_calculo_mm: diametroCalculo,
    tolerancia_por_lado_mm: toleranciaPorLado,
    ancho_efectivo_mm: anchoEfectivo,
    diametro_efectivo_mm: diametroEfectivo,
    ancho_caja_mm: anchoCaja,
    largo_caja_mm: largoCaja,
    alto_caja_mm: altoCaja,
    piezas_largo: piezasLargo,
    piezas_ancho: piezasAncho,
    niveles: niveles,
    cantidad_max: cantidadMax,
    aprovechamiento: aprovechamiento,
    criterio: criterio,
    estatus: estatus
  };
}

function escribirResultados(rows, limpiarPrimero) {
  writeRowsToSheet_('Resultados', rows, limpiarPrimero);
}

function escribirResultadosDetalle(rows, limpiarPrimero) {
  writeRowsToSheet_('Resultados_Detalle', rows, limpiarPrimero);
}

function writeRowsToSheet_(sheetName, rows, limpiarPrimero) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (limpiarPrimero) {
    sheet.clearContents();
  }

  if (!rows || !rows.length) {
    return;
  }

  const headers = Object.keys(rows[0]);
  const values = rows.map(r => headers.map(h => r[h]));

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);

  if (headers.includes('aprovechamiento')) {
    const col = headers.indexOf('aprovechamiento') + 1;
    sheet.getRange(2, col, values.length, 1).setNumberFormat('0.00%');
  }

  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}

function getSheetObjects_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('No existe la hoja: ' + sheetName);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => normalizeKey_(h));

  return values.slice(1)
    .filter(row => row.some(cell => cell !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((header, i) => {
        obj[header] = row[i];
      });
      return obj;
    });
}

function getParametros_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('No existe la hoja: ' + sheetName);
  }

  const values = sheet.getDataRange().getValues();
  const out = {};

  for (const row of values) {
    const key = normalizeKey_(row[0]);
    const value = row[1];

    if (key) {
      out[key] = value;
    }
  }

  return out;
}

function actualizarMarcaDeTiempo_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Parametros');
  if (!sheet) return;

  const values = sheet.getDataRange().getValues();

  for (let i = 0; i < values.length; i++) {
    const clave = String(values[i][0] || '').trim().toLowerCase();
    if (clave === 'ultima_actualizacion') {
      sheet.getRange(i + 1, 2).setValue(new Date());
      return;
    }
  }

  const nuevaFila = sheet.getLastRow() + 1;
  sheet.getRange(nuevaFila, 1).setValue('ultima_actualizacion');
  sheet.getRange(nuevaFila, 2).setValue(new Date());
}

function normalizeKey_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\w]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function toNumber_(value, fallback) {
  const num = Number(value);
  return isNaN(num) ? fallback : num;
}
