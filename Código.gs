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
    piezas_por_capa: mejor ? mejor.piezas_por_capa : 0,
    niveles: mejor ? mejor.niveles : 0,
    cantidad_max: mejor ? mejor.cantidad_max : 0,
    aprovechamiento: mejor ? mejor.aprovechamiento : 0,
    variante_seleccionada: mejor ? mejor.variante_seleccionada : '',
    score_final: mejor ? mejor.score_final : 0,
    repetible: mejor ? mejor.repetible : false,
    viable_altura: mejor ? mejor.viable_altura : false,
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
  let piezasPorCapa = 0;
  let niveles = 0;
  let cantidadMax = 0;
  let aprovechamiento = 0;
  let estatus = 'NO CABE';
  let scoreFinal = 0;
  let varianteSeleccionada = 'ninguna';
  let sobranteLargoMm = 0;
  let sobranteAnchoMm = 0;
  let sobranteAltoMm = 0;
  let repetible = false;
  let viableAltura = false;
  let layoutRows = [];
  let layoutOffsets = [];
  let motivoDescarte = '';
  let variantesEvaluadas = [];

  if (
    anchoEfectivo > 0 &&
    diametroEfectivo > 0 &&
    anchoCaja > 0 &&
    largoCaja > 0 &&
    altoCaja > 0
  ) {
    const variantes = generarVariantesRepetibles_(largoCaja, anchoCaja, diametroEfectivo);
    variantesEvaluadas = variantes.map(v =>
      evaluarVarianteVertical_(v, {
        largoCaja,
        anchoCaja,
        altoCaja,
        anchoEfectivo,
        diametroEfectivo,
        diametroCalculo,
        anchoRollo
      })
    );
    const mejorVariante = seleccionarMejorVariante_(variantesEvaluadas);

    if (mejorVariante) {
      piezasLargo = mejorVariante.piezas_largo;
      piezasAncho = mejorVariante.piezas_ancho;
      piezasPorCapa = mejorVariante.piezas_por_capa;
      niveles = mejorVariante.niveles;
      cantidadMax = mejorVariante.cantidad_total;
      aprovechamiento = mejorVariante.aprovechamiento_base;
      scoreFinal = mejorVariante.score_final;
      varianteSeleccionada = mejorVariante.nombre_variante;
      sobranteLargoMm = mejorVariante.sobrante_largo_mm;
      sobranteAnchoMm = mejorVariante.sobrante_ancho_mm;
      sobranteAltoMm = mejorVariante.sobrante_alto_mm;
      repetible = mejorVariante.repetible;
      viableAltura = mejorVariante.viable_altura;
      layoutRows = mejorVariante.layout_rows || [];
      layoutOffsets = mejorVariante.layout_offsets || [];
      estatus = 'OK';
    } else {
      motivoDescarte = 'Sin variante repetible/viable por altura.';
    }

    const volumenRollo = Math.PI * Math.pow(diametroCalculo / 2, 2) * anchoRollo;
    const volumenCaja = anchoCaja * largoCaja * altoCaja;
    aprovechamiento = volumenCaja > 0 ? (cantidadMax * volumenRollo) / volumenCaja : 0;
    if (cantidadMax <= 0) {
      estatus = 'NO CABE';
    }
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
    variante_seleccionada: varianteSeleccionada,
    piezas_por_capa: piezasPorCapa,
    piezas_largo: piezasLargo,
    piezas_ancho: piezasAncho,
    niveles: niveles,
    cantidad_max: cantidadMax,
    aprovechamiento: aprovechamiento,
    score_final: scoreFinal,
    sobrante_largo_mm: sobranteLargoMm,
    sobrante_ancho_mm: sobranteAnchoMm,
    sobrante_alto_mm: sobranteAltoMm,
    repetible: repetible,
    viable_altura: viableAltura,
    motivo_descarte: motivoDescarte,
    layout_rows: JSON.stringify(layoutRows),
    layout_offsets: JSON.stringify(layoutOffsets),
    variantes_evaluadas: JSON.stringify(variantesEvaluadas),
    criterio: criterio,
    estatus: estatus
  };
}

function generarVariantesRepetibles_(largoCaja, anchoCaja, diametroEfectivo) {
  const variantes = [];
  const filasGrid = Math.floor(anchoCaja / diametroEfectivo);
  const colsGrid = Math.floor(largoCaja / diametroEfectivo);
  variantes.push({
    nombre_variante: 'grid_recto',
    tipo: 'grid',
    paso_y_mm: diametroEfectivo,
    filas: filasGrid,
    cols: colsGrid
  });

  const pasoYTresbolillo = diametroEfectivo * 0.8660254038;
  const filasTresbolillo = Math.floor(((anchoCaja - diametroEfectivo) / pasoYTresbolillo) + 1);
  variantes.push({
    nombre_variante: 'grid_alternado_tresbolillo',
    tipo: 'tresbolillo',
    paso_y_mm: pasoYTresbolillo,
    filas: Math.max(0, filasTresbolillo),
    cols_par: Math.floor(largoCaja / diametroEfectivo),
    cols_impar: Math.floor((largoCaja - (diametroEfectivo / 2)) / diametroEfectivo)
  });

  variantes.push({
    nombre_variante: 'arranque_corrido_equivalente',
    tipo: 'arranque_corrido',
    paso_y_mm: pasoYTresbolillo,
    filas: Math.max(0, filasTresbolillo),
    cols_par: Math.floor((largoCaja - (diametroEfectivo / 2)) / diametroEfectivo),
    cols_impar: Math.floor(largoCaja / diametroEfectivo)
  });

  return variantes;
}

function evaluarVarianteVertical_(variante, ctx) {
  const niveles = Math.floor(ctx.altoCaja / ctx.anchoEfectivo);
  const viableAltura = cumpleAlturaCaja_(ctx.altoCaja, niveles, ctx.anchoEfectivo);
  const repetible = esPatronRepetible_(variante);
  const nombre = variante.nombre_variante;
  const out = {
    nombre_variante: nombre,
    piezas_largo: 0,
    piezas_ancho: 0,
    piezas_por_capa: 0,
    niveles: niveles,
    cantidad_total: 0,
    aprovechamiento_base: 0,
    sobrante_largo_mm: 0,
    sobrante_ancho_mm: 0,
    sobrante_alto_mm: Math.max(0, ctx.altoCaja - (niveles * ctx.anchoEfectivo)),
    score_final: 0,
    repetible: repetible,
    viable_altura: viableAltura,
    motivo_descarte: '',
    layout_rows: [],
    layout_offsets: []
  };

  if (!repetible) {
    out.motivo_descarte = 'Patrón no repetible operativamente.';
    return out;
  }
  if (!viableAltura || niveles < 1) {
    out.motivo_descarte = 'No cumple altura útil de caja.';
    return out;
  }

  if (variante.tipo === 'grid') {
    out.piezas_largo = Math.max(0, variante.cols);
    out.piezas_ancho = Math.max(0, variante.filas);
    out.layout_rows = Array(out.piezas_ancho).fill(out.piezas_largo);
    out.layout_offsets = Array(out.piezas_ancho).fill(0);
    out.sobrante_largo_mm = Math.max(0, ctx.largoCaja - (out.piezas_largo * ctx.diametroEfectivo));
    out.sobrante_ancho_mm = Math.max(0, ctx.anchoCaja - (out.piezas_ancho * ctx.diametroEfectivo));
  } else {
    const filas = Math.max(0, variante.filas);
    const offsetsCandidatos = [0, 0.25, 0.5, 0.75];
    let mejorOffset = null;

    for (const offsetX of offsetsCandidatos) {
      const layout = construirLayoutTresbolillo_(filas, ctx.largoCaja, ctx.diametroEfectivo, offsetX);
      const anchoOcupado = layout.piezas_ancho > 0
        ? (ctx.diametroEfectivo + ((layout.piezas_ancho - 1) * variante.paso_y_mm))
        : 0;
      const candidato = {
        nombre_variante: nombre,
        piezas_largo: layout.piezas_largo,
        piezas_ancho: layout.piezas_ancho,
        piezas_por_capa: layout.layout_rows.reduce((acc, v) => acc + v, 0),
        niveles: out.niveles,
        cantidad_total: 0,
        aprovechamiento_base: 0,
        sobrante_largo_mm: Math.max(0, ctx.largoCaja - layout.max_x_ocupada_mm),
        sobrante_ancho_mm: Math.max(0, ctx.anchoCaja - anchoOcupado),
        sobrante_alto_mm: out.sobrante_alto_mm,
        score_final: 0,
        repetible: true,
        viable_altura: true,
        layout_rows: layout.layout_rows,
        layout_offsets: layout.layout_offsets
      };
      candidato.cantidad_total = candidato.piezas_por_capa * candidato.niveles;
      const areaCaja = ctx.largoCaja * ctx.anchoCaja;
      const areaRollos = candidato.piezas_por_capa * (Math.PI * Math.pow(ctx.diametroCalculo / 2, 2));
      candidato.aprovechamiento_base = areaCaja > 0 ? areaRollos / areaCaja : 0;
      candidato.score_final = calcularScoreVariante_(candidato);

      if (!mejorOffset || esMejorCandidatoTresbolillo_(candidato, mejorOffset)) {
        mejorOffset = candidato;
      }
    }

    if (mejorOffset) {
      out.layout_rows = mejorOffset.layout_rows;
      out.layout_offsets = mejorOffset.layout_offsets;
      out.piezas_ancho = mejorOffset.piezas_ancho;
      out.piezas_largo = mejorOffset.piezas_largo;
      out.piezas_por_capa = mejorOffset.piezas_por_capa;
      out.cantidad_total = mejorOffset.cantidad_total;
      out.aprovechamiento_base = mejorOffset.aprovechamiento_base;
      out.sobrante_largo_mm = mejorOffset.sobrante_largo_mm;
      out.sobrante_ancho_mm = mejorOffset.sobrante_ancho_mm;
      out.score_final = mejorOffset.score_final;
      return out;
    }
  }

  out.piezas_por_capa = out.layout_rows.reduce((acc, v) => acc + v, 0);
  out.cantidad_total = out.piezas_por_capa * out.niveles;
  const areaCaja = ctx.largoCaja * ctx.anchoCaja;
  const areaRollos = out.piezas_por_capa * (Math.PI * Math.pow(ctx.diametroCalculo / 2, 2));
  out.aprovechamiento_base = areaCaja > 0 ? areaRollos / areaCaja : 0;
  out.score_final = calcularScoreVariante_(out);
  if (out.cantidad_total <= 0) {
    out.motivo_descarte = 'No caben rollos con patrón vertical repetible.';
  }
  return out;
}

function construirLayoutTresbolillo_(filas, largoCaja, diametroEfectivo, offsetInicial) {
  const rows = [];
  const offsets = [];
  let maxXOcupada = 0;

  for (let i = 0; i < filas; i++) {
    const offset = (offsetInicial + (i % 2 === 0 ? 0 : 0.5)) % 1;
    const espacioUtil = largoCaja - (offset * diametroEfectivo);
    const cols = espacioUtil > 0 ? Math.floor(espacioUtil / diametroEfectivo) : 0;
    rows.push(Math.max(0, cols));
    offsets.push(offset);
    const xOcupada = (offset * diametroEfectivo) + (cols * diametroEfectivo);
    if (xOcupada > maxXOcupada) maxXOcupada = xOcupada;
  }

  return {
    layout_rows: rows,
    layout_offsets: offsets,
    piezas_ancho: rows.length,
    piezas_largo: rows.length ? Math.max(...rows) : 0,
    max_x_ocupada_mm: maxXOcupada
  };
}

function esMejorCandidatoTresbolillo_(a, b) {
  if (a.cantidad_total !== b.cantidad_total) return a.cantidad_total > b.cantidad_total;
  if (a.piezas_por_capa !== b.piezas_por_capa) return a.piezas_por_capa > b.piezas_por_capa;
  if (a.aprovechamiento_base !== b.aprovechamiento_base) return a.aprovechamiento_base > b.aprovechamiento_base;
  const sobranteA = a.sobrante_largo_mm + a.sobrante_ancho_mm + a.sobrante_alto_mm;
  const sobranteB = b.sobrante_largo_mm + b.sobrante_ancho_mm + b.sobrante_alto_mm;
  if (sobranteA !== sobranteB) return sobranteA < sobranteB;
  return a.score_final > b.score_final;
}

function cumpleAlturaCaja_(altoCaja, niveles, anchoEfectivo) {
  return (niveles * anchoEfectivo) <= altoCaja;
}

function esPatronRepetible_(variante) {
  if (!variante || !variante.nombre_variante) return false;
  const permitidas = ['grid_recto', 'grid_alternado_tresbolillo', 'arranque_corrido_equivalente'];
  return permitidas.includes(variante.nombre_variante);
}

function calcularScoreVariante_(v) {
  const sobranteTotal = v.sobrante_largo_mm + v.sobrante_ancho_mm + v.sobrante_alto_mm;
  const simpleBonus = v.nombre_variante === 'grid_recto' ? 0.002 : 0;
  return (v.cantidad_total * 1000000) +
    (v.aprovechamiento_base * 1000) -
    sobranteTotal +
    simpleBonus;
}

function seleccionarMejorVariante_(variantes) {
  const candidatas = (variantes || []).filter(v =>
    v.repetible &&
    v.viable_altura &&
    v.cantidad_total > 0
  );
  if (!candidatas.length) return null;

  candidatas.sort((a, b) => {
    if (b.cantidad_total !== a.cantidad_total) return b.cantidad_total - a.cantidad_total;
    if (b.aprovechamiento_base !== a.aprovechamiento_base) return b.aprovechamiento_base - a.aprovechamiento_base;
    const sobranteA = a.sobrante_largo_mm + a.sobrante_ancho_mm + a.sobrante_alto_mm;
    const sobranteB = b.sobrante_largo_mm + b.sobrante_ancho_mm + b.sobrante_alto_mm;
    if (sobranteA !== sobranteB) return sobranteA - sobranteB;
    return b.score_final - a.score_final;
  });

  const mejor = candidatas[0];
  const segunda = candidatas[1];
  if (segunda) {
    const gapCantidad = Math.abs(mejor.cantidad_total - segunda.cantidad_total);
    const gapPct = mejor.cantidad_total > 0 ? gapCantidad / mejor.cantidad_total : 1;
    if (gapPct <= 0.03) {
      const preferidaSimple = [mejor, segunda].sort((a, b) => {
        const simpA = a.nombre_variante === 'grid_recto' ? 0 : 1;
        const simpB = b.nombre_variante === 'grid_recto' ? 0 : 1;
        if (simpA !== simpB) return simpA - simpB;
        return b.score_final - a.score_final;
      })[0];
      return preferidaSimple;
    }
  }
  return mejor;
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
