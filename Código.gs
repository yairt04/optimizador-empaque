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

const UMBRALES_SUGERENCIA_CAJA_ = {
  aprovechamiento_bajo: 0.72,
  sobrante_factor_alto: 0.75,
  mejora_min_cantidad: 1,
  mejora_min_aprovechamiento: 0.03,
  mejora_min_sobrante_ratio: 0.10,
  mejora_min_sobrante_abs_mm: 10
};

const CACHE_TTL_SEGUNDOS_ = 300;
const PERF_LOGS_ACTIVOS_ = true;
const CACHE_KEYS_ = {
  rollos: 'opt:rollos:v1',
  cajas: 'opt:cajas:v1',
  parametros: 'opt:parametros:v1',
  skus_web: 'opt:skus_web:v1'
};

function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const hoja = e.range.getSheet().getName();
    const hojasQueRecalculan = ['Parametros', 'Rollos', 'Cajas'];

    if (!hojasQueRecalculan.includes(hoja)) return;
    clearDataCaches_();

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
  const dataCtx = getDataCtxCached_();
  const rollos = dataCtx.rollos;

  const resultadosFinales = [];
  const resultadosDetalle = [];

  for (const rollo of rollos) {
    if (!rollo.codigo) continue;

    const matches = evaluarTodasLasCajas_(rollo, dataCtx);

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
  const dataCtx = getDataCtxCached_();
  const rollos = dataCtx.rollos;

  const resultadosFinales = [];
  const resultadosDetalle = [];

  for (const rollo of rollos) {
    if (!rollo.codigo) continue;

    const matches = evaluarTodasLasCajas_(rollo, dataCtx);

    matches.forEach(m => resultadosDetalle.push(m));

    const mejor = matches[0] || null;

    resultadosFinales.push(formatearResultadoFinal_(rollo, mejor));
  }

  escribirResultados(resultadosFinales, true);
  escribirResultadosDetalle(resultadosDetalle, true);
  actualizarMarcaDeTiempo_();
}

function simularSku(codigo, dataCtx) {
  const ctx = dataCtx || getDataCtxCached_();
  const rollos = ctx.rollos;

  const rollo = rollos.find(r => String(r.codigo) === String(codigo));
  if (!rollo) {
    throw new Error('No se encontró el SKU: ' + codigo);
  }

  const resultados = evaluarTodasLasCajas_(rollo, ctx);

  return {
    rollo,
    resultados,
    mejor: resultados[0] || null
  };
}

// Función para la web app
function simularSkuWeb(codigo) {
  timeStart_('simularSkuWeb.total');
  timeStart_('simularSkuWeb.carga_datos');
  const dataCtx = getDataCtxCached_();
  timeEnd_('simularSkuWeb.carga_datos');

  const data = simularSku(codigo, dataCtx);
  const parametros = dataCtx.parametros;

  timeStart_('simularSkuWeb.caja_ideal');
  const cajaIdeal = calcularCajaIdeal_(data.rollo, parametros, data.mejor);
  timeEnd_('simularSkuWeb.caja_ideal');

  const comparacionCajaIdeal = compararCajaIdealVsCatalogo_(data.mejor, cajaIdeal);
  const sugerencia = construirSugerenciaCaja_(data.resultados, data.mejor);

  timeStart_('simularSkuWeb.skus_web');
  const skus = getSkusWebCached_();
  timeEnd_('simularSkuWeb.skus_web');
  timeEnd_('simularSkuWeb.total');

  const visualizacion = crearPackingVisualization_(data.mejor);

  return {
    rollo: data.rollo,
    mejor: data.mejor,
    resultados: data.resultados,
    visualizacion: visualizacion,
    caja_ideal: cajaIdeal,
    comparacion_caja_ideal: comparacionCajaIdeal,
    sugerencia: sugerencia,
    skus
  };
}

function getSkusWeb() {
  return getSkusWebCached_();
}

function evaluarTodasLasCajas_(rollo, dataCtx) {
  timeStart_('evaluarTodasLasCajas_');
  const ctx = dataCtx || getDataCtxCached_();
  const cajas = ctx.cajas;
  const parametros = ctx.parametros;

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

  timeEnd_('evaluarTodasLasCajas_');
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

function calcularCajaIdeal_(rollo, parametros, mejorCatalogo) {
  timeStart_('calcularCajaIdeal_');
  const toleranciaPorLado = toNumber_(parametros?.tolerancia_por_lado_mm, 0.5);
  const usarDiametro = String(parametros?.usar_diametro || 'maximo').toLowerCase();
  const anchoRollo = toNumber_(rollo?.ancho_mm, 0);
  const diametroNominal = toNumber_(rollo?.diametro_nominal_mm, 0);
  const diametroMin = toNumber_(rollo?.diametro_min_mm, 0);
  const diametroMax = toNumber_(rollo?.diametro_max_mm, 0);

  let diametroCalculo = diametroNominal;
  if (usarDiametro === 'maximo') diametroCalculo = diametroMax || diametroNominal;
  if (usarDiametro === 'minimo') diametroCalculo = diametroMin || diametroNominal;
  if (usarDiametro === 'recomendado') diametroCalculo = Math.max(diametroNominal, diametroMax);

  const anchoEfectivo = anchoRollo + (2 * toleranciaPorLado);
  const diametroEfectivo = diametroCalculo + (2 * toleranciaPorLado);
  if (anchoEfectivo <= 0 || diametroEfectivo <= 0) return null;

  const margenLargo = toNumber_(parametros?.margen_seguridad_largo_mm, 6);
  const margenAncho = toNumber_(parametros?.margen_seguridad_ancho_mm, 6);
  const margenAlto = toNumber_(parametros?.margen_seguridad_alto_mm, 4);
  const alturaUtilMax = toNumber_(
    parametros?.altura_util_max_mm,
    Math.max(1, toNumber_(mejorCatalogo?.alto_caja_mm, Infinity))
  );

  const baseCapasCatalogo = Math.max(1, toNumber_(mejorCatalogo?.piezas_por_capa, 1));
  const baseNivelesCatalogo = Math.max(1, toNumber_(mejorCatalogo?.niveles, 1));
  const baseTotalCatalogo = Math.max(1, toNumber_(mejorCatalogo?.cantidad_max, 1));
  const baseColumnasCatalogo = Math.max(1, toNumber_(mejorCatalogo?.piezas_largo, Math.ceil(Math.sqrt(baseCapasCatalogo))));
  const baseFilasCatalogo = Math.max(1, toNumber_(mejorCatalogo?.piezas_ancho, Math.ceil(baseCapasCatalogo / baseColumnasCatalogo)));

  const maxPiezasPorCapa = Math.min(80, Math.max(baseCapasCatalogo + 10, Math.ceil(baseTotalCatalogo / baseNivelesCatalogo) + 4));
  const maxNiveles = Math.min(20, Math.max(baseNivelesCatalogo + 3, Math.ceil(baseTotalCatalogo / baseCapasCatalogo) + 2));
  const maxColumnas = Math.max(1, Math.ceil(maxPiezasPorCapa / Math.max(1, baseFilasCatalogo)) + 2);
  const maxFilas = Math.max(1, Math.ceil(maxPiezasPorCapa / Math.max(1, baseColumnasCatalogo)) + 2);

  const candidatas = [];
  const candidatasKeySet = {};
  const columnasRango = construirRangoCercano_(baseColumnasCatalogo, 3, 1, maxColumnas);
  const filasRango = construirRangoCercano_(baseFilasCatalogo, 3, 1, maxFilas);
  const nivelesRango = construirRangoCercano_(baseNivelesCatalogo, 2, 1, maxNiveles);

  const combinaciones = [];
  for (const columnas of columnasRango) {
    for (const filas of filasRango) {
      const piezasPorCapaEstimadas = columnas * filas;
      if (piezasPorCapaEstimadas > maxPiezasPorCapa) continue;

      for (const niveles of nivelesRango) {
        const altoSinMargen = niveles * anchoEfectivo;
        if (altoSinMargen > alturaUtilMax) continue;
        const altoIdealMinimo = altoSinMargen + margenAlto;
        if (altoIdealMinimo > alturaUtilMax) continue;

        combinaciones.push({
          columnas,
          filas,
          niveles,
          distancia: Math.abs(columnas - baseColumnasCatalogo) +
            Math.abs(filas - baseFilasCatalogo) +
            Math.abs(niveles - baseNivelesCatalogo)
        });
      }
    }
  }
  combinaciones.sort((a, b) => a.distancia - b.distancia);

  let mejorScore = -Infinity;
  let iteracionesSinMejora = 0;
  const limiteCombinaciones = 260;

  for (let i = 0; i < combinaciones.length && i < limiteCombinaciones; i++) {
    const item = combinaciones[i];
    const variantes = ['grid_recto', 'grid_alternado_tresbolillo', 'arranque_corrido_equivalente'];

    for (const nombreVariante of variantes) {
      const candidata = buildCandidataIdeal_(nombreVariante, {
        columnas: item.columnas,
        filas: item.filas,
        niveles: item.niveles,
        diametroEfectivo,
        anchoEfectivo,
        margenLargo,
        margenAncho,
        margenAlto,
        alturaUtilMax
      });
      if (!candidata) continue;
      pushCandidataIdealUnica_(candidatas, candidatasKeySet, candidata);

      if (candidata.viable_altura && candidata.repetible && candidata.cantidad_total > 0 && candidata.score_final > mejorScore) {
        mejorScore = candidata.score_final;
        iteracionesSinMejora = 0;
      }
    }

    iteracionesSinMejora++;
    if (mejorScore > -Infinity && iteracionesSinMejora >= 90 && item.distancia > 4) {
      break;
    }
  }

  const viables = candidatas.filter(c => c.viable_altura && c.repetible && c.cantidad_total > 0);
  if (!viables.length) {
    timeEnd_('calcularCajaIdeal_');
    return null;
  }

  viables.sort((a, b) => b.score_final - a.score_final);
  const mejor = viables[0];
  const volumenRollo = Math.PI * Math.pow(diametroCalculo / 2, 2) * anchoRollo;
  const volumenCaja = mejor.ancho_caja_mm * mejor.largo_caja_mm * mejor.alto_caja_mm;
  const aprovechamiento = volumenCaja > 0 ? (mejor.cantidad_total * volumenRollo) / volumenCaja : 0;

  const salida = {
    tipo_caja: 'CAJA_IDEAL_CALCULADA',
    orientacion: 'vertical_torre',
    patron: mejor.nombre_variante,
    variante_seleccionada: mejor.nombre_variante,
    ancho_rollo_mm: anchoRollo,
    diametro_calculo_mm: diametroCalculo,
    diametro_efectivo_mm: diametroEfectivo,
    ancho_efectivo_mm: anchoEfectivo,
    largo_caja_mm: mejor.largo_caja_mm,
    ancho_caja_mm: mejor.ancho_caja_mm,
    alto_caja_mm: mejor.alto_caja_mm,
    piezas_largo: mejor.piezas_largo,
    piezas_ancho: mejor.piezas_ancho,
    piezas_por_capa: mejor.piezas_por_capa,
    niveles: mejor.niveles,
    cantidad_max: mejor.cantidad_total,
    aprovechamiento: aprovechamiento,
    score_final: mejor.score_final,
    repetible: true,
    viable_altura: true,
    layout_rows: JSON.stringify(mejor.layout_rows || []),
    layout_offsets: JSON.stringify(mejor.layout_offsets || []),
    estatus: 'IDEAL'
  };
  timeEnd_('calcularCajaIdeal_');
  return salida;
}

function buildCandidataIdeal_(nombreVariante, ctx) {
  const SQRT_3_HALF = 0.8660254038;
  const columnas = Math.max(1, toNumber_(ctx.columnas, 1));
  const filas = Math.max(1, toNumber_(ctx.filas, 1));
  const niveles = Math.max(1, toNumber_(ctx.niveles, 1));
  const diametroEfectivo = toNumber_(ctx.diametroEfectivo, 0);
  const anchoEfectivo = toNumber_(ctx.anchoEfectivo, 0);
  if (diametroEfectivo <= 0 || anchoEfectivo <= 0) return null;

  let largoMin = 0;
  let anchoMin = 0;
  let layoutRows = [];
  let layoutOffsets = [];

  if (nombreVariante === 'grid_recto') {
    largoMin = (columnas * diametroEfectivo) + ctx.margenLargo;
    anchoMin = (filas * diametroEfectivo) + ctx.margenAncho;
    layoutRows = Array(filas).fill(columnas);
    layoutOffsets = Array(filas).fill(0);
  } else if (nombreVariante === 'grid_alternado_tresbolillo' || nombreVariante === 'arranque_corrido_equivalente') {
    const offsetInicial = nombreVariante === 'grid_alternado_tresbolillo' ? 0 : 0.5;
    const layout = construirLayoutTresbolilloConColumnas_(filas, columnas, offsetInicial);
    const maxAnchoFilas = layout.max_x_ocupada_factores_d * diametroEfectivo;
    largoMin = maxAnchoFilas + ctx.margenLargo;
    anchoMin = diametroEfectivo + ((filas - 1) * diametroEfectivo * SQRT_3_HALF) + ctx.margenAncho;
    layoutRows = layout.layout_rows;
    layoutOffsets = layout.layout_offsets;
  } else {
    return null;
  }

  const altoMin = (niveles * anchoEfectivo) + ctx.margenAlto;
  const largoRedondeado = redondearDimensionIndustrial_(largoMin);
  const anchoRedondeado = redondearDimensionIndustrial_(anchoMin);
  const altoRedondeado = redondearDimensionIndustrial_(altoMin);
  const viableAltura = altoRedondeado <= ctx.alturaUtilMax;
  const piezasPorCapa = layoutRows.reduce((acc, v) => acc + v, 0);
  const cantidadTotal = piezasPorCapa * niveles;
  const volumenCaja = largoRedondeado * anchoRedondeado * altoRedondeado;
  const densidad = volumenCaja > 0 ? cantidadTotal / volumenCaja : 0;

  return {
    nombre_variante: nombreVariante,
    piezas_largo: Math.max(...layoutRows),
    piezas_ancho: layoutRows.length,
    piezas_por_capa: piezasPorCapa,
    niveles: niveles,
    cantidad_total: cantidadTotal,
    largo_caja_mm: largoRedondeado,
    ancho_caja_mm: anchoRedondeado,
    alto_caja_mm: altoRedondeado,
    repetible: esPatronRepetible_({ nombre_variante: nombreVariante }),
    viable_altura: viableAltura,
    layout_rows: layoutRows,
    layout_offsets: layoutOffsets,
    score_final: calcularScoreCajaIdeal_(cantidadTotal, densidad, volumenCaja, nombreVariante)
  };
}

function pushCandidataIdealUnica_(candidatas, keySet, candidata) {
  if (!candidata) return;
  const key = [
    toNumber_(candidata.largo_caja_mm, 0),
    toNumber_(candidata.ancho_caja_mm, 0),
    toNumber_(candidata.alto_caja_mm, 0),
    candidata.nombre_variante || '',
    toNumber_(candidata.piezas_por_capa, 0),
    toNumber_(candidata.niveles, 0)
  ].join('|');

  if (keySet[key]) return;
  keySet[key] = true;
  candidatas.push(candidata);
}

function construirRangoCercano_(base, radio, min, max) {
  const centro = Math.max(min, Math.min(max, Math.round(base)));
  const out = [];
  for (let d = 0; d <= radio; d++) {
    const menos = centro - d;
    const mas = centro + d;
    if (menos >= min) out.push(menos);
    if (mas <= max && mas !== menos) out.push(mas);
  }
  return Array.from(new Set(out));
}

function construirLayoutTresbolilloConColumnas_(filas, columnas, offsetInicial) {
  const rows = [];
  const offsets = [];
  let maxXOcupadoFactoresD = 0;
  for (let i = 0; i < filas; i++) {
    const offset = (offsetInicial + (i % 2 === 0 ? 0 : 0.5)) % 1;
    rows.push(columnas);
    offsets.push(offset);
    const xOcupado = offset + columnas;
    if (xOcupado > maxXOcupadoFactoresD) maxXOcupadoFactoresD = xOcupado;
  }
  return {
    layout_rows: rows,
    layout_offsets: offsets,
    max_x_ocupada_factores_d: maxXOcupadoFactoresD
  };
}

function redondearDimensionIndustrial_(dimensionMm) {
  const dim = Math.max(0, toNumber_(dimensionMm, 0));
  if (dim <= 300) return Math.ceil(dim / 5) * 5;
  if (dim <= 800) return Math.ceil(dim / 10) * 10;
  return Math.ceil(dim / 20) * 20;
}

function calcularScoreCajaIdeal_(cantidadTotal, densidad, volumenCaja, nombreVariante) {
  const bonoSimplicidad = nombreVariante === 'grid_recto' ? 0.003 : 0;
  return (cantidadTotal * 1000000) + (densidad * 1000000000) - volumenCaja + bonoSimplicidad;
}

function compararCajaIdealVsCatalogo_(mejorCatalogo, cajaIdeal) {
  const cantidadCatalogo = toNumber_(mejorCatalogo?.cantidad_max, 0);
  const cantidadIdeal = toNumber_(cajaIdeal?.cantidad_max, 0);
  const aprovechamientoCatalogo = toNumber_(mejorCatalogo?.aprovechamiento, 0);
  const aprovechamientoIdeal = toNumber_(cajaIdeal?.aprovechamiento, 0);
  const volumenCatalogo = toNumber_(mejorCatalogo?.ancho_caja_mm, 0) *
    toNumber_(mejorCatalogo?.largo_caja_mm, 0) *
    toNumber_(mejorCatalogo?.alto_caja_mm, 0);
  const volumenIdeal = toNumber_(cajaIdeal?.ancho_caja_mm, 0) *
    toNumber_(cajaIdeal?.largo_caja_mm, 0) *
    toNumber_(cajaIdeal?.alto_caja_mm, 0);

  const deltaCantidad = cantidadIdeal - cantidadCatalogo;
  const deltaAprovechamiento = aprovechamientoIdeal - aprovechamientoCatalogo;
  const deltaVolumen = volumenIdeal - volumenCatalogo;
  const absDeltaAprovechamiento = Math.abs(deltaAprovechamiento);
  const cercaniaOptimo = (Math.abs(deltaCantidad) <= 1 && absDeltaAprovechamiento <= 0.01);
  const idealClaramenteMejor = deltaCantidad >= 2 || deltaAprovechamiento >= 0.03 || deltaVolumen <= (-0.1 * volumenCatalogo);

  return {
    caja_catalogo: mejorCatalogo?.tipo_caja || '',
    caja_ideal: cajaIdeal?.tipo_caja || 'CAJA_IDEAL_CALCULADA',
    cantidad_catalogo: cantidadCatalogo,
    cantidad_ideal: cantidadIdeal,
    delta_cantidad: deltaCantidad,
    aprovechamiento_catalogo: aprovechamientoCatalogo,
    aprovechamiento_ideal: aprovechamientoIdeal,
    delta_aprovechamiento: deltaAprovechamiento,
    volumen_catalogo_mm3: volumenCatalogo,
    volumen_ideal_mm3: volumenIdeal,
    delta_volumen_mm3: deltaVolumen,
    catalogo_cerca_optimo: cercaniaOptimo,
    ideal_claramente_mejor: idealClaramenteMejor,
    conclusion: idealClaramenteMejor
      ? 'La caja ideal calculada supera claramente a la mejor del catálogo.'
      : (cercaniaOptimo
        ? 'La mejor caja del catálogo está cerca del óptimo.'
        : 'La caja del catálogo es funcional, pero hay margen de ajuste hacia el óptimo.')
  };
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

function getDataCtxCached_() {
  return {
    rollos: getRollosCached_(),
    cajas: getCajasCached_(),
    parametros: getParametrosCached_()
  };
}

function getRollosCached_() {
  return getCachedJson_(CACHE_KEYS_.rollos, () => getSheetObjects_('Rollos'));
}

function getCajasCached_() {
  return getCachedJson_(CACHE_KEYS_.cajas, () => getSheetObjects_('Cajas'));
}

function getParametrosCached_() {
  return getCachedJson_(CACHE_KEYS_.parametros, () => getParametros_('Parametros'));
}

function getSkusWebCached_() {
  return getCachedJson_(CACHE_KEYS_.skus_web, () => {
    const rollos = getRollosCached_();
    return rollos.map(r => ({
      codigo: r.codigo,
      descripcion: r.descripcion || ''
    }));
  });
}

function getCachedJson_(cacheKey, fallbackFn) {
  const cache = CacheService.getScriptCache();
  const cachedRaw = cache.get(cacheKey);
  if (cachedRaw) {
    try {
      return JSON.parse(cachedRaw);
    } catch (error) {
      Logger.log('Cache JSON inválido para key=' + cacheKey + '. Se recarga desde hoja.');
    }
  }

  const freshData = fallbackFn();
  cache.put(cacheKey, JSON.stringify(freshData), CACHE_TTL_SEGUNDOS_);
  return freshData;
}

function clearDataCaches_() {
  CacheService.getScriptCache().removeAll([
    CACHE_KEYS_.rollos,
    CACHE_KEYS_.cajas,
    CACHE_KEYS_.parametros,
    CACHE_KEYS_.skus_web
  ]);
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

function timeStart_(label) {
  if (!PERF_LOGS_ACTIVOS_) return;
  try {
    console.time(label);
  } catch (error) {
    Logger.log('START ' + label + ' @ ' + new Date().toISOString());
  }
}

function timeEnd_(label) {
  if (!PERF_LOGS_ACTIVOS_) return;
  try {
    console.timeEnd(label);
  } catch (error) {
    Logger.log('END ' + label + ' @ ' + new Date().toISOString());
  }
}

function calcularSobranteTotal_(resultado) {
  if (!resultado) return 0;
  return toNumber_(resultado.sobrante_largo_mm, 0) +
    toNumber_(resultado.sobrante_ancho_mm, 0) +
    toNumber_(resultado.sobrante_alto_mm, 0);
}

function detectarDesperdicioAlto_(resultado) {
  if (!resultado) return false;
  const umbrales = UMBRALES_SUGERENCIA_CAJA_;
  const aprovechamiento = toNumber_(resultado.aprovechamiento, 0);
  const sobranteLargo = toNumber_(resultado.sobrante_largo_mm, 0);
  const sobranteAncho = toNumber_(resultado.sobrante_ancho_mm, 0);
  const sobranteAlto = toNumber_(resultado.sobrante_alto_mm, 0);
  const diametroEfectivo = toNumber_(resultado.diametro_efectivo_mm, 0);
  const anchoEfectivo = toNumber_(resultado.ancho_efectivo_mm, 0);

  return (
    aprovechamiento < umbrales.aprovechamiento_bajo ||
    sobranteLargo > (diametroEfectivo * umbrales.sobrante_factor_alto) ||
    sobranteAncho > (diametroEfectivo * umbrales.sobrante_factor_alto) ||
    sobranteAlto > (anchoEfectivo * umbrales.sobrante_factor_alto)
  );
}

function esMejoraSuficiente_(actual, candidata) {
  const umbrales = UMBRALES_SUGERENCIA_CAJA_;
  const cantidadActual = toNumber_(actual?.cantidad_max, 0);
  const cantidadCandidata = toNumber_(candidata?.cantidad_max, 0);
  const aprovechamientoActual = toNumber_(actual?.aprovechamiento, 0);
  const aprovechamientoCandidata = toNumber_(candidata?.aprovechamiento, 0);
  const sobranteActual = calcularSobranteTotal_(actual);
  const sobranteCandidata = calcularSobranteTotal_(candidata);

  const deltaCantidad = cantidadCandidata - cantidadActual;
  const deltaAprovechamiento = aprovechamientoCandidata - aprovechamientoActual;
  const deltaSobranteTotal = sobranteActual - sobranteCandidata;

  const mejoraCantidad = deltaCantidad >= umbrales.mejora_min_cantidad;
  const mejoraAprovechamiento = deltaAprovechamiento >= umbrales.mejora_min_aprovechamiento;
  const umbralSobrante = Math.max(
    umbrales.mejora_min_sobrante_abs_mm,
    sobranteActual * umbrales.mejora_min_sobrante_ratio
  );
  const mejoraSobrante = deltaCantidad >= 0 && deltaSobranteTotal >= umbralSobrante;

  let motivo = '';
  if (mejoraCantidad) {
    motivo = 'Mejora de cantidad máxima';
  } else if (mejoraAprovechamiento) {
    motivo = 'Mejora de aprovechamiento';
  } else if (mejoraSobrante) {
    motivo = 'Reducción clara de sobrante total sin empeorar cantidad';
  }

  return {
    mejora_suficiente: mejoraCantidad || mejoraAprovechamiento || mejoraSobrante,
    motivo: motivo,
    delta_cantidad: deltaCantidad,
    delta_aprovechamiento: deltaAprovechamiento,
    delta_sobrante_total: deltaSobranteTotal
  };
}

function construirSugerenciaCaja_(resultados, mejorActual) {
  const actual = mejorActual || null;
  const desperdicioAlto = detectarDesperdicioAlto_(actual);
  const base = {
    hay_sugerencia: false,
    motivo: desperdicioAlto
      ? 'No se encontró una alternativa mejor que justifique el cambio.'
      : 'La caja seleccionada ya está en rango saludable de aprovechamiento.',
    desperdicio_alto: desperdicioAlto,
    caja_actual: actual?.tipo_caja || '',
    caja_sugerida: '',
    variante_sugerida: '',
    cantidad_actual: toNumber_(actual?.cantidad_max, 0),
    cantidad_sugerida: toNumber_(actual?.cantidad_max, 0),
    aprovechamiento_actual: toNumber_(actual?.aprovechamiento, 0),
    aprovechamiento_sugerido: toNumber_(actual?.aprovechamiento, 0),
    delta_cantidad: 0,
    delta_aprovechamiento: 0,
    delta_sobrante_total: 0
  };

  if (!actual || !Array.isArray(resultados) || !resultados.length) {
    return base;
  }

  if (!desperdicioAlto) {
    return base;
  }

  const viables = resultados.filter(r =>
    r &&
    r.estatus === 'OK' &&
    r.repetible === true &&
    r.viable_altura === true &&
    r.tipo_caja !== actual.tipo_caja
  );

  viables.sort((a, b) => {
    if (b.cantidad_max !== a.cantidad_max) return b.cantidad_max - a.cantidad_max;
    if (b.aprovechamiento !== a.aprovechamiento) return b.aprovechamiento - a.aprovechamiento;
    return calcularSobranteTotal_(a) - calcularSobranteTotal_(b);
  });

  for (const candidata of viables) {
    const evaluacion = esMejoraSuficiente_(actual, candidata);
    if (!evaluacion.mejora_suficiente) continue;

    return {
      hay_sugerencia: true,
      motivo: evaluacion.motivo,
      desperdicio_alto: desperdicioAlto,
      caja_actual: actual.tipo_caja || '',
      caja_sugerida: candidata.tipo_caja || '',
      variante_sugerida: candidata.variante_seleccionada || '',
      cantidad_actual: toNumber_(actual.cantidad_max, 0),
      cantidad_sugerida: toNumber_(candidata.cantidad_max, 0),
      aprovechamiento_actual: toNumber_(actual.aprovechamiento, 0),
      aprovechamiento_sugerido: toNumber_(candidata.aprovechamiento, 0),
      delta_cantidad: evaluacion.delta_cantidad,
      delta_aprovechamiento: evaluacion.delta_aprovechamiento,
      delta_sobrante_total: evaluacion.delta_sobrante_total
    };
  }

  return base;
}
