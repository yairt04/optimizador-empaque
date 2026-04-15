/**
 * @typedef {Object} PackingVisualization
 * @property {{length:number,width:number,height:number}} box
 * @property {Array<{id:string,x:number,y:number,z:number,diameter:number,length:number,layer:number,orientation:("x"|"y"|"z")}>} rolls
 * @property {{totalRolls:number,usedVolume:number,freeVolume:number,occupancyRatio:number,totalLayers:number,pattern:string}} metrics
 */

/**
 * Adaptador puro: convierte un match del optimizador al contrato visual.
 * No modifica ni recalcula el algoritmo de selección de caja.
 * @param {Object} mejor
 * @returns {PackingVisualization}
 */
function crearPackingVisualization_(mejor) {
  if (!mejor) return crearVisualizacionVacia_();

  const box = {
    length: toNumberPv_(mejor.largo_caja_mm, 0),
    width: toNumberPv_(mejor.ancho_caja_mm, 0),
    height: toNumberPv_(mejor.alto_caja_mm, 0)
  };

  const diameter = Math.max(0, toNumberPv_(mejor.diametro_efectivo_mm, 0));
  const rollLength = Math.max(0, toNumberPv_(mejor.ancho_efectivo_mm, 0));
  const rows = normalizarFilasLayout_(mejor.layout_rows, Math.max(0, toNumberPv_(mejor.piezas_ancho, 0)), Math.max(0, toNumberPv_(mejor.piezas_largo, 0)));
  const offsets = normalizarOffsetsLayout_(mejor.layout_offsets, rows.length);
  const niveles = Math.max(0, toNumberPv_(mejor.niveles, 0));
  const pattern = String(mejor.variante_seleccionada || 'grid_recto');
  const useHoneycomb = pattern === 'grid_alternado_tresbolillo' || pattern === 'arranque_corrido_equivalente';
  const yStep = useHoneycomb ? diameter * 0.8660254038 : diameter;

  const layoutWidth = rows.reduce(function(maxWidth, cols, idx) {
    return Math.max(maxWidth, ((offsets[idx] || 0) + cols) * diameter);
  }, 0);
  const layoutDepth = rows.length ? diameter + (Math.max(0, rows.length - 1) * yStep) : 0;

  const originX = Math.max(0, (box.length - layoutWidth) / 2);
  const originY = Math.max(0, (box.width - layoutDepth) / 2);

  const rolls = [];
  for (let layerIndex = 0; layerIndex < niveles; layerIndex++) {
    const z = layerIndex * rollLength;
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
      const cols = rows[rowIndex];
      const rowOffset = (offsets[rowIndex] || 0) * diameter;
      for (let colIndex = 0; colIndex < cols; colIndex++) {
        rolls.push({
          id: `L${layerIndex + 1}-R${rowIndex + 1}-C${colIndex + 1}`,
          x: originX + rowOffset + (diameter / 2) + (colIndex * diameter),
          y: originY + (diameter / 2) + (rowIndex * yStep),
          z: z,
          diameter: diameter,
          length: rollLength,
          layer: layerIndex + 1,
          orientation: 'z'
        });
      }
    }
  }

  const boxVolume = box.length * box.width * box.height;
  const rollVolume = Math.PI * Math.pow(diameter / 2, 2) * rollLength;
  const usedVolume = rolls.length * rollVolume;
  const occupancyRatio = boxVolume > 0 ? Math.min(1, usedVolume / boxVolume) : 0;
  const freeVolume = Math.max(0, boxVolume - usedVolume);

  return {
    box: box,
    rolls: rolls,
    metrics: {
      totalRolls: rolls.length,
      usedVolume: usedVolume,
      freeVolume: freeVolume,
      occupancyRatio: occupancyRatio,
      totalLayers: niveles,
      pattern: pattern
    }
  };
}

function crearVisualizacionVacia_() {
  return {
    box: { length: 0, width: 0, height: 0 },
    rolls: [],
    metrics: {
      totalRolls: 0,
      usedVolume: 0,
      freeVolume: 0,
      occupancyRatio: 0,
      totalLayers: 0,
      pattern: 'sin_datos'
    }
  };
}

function normalizarFilasLayout_(rawRows, fallbackRowsCount, fallbackColsCount) {
  let rows = [];
  if (Array.isArray(rawRows)) {
    rows = rawRows;
  } else if (typeof rawRows === 'string' && rawRows.trim()) {
    try {
      const parsed = JSON.parse(rawRows);
      if (Array.isArray(parsed)) rows = parsed;
    } catch (error) {
      rows = [];
    }
  }

  if (!rows.length && fallbackRowsCount > 0 && fallbackColsCount > 0) {
    rows = Array(fallbackRowsCount).fill(fallbackColsCount);
  }

  return rows
    .map(function(v) { return Math.max(0, toNumberPv_(v, 0)); })
    .filter(function(v) { return v > 0; });
}

function normalizarOffsetsLayout_(rawOffsets, rowsLength) {
  let offsets = [];
  if (Array.isArray(rawOffsets)) {
    offsets = rawOffsets;
  } else if (typeof rawOffsets === 'string' && rawOffsets.trim()) {
    try {
      const parsed = JSON.parse(rawOffsets);
      if (Array.isArray(parsed)) offsets = parsed;
    } catch (error) {
      offsets = [];
    }
  }

  return Array(rowsLength).fill(0).map(function(_, idx) {
    return Math.max(0, toNumberPv_(offsets[idx], 0));
  });
}

function toNumberPv_(value, fallback) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}
