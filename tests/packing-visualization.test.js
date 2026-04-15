const fs = require('fs');
const vm = require('vm');
const path = require('path');

const source = fs.readFileSync(path.join(__dirname, '..', 'PackingVisualization.gs'), 'utf8');
const context = { console, Math, JSON, Number, Array, String };
vm.createContext(context);
vm.runInContext(source, context);

const createViz = context.crearPackingVisualization_;

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function testSingleRoll() {
  const viz = createViz({
    largo_caja_mm: 100,
    ancho_caja_mm: 100,
    alto_caja_mm: 100,
    diametro_efectivo_mm: 40,
    ancho_efectivo_mm: 30,
    piezas_largo: 1,
    piezas_ancho: 1,
    niveles: 1,
    variante_seleccionada: 'grid_recto'
  });
  assert(viz.rolls.length === 1, 'Debe crear 1 rollo');
  assert(viz.metrics.totalLayers === 1, 'Debe tener 1 capa');
}

function testMultipleLayers() {
  const viz = createViz({
    largo_caja_mm: 300,
    ancho_caja_mm: 200,
    alto_caja_mm: 300,
    diametro_efectivo_mm: 50,
    ancho_efectivo_mm: 70,
    piezas_largo: 2,
    piezas_ancho: 2,
    niveles: 3,
    variante_seleccionada: 'grid_recto'
  });
  assert(viz.rolls.length === 12, '2x2x3 debe producir 12 rollos');
  assert(viz.metrics.totalLayers === 3, 'Debe tener 3 capas');
}

function testEmptyCase() {
  const viz = createViz(null);
  assert(viz.rolls.length === 0, 'Caso vacío sin rollos');
  assert(viz.metrics.occupancyRatio === 0, 'Ocupación debe ser 0');
}

function testBoundaryCase() {
  const viz = createViz({
    largo_caja_mm: 100,
    ancho_caja_mm: 100,
    alto_caja_mm: 60,
    diametro_efectivo_mm: 100,
    ancho_efectivo_mm: 60,
    piezas_largo: 1,
    piezas_ancho: 1,
    niveles: 1,
    variante_seleccionada: 'grid_recto'
  });
  assert(viz.rolls.length === 1, 'Caso límite debe mantener 1 rollo');
  assert(viz.metrics.usedVolume > 0, 'Debe calcular volumen usado');
}

function run() {
  testSingleRoll();
  testMultipleLayers();
  testEmptyCase();
  testBoundaryCase();
  console.log('OK: packing-visualization tests passed');
}

run();
