const XLSX = require('xlsx');
const path = require('path');

const FILE = path.join(__dirname, 'parser_insert_turn.xlsx');
const OUT_FILE = process.argv[2]
  ? path.resolve(process.argv[2])
  : FILE;
const SHEET = 'Parser';
const MAP_SHEET = 'Map';
const LAST_ROW = 1001;

const SHAPE_MAP = {
  V: 'Ромб 35° (V***)',
  F: 'Ромб 50° (F***)',
  D: 'Ромб 55° (D***)',
  E: 'Ромб 75° (E***)',
  C: 'Ромб 80° (С***)',
  M: 'Ромб 86° (M***)',
  W: 'Ломаный треугольник 80° (W***)',
  T: 'Треугольник 60° (T***)',
  R: 'Круг (R***)',
  S: 'Квадрат 90° (S***)',
  H: 'Шестиугольник 120° (H***)',
  K: 'Параллелограмм 55° (K***)',
  B: 'Параллелограмм 82° (B***)',
  A: 'Параллелограмм 85° (A***)',
  Q: 'Восьмиугольник 135° (O***)',
  L: 'Прямоугольник 90° (L***)',
  P: 'Пятиугольник 108° (P***)',
};

const M2_MAP = {
  A: { m1: 'P', m2: 'M', m3: 'K' },
  P: { m1: 'P' },
  M: { m1: 'M' },
  S: { m1: 'S' },
  K: { m1: 'K' },
  H: { m1: 'H' },
  N: { m1: 'N' },
};

const MATERIAL_MAP = {
  1: 'CVD',
  2: 'CVD',
  6: 'CVD',
  7: 'CVD',
  3: 'PVD',
  4: 'PVD',
};

const KONSTRUKCIIA_MAP = {
  M: 'с отверстием/односторонний (***M)',
  G: 'с отверстием/двухсторонний (***G)',
  A: 'с отверстием/без стружколома (***A)',
  T: 'отверстие с зенковкой 40°-60°/односторонний (***T)',
  N: 'без отверстия/без стружколома (***N)',
  W: 'отверстие с зенковкой 40°-60°/без стружколома (***W)',
  H: 'отверстие с зенковкой 70°-90°/односторонний (***H)',
  X: 'иная (***X)',
};

const MAP = {
  shape: { startRow: 2, startCol: 3, data: SHAPE_MAP },
  material: { startRow: 2, startCol: 6, data: MATERIAL_MAP },
  m: { startRow: 2, startCol: 9, headers: ['Code', 'm1', 'm2', 'm3'], rows: M2_MAP },
  konstruktciia: { startRow: 2, startCol: 14, data: KONSTRUKCIIA_MAP },
};

function alloyActive(row) {
  return `OR(LEFT(H${row},1)="G",LEFT(H${row},1)="R")`;
}

function buildShapeFormula(row) {
  return `IF(A${row}="","",IFERROR(VLOOKUP(LEFT(A${row},1),Map!$D$2:$E$20,2,FALSE),""))`;
}

function buildM1Formula(row) {
  return `IF(OR(H${row}="",NOT(${alloyActive(row)})),"",IFERROR(VLOOKUP(MID(H${row},2,1),Map!$J$2:$M$10,2,FALSE),""))`;
}

function buildM2Formula(row) {
  return `IF(OR(H${row}="",NOT(${alloyActive(row)}),MID(H${row},2,1)<>"A"),"","M")`;
}

function buildM3Formula(row) {
  return `IF(OR(H${row}="",NOT(${alloyActive(row)}),MID(H${row},2,1)<>"A"),"","K")`;
}

function buildEmptyAlloyFormula(row) {
  return `IF(OR(H${row}="",NOT(${alloyActive(row)})),"","")`;
}

function buildMaterialFormula(row) {
  return `IF(OR(H${row}="",NOT(${alloyActive(row)})),"",IFERROR(VLOOKUP(MID(H${row},3,1),Map!$G$2:$H$10,2,FALSE),""))`;
}

function buildKonstruktciiaFormula(row) {
  return `IF(A${row}="","",IFERROR(VLOOKUP(MID(A${row},4,1),Map!$O$2:$P$10,2,FALSE),""))`;
}

function parseShape(art) {
  if (!art) return '';
  return SHAPE_MAP[art[0]] || '';
}

function parseKonstruktciia(art) {
  if (!art || art.length < 4) return '';
  return KONSTRUKCIIA_MAP[art[3]] || '';
}

function parseAlloy(alloy) {
  const result = {
    m1: '',
    m2: '',
    m3: '',
    m4: '',
    m5: '',
    m6: '',
    Material_plastiny: '',
  };
  if (!alloy) return result;

  const first = alloy[0];
  if (first !== 'G' && first !== 'R') return result;

  const mapping = M2_MAP[alloy[1]];
  if (mapping) {
    result.m1 = mapping.m1 || '';
    result.m2 = mapping.m2 || '';
    result.m3 = mapping.m3 || '';
  }

  const third = alloy[2];
  result.Material_plastiny = MATERIAL_MAP[third] || '';

  return result;
}

function setFormulaCell(ws, addr, formula, value) {
  const open = (formula.match(/\(/g) || []).length;
  const close = (formula.match(/\)/g) || []).length;
  if (open !== close) {
    throw new Error(`Unbalanced formula at ${addr}: (${open} vs ${close}) ${formula}`);
  }
  ws[addr] = { t: 's', f: formula, v: value };
}

function setCell(ws, r, c, value) {
  const addr = XLSX.utils.encode_cell({ r, c });
  if (value === '' || value === undefined || value === null) {
    delete ws[addr];
    return;
  }
  ws[addr] = { t: 's', v: String(value) };
}

function writeMapSheet(ws) {
  setCell(ws, 0, 3, 'Code');
  setCell(ws, 0, 4, 'shape');
  Object.entries(SHAPE_MAP).forEach(([code, label], index) => {
    const row = MAP.shape.startRow - 1 + index;
    setCell(ws, row, MAP.shape.startCol, code);
    setCell(ws, row, MAP.shape.startCol + 1, label);
  });

  setCell(ws, 0, 6, 'Code');
  setCell(ws, 0, 7, 'Material_plastiny');
  Object.entries(MATERIAL_MAP).forEach(([code, label], index) => {
    const row = MAP.material.startRow - 1 + index;
    setCell(ws, row, MAP.material.startCol, code);
    setCell(ws, row, MAP.material.startCol + 1, label);
  });

  MAP.m.headers.forEach((header, index) => {
    setCell(ws, 0, MAP.m.startCol + index, header);
  });
  Object.entries(M2_MAP).forEach(([code, values], index) => {
    const row = MAP.m.startRow - 1 + index;
    setCell(ws, row, MAP.m.startCol, code);
    setCell(ws, row, MAP.m.startCol + 1, values.m1 || '');
    setCell(ws, row, MAP.m.startCol + 2, values.m2 || '');
    setCell(ws, row, MAP.m.startCol + 3, values.m3 || '');
  });

  setCell(ws, 0, MAP.konstruktciia.startCol, 'Code');
  setCell(ws, 0, MAP.konstruktciia.startCol + 1, 'Konstruktciia_plastiny');
  Object.entries(KONSTRUKCIIA_MAP).forEach(([code, label], index) => {
    const row = MAP.konstruktciia.startRow - 1 + index;
    setCell(ws, row, MAP.konstruktciia.startCol, code);
    setCell(ws, row, MAP.konstruktciia.startCol + 1, label);
  });

  ws['!ref'] = 'A1:P20';
}

const wb = XLSX.readFile(FILE, { cellFormula: true, cellStyles: true });
const ws = wb.Sheets[SHEET];
const mapWs = wb.Sheets[MAP_SHEET];

writeMapSheet(mapWs);

const headers = [
  { col: 8, name: 'shape' },
  { col: 9, name: 'm1' },
  { col: 10, name: 'm2' },
  { col: 11, name: 'm3' },
  { col: 12, name: 'm4' },
  { col: 13, name: 'm5' },
  { col: 14, name: 'm6' },
  { col: 15, name: 'Material_plastiny' },
  { col: 16, name: 'Konstruktciia_plastiny' },
];

headers.forEach(({ col, name }) => {
  ws[XLSX.utils.encode_cell({ r: 0, c: col })] = { t: 's', v: name };
});

for (let row = 2; row <= LAST_ROW; row += 1) {
  const art = ws[`A${row}`]?.v || '';
  const alloy = ws[`H${row}`]?.v || '';
  const parsed = parseAlloy(String(alloy));
  const values = [
    parseShape(String(art)),
    parsed.m1,
    parsed.m2,
    parsed.m3,
    parsed.m4,
    parsed.m5,
    parsed.m6,
    parsed.Material_plastiny,
    parseKonstruktciia(String(art)),
  ];
  const formulas = [
    buildShapeFormula(row),
    buildM1Formula(row),
    buildM2Formula(row),
    buildM3Formula(row),
    buildEmptyAlloyFormula(row),
    buildEmptyAlloyFormula(row),
    buildEmptyAlloyFormula(row),
    buildMaterialFormula(row),
    buildKonstruktciiaFormula(row),
  ];

  formulas.forEach((formula, index) => {
    const addr = XLSX.utils.encode_cell({ r: row - 1, c: 8 + index });
    setFormulaCell(ws, addr, formula, values[index]);
  });
}

ws['!ref'] = `A1:Q${LAST_ROW}`;

XLSX.writeFile(wb, OUT_FILE);

console.log(`Updated ${path.basename(OUT_FILE)}`);
console.log('Added Konstruktciia_plastiny, fixed formulas via Map lookups');

for (const row of [2, 3, 4]) {
  const art = ws[`A${row}`]?.v;
  const alloy = ws[`H${row}`]?.v;
  console.log(
    row,
    art,
    '=> shape:',
    parseShape(String(art)),
    'konstr:',
    parseKonstruktciia(String(art)),
    'alloy:',
    parseAlloy(String(alloy)),
  );
}
