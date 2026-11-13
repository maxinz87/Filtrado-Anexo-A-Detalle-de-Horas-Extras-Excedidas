const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const workspace = 'c:\\Users\\maxin\\OneDrive\\Documentos\\Proyectos\\Justificacion hs extras';
const XLSX = require(path.join(workspace, 'libs/xlsx.full.min.js'));
const jspdfMod = require(path.join(workspace, 'libs/jspdf.umd.min.js'));
const autoTableMod = require(path.join(workspace, 'libs/jspdf.plugin.autotable.min.js'));
if (typeof autoTableMod.applyPlugin === 'function') {
  autoTableMod.applyPlugin(jspdfMod.jsPDF);
}

class MockElement {
  constructor(tagName, id = '') {
    this.tagName = tagName.toUpperCase();
    this.id = id;
    this.children = [];
    this.listeners = {};
    this.dataset = {};
    this.attributes = {};
    this.style = {};
    this.value = '';
    this.textContent = '';
    this.className = '';
    this.disabled = false;
    this.checked = false;
    this.indeterminate = false;
    this.selected = false;
    this.files = [];
    this.parentElement = null;
    this.selectionStart = 0;
    this._innerHTML = '';
    this.classList = {
      add: () => {},
      remove: () => {},
      toggle: () => {}
    };
  }
  appendChild(child) {
    child.parentElement = this;
    this.children.push(child);
    return child;
  }
  append(...children) {
    children.forEach((c) => this.appendChild(c));
  }
  addEventListener(event, fn) {
    if (!this.listeners[event]) this.listeners[event] = [];
    this.listeners[event].push(fn);
  }
  dispatchEvent(event) {
    const list = this.listeners[event.type] || [];
    list.forEach((fn) => fn(event));
    if (this.parentElement && event.bubbles) {
      this.parentElement.dispatchEvent(event);
    }
  }
  querySelector() { return null; }
  querySelectorAll() { return []; }
  setAttribute(k, v) { this.attributes[k] = v; }
}

class MockDocumentFragment {
  constructor() { this.children = []; }
  appendChild(c) { this.children.push(c); return c; }
}

const elementsById = {};
const ids = [
  'metadataForm', 'selectAnio', 'selectMes', 'inputExpte', 'expteHelp',
  'inputSector', 'inputJefe', 'inputObservacionesGlobal', 'btnApplyObs',
  'inputFile', 'fileNameDisplay', 'selectProgCriterion', 'criterionBadgeText',
  'chkRespectSheetFilter', 'btnProcessFile', 'btnDownloadSample',
  'sheetFilterInfoBar', 'sheetFilterDetails', 'statusBanner',
  'metricTotalRows', 'metricVisibleRows', 'metricFilteredRows',
  'btnToggleAllArt4', 'btnToggleAllArt7', 'btnExportPdf', 'tableBody'
];
for (const id of ids) {
  elementsById[id] = new MockElement(id.startsWith('select') ? 'select' : id.startsWith('btn') ? 'button' : 'input', id);
}
elementsById.tableBody = new MockElement('tbody', 'tableBody');

const masterCheckArt4 = new MockElement('input');
masterCheckArt4.dataset.column = 'incumpleArt4';
const masterCheckArt7 = new MockElement('input');
masterCheckArt7.dataset.column = 'incumpleArt7';

let domReadyCb = null;
const mockDocument = {
  getElementById: (id) => elementsById[id] || null,
  createElement: (tag) => new MockElement(tag),
  createDocumentFragment: () => new MockDocumentFragment(),
  querySelectorAll: (selector) => {
    if (selector === '.th-master-check') return [masterCheckArt4, masterCheckArt7];
    return [];
  },
  querySelector: (selector) => {
    if (selector === '.th-master-check[data-column="incumpleArt4"]') return masterCheckArt4;
    if (selector === '.th-master-check[data-column="incumpleArt7"]') return masterCheckArt7;
    return null;
  },
  addEventListener: (type, fn) => {
    if (type === 'DOMContentLoaded') domReadyCb = fn;
  }
};

let capturedAutoTableConfig = null;
let lastDocInstance = null;
const OriginalJsPDF = jspdfMod.jsPDF;
class TrackedJsPDF extends OriginalJsPDF {
  constructor(opts) {
    super(opts);
    lastDocInstance = this;
    const origAutoTable = this.autoTable;
    this.autoTable = (config) => {
      capturedAutoTableConfig = config;
      return origAutoTable.call(this, config);
    };
    this.save = (name) => {};
  }
}

class MockFileReader {
  readAsArrayBuffer(file) {
    queueMicrotask(() => {
      this.result = file._buffer;
      if (this.onload) this.onload({ target: this });
    });
  }
}

const sandbox = {
  window: {
    XLSX,
    jspdf: { jsPDF: TrackedJsPDF }
  },
  document: mockDocument,
  HTMLInputElement: MockElement,
  FileReader: MockFileReader,
  queueMicrotask,
  console,
  Date,
  Number,
  String,
  Boolean,
  Array,
  Map,
  Set,
  Math,
  Promise,
  Error,
  Uint8Array,
  Object
};

const ctx = vm.createContext(sandbox);
const appJsCode = fs.readFileSync(path.join(workspace, 'js/app.js'), 'utf-8');
vm.runInContext(appJsCode, ctx);

if (domReadyCb) {
  domReadyCb();
}

// Generate 35 realistic rows
const names = [
  'ACOSTA MARIO ALBERTO', 'BAEZ ROBERTO CARLOS', 'CANETE JUAN MANUEL',
  'DOMINGUEZ NESTOR FABIAN', 'ESPINOZA CLAUDIO', 'FERREYRA GUSTAVO',
  'GARCIA HECTOR OMAR', 'HERRERA LUIS ALBERTO', 'IBANEZ SERGIO DAMIAN',
  'JUAREZ MARCELO ALEJANDRO', 'LOPEZ CRISTIAN EDUARDO', 'MARTINEZ JORGE LUIS',
  'NUNEZ RUBEN DARIO', 'ORTIZ CARLOS SEBASTIAN', 'ROMERO ALCIDES DAVID',
  'LOVATO JULIAN OSMAR', 'FIOCCHI GUILLERMO', 'BENITEZ HUGO ROLANDO',
  'CALDERON JOSE LUIS', 'DIAZ ENRIQUE MANUEL', 'FALCON PABLO CESAR',
  'GOMEZ ROBERTO JAVIER', 'HERNANDEZ DANIEL OMAR', 'IGLESIAS MARIANO',
  'JARA SERGIO ARIEL', 'KRAUS OSVALDO ADOLFO', 'LUNA CHRISTIAN DARIO',
  'MEDINA FEDERICO GABRIEL', 'NAVARRO OSCAR ALFREDO', 'OLIVERA MARCOS ANTONIO',
  'PERALTA RAUL HORACIO', 'QUINTEROS SILVIO DAMIAN', 'RAMOS GUSTAVO FABIAN',
  'SANCHEZ WALTER EDUARDO', 'TOLEDO JUAN BAUTISTA'
];

const sampleSheetData = [
  ['N° personal', 'Apellido Nombre', 'Total Horas Programadas', 'Total Horas Emergencia', 'TOTAL HS PROGRAMADAS + EMERGENCIA']
];

names.forEach((name, i) => {
  sampleSheetData.push([String(60000 + i), name, 75.0, 12.5, 87.5]);
});

const ws = XLSX.utils.aoa_to_sheet(sampleSheetData);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Hoja1');
const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });

elementsById.inputJefe.value = 'Cantard Carlos - Leg. 59457';
elementsById.inputJefe.dispatchEvent({ type: 'input' });

elementsById.inputFile.files = [{ name: 'test_pagebreaks.xlsx', _buffer: buf }];
elementsById.inputFile.dispatchEvent({ type: 'change' });

setTimeout(() => {
  console.log('--- Triggering PDF Export ---');
  elementsById.btnExportPdf.dispatchEvent({ type: 'click' });

  assert(capturedAutoTableConfig, 'autoTable was called');
  console.log('rowPageBreak setting:', capturedAutoTableConfig.rowPageBreak);
  assert.strictEqual(capturedAutoTableConfig.rowPageBreak, 'avoid', 'rowPageBreak must be "avoid"');

  console.log('Total body rows passed to AutoTable:', capturedAutoTableConfig.body.length);
  assert.strictEqual(capturedAutoTableConfig.body.length, 35, 'Exactly 35 rows passed');

  const totalPages = lastDocInstance.internal.getNumberOfPages();
  console.log('Total pages generated:', totalPages);
  assert(totalPages >= 2, 'Should be multi-page document');

  const lastTable = lastDocInstance.lastAutoTable;
  console.log('Table body rows drawn:', lastTable.body.length);

  // Check that no row has spansMultiplePages = true
  const splitRows = lastTable.body.filter(r => r.spansMultiplePages);
  console.log('Any rows split across pages?', splitRows.length > 0);
  assert.strictEqual(splitRows.length, 0, 'No rows should be split across pages');

  // Check row index -1
  const negativeIndexRows = lastTable.body.filter(r => r.index === -1);
  console.log('Any row with index -1 (partial remainder row)?', negativeIndexRows.length > 0);
  assert.strictEqual(negativeIndexRows.length, 0, 'No row with index -1 should exist');

  console.log('ALL PAGE BREAK TESTS PASSED 100% SUCCESSFULLY!');
}, 50);
