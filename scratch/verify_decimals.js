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
  createDocumentFragment: () => ({ children: [], appendChild: function(c) { this.children.push(c); return c; } }),
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
const OriginalJsPDF = jspdfMod.jsPDF;
class TrackedJsPDF extends OriginalJsPDF {
  constructor(opts) {
    super(opts);
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
if (domReadyCb) domReadyCb();

const sampleSheetData = [
  ['N° personal', 'Apellido Nombre', 'Total Horas Programadas', 'Total Horas Emergencia', 'TOTAL HS PROGRAMADAS + EMERGENCIA'],
  ['83478', 'PONCE WALTER JAVIER', 77.22, 26.65, 103.87], // Art8: 0, Emerg: 26,65, Exc120: 0
  ['84442', 'MONJE MORENA', 77.1, 21.35, 98.45],          // Art8: 0, Emerg: 21,35, Exc120: 0
  ['11890', 'RODRIGUEZ CARLOS', 95, 30, 125],             // Art8: 15,00, Emerg: 30,00, Exc120: 5,00
  ['81153', 'LOVATO JULIAN OSMAR', 65, 7.47, 72.47]       // Art8: 0, Emerg: 7,47, Exc120: 0
];

const ws = XLSX.utils.aoa_to_sheet(sampleSheetData);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Hoja1');
const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });

elementsById.inputFile.files = [{ name: 'test_decimal_format.xlsx', _buffer: buf }];
elementsById.inputFile.dispatchEvent({ type: 'change' });

setTimeout(() => {
  elementsById.btnExportPdf.dispatchEvent({ type: 'click' });
  assert(capturedAutoTableConfig, 'autoTable called');

  console.log('--- Checking PDF Decimal Formatting ---');
  capturedAutoTableConfig.body.forEach((row, i) => {
    console.log('Row ' + i + ' (' + row[1] + '): Art8 = ' + row[2] + ' | Emerg = ' + row[3] + ' | Exc120 = ' + row[4]);
  });

  const ponce = capturedAutoTableConfig.body[0];
  assert.strictEqual(ponce[2], '0', 'Ponce Art8 should be "0"');
  assert.strictEqual(ponce[3], '26,65', 'Ponce Emerg should be "26,65"');
  assert.strictEqual(ponce[4], '0', 'Ponce Exc120 should be "0"');

  const monje = capturedAutoTableConfig.body[1];
  assert.strictEqual(monje[2], '0', 'Monje Art8 should be "0"');
  assert.strictEqual(monje[3], '21,35', 'Monje Emerg should be "21,35"');
  assert.strictEqual(monje[4], '0', 'Monje Exc120 should be "0"');

  const rodriguez = capturedAutoTableConfig.body[2];
  assert.strictEqual(rodriguez[2], '15,00', 'Rodriguez Art8 should be "15,00"');
  assert.strictEqual(rodriguez[3], '30,00', 'Rodriguez Emerg should be "30,00"');
  assert.strictEqual(rodriguez[4], '5,00', 'Rodriguez Exc120 should be "5,00"');

  const lovato = capturedAutoTableConfig.body[3];
  assert.strictEqual(lovato[2], '0', 'Lovato Art8 should be "0"');
  assert.strictEqual(lovato[3], '7,47', 'Lovato Emerg should be "7,47"');
  assert.strictEqual(lovato[4], '0', 'Lovato Exc120 should be "0"');

  console.log('SUCCESS: All decimal formats strictly follow (,) and 2 decimal places, and 0 without decimals!');
}, 50);

