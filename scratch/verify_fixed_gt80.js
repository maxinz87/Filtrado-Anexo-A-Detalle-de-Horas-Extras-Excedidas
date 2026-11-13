const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const workspace = 'c:\\Users\\maxin\\OneDrive\\Documentos\\Proyectos\\Justificacion hs extras';
const XLSX = require(path.join(workspace, 'libs/xlsx.full.min.js'));
const jspdfMod = require(path.join(workspace, 'libs/jspdf.umd.min.js'));
const autoTableMod = require(path.join(workspace, 'libs/jspdf.plugin.autotable.min.js'));
if (typeof autoTableMod.applyPlugin === 'function') autoTableMod.applyPlugin(jspdfMod.jsPDF);

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
    this.files = [];
    this.parentElement = null;
    this.classList = { add: () => {}, remove: () => {}, toggle: () => {} };
  }
  appendChild(child) { child.parentElement = this; this.children.push(child); return child; }
  append(...children) { children.forEach((c) => this.appendChild(c)); }
  addEventListener(event, fn) { (this.listeners[event] = this.listeners[event] || []).push(fn); }
  dispatchEvent(event) {
    (this.listeners[event.type] || []).forEach(fn => fn(event));
    if (this.parentElement && event.bubbles) this.parentElement.dispatchEvent(event);
  }
  querySelector() { return null; }
  querySelectorAll() { return []; }
  setAttribute(k, v) { this.attributes[k] = v; }
}

const elementsById = {};
const ids = [
  'metadataForm', 'selectAnio', 'selectMes', 'inputExpte', 'expteHelp',
  'inputSector', 'inputJefe', 'inputObservacionesGlobal', 'btnApplyObs',
  'inputFile', 'fileNameDisplay',
  'chkRespectSheetFilter', 'btnProcessFile', 'btnDownloadSample',
  'sheetFilterInfoBar', 'sheetFilterDetails', 'statusBanner',
  'metricTotalRows', 'metricVisibleRows', 'metricFilteredRows',
  'btnToggleAllArt4', 'btnToggleAllArt7', 'btnExportPdf', 'tableBody'
];
for (const id of ids) elementsById[id] = new MockElement(id.startsWith('select') ? 'select' : id.startsWith('btn') ? 'button' : 'input', id);
elementsById.tableBody = new MockElement('tbody', 'tableBody');

let domReadyCb = null;
const mockDocument = {
  getElementById: (id) => elementsById[id] || null,
  createElement: (tag) => new MockElement(tag),
  createDocumentFragment: () => ({ children: [], appendChild: function(c) { this.children.push(c); return c; } }),
  querySelectorAll: () => [],
  querySelector: () => null,
  addEventListener: (type, fn) => { if (type === 'DOMContentLoaded') domReadyCb = fn; }
};

const sandbox = {
  window: { XLSX, jspdf: jspdfMod },
  document: mockDocument,
  HTMLInputElement: MockElement,
  FileReader: class {
    readAsArrayBuffer(file) {
      queueMicrotask(() => {
        this.result = file._buffer;
        if (this.onload) this.onload({ target: this });
      });
    }
  },
  queueMicrotask, console, Date, Number, String, Boolean, Array, Map, Set, Math, Promise, Error, Uint8Array, Object
};

const ctx = vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(workspace, 'js/app.js'), 'utf-8'), ctx);
if (domReadyCb) domReadyCb();

// Test:
// Case 1: Person with hsProgramadas = 75, Emergencia = 0, Total = 75 -> DOES NOT COMPLY (>80 is false, Emerg>0 is false, Total>120 is false)
// Case 2: Person with hsProgramadas = 85, Emergencia = 0, Total = 85 -> COMPLIES (hsProgramadas > 80)
// Case 3: Person with hsProgramadas = 50, Emergencia = 10, Total = 60 -> COMPLIES (Emergencia > 0)
// Case 4: Person with hsProgramadas = 70, Emergencia = 0, Total = 130 -> COMPLIES (Total > 120)
// Case 5: Person with hsProgramadas = 80, Emergencia = 0, Total = 80 -> DOES NOT COMPLY (80 is not > 80)
const sampleSheetData = [
  ['N° personal', 'Apellido Nombre', 'Total Horas Programadas', 'Total Horas Emergencia', 'TOTAL HS PROGRAMADAS + EMERGENCIA'],
  ['1', 'CASO 1 NO CUMPLE', 75, 0, 75],
  ['2', 'CASO 2 CUMPLE HS PROG MAYOR A 80', 85, 0, 85],
  ['3', 'CASO 3 CUMPLE EMERG MAYOR A 0', 50, 10, 60],
  ['4', 'CASO 4 CUMPLE TOTAL MAYOR A 120', 70, 0, 130],
  ['5', 'CASO 5 NO CUMPLE 80 EXACTO', 80, 0, 80]
];

const ws = XLSX.utils.aoa_to_sheet(sampleSheetData);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Hoja1');
const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });

elementsById.inputFile.files = [{ name: 'test_fixed_gt80.xlsx', _buffer: buf }];
elementsById.inputFile.dispatchEvent({ type: 'change' });

setTimeout(() => {
  const trs = elementsById.tableBody.children;
  console.log('Filtered rows count in table:', trs.length);
  assert.strictEqual(trs.length, 3, 'Must filter exactly 3 rows (cases 2, 3 and 4)');

  const names = trs.map(tr => tr.children[1].textContent);
  console.log('Included agents:', names);
  assert.deepStrictEqual(names, [
    'CASO 2 CUMPLE HS PROG MAYOR A 80',
    'CASO 3 CUMPLE EMERG MAYOR A 0',
    'CASO 4 CUMPLE TOTAL MAYOR A 120'
  ]);

  console.log('ALL VERIFICATIONS PASSED: "> 80" is now the fixed condition!');
}, 50);

