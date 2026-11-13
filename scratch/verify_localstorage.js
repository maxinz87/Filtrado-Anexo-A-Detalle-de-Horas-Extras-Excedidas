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

const storageMap = new Map();
const mockLocalStorage = {
  getItem: (k) => storageMap.has(k) ? storageMap.get(k) : null,
  setItem: (k, v) => storageMap.set(k, String(v)),
  removeItem: (k) => storageMap.delete(k),
  clear: () => storageMap.clear()
};

function runTestSession(initialValues, simulateClickExport) {
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
    window: {
      XLSX,
      jspdf: jspdfMod,
      localStorage: mockLocalStorage
    },
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

  if (initialValues) {
    if (initialValues.sector !== undefined) {
      elementsById.inputSector.value = initialValues.sector;
      elementsById.inputSector.dispatchEvent({ type: 'input' });
    }
    if (initialValues.jefe !== undefined) {
      elementsById.inputJefe.value = initialValues.jefe;
      elementsById.inputJefe.dispatchEvent({ type: 'input' });
    }
    if (initialValues.obs !== undefined) {
      elementsById.inputObservacionesGlobal.value = initialValues.obs;
      elementsById.inputObservacionesGlobal.dispatchEvent({ type: 'input' });
    }
  }

  if (simulateClickExport) {
    elementsById.btnExportPdf.dispatchEvent({ type: 'click' });
  }

  return { elementsById };
}

console.log('--- TEST 1: Session 1 - User enters values and clicks btnExportPdf ---');
runTestSession({
  sector: 'Sector Transmisión y Redes',
  jefe: 'Ing. Carlos Cantard - Leg. 59457',
  obs: 'Trabajos de guardia y reposición extraordinaria de servicio'
}, true);

console.log('LocalStorage content:');
for (const [k, v] of storageMap.entries()) {
  console.log('  ' + k + ' => ' + v);
}

assert.strictEqual(storageMap.get('epe_anexo_a_inputSector'), 'Sector Transmisión y Redes');
assert.strictEqual(storageMap.get('epe_anexo_a_inputJefe'), 'Ing. Carlos Cantard - Leg. 59457');
assert.strictEqual(storageMap.get('epe_anexo_a_inputObservacionesGlobal'), 'Trabajos de guardia y reposición extraordinaria de servicio');

console.log('--- TEST 2: Session 2 - Reopening page restores saved values ---');
const session2 = runTestSession(null, false);
console.log('Restored inputSector:', session2.elementsById.inputSector.value);
console.log('Restored inputJefe:', session2.elementsById.inputJefe.value);
console.log('Restored inputObservacionesGlobal:', session2.elementsById.inputObservacionesGlobal.value);

assert.strictEqual(session2.elementsById.inputSector.value, 'Sector Transmisión y Redes');
assert.strictEqual(session2.elementsById.inputJefe.value, 'Ing. Carlos Cantard - Leg. 59457');
assert.strictEqual(session2.elementsById.inputObservacionesGlobal.value, 'Trabajos de guardia y reposición extraordinaria de servicio');

console.log('ALL LOCALSTORAGE TESTS PASSED 100% SUCCESSFULLY!');

