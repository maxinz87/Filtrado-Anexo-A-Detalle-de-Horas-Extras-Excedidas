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
      add: (c) => {
        const set = new Set(this.className.split(/\s+/).filter(Boolean));
        set.add(c);
        this.className = [...set].join(' ');
      },
      remove: (c) => {
        const set = new Set(this.className.split(/\s+/).filter(Boolean));
        set.delete(c);
        this.className = [...set].join(' ');
      },
      toggle: (c, force) => {
        if (force) this.classList.add(c);
        else this.classList.remove(c);
      }
    };
  }
  get innerHTML() { return this._innerHTML; }
  set innerHTML(v) {
    this._innerHTML = v;
    if (v === '') this.children = [];
  }
  appendChild(child) {
    if (child instanceof MockDocumentFragment) {
      for (const node of child.children) {
        node.parentElement = this;
        this.children.push(node);
      }
      child.children = [];
      return child;
    }
    child.parentElement = this;
    this.children.push(child);
    return child;
  }
  append(...nodes) {
    for (const n of nodes) this.appendChild(n);
  }
  setAttribute(k, v) { this.attributes[k] = String(v); }
  setSelectionRange(start) { this.selectionStart = start; }
  addEventListener(type, fn) {
    if (!this.listeners[type]) this.listeners[type] = [];
    this.listeners[type].push(fn);
  }
  dispatchEvent(event) {
    event.target = event.target || this;
    event.preventDefault = event.preventDefault || (() => { event.defaultPrevented = true; });
    const fns = this.listeners[event.type] || [];
    for (const fn of fns) fn(event);
    if (this.parentElement) {
      this.parentElement.dispatchEvent(event);
    }
  }
  querySelector(selector) {
    if (selector === '.check-state-badge') {
      return this.children.find((c) => c.className.includes('check-state-badge')) || null;
    }
    return null;
  }
  querySelectorAll(selector) {
    const results = [];
    const walk = (el) => {
      for (const c of el.children) {
        if (selector === 'input[data-field="observaciones"]' && c.tagName === 'INPUT' && c.dataset.field === 'observaciones') {
          results.push(c);
        }
        if (selector.startsWith('input[type="checkbox"][data-field=') && c.tagName === 'INPUT' && c.type === 'checkbox') {
          const matchField = selector.match(/data-field="([^"]+)"/);
          if (matchField && c.dataset.field === matchField[1]) {
            results.push(c);
          }
        }
        walk(c);
      }
    };
    walk(this);
    return results;
  }
}

class MockDocumentFragment extends MockElement {
  constructor() { super('#fragment'); }
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
elementsById.selectProgCriterion.value = 'lte80';
elementsById.chkRespectSheetFilter.checked = true;

const masterCheckArt4 = new MockElement('input');
masterCheckArt4.type = 'checkbox';
masterCheckArt4.dataset.column = 'incumpleArt4';
masterCheckArt4.className = 'th-master-check';

const masterCheckArt7 = new MockElement('input');
masterCheckArt7.type = 'checkbox';
masterCheckArt7.dataset.column = 'incumpleArt7';
masterCheckArt7.className = 'th-master-check';

let domReadyCb = null;
const mockDocument = {
  getElementById: (id) => elementsById[id] || null,
  createElement: (tag) => new MockElement(tag),
  createDocumentFragment: () => new MockDocumentFragment(),
  querySelectorAll: (selector) => {
    if (selector === '.th-master-check') {
      return [masterCheckArt4, masterCheckArt7];
    }
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

let lastSavedPdfRows = null;
const OriginalJsPDF = jspdfMod.jsPDF;
class TrackedJsPDF extends OriginalJsPDF {
  constructor(opts) {
    super(opts);
    const origAutoTable = this.autoTable;
    this.autoTable = (config) => {
      lastSavedPdfRows = config.body;
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
sandbox.window.window = sandbox.window;

const appCode = fs.readFileSync(path.join(workspace, 'js/app.js'), 'utf8');
vm.runInNewContext(appCode, sandbox);
domReadyCb();

// Create sample data matching the exact numbers in the user's screenshot!
// Row 1: PONCE WALTER JAVIER: prog 77.22 (diff 77.22-80 = -2.78), emerg 26.65, total 103.87 (diff 103.87-120 = -16.13)
// Row 2: MONJE MORENA: prog 77.1 (diff -2.9), emerg 21.35, total 98.45 (diff -21.55)
// Row 3: RODRIGUEZ: prog 95 (diff 15 > 0), emerg 30, total 125 (diff 5 > 0)
const sampleRows = [
  ['N° personal', 'Apellido Nombre', 'Total Horas Programadas', 'Total Horas Emergencia', 'TOTAL HS PROGRAMADAS + EMERGENCIA'],
  ['83478', 'PONCE WALTER JAVIER', 77.22, 26.65, 103.87],
  ['84442', 'MONJE MORENA', 77.1, 21.35, 98.45],
  ['11890', 'RODRIGUEZ CARLOS', 95, 30, 125]
];

const ws = XLSX.utils.aoa_to_sheet(sampleRows);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Hoja1');
const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });

elementsById.inputFile.files = [{ name: 'test_negatives_and_bulk.xlsx', _buffer: buf }];
elementsById.inputFile.dispatchEvent({ type: 'change' });

setTimeout(() => {
  console.log('--- TEST 1: Valores negativos reemplazados por 0 en la tabla ---');
  const trs = elementsById.tableBody.children;
  assert.strictEqual(trs.length, 3, '3 rows loaded');

  // Row 1 (PONCE): 77.22 horas programadas
  // Column 3 (Cant. hs. extras excedidas Art. 8) MUST BE '0', NOT '-2.78'
  const ponceArt8Text = trs[0].children[3].textContent;
  const ponceExc120Text = trs[0].children[6].textContent;
  console.log('PONCE: Art 8 =', ponceArt8Text, '(esperado 0) | Exc 120 =', ponceExc120Text, '(esperado 0)');
  assert.strictEqual(ponceArt8Text, '0', 'Art. 8 for Ponce must be 0, not negative');
  assert.strictEqual(ponceExc120Text, '0', 'Exc 120 for Ponce must be 0, not negative');

  // Row 2 (MONJE): 77.1 horas
  const monjeArt8Text = trs[1].children[3].textContent;
  const monjeExc120Text = trs[1].children[6].textContent;
  console.log('MONJE: Art 8 =', monjeArt8Text, '(esperado 0) | Exc 120 =', monjeExc120Text, '(esperado 0)');
  assert.strictEqual(monjeArt8Text, '0', 'Art. 8 for Monje must be 0, not negative');
  assert.strictEqual(monjeExc120Text, '0', 'Exc 120 for Monje must be 0, not negative');

  // Row 3 (RODRIGUEZ): 95 horas (diff = 15), total 125 (diff = 5)
  const rodArt8Text = trs[2].children[3].textContent;
  const rodExc120Text = trs[2].children[6].textContent;
  console.log('RODRIGUEZ: Art 8 =', rodArt8Text, '(esperado 15) | Exc 120 =', rodExc120Text, '(esperado 5)');
  assert.strictEqual(rodArt8Text, '15', 'Positive values remain untouched');
  assert.strictEqual(rodExc120Text, '5', 'Positive values remain untouched');

  console.log('--- TEST 2: Opción para cambiar masivamente el valor de las columnas Art. 4° y Art. 7° ---');
  // Check master checkbox for Art. 4° -> should check all rows
  masterCheckArt4.checked = true;
  masterCheckArt4.dispatchEvent({ type: 'change' });

  // Verify all 3 rows have Art 4 checked and badge "Si"
  const art4Checkboxes = elementsById.tableBody.querySelectorAll('input[type="checkbox"][data-field="incumpleArt4"]');
  assert.strictEqual(art4Checkboxes.length, 3);
  art4Checkboxes.forEach((cb) => {
    assert.strictEqual(cb.checked, true, 'Row checkbox checked');
    assert.strictEqual(cb.parentElement.children[1].textContent, 'Si', 'Badge says Si');
  });
  assert.strictEqual(masterCheckArt4.checked, true);
  assert.strictEqual(masterCheckArt4.indeterminate, false);
  console.log('Master check Art. 4°: Todas las filas marcadas como Si correctamente.');

  // Uncheck row 0 -> master checkbox becomes indeterminate
  art4Checkboxes[0].checked = false;
  art4Checkboxes[0].dispatchEvent({ type: 'change', target: art4Checkboxes[0] });
  assert.strictEqual(masterCheckArt4.indeterminate, true, 'Master check becomes indeterminate when mixed');
  console.log('Master check Art. 4°: Estado indeterminado detectado al desmarcar una fila.');

  // Test toolbar button for Art. 7°
  elementsById.btnToggleAllArt7.dispatchEvent({ type: 'click' });
  const art7Checkboxes = elementsById.tableBody.querySelectorAll('input[type="checkbox"][data-field="incumpleArt7"]');
  art7Checkboxes.forEach((cb) => {
    assert.strictEqual(cb.checked, true, 'Row checkbox in Art 7 checked');
    assert.strictEqual(cb.parentElement.children[1].textContent, 'Si');
  });
  console.log('Botón toolbar Art. 7°: Todas las filas marcadas como Si correctamente.');

  console.log('--- TEST 3: Verificación de valores en el PDF generado ---');
  elementsById.btnExportPdf.dispatchEvent({ type: 'click' });
  assert(lastSavedPdfRows, 'PDF rows captured');

  // Verify row 0 in PDF (PONCE): Art 8 is '0', Exc 120 is '0', Art 4 is 'No' (since we unchecked row 0), Art 7 is 'Si'
  const pdfPonce = lastSavedPdfRows[0];
  console.log('PDF Ponce:', pdfPonce);
  assert.strictEqual(pdfPonce[2], '0', 'PDF Art 8 for Ponce is 0, not negative');
  assert.strictEqual(pdfPonce[4], '0', 'PDF Exc 120 for Ponce is 0, not negative');
  assert.strictEqual(pdfPonce[5], 'No', 'PDF Art 4 for Ponce is No');
  assert.strictEqual(pdfPonce[6], 'Si', 'PDF Art 7 for Ponce is Si');

  // Verify row 1 in PDF (MONJE): Art 8 is '0', Exc 120 is '0', Art 4 is 'Si', Art 7 is 'Si'
  const pdfMonje = lastSavedPdfRows[1];
  console.log('PDF Monje:', pdfMonje);
  assert.strictEqual(pdfMonje[2], '0', 'PDF Art 8 for Monje is 0, not negative');
  assert.strictEqual(pdfMonje[4], '0', 'PDF Exc 120 for Monje is 0, not negative');
  assert.strictEqual(pdfMonje[5], 'Si', 'PDF Art 4 for Monje is Si');
  assert.strictEqual(pdfMonje[6], 'Si', 'PDF Art 7 for Monje is Si');

  // Verify row 2 in PDF (RODRIGUEZ): Art 8 is '15', Exc 120 is '5'
  const pdfRod = lastSavedPdfRows[2];
  console.log('PDF Rodriguez:', pdfRod);
  assert.strictEqual(pdfRod[2], '15');
  assert.strictEqual(pdfRod[4], '5');

  console.log('ALL TESTS PASSED WITH 100% SUCCESS!');
}, 100);

