/**
 * Generador Offline "Anexo A: Detalle de Horas Extras Excedidas" (EPE)
 * Formulario PR-RH-AP-010 (Versión 00 — Res. 077/2020)
 * Arquitectura Modular Vanilla JS ES6+ (100% Offline-First)
 */
(() => {
  'use strict';

  const MONTHS_ES = [
    'ENERO',
    'FEBRERO',
    'MARZO',
    'ABRIL',
    'MAYO',
    'JUNIO',
    'JULIO',
    'AGOSTO',
    'SEPTIEMBRE',
    'OCTUBRE',
    'NOVIEMBRE',
    'DICIEMBRE'
  ];

  const REQUIRED_COLUMNS = [
    'N° personal',
    'Apellido Nombre',
    'Total Horas Programadas',
    'Total Horas Emergencia',
    'TOTAL HS PROGRAMADAS + EMERGENCIA'
  ];

  /* ========================================================================
   * 1. StateManager: Estado centralizado de la aplicación
   * ======================================================================== */
  const StateManager = (() => {
    const now = new Date();
    // Mes anterior al mes actual (0 = enero -> mes anterior diciembre índice 11)
    const prevMonthIdx = (now.getMonth() + 11) % 12;
    // Año actual
    const defaultYear = now.getFullYear();
    const defaultMonth = MONTHS_ES[prevMonthIdx];

    const state = {
      metadata: {
        anio: defaultYear,
        mes: defaultMonth,
        expte: '',
        jefe: '',
        sector: '',
        observacionesGlobal: ''
      },
      respectSheetFilter: true, // Respeta filtros aplicados en el archivo (.xlsx/.ods)
      fileName: '',
      rawRows: [],
      totalRowsRead: 0,
      sheetFilterInfo: {
        hasActiveFilter: false,
        hasAutoFilter: false,
        hiddenRowCount: 0,
        visibleRowCount: 0
      },
      filteredRowsState: []
    };

    return {
      getState() {
        return state;
      },
      setMetadataField(key, value) {
        if (Object.prototype.hasOwnProperty.call(state.metadata, key)) {
          state.metadata[key] = value;
        }
      },
      setRespectSheetFilter(val) {
        state.respectSheetFilter = Boolean(val);
      },
      setParsedData({ fileName, rawRows, totalRowsRead, sheetFilterInfo, filteredRows }) {
        state.fileName = fileName;
        state.rawRows = rawRows;
        state.totalRowsRead = totalRowsRead;
        state.sheetFilterInfo = sheetFilterInfo;
        state.filteredRowsState = filteredRows;
      },
      setFilteredRows(filteredRows) {
        state.filteredRowsState = filteredRows;
      },
      updateRowField(index, field, value) {
        const row = state.filteredRowsState[index];
        if (row && Object.prototype.hasOwnProperty.call(row, field)) {
          row[field] = value;
        }
      },
      setAllColumnCheckboxes(field, value) {
        const boolVal = Boolean(value);
        for (const row of state.filteredRowsState) {
          row[field] = boolVal;
        }
      },
      toggleAllColumnCheckboxes(field) {
        const allChecked =
          state.filteredRowsState.length > 0 &&
          state.filteredRowsState.every((r) => Boolean(r[field]));
        const targetVal = !allChecked;
        for (const row of state.filteredRowsState) {
          row[field] = targetVal;
        }
        return targetVal;
      },
      applyGlobalObservaciones(text) {
        state.metadata.observacionesGlobal = text;
        for (const row of state.filteredRowsState) {
          row.observaciones = text;
        }
      }
    };
  })();

  /* ========================================================================
   * 2. SpreadsheetParser: Lectura (.xlsx/.ods), normalización y filtrado
   * ======================================================================== */
  const SpreadsheetParser = (() => {
    /**
     * Normaliza nombres de encabezado recortando extremos y colapsando espacios múltiples.
     */
    function normalizeHeader(value) {
      return String(value ?? '').trim().replace(/\s+/g, ' ');
    }

    /**
     * Convierte de forma segura celdas numéricas o cadenas con coma/punto decimal a Number.
     * Devuelve 0 si la celda está vacía, nula o no es numérica.
     */
    function parseNumericCell(val) {
      if (typeof val === 'number') {
        return Number.isFinite(val) ? val : 0;
      }
      if (val === null || val === undefined) {
        return 0;
      }

      const raw = String(val).trim().replace(/\s+/g, '');
      if (!raw) {
        return 0;
      }

      let normalized = raw;
      const hasComma = normalized.includes(',');
      const hasDot = normalized.includes('.');

      if (hasComma && hasDot) {
        if (normalized.lastIndexOf(',') > normalized.lastIndexOf('.')) {
          normalized = normalized.replace(/\./g, '').replace(',', '.');
        } else {
          normalized = normalized.replace(/,/g, '');
        }
      } else if (hasComma) {
        normalized = normalized.replace(',', '.');
      }

      const parsed = Number(normalized);
      return Number.isFinite(parsed) ? parsed : 0;
    }

    /**
     * Formatea valores numéricos para visualización en pantalla.
     * Muestra entero o decimales sin ceros sobrantes, o '0' si es <= 0.
     */
    function formatHours(value) {
      const rounded = Math.round(value * 100) / 100;
      if (Object.is(rounded, -0) || rounded <= 0) {
        return '0';
      }
      return Number.isInteger(rounded) ? String(rounded) : String(rounded);
    }

    /**
     * Formatea valores numéricos para las columnas de horas del archivo PDF:
     * - Si el valor es 0 (o <= 0): devuelve exactamente '0' (sin decimales).
     * - Si el valor es mayor a 0: lleva coma (,) como separador decimal y exactamente 2 decimales (ej. 7,47; 15,00).
     */
    function formatHoursPdf(value) {
      const num = Number(value);
      if (!Number.isFinite(num) || num <= 0) {
        return '0';
      }
      const rounded = Math.round(num * 100) / 100;
      if (rounded <= 0) {
        return '0';
      }
      return rounded.toFixed(2).replace('.', ',');
    }

    /**
     * Determina si una fila está completamente vacía respecto a las columnas requeridas.
     */
    function isRowEmpty(rowArray, colMap) {
      if (!Array.isArray(rowArray)) {
        return true;
      }
      return REQUIRED_COLUMNS.every((colName) => {
        const cell = rowArray[colMap[colName]];
        return cell === null || cell === undefined || String(cell).trim() === '';
      });
    }

    /**
     * Verifica si una fila de la hoja está oculta/filtrada según las propiedades de SheetJS.
     */
    function isRowHidden(worksheet, rowIndex) {
      if (!worksheet || !worksheet['!rows']) {
        return false;
      }
      const rowInfo = worksheet['!rows'][rowIndex];
      if (!rowInfo) {
        return false;
      }
      return rowInfo.hidden === true || rowInfo.hpx === 0 || rowInfo.hpt === 0;
    }

    /**
     * Para archivos .ods: LibreOffice almacena la visibilidad de filtros en content.xml
     * como table:visibility="filter" o "collapse". Esta función inspecciona el archivo
     * y puebla worksheet['!rows'] para que el motor reconozca los filtros aplicados en Calc.
     */
    function populateOdsHiddenRows(workbook, buffer) {
      try {
        if (!window.XLSX || !window.XLSX.CFB) {
          return;
        }
        const cfb = window.XLSX.CFB.read(buffer, { type: 'array' });
        const entry =
          window.XLSX.CFB.find(cfb, 'Root Entry/content.xml') ||
          window.XLSX.CFB.find(cfb, 'content.xml');
        if (!entry || !entry.content) {
          return;
        }

        const uint8 =
          entry.content instanceof Uint8Array
            ? entry.content
            : new Uint8Array(entry.content);
        const text =
          typeof TextDecoder !== 'undefined'
            ? new TextDecoder('utf-8').decode(uint8)
            : String.fromCharCode.apply(null, uint8);

        // Estilos que definen visibilidad oculta/colapsada
        const hiddenStyles = new Set();
        const styleRegex = /<style:style\b([^>]*)>([\s\S]*?)<\/style:style>/gi;
        let sMatch;
        while ((sMatch = styleRegex.exec(text)) !== null) {
          const sAttrs = sMatch[1];
          const sBody = sMatch[2];
          const nameMatch = sAttrs.match(/style:name="([^"]+)"/i);
          if (
            nameMatch &&
            /table:(visibility="filter"|visibility="collapse"|display="false")/i.test(
              sBody + ' ' + sAttrs
            )
          ) {
            hiddenStyles.add(nameMatch[1]);
          }
        }

        const tableMatch = text.match(/<table:table\b[^>]*>([\s\S]*?)<\/table:table>/i);
        if (!tableMatch) {
          return;
        }
        const tableContent = tableMatch[1];

        const rowRegex = /<table:table-row\b([^>]*)>/gi;
        let match;
        let rowIndex = 0;
        const firstSheetName = workbook.SheetNames[0];
        const ws = workbook.Sheets[firstSheetName];
        if (!ws) {
          return;
        }
        if (!ws['!rows']) {
          ws['!rows'] = [];
        }

        while ((match = rowRegex.exec(tableContent)) !== null) {
          const attrs = match[1];
          const repeatMatch = attrs.match(/table:number-rows-repeated="([0-9]+)"/i);
          const repeatCount = repeatMatch ? parseInt(repeatMatch[1], 10) : 1;

          const isVisibilityHidden = /table:visibility="(filter|collapse|hidden)"/i.test(attrs);
          const styleMatch = attrs.match(/table:style-name="([^"]+)"/i);
          const isStyleHidden = styleMatch ? hiddenStyles.has(styleMatch[1]) : false;
          const isHidden = isVisibilityHidden || isStyleHidden;

          const effectiveRepeat = Math.min(repeatCount, 10000);
          for (let k = 0; k < effectiveRepeat; k += 1) {
            if (!ws['!rows'][rowIndex]) {
              ws['!rows'][rowIndex] = {};
            }
            if (isHidden) {
              ws['!rows'][rowIndex].hidden = true;
            }
            rowIndex += 1;
          }
        }
      } catch (err) {
        console.warn('No se pudo inspeccionar content.xml para filtros ODS:', err);
      }
    }

    /**
     * Lee el archivo local (.xlsx o .ods) en ArrayBuffer y extrae las filas base de la 1ra hoja,
     * detectando los filtros aplicados (AutoFiltro de Excel / Calc) y la visibilidad de cada fila.
     */
    function parseFile(file) {
      return new Promise((resolve, reject) => {
        if (typeof window.XLSX === 'undefined') {
          reject(new Error('No se encontró la librería SheetJS (./libs/xlsx.full.min.js).'));
          return;
        }

        const reader = new FileReader();

        reader.onerror = () => {
          reject(new Error('No se pudo leer el archivo seleccionado desde el disco.'));
        };

        reader.onload = (event) => {
          try {
            const buffer = new Uint8Array(event.target.result);
            // IMPORTANTE: cellStyles: true es necesario para que SheetJS capture '!rows' con { hidden: true }
            const workbook = window.XLSX.read(buffer, {
              type: 'array',
              cellStyles: true
            });

            if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
              reject(new Error('El archivo no contiene ninguna hoja de cálculo.'));
              return;
            }

            const isOds = /\.ods$/i.test(file.name);
            if (isOds) {
              populateOdsHiddenRows(workbook, buffer);
            }

            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // blankrows: true preserva correspondencia 1 a 1 de índices entre matrix y worksheet['!rows']
            const matrix = window.XLSX.utils.sheet_to_json(worksheet, {
              header: 1,
              defval: '',
              blankrows: true
            });

            if (!matrix || matrix.length === 0) {
              reject(new Error('La primera hoja del archivo está vacía.'));
              return;
            }

            // Hallar la fila de encabezado (primera fila no vacía)
            let headerRowIndex = 0;
            while (
              headerRowIndex < matrix.length &&
              (!Array.isArray(matrix[headerRowIndex]) || matrix[headerRowIndex].length === 0)
            ) {
              headerRowIndex += 1;
            }

            if (headerRowIndex >= matrix.length) {
              reject(new Error('No se encontró fila de encabezados en el archivo.'));
              return;
            }

            const headerRow = matrix[headerRowIndex].map(normalizeHeader);
            const colMap = {};
            const missingCols = [];

            for (const requiredCol of REQUIRED_COLUMNS) {
              const target = normalizeHeader(requiredCol);
              const idx = headerRow.findIndex((h) => h === target);
              if (idx === -1) {
                missingCols.push(requiredCol);
              } else {
                colMap[requiredCol] = idx;
              }
            }

            if (missingCols.length > 0) {
              reject(
                new Error(
                  `Faltan las siguientes columnas obligatorias en la primera fila: ${missingCols.join(', ')}.`
                )
              );
              return;
            }

            const hasAutoFilter = Boolean(worksheet['!autofilter'] && worksheet['!autofilter'].ref);
            let hiddenRowCount = 0;
            let visibleRowCount = 0;
            const rawRows = [];

            for (let i = headerRowIndex + 1; i < matrix.length; i += 1) {
              const rowArray = matrix[i];
              if (!Array.isArray(rowArray) || isRowEmpty(rowArray, colMap)) {
                continue;
              }

              const isHidden = isRowHidden(worksheet, i);
              if (isHidden) {
                hiddenRowCount += 1;
              } else {
                visibleRowCount += 1;
              }

              const legajoRaw = rowArray[colMap['N° personal']];
              const apellidoRaw = rowArray[colMap['Apellido Nombre']];
              const hsProgramadas = parseNumericCell(rowArray[colMap['Total Horas Programadas']]);
              const hsEmergencia = parseNumericCell(rowArray[colMap['Total Horas Emergencia']]);
              const hsTotalProgMasEmerg = parseNumericCell(
                rowArray[colMap['TOTAL HS PROGRAMADAS + EMERGENCIA']]
              );

              rawRows.push({
                sourceRowIndex: i + 1,
                sheetRowIndex: i,
                isVisibleInSheet: !isHidden,
                legajo: String(legajoRaw ?? '').trim(),
                apellidoNombre: String(apellidoRaw ?? '').trim(),
                hsProgramadas,
                hsEmergencia,
                hsTotalProgMasEmerg
              });
            }

            const hasActiveFilter = hasAutoFilter || hiddenRowCount > 0;
            const sheetFilterInfo = {
              hasActiveFilter,
              hasAutoFilter,
              hiddenRowCount,
              visibleRowCount
            };

            resolve({
              totalRowsRead: rawRows.length,
              rawRows,
              sheetFilterInfo
            });
          } catch (err) {
            reject(new Error(`Error al procesar el archivo de planilla: ${err.message}`));
          }
        };

        reader.readAsArrayBuffer(file);
      });
    }

    /**
     * Aplica simultáneamente el respeto a los filtros del archivo (.xlsx/.ods)
     * y las condiciones de horas extras (Res. 077/2020) con lógica OR:
     * - Condición 1: Total Horas Programadas > 80 (Excedidas Art. 8°)
     * - Condición 2: Total Horas Emergencia > 0
     * - Condición 3: TOTAL HS PROGRAMADAS + EMERGENCIA > 120
     *
     * IMPORTANTE: Si las operaciones de cálculo (hsProgramadas - 80) o (hsTotal - 120)
     * dan valores negativos, se coloca exactamente 0.
     */
    function filterRows(rawRows, options = {}, previousFiltered = []) {
      const {
        respectSheetFilter = true
      } = options;

      const previousByKey = new Map();
      for (const prev of previousFiltered) {
        previousByKey.set(`${prev.sourceRowIndex}::${prev.legajo}`, prev);
      }

      const filtered = [];
      for (const item of rawRows) {
        // 1. Respetar filtros del archivo: si está habilitado, descartar filas ocultadas por el filtro
        if (respectSheetFilter && !item.isVisibleInSheet) {
          continue;
        }

        // 2. Filtrar por las condiciones marcadas en la reglamentación (OR Lógico):
        // Total Horas Programadas > 80, Total Horas Emergencia > 0 o TOTAL > 120
        const condProgramadas = item.hsProgramadas > 80;
        const condEmergencia = item.hsEmergencia > 0;
        const condTotal = item.hsTotalProgMasEmerg > 120;

        const cumpleCondicion = condProgramadas || condEmergencia || condTotal;
        if (!cumpleCondicion) {
          continue;
        }

        const prev = previousByKey.get(`${item.sourceRowIndex}::${item.legajo}`);

        // Regla: si el resultado de la operación es negativo, colocar 0
        const rawArt8 = item.hsProgramadas - 80;
        const excedidasArt8 = rawArt8 > 0 ? rawArt8 : 0;

        const raw120 = item.hsTotalProgMasEmerg - 120;
        const exceden120 = raw120 > 0 ? raw120 : 0;

        filtered.push({
          sourceRowIndex: item.sourceRowIndex,
          sheetRowIndex: item.sheetRowIndex,
          isVisibleInSheet: item.isVisibleInSheet,
          legajo: item.legajo,
          apellidoNombre: item.apellidoNombre,
          hsProgramadas: item.hsProgramadas,
          excedidasArt8,
          hsEmergencia: item.hsEmergencia,
          hsTotalProgMasEmerg: item.hsTotalProgMasEmerg,
          exceden120,
          incumpleArt4: prev ? prev.incumpleArt4 : false,
          incumpleArt7: prev ? prev.incumpleArt7 : false,
          observaciones: prev ? prev.observaciones : ''
        });
      }

      return filtered;
    }

    return {
      parseNumericCell,
      formatHours,
      formatHoursPdf,
      parseFile,
      filterRows
    };
  })();

  /* ========================================================================
   * 3. TableView: Renderizado interactivo y delegación de eventos
   * ======================================================================== */
  const TableView = (() => {
    let tbodyEl = null;

    function createTextCell(text, className = '') {
      const td = document.createElement('td');
      if (className) {
        td.className = className;
      }
      td.textContent = text;
      return td;
    }

    function createCheckboxCell(rowIndex, field, checked, labelTitle) {
      const td = document.createElement('td');
      td.className = 'cell-check';

      const label = document.createElement('label');
      label.className = 'check-toggle-label';
      label.title = labelTitle;

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = Boolean(checked);
      checkbox.dataset.index = String(rowIndex);
      checkbox.dataset.field = field;

      const stateBadge = document.createElement('span');
      stateBadge.className = `check-state-badge${checked ? ' is-yes' : ''}`;
      stateBadge.textContent = checked ? 'Si' : 'No';

      label.append(checkbox, stateBadge);
      td.appendChild(label);
      return td;
    }

    function createObsCell(rowIndex, value) {
      const td = document.createElement('td');
      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'row-obs-input';
      input.value = value ?? '';
      input.placeholder = 'Observaciones...';
      input.dataset.index = String(rowIndex);
      input.dataset.field = 'observaciones';
      input.setAttribute('aria-label', `Observaciones fila ${rowIndex + 1}`);
      td.appendChild(input);
      return td;
    }

    function render(rows, hasLoadedFile) {
      if (!tbodyEl) {
        return;
      }

      tbodyEl.innerHTML = '';

      if (!rows || rows.length === 0) {
        const tr = document.createElement('tr');
        tr.className = 'empty-row';
        const td = document.createElement('td');
        td.colSpan = 10;
        td.textContent = hasLoadedFile
          ? 'Ninguna fila cumplió con las condiciones de filtrado configuradas.'
          : 'No hay registros cargados aún. Seleccione un archivo .xlsx o .ods para visualizar y editar los datos.';
        tr.appendChild(td);
        tbodyEl.appendChild(tr);
        return;
      }

      const fragment = document.createDocumentFragment();

      rows.forEach((row, index) => {
        const tr = document.createElement('tr');
        tr.dataset.rowIndex = String(index);

        // Garantizar que si el cálculo es <= 0 se muestre '0'
        const art8Number = row.excedidasArt8 > 0 ? row.excedidasArt8 : 0;
        const exc120Number = row.exceden120 > 0 ? row.exceden120 : 0;
        const art8Str = SpreadsheetParser.formatHours(art8Number);
        const exc120Str = SpreadsheetParser.formatHours(exc120Number);

        tr.appendChild(createTextCell(row.legajo));
        tr.appendChild(createTextCell(row.apellidoNombre));
        tr.appendChild(createTextCell(SpreadsheetParser.formatHours(row.hsProgramadas), 'cell-num'));
        tr.appendChild(
          createTextCell(
            art8Str,
            `cell-num cell-calc${art8Number > 0 ? ' is-positive' : ''}`
          )
        );
        tr.appendChild(createTextCell(SpreadsheetParser.formatHours(row.hsEmergencia), 'cell-num'));
        tr.appendChild(createTextCell(SpreadsheetParser.formatHours(row.hsTotalProgMasEmerg), 'cell-num'));
        tr.appendChild(
          createTextCell(
            exc120Str,
            `cell-num cell-calc${exc120Number > 0 ? ' is-positive' : ''}`
          )
        );
        tr.appendChild(
          createCheckboxCell(
            index,
            'incumpleArt4',
            row.incumpleArt4,
            'Incumple Art. 4° Res. 077/2020'
          )
        );
        tr.appendChild(
          createCheckboxCell(
            index,
            'incumpleArt7',
            row.incumpleArt7,
            'Incumple Art. 7° Res. 077/2020'
          )
        );
        tr.appendChild(createObsCell(index, row.observaciones));

        fragment.appendChild(tr);
      });

      tbodyEl.appendChild(fragment);
    }

    function updateAllCheckboxesInDom(field, checked) {
      if (!tbodyEl) {
        return;
      }
      const checkboxes = tbodyEl.querySelectorAll(`input[type="checkbox"][data-field="${field}"]`);
      checkboxes.forEach((cb) => {
        cb.checked = checked;
        const badge = cb.parentElement?.querySelector('.check-state-badge');
        if (badge) {
          badge.textContent = checked ? 'Si' : 'No';
          badge.classList.toggle('is-yes', checked);
        }
      });
    }

    function syncObservacionesInputs(text) {
      if (!tbodyEl) {
        return;
      }
      const inputs = tbodyEl.querySelectorAll('input[data-field="observaciones"]');
      inputs.forEach((input) => {
        input.value = text;
      });
    }

    function init(tbodyElement, onCheckboxChangeCallback) {
      tbodyEl = tbodyElement;

      tbodyEl.addEventListener('change', (event) => {
        const target = event.target;
        if (!(target instanceof HTMLInputElement) || target.type !== 'checkbox') {
          return;
        }
        const index = Number(target.dataset.index);
        const field = target.dataset.field;
        if (!Number.isInteger(index) || !field) {
          return;
        }

        StateManager.updateRowField(index, field, target.checked);

        const badge = target.parentElement?.querySelector('.check-state-badge');
        if (badge) {
          badge.textContent = target.checked ? 'Si' : 'No';
          badge.classList.toggle('is-yes', target.checked);
        }

        if (typeof onCheckboxChangeCallback === 'function') {
          onCheckboxChangeCallback(field);
        }
      });

      tbodyEl.addEventListener('input', (event) => {
        const target = event.target;
        if (!(target instanceof HTMLInputElement) || target.dataset.field !== 'observaciones') {
          return;
        }
        const index = Number(target.dataset.index);
        if (!Number.isInteger(index)) {
          return;
        }
        StateManager.updateRowField(index, 'observaciones', target.value);
      });
    }

    return {
      init,
      render,
      updateAllCheckboxesInDom,
      syncObservacionesInputs
    };
  })();

  /* ========================================================================
   * 4. PdfExporter: Generación exacta del Formulario Anexo A (PR-RH-AP-010)
   * ======================================================================== */
  const PdfExporter = (() => {
    /**
     * Dibuja una línea punteada horizontal entre startX y endX en la coordenada Y indicada.
     */
    function drawDottedLine(doc, startX, endX, y) {
      doc.setDrawColor(0, 0, 0);
      doc.setLineWidth(0.2);
      doc.setLineDashPattern([0.55, 0.85], 0);
      doc.line(startX, y, endX, y);
      doc.setLineDashPattern([], 0);
    }

    /**
     * Dibuja el logotipo vectorial simplificado de EPE Santa Fe en el bloque superior izquierdo.
     */
const LOGO_EPE_BLANCO_DATA = 'data:image/jpeg;base64,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';

    function drawEpeLogo(doc, boxX, boxY, boxW, boxH = 19) {
      try {
        const maxImgW = 26;
        const maxImgH = 15;
        const finalH = Math.min(maxImgW * (175 / 258), maxImgH);
        const finalW = finalH * (258 / 175);
        const imgX = boxX + (boxW - finalW) / 2;
        const imgY = boxY + (boxH - finalH) / 2;
        doc.addImage(LOGO_EPE_BLANCO_DATA, 'JPEG', imgX, imgY, finalW, finalH);
      } catch (err) {
        const centerX = boxX + boxW / 2;
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(16);
        doc.setTextColor(0, 75, 135);
        doc.text('EPE', centerX, boxY + 8.5, { align: 'center' });
        doc.setFillColor(0, 75, 135);
        doc.rect(boxX + 4, boxY + 10.2, boxW - 11, 1.3, 'F');
        doc.triangle(boxX + boxW - 7, boxY + 9.7, boxX + boxW - 3.5, boxY + 10.85, boxX + boxW - 7, boxY + 12.0, 'F');
        doc.setFontSize(6.5);
        doc.text('Energía de Santa Fe', centerX, boxY + 15.8, { align: 'center' });
        doc.setTextColor(0, 0, 0);
      }
    }

    /**
     * Dibuja la cabecera institucional, título, campos del formulario y pie de página.
     */
    function drawPageChrome(doc, metadata) {
      const marginX = 15;
      const fullW = 180;
      const topY = 12;
      const topBandH = 19;
      const stripH = 5;
      const leftW = 32;
      const centerW = 110;
      const rightW = 38;

      const xCenter = marginX + leftW;
      const xRight = xCenter + centerW;
      const stripY = topY + topBandH;

      // 5.1 Caja de Encabezado Institucional (180 x 24 mm)
      doc.setDrawColor(0, 0, 0);
      doc.setLineWidth(0.3);
      doc.rect(marginX, topY, fullW, topBandH + stripH);

      // Líneas divisorias internas de la caja superior
      doc.line(marginX, stripY, marginX + fullW, stripY);
      doc.line(xCenter, topY, xCenter, stripY);
      doc.line(xRight, topY, xRight, stripY);

      // Bloque Izquierdo: Logo EPE
      drawEpeLogo(doc, marginX, topY, leftW);

      // Bloque Central: Textos institucionales
      const centerTextX = xCenter + centerW / 2;
      doc.setTextColor(0, 0, 0);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(9.5);
      doc.text('EMPRESA PROVINCIAL DE LA ENERGÍA DE SANTA FE', centerTextX, topY + 6.0, {
        align: 'center'
      });
      doc.setFontSize(9);
      doc.text('GERENCIA DE RECURSOS HUMANOS', centerTextX, topY + 11.0, { align: 'center' });
      doc.text('MANUAL DE PROCEDIMIENTOS', centerTextX, topY + 16.0, { align: 'center' });

      // Bloque Derecho: 3 celdas horizontales
      const cellH = topBandH / 3;
      doc.line(xRight, topY + cellH, xRight + rightW, topY + cellH);
      doc.line(xRight, topY + cellH * 2, xRight + rightW, topY + cellH * 2);

      doc.setFontSize(7.5);
      doc.setFont('helvetica', 'bold');
      doc.text('PR-RH-AP-010', xRight + rightW / 2, topY + 4.3, { align: 'center' });

      doc.setFont('helvetica', 'normal');
      doc.text('Versión:00', xRight + 2, topY + cellH + 4.3);
      doc.text('Fecha de emisión:', xRight + 2, topY + cellH * 2 + 4.3);

      // Franja Inferior del Encabezado
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(8.5);
      doc.text(
        'JUSTIFICACIONES DE EXCESOS REGLAMENTACIÓN RES. 077/2020',
        marginX + fullW / 2,
        stripY + 3.6,
        { align: 'center' }
      );

      // 5.2 Título y Campos de Cabecera del Anexo A
      const titleText = 'ANEXO A: DETALLE DE HORAS EXTRAS EXCEDIDAS';
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(10);
      doc.text(titleText, 105, 45, { align: 'center' });

      // Subrayado exacto debajo de "ANEXO A:"
      const titleWidth = doc.getTextWidth(titleText);
      const prefixWidth = doc.getTextWidth('ANEXO A:');
      const titleStartX = 105 - titleWidth / 2;
      doc.setLineWidth(0.35);
      doc.line(titleStartX, 45.9, titleStartX + prefixWidth, 45.9);

      // Campo "Cde. Expte. N°:" (y = 55)
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(9);
      doc.text('Cde. Expte. N°:', 112, 55);
      drawDottedLine(doc, 138.5, 195, 55.3);
      if (metadata.expte) {
        doc.setFont('helvetica', 'normal');
        doc.text(String(metadata.expte), 140, 54.4, { maxWidth: 54 });
      }

      // Campo "SECTOR:" (y = 62)
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(9);
      doc.text('SECTOR:', 15, 62);
      drawDottedLine(doc, 31, 195, 62.3);
      if (metadata.sector) {
        doc.setFont('helvetica', 'normal');
        doc.text(String(metadata.sector), 33, 61.4, { maxWidth: 160 });
      }

      // Fila de Período (y = 69)
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(9);
      doc.text('AÑO:', 15, 69);
      drawDottedLine(doc, 24.5, 43, 69.3);
      doc.setFont('helvetica', 'normal');
      doc.text(String(metadata.anio ?? ''), 33.75, 68.4, { align: 'center' });

      doc.setFont('helvetica', 'bold');
      doc.text('MES:', 46.5, 69);
      drawDottedLine(doc, 55.5, 92, 69.3);
      doc.setFont('helvetica', 'normal');
      doc.text(String(metadata.mes ?? ''), 73.75, 68.4, { align: 'center' });

      doc.setFont('helvetica', 'bold');
      doc.text('DE EJECUCIÓN DE HORAS EXTRAS TRABAJADAS', 94.5, 69);

      // 5.4 Pie de Página Institucional (y = 268, ancho 180 mm, alto 11 mm)
      const footY = 268;
      const footH = 11;
      const footRowH = 5.5;
      const footRow2Y = footY + footRowH;

      doc.setDrawColor(0, 0, 0);
      doc.setLineWidth(0.3);
      doc.rect(marginX, footY, fullW, footH);
      doc.line(marginX, footRow2Y, marginX + fullW, footRow2Y);

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7.5);
      doc.text('FECHA DE ÚLTIMA REVISIÓN: 20-05-2026', marginX + 2, footY + 3.8);

      // 6 columnas de la fila inferior (26 + 34 + 28 + 32 + 26 + 34 = 180 mm)
      const colWidths = [26, 34, 28, 32, 26, 34];
      const colLabels = ['Realizado:', 'AAP', 'Supervisado:', 'AAP/ARL', 'Aprobado:', 'Gcia. RRHH'];
      const colCentered = [false, true, false, true, false, true];

      let cursorX = marginX;
      colWidths.forEach((w, idx) => {
        if (idx > 0) {
          doc.line(cursorX, footRow2Y, cursorX, footY + footH);
        }
        const textY = footRow2Y + 3.8;
        if (colCentered[idx]) {
          doc.text(colLabels[idx], cursorX + w / 2, textY, { align: 'center' });
        } else {
          doc.text(colLabels[idx], cursorX + 2, textY);
        }
        cursorX += w;
      });
    }

    function exportToPdf(appState) {
      if (!window.jspdf || !window.jspdf.jsPDF) {
        throw new Error('La librería jsPDF (./libs/jspdf.umd.min.js) no está cargada.');
      }

      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({
        orientation: 'portrait',
        unit: 'mm',
        format: 'a4'
      });

      if (typeof doc.autoTable !== 'function') {
        throw new Error(
          'El plugin jsPDF-AutoTable (./libs/jspdf.plugin.autotable.min.js) no está cargado.'
        );
      }

      const { metadata, filteredRowsState } = appState;
      const jefeValue = String(metadata.jefe ?? '').trim();

      // Filtrar exclusivamente filas que posean datos válidos (evitar cualquier fila vacía o nula)
      const validRows = (filteredRowsState || []).filter((row) => {
        if (!row) return false;
        const hasLegajo = String(row.legajo ?? '').trim() !== '';
        const hasNombre = String(row.apellidoNombre ?? '').trim() !== '';
        return hasLegajo || hasNombre;
      });

      // Regla de horas en PDF:
      // Columnas: Art. 8°, Emergencia y Exceden 120
      // - Si el valor es <= 0: exactamente '0' (sin decimales).
      // - Si es mayor a 0: parte decimal con coma (,) y 2 decimales fijos (ej: 7,47; 15,00).
      const bodyRows = validRows.length > 0
        ? validRows.map((row) => {
            const art8Val = row.excedidasArt8 > 0 ? row.excedidasArt8 : 0;
            const exc120Val = row.exceden120 > 0 ? row.exceden120 : 0;

            return [
              String(row.legajo ?? '').trim(),
              String(row.apellidoNombre ?? '').trim(),
              SpreadsheetParser.formatHoursPdf(art8Val),
              SpreadsheetParser.formatHoursPdf(row.hsEmergencia),
              SpreadsheetParser.formatHoursPdf(exc120Val),
              row.incumpleArt4 ? 'Si' : 'No',
              row.incumpleArt7 ? 'Si' : 'No',
              jefeValue,
              String(row.observaciones ?? '').trim()
            ];
          })
        : [['', '', '', '', '', '', '', '', '']];

      const headColumns = [[
        'Legajo',
        'Apellido y Nombres',
        'Cantidad de Horas extras excedidas Art. 8°',
        'Cantidad de horas extras en emergencia',
        'Cantidad de horas extras que exceden las 120',
        'Incumple Art. 4° Res. 077/2020',
        'Incumple Art. 7° Res. 077/2020',
        'Jefe a cargo del agente (legajo / Apellido y Nombres)',
        'Observaciones'
      ]];

      doc.autoTable({
        head: headColumns,
        body: bodyRows,
        startY: 74,
        margin: { top: 74, left: 15, right: 15, bottom: 32 },
        tableWidth: 180,
        theme: 'grid',
        rowPageBreak: 'avoid',
        styles: {
          font: 'helvetica',
          fontSize: 7.2,
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.25,
          cellPadding: { top: 1.4, right: 1.2, bottom: 1.4, left: 1.2 },
          valign: 'middle',
          overflow: 'linebreak'
        },
        headStyles: {
          fillColor: [217, 217, 217],
          textColor: [0, 0, 0],
          fontStyle: 'bold',
          halign: 'center',
          valign: 'middle',
          fontSize: 6.9,
          minCellHeight: 16
        },
        bodyStyles: {
          minCellHeight: 11,
          fillColor: [255, 255, 255]
        },
        columnStyles: {
          0: { cellWidth: 15, halign: 'center' },
          1: { cellWidth: 29, halign: 'left' },
          2: { cellWidth: 18, halign: 'center' },
          3: { cellWidth: 18, halign: 'center' },
          4: { cellWidth: 18, halign: 'center' },
          5: { cellWidth: 15, halign: 'center' },
          6: { cellWidth: 15, halign: 'center' },
          7: { cellWidth: 26, halign: 'left' },
          8: { cellWidth: 26, halign: 'left' }
        },
        didDrawPage: () => {
          drawPageChrome(doc, metadata);
        }
      });

      const expteSlug = metadata.expte ? `_${metadata.expte}` : '';
      const fileName = `Anexo_A_Horas_Extras_${metadata.anio}_${metadata.mes}${expteSlug}.pdf`;
      doc.save(fileName);
      return fileName;
    }

    return {
      exportToPdf
    };
  })();

  /* ========================================================================
   * 5. Controlador Principal de la UI e Inicialización
   * ======================================================================== */
  function initApp() {
    const dom = {
      metadataForm: document.getElementById('metadataForm'),
      selectAnio: document.getElementById('selectAnio'),
      selectMes: document.getElementById('selectMes'),
      inputExpte: document.getElementById('inputExpte'),
      expteHelp: document.getElementById('expteHelp'),
      inputSector: document.getElementById('inputSector'),
      inputJefe: document.getElementById('inputJefe'),
      inputObservacionesGlobal: document.getElementById('inputObservacionesGlobal'),
      btnApplyObs: document.getElementById('btnApplyObs'),
      inputFile: document.getElementById('inputFile'),
      fileNameDisplay: document.getElementById('fileNameDisplay'),
      chkRespectSheetFilter: document.getElementById('chkRespectSheetFilter'),
      btnProcessFile: document.getElementById('btnProcessFile'),
      sheetFilterInfoBar: document.getElementById('sheetFilterInfoBar'),
      sheetFilterDetails: document.getElementById('sheetFilterDetails'),
      statusBanner: document.getElementById('statusBanner'),
      metricTotalRows: document.getElementById('metricTotalRows'),
      metricVisibleRows: document.getElementById('metricVisibleRows'),
      metricFilteredRows: document.getElementById('metricFilteredRows'),
      btnToggleAllArt4: document.getElementById('btnToggleAllArt4'),
      btnToggleAllArt7: document.getElementById('btnToggleAllArt7'),
      masterChecks: document.querySelectorAll('.th-master-check'),
      btnExportPdf: document.getElementById('btnExportPdf'),
      tableBody: document.getElementById('tableBody')
    };

    // Poblar selector de años (año actual seleccionado por defecto)
    const currentYear = new Date().getFullYear();
    const minYear = Math.min(2020, currentYear);
    const maxYear = Math.max(2035, currentYear + 5);
    for (let y = minYear; y <= maxYear; y += 1) {
      const opt = document.createElement('option');
      opt.value = String(y);
      opt.textContent = String(y);
      if (y === currentYear) {
        opt.selected = true;
      }
      dom.selectAnio.appendChild(opt);
    }

    // Poblar selector de meses (mes anterior al actual seleccionado por defecto)
    const now = new Date();
    const prevMonthIdx = (now.getMonth() + 11) % 12;
    const defaultMonthName = MONTHS_ES[prevMonthIdx];
    for (const mes of MONTHS_ES) {
      const opt = document.createElement('option');
      opt.value = mes;
      opt.textContent = mes;
      if (mes === defaultMonthName) {
        opt.selected = true;
      }
      dom.selectMes.appendChild(opt);
    }

    function setStatus(message, type = 'info') {
      dom.statusBanner.className = `status-banner is-${type}`;
      dom.statusBanner.textContent = message;
    }

    // Claves de almacenamiento para localStorage
    const STORAGE_KEYS = {
      sector: 'epe_anexo_a_inputSector',
      jefe: 'epe_anexo_a_inputJefe',
      observacionesGlobal: 'epe_anexo_a_inputObservacionesGlobal'
    };

    function saveInputsToLocalStorage() {
      try {
        if (typeof window.localStorage === 'undefined') {
          return;
        }
        window.localStorage.setItem(STORAGE_KEYS.sector, dom.inputSector.value);
        window.localStorage.setItem(STORAGE_KEYS.jefe, dom.inputJefe.value);
        window.localStorage.setItem(STORAGE_KEYS.observacionesGlobal, dom.inputObservacionesGlobal.value);
      } catch (e) {
        // Ignorar posibles restricciones de cuota o modo incógnito restringido
      }
    }

    function loadInputsFromLocalStorage() {
      try {
        if (typeof window.localStorage === 'undefined') {
          return;
        }
        const savedSector = window.localStorage.getItem(STORAGE_KEYS.sector);
        if (savedSector !== null) {
          dom.inputSector.value = savedSector;
          StateManager.setMetadataField('sector', savedSector.trim());
        }

        const savedJefe = window.localStorage.getItem(STORAGE_KEYS.jefe);
        if (savedJefe !== null) {
          dom.inputJefe.value = savedJefe;
          StateManager.setMetadataField('jefe', savedJefe.trim());
        }

        const savedObs = window.localStorage.getItem(STORAGE_KEYS.observacionesGlobal);
        if (savedObs !== null) {
          dom.inputObservacionesGlobal.value = savedObs;
          StateManager.setMetadataField('observacionesGlobal', savedObs);
        }
      } catch (e) {
        // Ignorar posibles errores de acceso
      }
    }

    // Restaurar valores guardados de localStorage si existen
    loadInputsFromLocalStorage();

    function syncMasterCheckboxes() {
      const state = StateManager.getState();
      const rows = state.filteredRowsState;
      const hasRows = rows.length > 0;

      if (dom.btnToggleAllArt4) {
        dom.btnToggleAllArt4.disabled = !hasRows;
      }
      if (dom.btnToggleAllArt7) {
        dom.btnToggleAllArt7.disabled = !hasRows;
      }

      ['incumpleArt4', 'incumpleArt7'].forEach((field) => {
        const master = document.querySelector(`.th-master-check[data-column="${field}"]`);
        if (!master) {
          return;
        }
        if (!hasRows) {
          master.checked = false;
          master.indeterminate = false;
          master.disabled = true;
          return;
        }
        master.disabled = false;
        const checkedCount = rows.filter((r) => Boolean(r[field])).length;
        if (checkedCount === rows.length) {
          master.checked = true;
          master.indeterminate = false;
        } else if (checkedCount === 0) {
          master.checked = false;
          master.indeterminate = false;
        } else {
          master.checked = false;
          master.indeterminate = true;
        }
      });
    }

    TableView.init(dom.tableBody, () => {
      syncMasterCheckboxes();
    });

    function refreshMetricsAndTable() {
      const state = StateManager.getState();
      const hasFile = Boolean(state.fileName);

      dom.metricTotalRows.textContent = String(state.totalRowsRead);
      dom.metricVisibleRows.textContent = String(state.sheetFilterInfo.visibleRowCount);
      dom.metricFilteredRows.textContent = String(state.filteredRowsState.length);
      dom.btnExportPdf.disabled = state.filteredRowsState.length === 0;

      TableView.render(state.filteredRowsState, hasFile);
      syncMasterCheckboxes();
    }

    function executeFilterUpdate() {
      const state = StateManager.getState();
      if (state.rawRows.length === 0) {
        return;
      }

      const filtered = SpreadsheetParser.filterRows(
        state.rawRows,
        {
          respectSheetFilter: state.respectSheetFilter
        },
        state.filteredRowsState
      );

      StateManager.setFilteredRows(filtered);
      refreshMetricsAndTable();

      const filterDesc = state.respectSheetFilter && state.sheetFilterInfo.hasActiveFilter
        ? 'Filtro del archivo + Reglas de Horas (Programadas > 80 / Emergencia > 0 / Total > 120)'
        : 'Reglas de Horas (Programadas > 80 / Emergencia > 0 / Total > 120)';

      setStatus(
        `Filtro actualizado (${filterDesc}): ${filtered.length} filas en tabla.`,
        filtered.length > 0 ? 'success' : 'warning'
      );
    }

    // Evitar envío accidental del formulario con Enter
    dom.metadataForm.addEventListener('submit', (event) => {
      event.preventDefault();
    });

    // Sincronización de metadatos
    dom.selectAnio.addEventListener('change', () => {
      StateManager.setMetadataField('anio', Number(dom.selectAnio.value));
    });

    dom.selectMes.addEventListener('change', () => {
      StateManager.setMetadataField('mes', dom.selectMes.value);
    });

    dom.inputSector.addEventListener('input', () => {
      StateManager.setMetadataField('sector', dom.inputSector.value.trim());
    });

    dom.inputJefe.addEventListener('input', () => {
      StateManager.setMetadataField('jefe', dom.inputJefe.value.trim());
    });

    // Restricción estricta sin espacios en blanco para "Cde. Expte N°"
    dom.inputExpte.addEventListener('keydown', (event) => {
      if (event.key === ' ' || event.code === 'Space') {
        event.preventDefault();
        dom.expteHelp.textContent = 'No se permiten espacios en blanco en Cde. Expte N°.';
        dom.expteHelp.classList.add('is-warning');
      }
    });

    function sanitizeExpteInput() {
      const original = dom.inputExpte.value;
      const cleaned = original.replace(/\s+/g, '');
      if (original !== cleaned) {
        const cursorPos = dom.inputExpte.selectionStart ?? cleaned.length;
        const removedBeforeCursor = (original.slice(0, cursorPos).match(/\s/g) || []).length;
        dom.inputExpte.value = cleaned;
        const nextPos = Math.max(0, cursorPos - removedBeforeCursor);
        dom.inputExpte.setSelectionRange(nextPos, nextPos);
        dom.expteHelp.textContent = 'Se eliminaron espacios en blanco automáticamente.';
        dom.expteHelp.classList.add('is-warning');
      } else {
        dom.expteHelp.textContent = 'Cadena continua sin espacios en blanco.';
        dom.expteHelp.classList.remove('is-warning');
      }
      StateManager.setMetadataField('expte', cleaned);
    }

    dom.inputExpte.addEventListener('input', sanitizeExpteInput);
    dom.inputExpte.addEventListener('paste', () => {
      queueMicrotask(sanitizeExpteInput);
    });

    // Botón de carga masiva de Observaciones
    dom.btnApplyObs.addEventListener('click', () => {
      const state = StateManager.getState();
      const text = dom.inputObservacionesGlobal.value;

      if (state.filteredRowsState.length === 0) {
        setStatus(
          'Primero cargue una planilla con registros filtrados para aplicar observaciones en la tabla.',
          'warning'
        );
        return;
      }

      StateManager.applyGlobalObservaciones(text);
      TableView.syncObservacionesInputs(text);
      setStatus(
        `Observaciones aplicadas a las ${state.filteredRowsState.length} filas de la tabla.`,
        'success'
      );
    });

    // Procesamiento del archivo
    async function processSelectedFile() {
      const file = dom.inputFile.files && dom.inputFile.files[0];
      if (!file) {
        setStatus('Seleccione un archivo .xlsx o .ods antes de procesar.', 'warning');
        return;
      }

      try {
        setStatus(`Procesando "${file.name}"...`, 'info');
        const { totalRowsRead, rawRows, sheetFilterInfo } = await SpreadsheetParser.parseFile(file);

        if (sheetFilterInfo.hasActiveFilter) {
          dom.sheetFilterInfoBar.style.display = 'flex';
          dom.sheetFilterDetails.textContent = `${sheetFilterInfo.visibleRowCount} filas visibles seleccionadas por el filtro de la planilla / ${sheetFilterInfo.hiddenRowCount} filas ocultas ignoradas.`;
        } else {
          dom.sheetFilterInfoBar.style.display = 'none';
        }

        const filtered = SpreadsheetParser.filterRows(rawRows, {
          respectSheetFilter: StateManager.getState().respectSheetFilter
        });

        StateManager.setParsedData({
          fileName: file.name,
          rawRows,
          totalRowsRead,
          sheetFilterInfo,
          filteredRows: filtered
        });

        refreshMetricsAndTable();

        if (sheetFilterInfo.hasActiveFilter && StateManager.getState().respectSheetFilter) {
          setStatus(
            `Filtro detectado en el archivo: se respetó el filtro aplicado (${sheetFilterInfo.visibleRowCount} visibles, ${sheetFilterInfo.hiddenRowCount} ocultas) y se filtraron ${filtered.length} filas que cumplen las condiciones de horas extras (Res. 077/2020: Programadas > 80 / Emergencia > 0 / Total > 120).`,
            'success'
          );
        } else if (filtered.length === 0) {
          setStatus(
            `Archivo "${file.name}" procesado (${totalRowsRead} filas leídas), pero ninguna fila cumplió las condiciones de filtrado.`,
            'warning'
          );
        } else {
          setStatus(
            `Archivo "${file.name}" cargado correctamente: ${filtered.length} de ${totalRowsRead} filas cumplen las condiciones de horas extras (Res. 077/2020: Programadas > 80 / Emergencia > 0 / Total > 120).`,
            'success'
          );
        }
      } catch (error) {
        StateManager.setParsedData({
          fileName: '',
          rawRows: [],
          totalRowsRead: 0,
          sheetFilterInfo: {
            hasActiveFilter: false,
            hasAutoFilter: false,
            hiddenRowCount: 0,
            visibleRowCount: 0
          },
          filteredRows: []
        });
        dom.sheetFilterInfoBar.style.display = 'none';
        refreshMetricsAndTable();
        setStatus(error.message || 'Ocurrió un error al analizar el archivo.', 'error');
      }
    }

    dom.inputFile.addEventListener('change', () => {
      const file = dom.inputFile.files && dom.inputFile.files[0];
      dom.fileNameDisplay.textContent = file ? file.name : 'Ningún archivo seleccionado';
      dom.btnProcessFile.disabled = !file;
      if (file) {
        processSelectedFile();
      }
    });

    dom.btnProcessFile.addEventListener('click', processSelectedFile);

    // Toggle de respeto a los filtros del archivo
    dom.chkRespectSheetFilter.addEventListener('change', () => {
      StateManager.setRespectSheetFilter(dom.chkRespectSheetFilter.checked);
      executeFilterUpdate();
    });

    // Master checkboxes de cabecera para Art. 4° y Art. 7°
    dom.masterChecks.forEach((master) => {
      master.addEventListener('change', (e) => {
        const field = e.target.dataset.column;
        if (!field) {
          return;
        }
        const isChecked = e.target.checked;
        StateManager.setAllColumnCheckboxes(field, isChecked);
        TableView.updateAllCheckboxesInDom(field, isChecked);
        syncMasterCheckboxes();
      });
    });

    // Botones rápidos de alternar todos en toolbar para Art. 4° y Art. 7°
    if (dom.btnToggleAllArt4) {
      dom.btnToggleAllArt4.addEventListener('click', () => {
        const newVal = StateManager.toggleAllColumnCheckboxes('incumpleArt4');
        TableView.updateAllCheckboxesInDom('incumpleArt4', newVal);
        syncMasterCheckboxes();
      });
    }

    if (dom.btnToggleAllArt7) {
      dom.btnToggleAllArt7.addEventListener('click', () => {
        const newVal = StateManager.toggleAllColumnCheckboxes('incumpleArt7');
        TableView.updateAllCheckboxesInDom('incumpleArt7', newVal);
        syncMasterCheckboxes();
      });
    }

    // Exportación a PDF Anexo A
    dom.btnExportPdf.addEventListener('click', () => {
      try {
        // Guardar en localStorage los valores de los controles solicitados
        saveInputsToLocalStorage();

        const state = StateManager.getState();
        if (state.filteredRowsState.length === 0) {
          setStatus('No hay registros filtrados para exportar a PDF.', 'warning');
          return;
        }
        const generatedFileName = PdfExporter.exportToPdf(state);
        setStatus(`PDF generado y descargado correctamente: ${generatedFileName}`, 'success');
      } catch (error) {
        setStatus(`Error al exportar PDF: ${error.message}`, 'error');
      }
    });

    // Verificar carga de librerías offline al inicio
    if (
      typeof window.XLSX === 'undefined' ||
      !window.jspdf ||
      typeof window.jspdf.jsPDF !== 'function'
    ) {
      setStatus(
        'Advertencia: alguna librería local en ./libs/ no se cargó correctamente. Verifique la carpeta libs.',
        'error'
      );
    }
  }

  document.addEventListener('DOMContentLoaded', initApp);
})();
