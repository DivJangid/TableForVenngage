(function () {
  'use strict';

  // ===== Inline styles for the generated WordPress table =====
  // All properties use !important to prevent WordPress theme CSS from overriding
  const TABLE_STYLES = {
    wrapper: [
      'overflow: hidden !important',
      'border-radius: 10px !important',
      'border: 1px solid #D1D5DB !important',
      'margin: 0 !important',
      'padding: 0 !important',
      'display: flex !important',
      'flex-direction: column !important',
      'margin-block-start: 0 !important',
      'margin-block-end: 0 !important'
    ].join('; '),

    table: [
      'border-collapse: collapse !important',
      'border-spacing: 0 !important',
      'width: 100% !important',
      'max-width: 100% !important',
      'font-family: "Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important',
      'font-size: 14px !important',
      'line-height: 1.5 !important',
      'color: #1a1a2e !important',
      'margin: 0 !important',
      'padding: 0 !important',
      'min-width: 320px !important',
      'background: none !important',
      'border: none !important',
      'margin-bottom: 0 !important',
      'margin-top: 0 !important'
    ].join('; '),

    headerRow: 'background-color: #0B3D91 !important; background: #0B3D91 !important',

    headerCell: [
      'padding: 14px 16px !important',
      'text-align: left !important',
      'font-weight: 600 !important',
      'color: #FFFFFF !important',
      'border-right: 1px solid rgba(255, 255, 255, 0.3) !important',
      'border-bottom: 2px solid #094080 !important',
      'border-top: none !important',
      'border-left: none !important',
      'font-size: 15px !important',
      'text-transform: none !important',
      'letter-spacing: 0 !important',
      'word-wrap: break-word !important',
      'overflow-wrap: break-word !important',
      'background-color: #0B3D91 !important',
      'background: #0B3D91 !important'
    ].join('; '),

    headerCellLast: [
      'padding: 14px 16px !important',
      'text-align: left !important',
      'font-weight: 600 !important',
      'color: #FFFFFF !important',
      'border-right: none !important',
      'border-bottom: 2px solid #094080 !important',
      'border-top: none !important',
      'border-left: none !important',
      'font-size: 15px !important',
      'text-transform: none !important',
      'letter-spacing: 0 !important',
      'word-wrap: break-word !important',
      'overflow-wrap: break-word !important',
      'background-color: #0B3D91 !important',
      'background: #0B3D91 !important'
    ].join('; '),

    bodyRowEven: 'background-color: #FFFFFF !important; background: #FFFFFF !important',

    bodyRowOdd: 'background-color: #F5F8FF !important; background: #F5F8FF !important',

  };

  // Build body cell style dynamically based on position
  function buildBodyCellStyle(isLastCol, isLastRow, isBold) {
    const styles = [
      'padding: 12px 16px !important',
      'text-align: left !important',
      isLastRow ? 'border-bottom: none !important' : 'border-bottom: 1px solid #E5E7EB !important',
      isLastCol ? 'border-right: none !important' : 'border-right: 1px solid #E5E7EB !important',
      'border-top: none !important',
      'border-left: none !important',
      'color: #1a1a2e !important',
      'font-size: 14px !important',
      'word-wrap: break-word !important',
      'overflow-wrap: break-word !important',
      'background: inherit !important'
    ];
    if (isBold) {
      styles.push('font-weight: 600 !important');
    }
    return styles.join('; ');
  }

  // ===== State =====
  const state = {
    rows: 0,
    cols: 0,
    data: [],
    boldCol1: false
  };

  // ===== DOM references =====
  const $ = (id) => document.getElementById(id);

  // ===== Utilities =====
  function escapeHTML(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  function debounce(fn, delay) {
    let timer = null;
    return function (...args) {
      clearTimeout(timer);
      timer = setTimeout(() => fn.apply(this, args), delay);
    };
  }

  function clamp(val, min, max) {
    return Math.min(Math.max(val, min), max);
  }

  // ===== Sync state from editor DOM =====
  function syncStateFromEditor() {
    const table = document.querySelector('#editor-table-wrapper table');
    if (!table) {
      state.data = [];
      state.rows = 0;
      state.cols = 0;
      return;
    }
    state.data = [];
    const allRows = table.querySelectorAll('tr');
    allRows.forEach((tr) => {
      const rowData = [];
      tr.querySelectorAll('th, td').forEach((cell) => {
        rowData.push(cell.textContent.trim());
      });
      state.data.push(rowData);
    });
    state.rows = state.data.length;
    state.cols = state.data.length > 0 ? state.data[0].length : 0;
  }

  // ===== Generate editable table =====
  function generateTable(rows, cols) {
    rows = clamp(rows, 2, 51); // at least 1 header + 1 data row
    cols = clamp(cols, 1, 20);

    const wrapper = $('editor-table-wrapper');
    wrapper.innerHTML = '';

    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');

    // Header row
    const headerTr = document.createElement('tr');
    for (let c = 0; c < cols; c++) {
      const th = document.createElement('th');
      th.contentEditable = 'true';
      th.textContent = 'Header ' + (c + 1);
      headerTr.appendChild(th);
    }
    thead.appendChild(headerTr);

    // Body rows
    for (let r = 1; r < rows; r++) {
      const tr = document.createElement('tr');
      for (let c = 0; c < cols; c++) {
        const td = document.createElement('td');
        td.contentEditable = 'true';
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }

    table.appendChild(thead);
    table.appendChild(tbody);
    wrapper.appendChild(table);

    syncStateFromEditor();
    updatePreview();
  }

  // ===== Populate editor from 2D data array =====
  function populateEditorFromData(data) {
    if (data.length === 0) return;
    const rows = data.length;
    const cols = data[0].length;

    generateTable(rows, cols);

    const table = document.querySelector('#editor-table-wrapper table');
    const allRows = table.querySelectorAll('tr');

    allRows.forEach((tr, r) => {
      const cells = tr.querySelectorAll('th, td');
      cells.forEach((cell, c) => {
        cell.textContent = data[r] && data[r][c] ? data[r][c] : '';
      });
    });

    syncStateFromEditor();
    updatePreview();
  }

  // ===== Add / Delete rows and columns =====
  function addRow() {
    const tbody = document.querySelector('#editor-table-wrapper tbody');
    if (!tbody) return;

    syncStateFromEditor();
    const tr = document.createElement('tr');
    for (let c = 0; c < state.cols; c++) {
      const td = document.createElement('td');
      td.contentEditable = 'true';
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
    updatePreview();
  }

  function addColumn() {
    const table = document.querySelector('#editor-table-wrapper table');
    if (!table) return;

    syncStateFromEditor();
    const allRows = table.querySelectorAll('tr');
    const newColIndex = state.cols + 1;

    allRows.forEach((tr, i) => {
      if (i === 0) {
        const th = document.createElement('th');
        th.contentEditable = 'true';
        th.textContent = 'Header ' + newColIndex;
        tr.appendChild(th);
      } else {
        const td = document.createElement('td');
        td.contentEditable = 'true';
        tr.appendChild(td);
      }
    });

    updatePreview();
  }

  function deleteRow() {
    const tbody = document.querySelector('#editor-table-wrapper tbody');
    if (!tbody) return;

    const bodyRows = tbody.querySelectorAll('tr');
    if (bodyRows.length <= 1) return; // keep at least 1 data row

    tbody.removeChild(bodyRows[bodyRows.length - 1]);
    updatePreview();
  }

  function deleteColumn() {
    const table = document.querySelector('#editor-table-wrapper table');
    if (!table) return;

    syncStateFromEditor();
    if (state.cols <= 1) return; // keep at least 1 column

    const allRows = table.querySelectorAll('tr');
    allRows.forEach((tr) => {
      const cells = tr.querySelectorAll('th, td');
      if (cells.length > 0) {
        tr.removeChild(cells[cells.length - 1]);
      }
    });

    updatePreview();
  }

  // ===== Parse spreadsheet paste (tab-separated) =====
  function parseSpreadsheetPaste(text) {
    if (!text || !text.trim()) return [];

    const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');

    // Remove trailing empty lines
    while (lines.length > 0 && lines[lines.length - 1].trim() === '') {
      lines.pop();
    }

    if (lines.length === 0) return [];

    const data = lines.map((line) => line.split('\t').map((cell) => cell.trim()));

    // Normalize column count to the maximum found
    const maxCols = Math.max(...data.map((row) => row.length));
    data.forEach((row) => {
      while (row.length < maxCols) {
        row.push('');
      }
    });

    return data;
  }

  // ===== Generate inline-styled HTML for WordPress =====
  // Returns minified HTML (no newlines) to prevent WordPress wpautop from injecting <p> tags
  function generateStyledHTML() {
    syncStateFromEditor();

    if (state.data.length === 0 || state.cols === 0) {
      return '<!-- No table data -->';
    }

    const lastCol = state.cols - 1;
    const lastRow = state.rows - 1;
    const parts = [];

    parts.push('<div style="' + TABLE_STYLES.wrapper + '">');
    parts.push('<table style="' + TABLE_STYLES.table + '">');

    // Header
    parts.push('<thead style="margin:0!important;padding:0!important;border:none!important;">');
    parts.push('<tr style="' + TABLE_STYLES.headerRow + '">');
    for (let c = 0; c < state.cols; c++) {
      const text = escapeHTML(state.data[0][c] || '');
      const thStyle = (c === lastCol) ? TABLE_STYLES.headerCellLast : TABLE_STYLES.headerCell;
      parts.push('<th style="' + thStyle + '">' + text + '</th>');
    }
    parts.push('</tr>');
    parts.push('</thead>');

    // Body
    parts.push('<tbody style="margin:0!important;padding:0!important;border:none!important;">');
    for (let r = 1; r < state.rows; r++) {
      const isOdd = (r - 1) % 2 === 1;
      const rowStyle = isOdd ? TABLE_STYLES.bodyRowOdd : TABLE_STYLES.bodyRowEven;
      parts.push('<tr style="' + rowStyle + '">');
      for (let c = 0; c < state.cols; c++) {
        const text = escapeHTML(state.data[r] ? (state.data[r][c] || '') : '');
        const cellStyle = buildBodyCellStyle(c === lastCol, r === lastRow, c === 0 && state.boldCol1);
        parts.push('<td style="' + cellStyle + '">' + text + '</td>');
      }
      parts.push('</tr>');
    }
    parts.push('</tbody>');

    parts.push('</table>');
    parts.push('</div>');

    // Join with no whitespace — prevents WordPress wpautop from injecting <p> tags
    return parts.join('');
  }

  // ===== Update preview and output =====
  function updatePreview() {
    const html = generateStyledHTML();
    $('preview-container').innerHTML = html;
    $('output-code').textContent = html;
  }

  // ===== Copy to clipboard =====
  async function copyToClipboard() {
    const html = $('output-code').textContent;
    const feedback = $('copy-feedback');

    try {
      await navigator.clipboard.writeText(html);
      showCopyFeedback(feedback);
    } catch (err) {
      // Fallback for older browsers or non-HTTPS
      const textarea = document.createElement('textarea');
      textarea.value = html;
      textarea.style.position = 'fixed';
      textarea.style.opacity = '0';
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
      showCopyFeedback(feedback);
    }
  }

  function showCopyFeedback(el) {
    el.textContent = 'Copied!';
    el.classList.add('visible');
    setTimeout(() => {
      el.textContent = '';
      el.classList.remove('visible');
    }, 2000);
  }

  // ===== Event binding =====
  function bindEvents() {
    // Generate table from scratch
    $('btn-generate').addEventListener('click', () => {
      const rows = clamp(parseInt($('input-rows').value, 10) || 3, 1, 50);
      const cols = clamp(parseInt($('input-cols').value, 10) || 3, 1, 20);
      generateTable(rows + 1, cols); // +1 for header row
    });

    // Import pasted spreadsheet data
    $('btn-parse-paste').addEventListener('click', () => {
      const text = $('paste-area').value;
      const data = parseSpreadsheetPaste(text);
      if (data.length > 0) {
        populateEditorFromData(data);
      }
    });

    // Add / delete rows and columns
    $('btn-add-row').addEventListener('click', addRow);
    $('btn-add-col').addEventListener('click', addColumn);
    $('btn-del-row').addEventListener('click', deleteRow);
    $('btn-del-col').addEventListener('click', deleteColumn);

    // Bold column 1 toggle
    $('chk-bold-col1').addEventListener('change', (e) => {
      state.boldCol1 = e.target.checked;
      updatePreview();
    });

    // Copy HTML
    $('btn-copy-html').addEventListener('click', copyToClipboard);

    // Live preview on cell edits (debounced)
    $('editor-table-wrapper').addEventListener('input', debounce(updatePreview, 300));

    // Tab navigation between cells
    $('editor-table-wrapper').addEventListener('keydown', (e) => {
      if (e.key === 'Tab') {
        e.preventDefault();
        const cells = Array.from(
          document.querySelectorAll('#editor-table-wrapper th, #editor-table-wrapper td')
        );
        const currentIndex = cells.indexOf(document.activeElement);
        if (currentIndex === -1) return;
        const nextIndex = e.shiftKey
          ? Math.max(0, currentIndex - 1)
          : Math.min(cells.length - 1, currentIndex + 1);
        cells[nextIndex].focus();
      }
    });
  }

  // ===== Init =====
  function init() {
    bindEvents();
    generateTable(4, 3); // Default: 3 data rows + 1 header, 3 columns
  }

  document.addEventListener('DOMContentLoaded', init);
})();
