/* ═══════════════════════════════════════════════════════════
   card-ext.js — Extract Text  (TRIM + MID + FIND)
   • Separate IFNA / IFERROR fields
   • Preset editor with auto-detected delimiter (, | or newline)
   • Live preview  •  Explain Formula  •  Export All
   ═══════════════════════════════════════════════════════════ */
'use strict';

const EXT_DEFAULT_PRESETS = ['ACC', 'Safety Button', 'Panic', 'Door', 'Buzzer'];
const EXT_STORAGE_KEY     = 'ahmos_ext_presets';

// ══════════════════════════════════════════════════════════
//  PRESETS
// ══════════════════════════════════════════════════════════

function ext_loadPresets() {
  try {
    const saved = JSON.parse(localStorage.getItem(EXT_STORAGE_KEY));
    return Array.isArray(saved) && saved.length ? saved : [...EXT_DEFAULT_PRESETS];
  } catch(e) { return [...EXT_DEFAULT_PRESETS]; }
}
function ext_savePresets(list) {
  try { localStorage.setItem(EXT_STORAGE_KEY, JSON.stringify(list)); } catch(e) {}
}

/** DOM-property approach — safe for spaces, &, quotes, any characters */
function ext_renderDropdown() {
  const sel = document.getElementById('ext-searchvalue-dropdown');
  if (!sel) return;
  const prev = sel.value;
  sel.innerHTML = '';
  const add = (val, label) => {
    const o = document.createElement('option');
    o.value = val; o.textContent = label; sel.appendChild(o);
  };
  add('',    '— Select a preset —');
  add('ALL', '🔥 ALL — one formula per preset');
  ext_loadPresets().forEach(p => add(p, p));
  if ([...sel.options].some(o => o.value === prev)) sel.value = prev;
}

function ext_openPresetEditor() {
  const current = ext_loadPresets().join('\n');
  const result = prompt(
    'Enter one preset per line — OR separate with commas or pipes (|).\n' +
    'Multi-word values like "Weight Sensor" are fully supported.\n' +
    'Blank entries are ignored.\n\n' +
    'Examples:\n  ACC, Safety Button, Panic\n  ACC | Weight Sensor | Door\n  (one per line)',
    current
  );
  if (result === null) return;
  // Auto-detect delimiter: if result has newlines use newline, else try | then ,
  let parts;
  // Priority: newline > pipe > comma.
  // Comma-split is always valid because the prompt explicitly shows it as a delimiter option.
  if      (result.includes('\n')) parts = result.split('\n');
  else if (result.includes('|'))  parts = result.split('|');
  else if (result.includes(','))  parts = result.split(',');
  else                            parts = [result]; // single value, no delimiter
  const updated = parts.map(s => s.trim()).filter(Boolean);
  if (!updated.length) { alert('Preset list cannot be empty.'); return; }
  ext_savePresets(updated);
  ext_renderDropdown();
}

// ══════════════════════════════════════════════════════════
//  PURE FORMULA BUILDER
// ══════════════════════════════════════════════════════════

function ext_buildFormulas(cell, sep, searchValues, ifnaVal, ierrVal, inclTerm) {
  const sepEsc   = sep.replace(/"/g, '""');
  const naText   = (ifnaVal  || '').replace(/"/g, '""');
  const errText  = (ierrVal  || '').replace(/"/g, '""');

  const buildOne = sv => {
    const svEsc = sv.replace(/"/g, '""');
    let core;
    if (inclTerm) {
      core =
        `TRIM(MID(${cell},` +
          `FIND("${svEsc}",${cell}),` +
          `FIND("${sepEsc}",${cell},FIND("${svEsc}",${cell}))` +
          `-FIND("${svEsc}",${cell})))`;
    } else {
      const slen = sv.length;
      core =
        `TRIM(MID(${cell},` +
          `FIND("${svEsc}",${cell})+${slen},` +
          `FIND("${sepEsc}",${cell},FIND("${svEsc}",${cell})+${slen})` +
          `-(FIND("${svEsc}",${cell})+${slen})))`;
    }
    // Correct order: IFNA wraps core first (catches #N/A),
    // then IFERROR wraps outside (catches #VALUE! and all other errors).
    // Result: IFERROR(IFNA(core, na_val), err_val)
    let f = core;
    if (ifnaVal)  f = `IFNA(${f},"${naText}")`;
    if (ierrVal)  f = `IFERROR(${f},"${errText}")`;
    return `=${f}`;
  };

  const lines = [];
  searchValues.forEach((sv, i) => {
    if (searchValues.length > 1) lines.push(`// "${sv}"`);
    lines.push(buildOne(sv));
    if (searchValues.length > 1 && i < searchValues.length - 1) lines.push('');
  });
  return lines.join('\n');
}

// ══════════════════════════════════════════════════════════
//  LIVE PREVIEW
// ══════════════════════════════════════════════════════════

function ext_previewBuildFn() {
  const cell     = normaliseCellRef(document.getElementById('ext-celladdress')?.value || '');
  const sep      = (document.getElementById('ext-separator')?.value || '').trim();
  const dropdown = document.getElementById('ext-searchvalue-dropdown')?.value || '';
  const manual   = (document.getElementById('ext-searchvalue-manual')?.value || '').trim();
  const ifnaVal  = (document.getElementById('ext-ifna')?.value  || '').trim();
  const ierrVal  = (document.getElementById('ext-iferror')?.value || '').trim();
  const inclTerm = document.getElementById('ext-include-searchterm')?.checked ?? true;

  if (!cell || !/^[A-Z]{1,3}\d{1,7}$/.test(cell)) return null;
  if (!sep) return null;

  let sv;
  if      (manual)           sv = [manual];
  else if (dropdown === 'ALL') { const all = ext_loadPresets(); sv = [all[0]]; }
  else if (dropdown)          sv = [dropdown];
  else                        return null;

  return ext_buildFormulas(cell, sep, sv, ifnaVal, ierrVal, inclTerm);
}

// ══════════════════════════════════════════════════════════
//  EXPLANATION
// ══════════════════════════════════════════════════════════

function ext_explainBuildFn() {
  const inclTerm = document.getElementById('ext-include-searchterm')?.checked ?? true;
  // Order matches the formula from outside-in: IFERROR wraps IFNA wraps TRIM(MID(FIND,FIND))
  const funcs = ['IFERROR','IFNA','TRIM','MID','FIND'];
  const notes = [inclTerm
    ? '📌 <strong>Include term ON:</strong> MID starts <em>at</em> the search term position — "Weight Sensor 42kg" with term "Weight Sensor " returns "Weight Sensor 42kg" (up to the separator).'
    : '📌 <strong>Include term OFF:</strong> MID skips past the search term by adding its character length to the start position — "Weight Sensor 42kg" returns only "42kg".',
  ];
  return buildExplainHTML(funcs,
    'Finds a <strong>search term</strong> inside a cell and extracts the text ' +
    (inclTerm ? 'starting from that term' : 'immediately after that term') +
    ' up to a separator character.',
    notes);
}

// ══════════════════════════════════════════════════════════
//  RESET
// ══════════════════════════════════════════════════════════

function ext_reset() {
  resetFormByPrefix('ext', () => {
    document.getElementById('ext-separator').value            = ',';
    document.getElementById('ext-include-searchterm').checked = true;
    document.getElementById('ext-searchvalue-dropdown').selectedIndex = 0;
  });
  const pr = document.getElementById('ext-live-preview');
  if (pr) { pr.textContent = '— fill in the fields above to see a live preview —'; pr.className = 'live-preview-box hint'; }
  const panel = document.getElementById('ext-explain-panel');
  if (panel) { panel.classList.remove('open'); panel.innerHTML = ''; }
  const btn = document.getElementById('ext-explain-btn');
  if (btn) btn.textContent = '💡 Explain Formula';
}

// ══════════════════════════════════════════════════════════
//  INIT
// ══════════════════════════════════════════════════════════

document.addEventListener('DOMContentLoaded', () => {
  ext_renderDropdown();
  ['ext-celladdress','ext-separator','ext-searchvalue-manual','ext-ifna','ext-iferror']
    .forEach(id => document.getElementById(id)
      ?.addEventListener('input', () => schedulePreview('ext', ext_previewBuildFn)));
  ['ext-searchvalue-dropdown','ext-include-searchterm']
    .forEach(id => document.getElementById(id)
      ?.addEventListener('change', () => schedulePreview('ext', ext_previewBuildFn)));
});

// ══════════════════════════════════════════════════════════
//  MAIN GENERATOR
// ══════════════════════════════════════════════════════════

function generateExtractText() {
  const cellRaw  = document.getElementById('ext-celladdress').value.trim();
  const sep      = document.getElementById('ext-separator').value.trim();
  const dropdown = document.getElementById('ext-searchvalue-dropdown').value;
  const manual   = document.getElementById('ext-searchvalue-manual').value.trim();
  const ifnaVal  = document.getElementById('ext-ifna').value.trim();
  const ierrVal  = document.getElementById('ext-iferror').value.trim();
  const inclTerm = document.getElementById('ext-include-searchterm').checked;

  const cell = normaliseCellRef(cellRaw);
  if (!cell)                              { alert('Please enter the cell address (e.g., E2).'); return; }
  if (!/^[A-Z]{1,3}\d{1,7}$/.test(cell)) { alert(`"${cellRaw}" is not a valid cell address.\nExamples: E2  AB10  C5`); return; }
  if (!sep)                               { alert('Please enter a separator character.'); return; }

  let searchValues = [];
  if      (manual)            searchValues = [manual];
  else if (dropdown === 'ALL') searchValues = ext_loadPresets();
  else if (dropdown)          searchValues = [dropdown];
  else { alert('Select a preset or type a manual search value.'); return; }

  const output = ext_buildFormulas(cell, sep, searchValues, ifnaVal, ierrVal, inclTerm);
  document.getElementById('ext-out').value = output;
  document.getElementById('ext-output-section').classList.add('visible');

  // Refresh explanation if open
  const panel = document.getElementById('ext-explain-panel');
  if (panel?.classList.contains('open')) panel.innerHTML = ext_explainBuildFn();

  document.getElementById('ext-output-section').scrollIntoView({ behavior:'smooth', block:'nearest' });
}
