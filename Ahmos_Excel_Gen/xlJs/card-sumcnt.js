/* ═══════════════════════════════════════════════════════════
   card-sumcnt.js — SUMIF / SUMIFS / COUNTIF / COUNTIFS
   (V6.2 — full review pass)
   • Force text coercion: &"" on BOTH criteria range AND criteria value
   • Event delegation on dynamic criteria containers (SUMIFS / COUNTIFS)
   • Counter reset on _resetCriteria
   • Remove button triggers preview
   • Live preview  •  Explain  •  Export All
   ═══════════════════════════════════════════════════════════ */
'use strict';

// ══════════════════════════════════════════════════════════
//  CRITERIA GROUP HELPERS
// ══════════════════════════════════════════════════════════

function _getCriteriaGroups(prefix, sheetPath, lock, addStr) {
  const ranges = document.querySelectorAll(`.${prefix}-crit-range`);
  const cells  = document.querySelectorAll(`.${prefix}-crit-cell`);
  const lc     = lock ? '$' : '';
  const groups = [];
  ranges.forEach((rangeEl, i) => {
    const rRaw = rangeEl.value.trim().toUpperCase();
    const cRaw = (cells[i]?.value || '').trim();
    if (!rRaw || !cRaw) return;
    const rc = parseRange(rRaw);
    // Always full-column for SUM/COUNT functions
    const rangeRef    = `${sheetPath}${lc}${rc.startCol}:${lc}${rc.endCol}`;
    const cellVal     = resolveMatchValue(cRaw);
    // Both sides: range coerced with &"", cell-ref criteria also coerced
    const modernRange = addStr ? `${rangeRef}&""` : rangeRef;
    const modernCrit  = (addStr && cellVal.type === 'cell')
      ? `${cellVal.formula}&""` : cellVal.formula;
    groups.push({ rangeRef, modernRange, cellVal, modernCrit });
  });
  return groups;
}

// Counter so "Add Criteria" labels increment correctly; reset on _resetCriteria
const _critCounters = {};

function _addCriteriaGroup(prefix, n, previewFn) {
  const ctr = document.getElementById(`${prefix}-criteria-container`);
  if (!ctr) return;
  const div = document.createElement('div');
  div.className = 'criteria-group';
  // Remove button calls schedulePreview after removal via event on the container
  div.innerHTML = `
    <button class="remove-criteria-btn" onclick="this.closest('.criteria-group').remove();schedulePreview('${prefix}',${prefix}_previewBuildFn)">✕ Remove</button>
    <div class="g2">
      <div class="input-group">
        <label>Criteria Range ${n}</label>
        <input type="text" class="${prefix}-crit-range" placeholder="e.g., B:B">
      </div>
      <div class="input-group">
        <label>Criteria Cell / Value ${n}</label>
        <input type="text" class="${prefix}-crit-cell" placeholder="e.g., B2">
      </div>
    </div>`;
  ctr.appendChild(div);
}

function _resetCriteria(prefix) {
  _critCounters[prefix] = 1;             // ← reset counter so next "Add" gives label 2
  const ctr = document.getElementById(`${prefix}-criteria-container`);
  if (!ctr) return;
  ctr.innerHTML = `
    <div class="criteria-group">
      <div class="g2">
        <div class="input-group">
          <label>Criteria Range 1</label>
          <input type="text" class="${prefix}-crit-range" placeholder="e.g., A:A">
        </div>
        <div class="input-group">
          <label>Criteria Cell / Value 1</label>
          <input type="text" class="${prefix}-crit-cell" placeholder="e.g., A2">
        </div>
      </div>
    </div>`;
}

function _bumpCriteria(prefix) {
  _critCounters[prefix] = (_critCounters[prefix] || 1) + 1;
  _addCriteriaGroup(prefix, _critCounters[prefix]);
}

/** Event delegation on a criteria container — fires preview for any input inside */
function _initCriteriaDelegate(prefix, buildFn) {
  const ctr = document.getElementById(`${prefix}-criteria-container`);
  if (ctr) ctr.addEventListener('input', () => schedulePreview(prefix, buildFn));
}

// ══════════════════════════════════════════════════════════
//  SHARED EXPLAIN
// ══════════════════════════════════════════════════════════

function _sumcntExplain(funcs, intro, addStr) {
  const notes = [
    addStr
      ? '📝 <strong>Force text ON — both sides:</strong> Both the criteria range and the criteria value are coerced with <code>&amp;""</code>. Prevents missed matches when your data has numbers stored as text (or vice versa).'
      : '📝 <strong>Force text OFF:</strong> No coercion applied. If you get 0 when you expect a result, your criteria type may not match your data — turn Force text ON.',
  ];
  return buildExplainHTML(funcs, intro, notes);
}

// ══════════════════════════════════════════════════════════
//  SUMIF
// ══════════════════════════════════════════════════════════

const SIF_TABS = ['modern','legacy'];
function sif_switchTab(tab) { switchTab('sif', tab, SIF_TABS); }
function sif_reset() {
  resetFormByPrefix('sif');
  const pr = document.getElementById('sif-live-preview');
  if (pr) { pr.textContent = '— fill in the fields above to see a live preview —'; pr.className = 'live-preview-box hint'; }
  const panel = document.getElementById('sif-explain-panel');
  if (panel) { panel.classList.remove('open'); panel.innerHTML = ''; }
  const btn = document.getElementById('sif-explain-btn');
  if (btn) btn.textContent = '💡 Explain Formula';
}

function sif_buildAll() {
  const filePath  = document.getElementById('sif-filepath').value.trim();
  const sheetName = document.getElementById('sif-sheetname').value.trim();
  const rangeRaw  = document.getElementById('sif-range').value.trim().toUpperCase();
  const critRaw   = document.getElementById('sif-criteria').value.trim();
  const sumRaw    = document.getElementById('sif-sumrange').value.trim().toUpperCase();
  const ierrVal   = document.getElementById('sif-error').value.trim();
  const lock      = document.getElementById('sif-lock').checked;
  const addStr    = document.getElementById('sif-addstring').checked;

  if (!rangeRaw || !critRaw || !sumRaw) return null;
  const rv = validateRange(rangeRaw, 'Criteria Range'); if (!rv.ok) return null;
  const sv = validateRange(sumRaw,   'Sum Range');      if (!sv.ok) return null;
  const sp = buildSheetPath(filePath, sheetName);       if (sp === null) return null;

  const lc  = lock ? '$' : '';
  const rc  = parseRange(rangeRaw);
  const sc  = parseRange(sumRaw);
  const cv  = resolveMatchValue(critRaw);
  const err = (ierrVal || '0').replace(/"/g, '""');

  const rangeRef = `${sp}${lc}${rc.startCol}:${lc}${rc.endCol}`;
  const sumRef   = `${sp}${lc}${sc.startCol}:${lc}${sc.endCol}`;
  const mRange   = addStr ? `${rangeRef}&""` : rangeRef;
  const mCrit    = (addStr && cv.type === 'cell') ? `${cv.formula}&""` : cv.formula;

  return {
    modern : `=IFERROR(SUMIF(${mRange},${mCrit},${sumRef}),"${err}")`,
    legacy : `=IFERROR(SUMIF(${rangeRef},${cv.formula},${sumRef}),"${err}")`,
    addStr,
  };
}

function sif_previewBuildFn() {
  try { const r = sif_buildAll(); return r?.modern || null; } catch(e) { return null; }
}
function sif_explainBuildFn() {
  return _sumcntExplain(['IFERROR','SUMIF'],
    'Sums values in a sum range for every row where the criteria range matches the given value.',
    document.getElementById('sif-addstring')?.checked);
}

document.addEventListener('DOMContentLoaded', () => {
  ['sif-filepath','sif-sheetname','sif-range','sif-criteria','sif-sumrange','sif-error']
    .forEach(id => document.getElementById(id)
      ?.addEventListener('input', () => schedulePreview('sif', sif_previewBuildFn)));
  ['sif-lock','sif-addstring'].forEach(id =>
    document.getElementById(id)
      ?.addEventListener('change', () => schedulePreview('sif', sif_previewBuildFn)));
});

function generateSumIf() {
  const rangeRaw = document.getElementById('sif-range').value.trim();
  const critRaw  = document.getElementById('sif-criteria').value.trim();
  const sumRaw   = document.getElementById('sif-sumrange').value.trim();
  if (!rangeRaw || !critRaw || !sumRaw) {
    alert('Please fill in: Criteria Range, Criteria Value, and Sum Range.'); return;
  }
  const rv = validateRange(rangeRaw, 'Criteria Range'); if (!rv.ok) { alert(rv.error); return; }
  const sv = validateRange(sumRaw,   'Sum Range');      if (!sv.ok) { alert(sv.error); return; }

  const r = sif_buildAll();
  if (!r) { alert('Could not build formula — check your inputs.'); return; }

  document.getElementById('sif-out-modern').value        = r.modern;
  document.getElementById('sif-out-legacy').value        = r.legacy;
  document.getElementById('sif-desc-modern').textContent = r.addStr
    ? 'Force text ON — both sides: criteria range &"" + criteria value &"". Excel 2003+.'
    : 'Standard SUMIF, no coercion. Excel 2003+.';
  document.getElementById('sif-desc-legacy').textContent = 'Standard SUMIF, no coercion. Excel 2003+.';
  document.getElementById('sif-output-section').classList.add('visible');
  sif_switchTab('modern');
  const panel = document.getElementById('sif-explain-panel');
  if (panel?.classList.contains('open')) panel.innerHTML = sif_explainBuildFn();
  document.getElementById('sif-output-section').scrollIntoView({ behavior:'smooth', block:'nearest' });
}

// ══════════════════════════════════════════════════════════
//  SUMIFS
// ══════════════════════════════════════════════════════════

const SFS_TABS = ['modern','legacy'];
function sfs_switchTab(tab) { switchTab('sfs', tab, SFS_TABS); }
function sfs_addCriteria()  { _bumpCriteria('sfs'); }
function sfs_reset() {
  resetFormByPrefix('sfs', () => _resetCriteria('sfs'));
  const pr = document.getElementById('sfs-live-preview');
  if (pr) { pr.textContent = '— fill in the fields above to see a live preview —'; pr.className = 'live-preview-box hint'; }
  const panel = document.getElementById('sfs-explain-panel');
  if (panel) { panel.classList.remove('open'); panel.innerHTML = ''; }
  const btn = document.getElementById('sfs-explain-btn');
  if (btn) btn.textContent = '💡 Explain Formula';
}

function sfs_buildAll() {
  const filePath  = document.getElementById('sfs-filepath').value.trim();
  const sheetName = document.getElementById('sfs-sheetname').value.trim();
  const sumRaw    = document.getElementById('sfs-sumrange').value.trim().toUpperCase();
  const ierrVal   = document.getElementById('sfs-error').value.trim();
  const lock      = document.getElementById('sfs-lock').checked;
  const addStr    = document.getElementById('sfs-addstring').checked;

  if (!sumRaw) return null;
  const sv = validateRange(sumRaw, 'Sum Range'); if (!sv.ok) return null;
  const sp = buildSheetPath(filePath, sheetName); if (sp === null) return null;

  const lc     = lock ? '$' : '';
  const sc     = parseRange(sumRaw);
  const sumRef = `${sp}${lc}${sc.startCol}:${lc}${sc.endCol}`;
  const err    = (ierrVal || '0').replace(/"/g, '""');
  const groups = _getCriteriaGroups('sfs', sp, lock, addStr);
  if (!groups.length) return null;

  const modernPairs = groups.map(g => `${g.modernRange},${g.modernCrit}`).join(',');
  const legacyPairs = groups.map(g => `${g.rangeRef},${g.cellVal.formula}`).join(',');

  return {
    modern : `=IFERROR(SUMIFS(${sumRef},${modernPairs}),"${err}")`,
    legacy : `=IFERROR(SUMIFS(${sumRef},${legacyPairs}),"${err}")`,
    addStr,
  };
}

function sfs_previewBuildFn() {
  try { const r = sfs_buildAll(); return r?.modern || null; } catch(e) { return null; }
}
function sfs_explainBuildFn() {
  return _sumcntExplain(['IFERROR','SUMIFS'],
    'Sums values where ALL criteria conditions are met simultaneously (AND logic across multiple criteria ranges). Excel 2007+.',
    document.getElementById('sfs-addstring')?.checked);
}

document.addEventListener('DOMContentLoaded', () => {
  ['sfs-filepath','sfs-sheetname','sfs-sumrange','sfs-error']
    .forEach(id => document.getElementById(id)
      ?.addEventListener('input', () => schedulePreview('sfs', sfs_previewBuildFn)));
  ['sfs-lock','sfs-addstring'].forEach(id =>
    document.getElementById(id)
      ?.addEventListener('change', () => schedulePreview('sfs', sfs_previewBuildFn)));
  _initCriteriaDelegate('sfs', sfs_previewBuildFn);  // ← event delegation for dynamic inputs
});

function generateSumIfs() {
  const sumRaw    = document.getElementById('sfs-sumrange').value.trim();
  const filePath  = document.getElementById('sfs-filepath').value.trim();
  const sheetName = document.getElementById('sfs-sheetname').value.trim();
  if (!sumRaw) { alert('Please fill in the Sum Range.'); return; }
  const sv = validateRange(sumRaw, 'Sum Range'); if (!sv.ok) { alert(sv.error); return; }
  const sp = buildSheetPath(filePath, sheetName); if (sp === null) { alert(processFilePath(filePath).error); return; }

  const r = sfs_buildAll();
  if (!r) { alert('Fill in at least one Criteria Range and Value.'); return; }

  document.getElementById('sfs-out-modern').value        = r.modern;
  document.getElementById('sfs-out-legacy').value        = r.legacy;
  document.getElementById('sfs-desc-modern').textContent = r.addStr
    ? 'Force text ON — both sides: criteria ranges &"" + criteria values &"". Excel 2007+.'
    : 'Standard SUMIFS, no coercion. Excel 2007+.';
  document.getElementById('sfs-desc-legacy').textContent = 'Standard SUMIFS, no coercion. Excel 2007+.';
  document.getElementById('sfs-output-section').classList.add('visible');
  sfs_switchTab('modern');
  const panel = document.getElementById('sfs-explain-panel');
  if (panel?.classList.contains('open')) panel.innerHTML = sfs_explainBuildFn();
  document.getElementById('sfs-output-section').scrollIntoView({ behavior:'smooth', block:'nearest' });
}

// ══════════════════════════════════════════════════════════
//  COUNTIF
// ══════════════════════════════════════════════════════════

const CIF_TABS = ['modern','legacy'];
function cif_switchTab(tab) { switchTab('cif', tab, CIF_TABS); }
function cif_reset() {
  resetFormByPrefix('cif');
  const pr = document.getElementById('cif-live-preview');
  if (pr) { pr.textContent = '— fill in the fields above to see a live preview —'; pr.className = 'live-preview-box hint'; }
  const panel = document.getElementById('cif-explain-panel');
  if (panel) { panel.classList.remove('open'); panel.innerHTML = ''; }
  const btn = document.getElementById('cif-explain-btn');
  if (btn) btn.textContent = '💡 Explain Formula';
}

function cif_buildAll() {
  const filePath  = document.getElementById('cif-filepath').value.trim();
  const sheetName = document.getElementById('cif-sheetname').value.trim();
  const rangeRaw  = document.getElementById('cif-range').value.trim().toUpperCase();
  const critRaw   = document.getElementById('cif-criteria').value.trim();
  const ierrVal   = document.getElementById('cif-error').value.trim();
  const lock      = document.getElementById('cif-lock').checked;
  const addStr    = document.getElementById('cif-addstring').checked;

  if (!rangeRaw || !critRaw) return null;
  const rv = validateRange(rangeRaw, 'Range to Count'); if (!rv.ok) return null;
  const sp = buildSheetPath(filePath, sheetName);       if (sp === null) return null;

  const lc  = lock ? '$' : '';
  const rc  = parseRange(rangeRaw);
  const cv  = resolveMatchValue(critRaw);
  const err = (ierrVal || '0').replace(/"/g, '""');

  const rangeRef = `${sp}${lc}${rc.startCol}:${lc}${rc.endCol}`;
  const mRange   = addStr ? `${rangeRef}&""` : rangeRef;
  const mCrit    = (addStr && cv.type === 'cell') ? `${cv.formula}&""` : cv.formula;

  return {
    modern : `=IFERROR(COUNTIF(${mRange},${mCrit}),"${err}")`,
    legacy : `=IFERROR(COUNTIF(${rangeRef},${cv.formula}),"${err}")`,
    addStr,
  };
}

function cif_previewBuildFn() {
  try { const r = cif_buildAll(); return r?.modern || null; } catch(e) { return null; }
}
function cif_explainBuildFn() {
  return _sumcntExplain(['IFERROR','COUNTIF'],
    'Counts cells in a range that match a single condition. Returns 0 when nothing matches.',
    document.getElementById('cif-addstring')?.checked);
}

document.addEventListener('DOMContentLoaded', () => {
  ['cif-filepath','cif-sheetname','cif-range','cif-criteria','cif-error']
    .forEach(id => document.getElementById(id)
      ?.addEventListener('input', () => schedulePreview('cif', cif_previewBuildFn)));
  ['cif-lock','cif-addstring'].forEach(id =>
    document.getElementById(id)
      ?.addEventListener('change', () => schedulePreview('cif', cif_previewBuildFn)));
});

function generateCountIf() {
  const rangeRaw = document.getElementById('cif-range').value.trim();
  const critRaw  = document.getElementById('cif-criteria').value.trim();
  if (!rangeRaw || !critRaw) {
    alert('Please fill in: Range to Count and Criteria Value.'); return;
  }
  const rv = validateRange(rangeRaw, 'Range to Count'); if (!rv.ok) { alert(rv.error); return; }

  const r = cif_buildAll();
  if (!r) { alert('Could not build formula — check your inputs.'); return; }

  document.getElementById('cif-out-modern').value        = r.modern;
  document.getElementById('cif-out-legacy').value        = r.legacy;
  document.getElementById('cif-desc-modern').textContent = r.addStr
    ? 'Force text ON — both sides: range &"" + criteria value &"". Excel 2003+.'
    : 'Standard COUNTIF, no coercion. Excel 2003+.';
  document.getElementById('cif-desc-legacy').textContent = 'Standard COUNTIF, no coercion. Excel 2003+.';
  document.getElementById('cif-output-section').classList.add('visible');
  cif_switchTab('modern');
  const panel = document.getElementById('cif-explain-panel');
  if (panel?.classList.contains('open')) panel.innerHTML = cif_explainBuildFn();
  document.getElementById('cif-output-section').scrollIntoView({ behavior:'smooth', block:'nearest' });
}

// ══════════════════════════════════════════════════════════
//  COUNTIFS
// ══════════════════════════════════════════════════════════

const CFS_TABS = ['modern','legacy'];
function cfs_switchTab(tab) { switchTab('cfs', tab, CFS_TABS); }
function cfs_addCriteria()  { _bumpCriteria('cfs'); }
function cfs_reset() {
  resetFormByPrefix('cfs', () => _resetCriteria('cfs'));
  const pr = document.getElementById('cfs-live-preview');
  if (pr) { pr.textContent = '— fill in the fields above to see a live preview —'; pr.className = 'live-preview-box hint'; }
  const panel = document.getElementById('cfs-explain-panel');
  if (panel) { panel.classList.remove('open'); panel.innerHTML = ''; }
  const btn = document.getElementById('cfs-explain-btn');
  if (btn) btn.textContent = '💡 Explain Formula';
}

function cfs_buildAll() {
  const filePath  = document.getElementById('cfs-filepath').value.trim();
  const sheetName = document.getElementById('cfs-sheetname').value.trim();
  const ierrVal   = document.getElementById('cfs-error').value.trim();
  const lock      = document.getElementById('cfs-lock').checked;
  const addStr    = document.getElementById('cfs-addstring').checked;

  const sp = buildSheetPath(filePath, sheetName); if (sp === null) return null;
  const err    = (ierrVal || '0').replace(/"/g, '""');
  const groups = _getCriteriaGroups('cfs', sp, lock, addStr);
  if (!groups.length) return null;

  const modernPairs = groups.map(g => `${g.modernRange},${g.modernCrit}`).join(',');
  const legacyPairs = groups.map(g => `${g.rangeRef},${g.cellVal.formula}`).join(',');

  return {
    modern : `=IFERROR(COUNTIFS(${modernPairs}),"${err}")`,
    legacy : `=IFERROR(COUNTIFS(${legacyPairs}),"${err}")`,
    addStr,
  };
}

function cfs_previewBuildFn() {
  try { const r = cfs_buildAll(); return r?.modern || null; } catch(e) { return null; }
}
function cfs_explainBuildFn() {
  return _sumcntExplain(['IFERROR','COUNTIFS'],
    'Counts rows where ALL criteria conditions are met simultaneously (AND logic). Excel 2007+.',
    document.getElementById('cfs-addstring')?.checked);
}

document.addEventListener('DOMContentLoaded', () => {
  ['cfs-filepath','cfs-sheetname','cfs-error']
    .forEach(id => document.getElementById(id)
      ?.addEventListener('input', () => schedulePreview('cfs', cfs_previewBuildFn)));
  ['cfs-lock','cfs-addstring'].forEach(id =>
    document.getElementById(id)
      ?.addEventListener('change', () => schedulePreview('cfs', cfs_previewBuildFn)));
  _initCriteriaDelegate('cfs', cfs_previewBuildFn);  // ← event delegation for dynamic inputs
});

function generateCountIfs() {
  const filePath  = document.getElementById('cfs-filepath').value.trim();
  const sheetName = document.getElementById('cfs-sheetname').value.trim();
  const sp = buildSheetPath(filePath, sheetName);
  if (sp === null) { alert(processFilePath(filePath).error); return; }

  const r = cfs_buildAll();
  if (!r) { alert('Fill in at least one Criteria Range and Value.'); return; }

  document.getElementById('cfs-out-modern').value        = r.modern;
  document.getElementById('cfs-out-legacy').value        = r.legacy;
  document.getElementById('cfs-desc-modern').textContent = r.addStr
    ? 'Force text ON — both sides: criteria ranges &"" + criteria values &"". Excel 2007+.'
    : 'Standard COUNTIFS, no coercion. Excel 2007+.';
  document.getElementById('cfs-desc-legacy').textContent = 'Standard COUNTIFS, no coercion. Excel 2007+.';
  document.getElementById('cfs-output-section').classList.add('visible');
  cfs_switchTab('modern');
  const panel = document.getElementById('cfs-explain-panel');
  if (panel?.classList.contains('open')) panel.innerHTML = cfs_explainBuildFn();
  document.getElementById('cfs-output-section').scrollIntoView({ behavior:'smooth', block:'nearest' });
}
