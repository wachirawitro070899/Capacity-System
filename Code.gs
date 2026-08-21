const SPREADSHEET_ID = '1PvNRR9uekifW7O-cmxh78bOWpYMuYhk3VtK76FGy6BE';
const SOURCE_GID = 1541790103;
const MAX_HEADER_SCAN_ROWS = 40;

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || '').trim().toLowerCase();
  if (action === 'capacity') {
    const data = getCapacityData();
    const callback = String((e && e.parameter && e.parameter.callback) || '').trim();
    const json = JSON.stringify(data);
    if (callback && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
      return ContentService.createTextOutput(callback + '(' + json + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    service: 'Capacity System API',
    message: 'Use ?action=capacity'
  })).setMimeType(ContentService.MimeType.JSON);
}

function getCapacityData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const records = [];
    const machines = [];
    const diagnostics = [];

    ss.getSheets().forEach(sheet => {
      const machine = String(sheet.getName() || '').trim();
      const values = sheet.getDataRange().getDisplayValues();
      if (!machine || !values.length) return;

      const header = detectHeader_(values);
      if (!header) {
        diagnostics.push({ machine, status: 'skipped', reason: 'header_not_found' });
        return;
      }

      let count = 0;
      let lastPartNo = '';
      let lastPartName = '';
      let lastProcess = '';

      for (let r = header.row + 1; r < values.length; r++) {
        const row = values[r] || [];
        if (!row.some(v => clean_(v))) continue;

        let partNo = clean_(getByAlias_(row, header.map, ALIASES.partNo));
        let partName = clean_(getByAlias_(row, header.map, ALIASES.partName));
        let process = clean_(getByAlias_(row, header.map, ALIASES.process));
        let step = clean_(getByAlias_(row, header.map, ALIASES.step));

        if (!partNo && (process || step)) partNo = lastPartNo;
        if (!partName && partNo === lastPartNo) partName = lastPartName;
        if (!process && step) process = lastProcess;

        if (partNo) lastPartNo = partNo;
        if (partName) lastPartName = partName;
        if (process) lastProcess = process;

        if (!partNo && !process && !step) continue;

        const ctRaw = getByAlias_(row, header.map, ALIASES.ct);
        const ct = number_(ctRaw);
        const outputCycle = positive_(number_(getByAlias_(row, header.map, ALIASES.outputCycle)), 1);
        const efficiency = percent_(getByAlias_(row, header.map, ALIASES.efficiency), 100);
        const hoursPerShift = positive_(number_(getByAlias_(row, header.map, ALIASES.hoursPerShift)), 8);
        const shiftsPerDay = positive_(number_(getByAlias_(row, header.map, ALIASES.shiftsPerDay)), 2);

        records.push({
          machine,
          sheetId: sheet.getSheetId(),
          row: r + 1,
          partNo,
          partName,
          process,
          step,
          ct,
          outputCycle,
          efficiency,
          eff: efficiency,
          hoursPerShift,
          shiftsPerDay,
          status: clean_(getByAlias_(row, header.map, ALIASES.status)) || 'Active',
          remark: clean_(getByAlias_(row, header.map, ALIASES.remark))
        });
        count++;
      }

      if (count) machines.push(machine);
      diagnostics.push({ machine, status: count ? 'ok' : 'empty', headerRow: header.row + 1, rows: count });
    });

    return {
      ok: true,
      spreadsheetId: SPREADSHEET_ID,
      sourceGid: SOURCE_GID,
      spreadsheetName: ss.getName(),
      machineSheets: [...new Set(machines)],
      records,
      diagnostics,
      generatedAt: new Date().toISOString()
    };
  } catch (err) {
    return {
      ok: false,
      error: String(err && err.message ? err.message : err),
      records: [],
      machineSheets: [],
      diagnostics: [],
      generatedAt: new Date().toISOString()
    };
  }
}

const ALIASES = {
  partNo: ['Part No.', 'Part No', 'PartNo', 'Part Number', 'Part_Number', 'Part', 'Material No', 'Item No', 'Product No'],
  partName: ['Part Name', 'PartName', 'Part Description', 'Description', 'Product Name', 'Item Name', 'Name'],
  process: ['Process', 'Operation', 'Process Name', 'Operation Name', 'Process/Operation', 'Secondary Process'],
  step: ['Step', 'Process Step', 'Operation Step', 'Step No', 'Step No.', 'Process No', 'Process No.', 'Sequence', 'Seq'],
  ct: ['CT (sec/pc)', 'CT (sec)', 'CT (s)', 'CT', 'Cycle Time', 'Cycle Time (sec)', 'Cycle Time (s)', 'Time (sec)', 'Production Time (min)', 'Production Time (minutes)', '生产工时（分钟）', '生产工时(分钟)'],
  outputCycle: ['Output/Cycle', 'Output per Cycle', 'Output / Cycle', 'Qty/Cycle', 'Output per cycle #1', 'No. of unit #1', 'No of unit #1', 'No. of unit', 'No of unit'],
  efficiency: ['Efficiency %', 'Eff %', 'Efficiency', 'Eff', 'Eff % #1', 'Efficiency #1'],
  hoursPerShift: ['Working Hours/Shift', 'Hours/Shift', 'Hours', 'Daily Working Hrs #1', 'Daily Working Hrs', 'Working Hours'],
  shiftsPerDay: ['Shifts/Day', 'Shift/Day', 'Shifts', 'No. of Shift', 'No of Shift'],
  status: ['Status', 'Active/Inactive'],
  remark: ['Remark', 'Remarks', 'Note', 'Notes', 'Comment']
};

function detectHeader_(values) {
  let best = null;
  const max = Math.min(MAX_HEADER_SCAN_ROWS, values.length);

  for (let r = 0; r < max; r++) {
    const map = {};
    (values[r] || []).forEach((v, c) => {
      const key = norm_(v);
      if (key && map[key] === undefined) map[key] = c;
    });

    const found = {};
    Object.keys(ALIASES).forEach(key => {
      const col = aliasCol_(map, ALIASES[key]);
      if (col !== undefined) found[key] = col;
    });

    let score = 0;
    if (found.partNo !== undefined) score += 10;
    if (found.partName !== undefined) score += 2;
    if (found.process !== undefined) score += 6;
    if (found.step !== undefined) score += 5;
    if (found.ct !== undefined) score += 5;
    if (found.outputCycle !== undefined) score += 1;
    if (found.efficiency !== undefined) score += 1;

    const valid = found.partNo !== undefined || (found.process !== undefined && found.step !== undefined);
    if (valid && (!best || score > best.score)) best = { row: r, map, score };
  }
  return best;
}

function aliasCol_(map, aliases) {
  for (const alias of aliases) {
    const key = norm_(alias);
    if (map[key] !== undefined) return map[key];
  }
  const entries = Object.keys(map);
  for (const alias of aliases) {
    const target = loose_(alias);
    for (const key of entries) {
      if (loose_(key) === target) return map[key];
    }
  }
}

function getByAlias_(row, map, aliases) {
  const col = aliasCol_(map, aliases);
  return col === undefined ? '' : (row[col] ?? '');
}

function norm_(v) {
  return String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').trim().toLowerCase().replace(/[_-]+/g, ' ').replace(/\s+/g, ' ');
}

function loose_(v) {
  return norm_(v).replace(/#\s*\d+/g, '').replace(/[()./%]/g, ' ').replace(/\s+/g, ' ').trim();
}

function clean_(v) { return String(v == null ? '' : v).trim(); }

function number_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const s = clean_(v).replace(/,/g, '').replace(/%/g, '');
  const n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

function positive_(n, fallback) { return n > 0 ? n : fallback; }

function percent_(v, fallback) {
  const s = clean_(v);
  if (!s) return fallback;
  let n = number_(s);
  if (n > 0 && n <= 1 && s.indexOf('%') === -1) n *= 100;
  return n > 0 ? n : fallback;
}
