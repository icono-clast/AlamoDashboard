// ============================================================
// Alamo MPD Dashboard — Google Apps Script Web App v3
// Role detail table added. Original data pipeline preserved.
// ============================================================
//
// SHEET TABS REQUIRED (same as v2 — nothing new needed):
//   "time"        — Kantata Time export (Name tab), paste from row 1
//   "allocations" — Kantata Allocations export (Name tab), paste from row 1
//   "tasks"       — Kantata Tasks export (Name tab), paste from row 1
//   "config"      — key/value pairs (see below)
//
// OPTIONAL: add a "budget-plan-int" and "budget-plan-lin" tab
//   (paste sheets 44172716 / 44923833 from Alamo_MPD_Budgets_per_WO.xlsx)
//   If present, the role detail table will show budget vs actual vs allocation.
//   If absent, the table still shows actual vs allocation with no budget column.
//
// CONFIG TAB (columns A + B, no header row):
//   int_contract_value    2710422
//   int_labor_budget      1912056
//   lin_contract_value    3589699
//   lin_labor_budget      1402256
//   overhead_multiplier   2.35
//   data_as_of            May 7, 2026
//
// ============================================================

const PROJ_INT = 'Alamo | MPD | VCM Interactive';
const PROJ_LIN = 'Alamo | MPD | VCM Linear';

// Task IDs for subs classification (no project col in tasks export)
const INT_SUB_IDS = ['928973691','928973692','928973693','928973694','928973695'];
const LIN_SUB_IDS = ['937455779','937455780','937455781','937455782','937455783','944222199','948222969'];

// Project timeline bounds for flat budget distribution
const PROJECT_BOUNDS = {
  int: { start: new Date('2025-10-01'), end: new Date('2027-11-30') },
  lin: { start: new Date('2025-10-01'), end: new Date('2028-03-31') }
};

// ── ENTRY POINT ───────────────────────────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify(buildDashboardData()))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── MAIN DATA BUILDER ─────────────────────────────────────────────────────────
const SPREADSHEET_ID = '1OxpBklTrYUxpPfS2BJjG1CWTNoHmzyOA9GesJqyu7To';

function buildDashboardData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const config = readConfig(ss);

  // Cost Actual / Cost Allocated fields from Kantata exports are already loaded
  // (Kantata applies PET overhead internally). Multiplier should be 1 unless
  // the workspace is configured to export raw base costs.
  const mult    = config.overhead_multiplier || 1;
  const asOf    = config.data_as_of || Utilities.formatDate(new Date(), 'America/New_York', 'MMM d, yyyy');
  const todayMs = new Date().getTime();

  // Parse export tabs (identical to v2)
  const timeRows  = readSheetFromMarker(ss, 'time',        'Time Entry: ID');
  const allocRows = readSheetFromMarker(ss, 'allocations', 'Resource Allocation Day: ID');
  const taskRows  = readSheetFromMarker(ss, 'tasks',       'Task: Name');

  // Parse optional budget plan tabs (new in v3)
  const budgetInt = parseBudgetTab(ss, 'budget-plan-int');
  const budgetLin = parseBudgetTab(ss, 'budget-plan-lin');

  // Subs budgets from tasks (identical to v2)
  const subsBudgets = computeSubsBudgets(taskRows);

  // Labor actuals from time tab (identical to v2 — Cost Actual × mult)
  const laborActuals = {
    int: sumLaborActuals(timeRows, PROJ_INT, mult),
    lin: sumLaborActuals(timeRows, PROJ_LIN, mult),
  };

  // Future allocations (identical to v2)
  const futureAlloc = {
    int: sumFutureAlloc(allocRows, PROJ_INT, todayMs, mult),
    lin: sumFutureAlloc(allocRows, PROJ_LIN, todayMs, mult),
  };
  futureAlloc.all = {
    laborCost: futureAlloc.int.laborCost + futureAlloc.lin.laborCost,
    hrs:       futureAlloc.int.hrs       + futureAlloc.lin.hrs,
  };

  // KPI objects (identical to v2)
  function buildKD(key) {
    const laborEAC      = laborActuals[key] + futureAlloc[key].laborCost;
    const totalEAC      = laborEAC + subsBudgets[key];
    const contractValue = Number(config[key + '_contract_value'] || 0);
    const laborBudget   = Number(config[key + '_labor_budget']   || 0);
    const eacMargin     = contractValue - totalEAC;
    const eacMarginPct  = contractValue > 0
      ? Math.round(eacMargin / contractValue * 1000) / 10
      : 0;
    return { contractValue, laborBudget, subsBudget: subsBudgets[key],
             laborActuals: laborActuals[key], laborEAC, totalEAC,
             eacMargin, eacMarginPct, samPct: 8 };
  }

  const kd = { int: buildKD('int'), lin: buildKD('lin') };

  // Monthly chart series (identical to v2)
  const monthly = {
    int: buildMonthly(timeRows, allocRows, PROJ_INT, todayMs, mult),
    lin: buildMonthly(timeRows, allocRows, PROJ_LIN, todayMs, mult),
  };
  monthly.all = mergeMonthly(monthly.int, monthly.lin);

  // Role detail (new in v3)
  const roleDetail = {
    int: buildRoleDetail('int', timeRows, allocRows, budgetInt, PROJ_INT, todayMs, mult),
    lin: buildRoleDetail('lin', timeRows, allocRows, budgetLin, PROJ_LIN, todayMs, mult),
  };
  roleDetail.all = mergeRoleDetails(roleDetail.int, roleDetail.lin);

  // Validation (identical to v2)
  const validation = buildValidation(kd, config);

  // Strip weekly breakdown from payload — too large for Apps Script response.
  // Weekly view is not yet wired in the HTML anyway; monthly is the default.
  ['int','lin','all'].forEach(key => {
    if (!roleDetail[key]) return;
    delete roleDetail[key].allWeeks;
    roleDetail[key].roles.forEach(r => {
      delete r.budgetByWeek;
      delete r.actualByWeek;
      delete r.allocByWeek;
    });
  });

  const payload = { kd, future: futureAlloc, monthly, roleDetail, meta: { asOf, mult }, validation };
  const payloadSize = JSON.stringify(payload).length;
  Logger.log('Payload size: ' + Math.round(payloadSize / 1024) + ' KB');
  return payload;
}

// ── ROLE DETAIL (new in v3) ───────────────────────────────────────────────────
function buildRoleDetail(projKey, timeRows, allocRows, budgetRows, projName, todayMs, mult) {
  const bounds      = PROJECT_BOUNDS[projKey];
  const totalMonths = monthsBetween(bounds.start, bounds.end);
  const totalWeeks  = weeksBetween(bounds.start, bounds.end);
  const hasBudget   = budgetRows.length > 0;

  // Collect all roles across all three sources
  const roleSet = new Set();
  budgetRows.forEach(r => { if (r.role) roleSet.add(r.role); });
  timeRows.forEach(r  => {
    if (String(r['Project: Name'] || '').trim() === projName && r['Role: Name'])
      roleSet.add(String(r['Role: Name']).trim());
  });
  allocRows.forEach(r => {
    if (String(r['Project: Name'] || '').trim() === projName && r['Role: Name'])
      roleSet.add(String(r['Role: Name']).trim());
  });

  const roles = [];
  roleSet.forEach(role => {
    // Budget (from optional budget-plan tab)
    const bRows       = budgetRows.filter(r => r.role === role);
    const budgetFees  = bRows.reduce((s, r) => s + (r.kProjCost * mult || 0), 0);
    const budgetHours = bRows.reduce((s, r) => s + (r.hours || 0), 0);

    // Actuals by month and week
    const actualByMonth = {}, actualByWeek = {};
    timeRows.forEach(r => {
      if (String(r['Project: Name'] || '').trim() !== projName) return;
      if (String(r['Role: Name']    || '').trim() !== role)     return;
      const d = parseDate(r['Date (Shared)']);
      if (!d) return;
      const mk = toYM(d), wk = toISOWeek(d);
      const cost = parseNum(r['Cost Actual']) * mult;
      const hrs  = parseNum(r['Hours Actual']);
      if (!actualByMonth[mk]) actualByMonth[mk] = { fees:0, hours:0 };
      actualByMonth[mk].fees  += cost;
      actualByMonth[mk].hours += hrs;
      if (!actualByWeek[wk]) actualByWeek[wk] = { fees:0, hours:0 };
      actualByWeek[wk].fees  += cost;
      actualByWeek[wk].hours += hrs;
    });

    // Future allocations by month and week
    const allocByMonth = {}, allocByWeek = {};
    allocRows.forEach(r => {
      if (String(r['Project: Name'] || '').trim() !== projName) return;
      if (String(r['Role: Name']    || '').trim() !== role)     return;
      const d = parseDate(r['Date (Shared)']);
      if (!d || d.getTime() < todayMs) return;
      const mk = toYM(d), wk = toISOWeek(d);
      const cost = parseNum(r['Cost Allocated']) * mult;
      const hrs  = parseNum(r['Hours Allocated']);
      if (!allocByMonth[mk]) allocByMonth[mk] = { fees:0, hours:0 };
      allocByMonth[mk].fees  += cost;
      allocByMonth[mk].hours += hrs;
      if (!allocByWeek[wk]) allocByWeek[wk] = { fees:0, hours:0 };
      allocByWeek[wk].fees  += cost;
      allocByWeek[wk].hours += hrs;
    });

    // Flat budget distribution across project timeline
    const budgetByMonth = {}, budgetByWeek = {};
    if (hasBudget && budgetFees > 0) {
      const perMonth = budgetFees / totalMonths;
      const perWeek  = budgetFees / totalWeeks;
      eachMonth(bounds.start, bounds.end).forEach(mk => { budgetByMonth[mk] = perMonth; });
      eachWeek(bounds.start,  bounds.end).forEach(wk => { budgetByWeek[wk]  = perWeek; });
    }

    const actualTotal = Object.values(actualByMonth).reduce((s,v) => s + v.fees, 0);
    const allocTotal  = Object.values(allocByMonth).reduce((s,v)  => s + v.fees, 0);
    const eac         = actualTotal + allocTotal;
    const variance    = hasBudget ? budgetFees - eac : null;

    roles.push({ role, budgetFees, budgetHours, hasBudget, actualTotal, allocTotal, eac, variance,
                 budgetByMonth, budgetByWeek, actualByMonth, actualByWeek, allocByMonth, allocByWeek });
  });

  roles.sort((a,b) => (b.budgetFees - a.budgetFees) || (b.actualTotal - a.actualTotal) || a.role.localeCompare(b.role));

  return {
    roles,
    hasBudget,
    allMonths: collectAllPeriodKeys(roles, 'Month'),
    allWeeks:  collectAllPeriodKeys(roles, 'Week')
  };
}

function collectAllPeriodKeys(roles, type) {
  const keySet = new Set();
  const [bk,ak,lk] = type === 'Month'
    ? ['budgetByMonth','actualByMonth','allocByMonth']
    : ['budgetByWeek', 'actualByWeek', 'allocByWeek'];
  roles.forEach(r => {
    Object.keys(r[bk]).forEach(k => keySet.add(k));
    Object.keys(r[ak]).forEach(k => keySet.add(k));
    Object.keys(r[lk]).forEach(k => keySet.add(k));
  });
  return Array.from(keySet).sort();
}

function mergeRoleDetails(intD, linD) {
  const map = {};
  function absorb(detail) {
    detail.roles.forEach(r => {
      if (!map[r.role]) {
        map[r.role] = JSON.parse(JSON.stringify(r));
      } else {
        const m = map[r.role];
        m.budgetFees  += r.budgetFees;
        m.budgetHours += r.budgetHours;
        m.actualTotal += r.actualTotal;
        m.allocTotal  += r.allocTotal;
        m.eac         += r.eac;
        if (m.variance !== null && r.variance !== null) m.variance += r.variance;
        else m.variance = null;
        ['budgetByMonth','budgetByWeek'].forEach(key => {
          Object.entries(r[key]).forEach(([k,v]) => { m[key][k] = (m[key][k]||0) + v; });
        });
        ['actualByMonth','actualByWeek','allocByMonth','allocByWeek'].forEach(key => {
          Object.entries(r[key]).forEach(([k,v]) => {
            if (!m[key][k]) m[key][k] = { fees:0, hours:0 };
            m[key][k].fees  += v.fees  || 0;
            m[key][k].hours += v.hours || 0;
          });
        });
      }
    });
  }
  absorb(intD);
  absorb(linD);
  const roles = Object.values(map).sort((a,b) => (b.budgetFees - a.budgetFees) || (b.actualTotal - a.actualTotal) || a.role.localeCompare(b.role));
  const hasBudget = intD.hasBudget || linD.hasBudget;
  return { roles, hasBudget, allMonths: collectAllPeriodKeys(roles,'Month'), allWeeks: collectAllPeriodKeys(roles,'Week') };
}

// ── BUDGET PLAN TABS (optional, new in v3) ────────────────────────────────────
// Paste sheet 44172716 → tab "budget-plan-int", sheet 44923833 → tab "budget-plan-lin"
// Structure: blank row 1, title row 2, year row 3, header row 4 (col B="Team Member"),
//            data rows. Col B=name, Col C=Team Role, Col E=Hours, Col F=K Proj Cost
function parseBudgetTab(ss, tabName) {
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) { Logger.log('Budget tab absent (optional): ' + tabName); return []; }
  const rows = sheet.getDataRange().getValues();
  let hi = -1;
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === 'Team Member') { hi = i; break; }
  }
  if (hi < 0) { Logger.log('Budget header not found in ' + tabName); return []; }
  const results = [];
  for (let i = hi + 1; i < rows.length; i++) {
    const r         = rows[i];
    const role      = String(r[2] || '').trim();   // col C
    const hours     = parseNum(r[4]);              // col E
    const kProjCost = parseNum(r[5]);              // col F
    if (!role || role.toLowerCase() === 'subcontractors') break;
    if (!hours && !kProjCost) continue;
    results.push({ role, hours, kProjCost });
  }
  Logger.log('Budget rows for ' + tabName + ': ' + results.length);
  return results;
}

// ── PERIOD HELPERS (new in v3) ────────────────────────────────────────────────
function monthsBetween(start, end) {
  return (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth()) + 1;
}
function weeksBetween(start, end) {
  return Math.ceil((end - start) / (7 * 24 * 60 * 60 * 1000));
}
function eachMonth(start, end) {
  const months = [], d = new Date(start.getFullYear(), start.getMonth(), 1);
  while (d <= end) { months.push(toYM(d)); d.setMonth(d.getMonth()+1); }
  return months;
}
function eachWeek(start, end) {
  const weeks = [], d = new Date(start);
  d.setDate(d.getDate() - (d.getDay()+6) % 7);
  while (d <= end) { weeks.push(toISOWeek(d)); d.setDate(d.getDate()+7); }
  return [...new Set(weeks)];
}
function toISOWeek(d) {
  const date = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  date.setDate(date.getDate() + 3 - (date.getDay()+6) % 7);
  const week1 = new Date(date.getFullYear(), 0, 4);
  const wn = 1 + Math.round(((date - week1) / 86400000 - 3 + (week1.getDay()+6) % 7) / 7);
  return date.getFullYear() + '-W' + String(wn).padStart(2,'0');
}

// ── EVERYTHING BELOW IS IDENTICAL TO v2 ──────────────────────────────────────

function computeSubsBudgets(taskRows) {
  const subs = { int: 0, lin: 0 };
  taskRows.forEach(r => {
    const budget = parseNum(r['Task Estimated Budget']);
    if (!budget || budget <= 0) return;
    const id = String(r['Task: ID'] || '').trim();
    if      (INT_SUB_IDS.indexOf(id) !== -1) subs.int += budget;
    else if (LIN_SUB_IDS.indexOf(id) !== -1) subs.lin += budget;
    else Logger.log('Unclassified task: ' + r['Task: Name'] + ' (ID:' + id + ')');
  });
  Logger.log('Subs — INT: $' + subs.int + '  LIN: $' + subs.lin);
  return subs;
}

function sumLaborActuals(timeRows, projectName, mult) {
  let total = 0;
  timeRows.forEach(r => {
    if (String(r['Project: Name'] || '').trim() !== projectName) return;
    total += parseNum(r['Cost Actual']) * mult;
  });
  return Math.round(total);
}

function sumFutureAlloc(allocRows, projectName, todayMs, mult) {
  let cost = 0, hrs = 0;
  allocRows.forEach(r => {
    if (String(r['Project: Name'] || '').trim() !== projectName) return;
    const d = parseDate(r['Date (Shared)']);
    if (!d || d.getTime() < todayMs) return;
    cost += parseNum(r['Cost Allocated']) * mult;
    hrs  += parseNum(r['Hours Allocated']);
  });
  return { laborCost: Math.round(cost), hrs: Math.round(hrs * 10) / 10 };
}

function buildMonthly(timeRows, allocRows, projectName, todayMs, mult) {
  const actByYM = {}, allocByYM = {};
  timeRows.forEach(r => {
    if (String(r['Project: Name'] || '').trim() !== projectName) return;
    const d = parseDate(r['Date (Shared)']);
    if (!d) return;
    const ym = toYM(d);
    if (!actByYM[ym]) actByYM[ym] = { cost:0, hrs:0 };
    actByYM[ym].cost += parseNum(r['Cost Actual']) * mult;
    actByYM[ym].hrs  += parseNum(r['Hours Actual']);
  });
  allocRows.forEach(r => {
    if (String(r['Project: Name'] || '').trim() !== projectName) return;
    const d = parseDate(r['Date (Shared)']);
    if (!d) return;
    const ym = toYM(d);
    if (!allocByYM[ym]) allocByYM[ym] = { cost:0, hrs:0 };
    allocByYM[ym].cost += parseNum(r['Cost Allocated']) * mult;
    allocByYM[ym].hrs  += parseNum(r['Hours Allocated']);
  });
  const todayYM = toYM(new Date(todayMs));
  return [...new Set([...Object.keys(actByYM), ...Object.keys(allocByYM)])].sort()
    .map(ym => ({
      ym,
      isFuture:   ym >= todayYM,
      actualCost: Math.round((actByYM[ym]   || {cost:0}).cost),
      allocCost:  Math.round((allocByYM[ym] || {cost:0}).cost),
      actualHrs:  Math.round(((actByYM[ym]  || {hrs:0}).hrs) * 10) / 10,
      allocHrs:   Math.round(((allocByYM[ym]|| {hrs:0}).hrs) * 10) / 10,
    }));
}

function mergeMonthly(a, b) {
  const map = {};
  [...a, ...b].forEach(r => {
    if (!map[r.ym]) map[r.ym] = { ym:r.ym, isFuture:r.isFuture, actualCost:0, allocCost:0, actualHrs:0, allocHrs:0 };
    map[r.ym].actualCost += r.actualCost;
    map[r.ym].allocCost  += r.allocCost;
    map[r.ym].actualHrs  += r.actualHrs;
    map[r.ym].allocHrs   += r.allocHrs;
  });
  return Object.values(map).sort((a,b) => a.ym.localeCompare(b.ym));
}

function buildValidation(kd, config) {
  const checks = [
    ['int','laborActuals','int_labor_actuals'],
    ['int','laborEAC',    'int_labor_eac'],
    ['int','totalEAC',    'int_total_eac'],
    ['int','eacMargin',   'int_eac_margin'],
    ['lin','laborActuals','lin_labor_actuals'],
    ['lin','laborEAC',    'lin_labor_eac'],
    ['lin','totalEAC',    'lin_total_eac'],
    ['lin','eacMargin',   'lin_eac_margin'],
  ];
  return checks
    .filter(([proj, field, cfgKey]) => config[cfgKey] !== undefined)
    .map(([proj, field, cfgKey]) => {
      const computed = kd[proj][field];
      const kPdf     = Number(config[cfgKey]);
      const delta    = computed - kPdf;
      const pct      = kPdf !== 0 ? Math.round(delta / Math.abs(kPdf) * 1000) / 10 : null;
      Logger.log(cfgKey + ': computed=' + computed + '  K-PDF=' + kPdf + '  Δ=' + delta);
      return { label: cfgKey, computed, kPdf, delta, pct };
    });
}

function readSheetFromMarker(ss, tabName, markerText) {
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) { Logger.log('Tab not found: ' + tabName); return []; }
  const data = sheet.getDataRange().getValues();
  let headerRow = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i].join('|').indexOf(markerText) !== -1) { headerRow = i; break; }
  }
  if (headerRow === -1) { Logger.log('Marker not found: ' + markerText + ' in ' + tabName); return []; }
  const headers = data[headerRow].map(h => String(h).trim());
  const rows = [];
  for (let i = headerRow + 1; i < data.length; i++) {
    if (data[i].every(c => c === '' || c === null)) continue;
    const obj = {};
    headers.forEach((h, j) => { obj[h] = data[i][j]; });
    rows.push(obj);
  }
  return rows;
}

function readConfig(ss) {
  const sheet = ss.getSheetByName('config');
  if (!sheet) return {};
  const cfg = {};
  sheet.getDataRange().getValues().forEach(row => {
    const key = String(row[0]).trim();
    const val = row[1];
    if (!key) return;
    cfg[key] = (typeof val === 'number' || (!isNaN(Number(val)) && String(val).trim() !== ''))
      ? Number(val) : String(val);
  });
  return cfg;
}

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}
function toYM(d) {
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0');
}
function parseNum(val) {
  if (typeof val === 'number') return val;
  if (!val && val !== 0) return 0;
  const n = parseFloat(String(val).replace(/[$,\s]/g,''));
  return isNaN(n) ? 0 : n;
}