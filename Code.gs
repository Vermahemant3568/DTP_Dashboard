/* ================================================================
   Code.gs — DTP Project Tracker
   Single source of truth for: constants, helpers, calculations,
   dashboard data, dropdown data, batch load.
================================================================ */

const SS_ID       = "1-jcai8DeFYpHmfzWGjcYmA4Ehbt0ASO3Ue240P0HGM4";
const SH_PROJECTS = "Projects";
const SH_TASKS    = "Tasks";
const SH_REVISIONS= "Revisions";
const SH_VENDORS  = "Vendors";
const SH_TEAM     = "Team";
const SH_SNAPSHOT = "MonthlySnapshot";

/* ── TASK column indices (0-based) ── */
const TC = {
  ID:0, PROJECT_ID:1, CLIENT:2, PROJECT_NAME:3,
  TASK_TYPE:4, WORK_TYPE:5, ASSIGNED_TO:6, VENDOR_NAME:7,
  LANGUAGE:8, SOURCE_PAGES:9, FINAL_PAGES:10, LANG_COUNT:11,
  STATUS:12, PRIORITY:13, START_DATE:14, DELIVERY_DATE:15,
  COMPLETED_DATE:16, SOURCE_LINK:17, DELIVERABLE_LINK:18, NOTES:19,
  CREATED_AT:20, UPDATED_AT:21,
  RATE_PER_PAGE:22, CURRENCY:23, PAYMENT_STATUS:24
};

/* ── REVISION column indices (0-based) ── */
const RC = {
  ID:0, PROJECT_ID:1, TASK_ID:2, PROJECT_NAME:3,
  REV_NUMBER:4, REV_TYPE:5, LANGUAGE:6, REV_PAGES:7,
  WORK_TYPE:8, ASSIGNED_TO:9, VENDOR_NAME:10, STATUS:11,
  REV_DATE:12, DELIVERY_DATE:13, COMPLETED_DATE:14, NOTES:15,
  CREATED_AT:16, UPDATED_AT:17,
  RATE_PER_PAGE:18, CURRENCY:19, PAYMENT_STATUS:20
};

/* ── PROJECT column indices (0-based) ── */
/* Relationships:
   Projects (1) ──< Tasks    (many)  via Tasks.PROJECT_ID
   Projects (1) ──< Revisions(many)  via Revisions.PROJECT_ID
   Tasks    (1) ──< Revisions(many)  via Revisions.TASK_ID  (optional)
   Tasks.ASSIGNED_TO  → Team.Name   (In-House work)
   Tasks.VENDOR_NAME  → Vendors.Name(Vendor work)
   Revisions.ASSIGNED_TO → Team.Name
   Revisions.VENDOR_NAME → Vendors.Name
*/
const PC = {
  ID:0, CLIENT:1, PROJECT_NAME:2, COORDINATOR:3,
  SOURCE_LANG:4, TARGET_LANGS:5, LANG_COUNT:6, SOURCE_PAGES:7,
  PRIORITY:8, STATUS:9, RECEIVED_DATE:10,
  NOTES:11, CREATED_AT:12, UPDATED_AT:13
};

/* ── VENDOR column indices (0-based) ── */
const VC = {
  ID:0, NAME:1, CONTACT:2, EMAIL:3,
  PHONE:4, SPECIALIZATION:5, LANGUAGES:6, RATE_PER_PAGE:7,
  CURRENCY:8, STATUS:9, NOTES:10, CREATED_AT:11, UPDATED_AT:12
};

/* ── TEAM column indices (0-based) ── */
const MC = {
  ID:0, NAME:1, ROLE:2, EMAIL:3,
  PHONE:4, SPECIALIZATION:5, STATUS:6, RATE_PER_PAGE:7,
  CURRENCY:8, CREATED_AT:9
};

/* ── ENTRY POINT ── */
function doGet() {
  return HtmlService.createTemplateFromFile("Layout")
    .evaluate()
    .setTitle("DTP Project Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
function loadPageFile(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ── CORE HELPERS ── */
function _ss()  { return SpreadsheetApp.openById(SS_ID); }
function _sh(n) { return _ss().getSheetByName(n); }

function _id(prefix) {
  try {
    var sheetName = prefix === "PRJ" ? SH_PROJECTS
                  : prefix === "TSK" ? SH_TASKS
                  : prefix === "REV" ? SH_REVISIONS
                  : prefix === "VND" ? SH_VENDORS
                  : prefix === "MBR" ? SH_TEAM : null;
    var max = 0;
    if (sheetName) {
      _sheetRows(sheetName).forEach(function(r) {
        var m = String(r[0]).match(/^[A-Z]+-0*(\d+)$/);
        if (m) max = Math.max(max, parseInt(m[1], 10));
      });
    }
    var next = max + 1;
    return prefix + "-" + (next < 1000 ? String(next).padStart(3, "0") : String(next));
  } catch(e) {
    return prefix + "-" + Date.now();
  }
}

function _fmt(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy");
  return (v === null || v === undefined) ? "" : v;
}
function _fmtRow(row) { return row.map(_fmt); }

function _findRow(sheet, id) {
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) return { row: data[i], index: i + 1 };
  }
  return null;
}

function _sheetRows(name) {
  const sh = _sh(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getDataRange().getValues().slice(1).filter(r => r[0] !== "");
}

function _inMonth(v, month, year) {
  if (!v) return false;
  if (typeof v === "string") {
    var dmy = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (dmy) v = new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]));
    else v = new Date(v);
  }
  const d = v instanceof Date ? v : new Date(v);
  return !isNaN(d) && d.getMonth() === month && d.getFullYear() === year;
}

/* ================================================================
   CALCULATIONS
   All payment and stats logic is in Calculations.gs.
   Functions available: _calcPaymentStats, applyFilters,
   calculateVendorStats, calculateTeamMemberStats,
   calculateProjectPaymentSummary, calculateDashboardPaymentTotals
================================================================ */

/* ================================================================
   BATCH LOAD
================================================================ */
function getAllData() {
  try {
    const ss = _ss();
    function _read(name) {
      const sh = ss.getSheetByName(name);
      if (!sh || sh.getLastRow() < 2) return [];
      return sh.getDataRange().getValues().slice(1).filter(r => r[0] !== "").map(_fmtRow);
    }
    return {
      projects:  _read(SH_PROJECTS),
      tasks:     _read(SH_TASKS),
      revisions: _read(SH_REVISIONS),
      vendors:   _read(SH_VENDORS),
      team:      _read(SH_TEAM)
    };
  } catch (e) {
    console.error("getAllData:", e);
    return { projects: [], tasks: [], revisions: [], vendors: [], team: [] };
  }
}

/* ================================================================
   MONTHLY SUMMARY
================================================================ */
function getMonthlySummary(params) {
  try {
    const now   = new Date();
    const year  = (params && params.year)  ? Number(params.year)  : now.getFullYear();
    const month = (params && params.month) ? Number(params.month) - 1 : now.getMonth();

    const tasks     = (params && params._tasks)     ? params._tasks     : _sheetRows(SH_TASKS);
    const revisions = (params && params._revisions) ? params._revisions : _sheetRows(SH_REVISIONS);

    const monthTasks = tasks.filter(r => _inMonth(r[TC.START_DATE] || r[TC.CREATED_AT], month, year));
    const monthRevs  = revisions.filter(r => _inMonth(r[RC.REV_DATE] || r[RC.CREATED_AT], month, year));

    const TASK_TYPES = ["Main DTP", "Pre-Engineering", "Bilingual Creation", "Key Insertion of PDFs", "Formatting", "Others"];
    const TASK_KEYS  = { "Main DTP":"mainDTP", "Pre-Engineering":"preEng",
                         "Bilingual Creation":"bilingual", "Key Insertion of PDFs":"keyInsertion",
                         "Formatting":"formatting", "Others":"others" };

    const byTaskType = {};
    TASK_TYPES.forEach(function(t) {
      const rows = monthTasks.filter(r => r[TC.TASK_TYPE] === t);
      const ih   = rows.filter(r => r[TC.WORK_TYPE] === "In-House").reduce((s,r) => s+(Number(r[TC.FINAL_PAGES])||0), 0);
      const vd   = rows.filter(r => r[TC.WORK_TYPE] === "Vendor").reduce((s,r)   => s+(Number(r[TC.FINAL_PAGES])||0), 0);
      byTaskType[TASK_KEYS[t]] = { label:t, total:ih+vd, inHouse:ih, vendor:vd, count:rows.length };
    });

    const revIHRows = monthRevs.filter(r => r[RC.WORK_TYPE] === "In-House");
    const revVDRows = monthRevs.filter(r => r[RC.WORK_TYPE] === "Vendor");
    const revIH    = revIHRows.reduce((s,r) => s+(Number(r[RC.REV_PAGES])||0), 0);
    const revVD    = revVDRows.reduce((s,r) => s+(Number(r[RC.REV_PAGES])||0), 0);
    byTaskType.revisions = { label:"Revisions", total:revIH+revVD, inHouse:revIH, vendor:revVD, count:monthRevs.length, inHouseCount:revIHRows.length, vendorCount:revVDRows.length };

    const allTaskPages = monthTasks.reduce((s,r) => s+(Number(r[TC.FINAL_PAGES])||0), 0);
    const allRevPages  = monthRevs.reduce((s,r)  => s+(Number(r[RC.REV_PAGES])||0), 0);
    const grandTotal   = allTaskPages + allRevPages;
    const totalInHouse = monthTasks.filter(r => r[TC.WORK_TYPE]==="In-House").reduce((s,r) => s+(Number(r[TC.FINAL_PAGES])||0), 0) + revIH;
    const totalVendor  = monthTasks.filter(r => r[TC.WORK_TYPE]==="Vendor").reduce((s,r)   => s+(Number(r[TC.FINAL_PAGES])||0), 0) + revVD;

    const monthLabel = Utilities.formatDate(new Date(year, month, 1), Session.getScriptTimeZone(), "MMMM yyyy");

    return {
      month: monthLabel, year, monthIndex: month + 1,
      byTaskType, grandTotal, totalInHouse, totalVendor,
      taskCount: monthTasks.length, revisionCount: monthRevs.length
    };
  } catch (e) {
    console.error("getMonthlySummary:", e);
    return { byTaskType:{}, grandTotal:0, totalInHouse:0, totalVendor:0, month:"", taskCount:0, revisionCount:0 };
  }
}

/* ================================================================
   DASHBOARD DATA
================================================================ */
function getDashboardData(params) {
  try {
    const now   = new Date();
    const year  = (params && params.year)  ? Number(params.year)  : now.getFullYear();
    const month = (params && params.month) ? Number(params.month) - 1 : now.getMonth();

    const rawTasks     = _sheetRows(SH_TASKS);
    const rawRevisions = _sheetRows(SH_REVISIONS);
    const allTasks     = rawTasks.map(_fmtRow);
    const allRevisions = rawRevisions.map(_fmtRow);
    const allProjects  = _sheetRows(SH_PROJECTS).map(_fmtRow);

    /* Build project → total pages map for active/in-progress page counts */
    const rawTasksForProj     = rawTasks;
    const rawRevisionsForProj = rawRevisions;
    const projPageMap = {};
    rawTasksForProj.forEach(function(t) {
      const pid = String(t[TC.PROJECT_ID] || "").trim();
      if (!pid) return;
      projPageMap[pid] = (projPageMap[pid] || 0) + (Number(t[TC.FINAL_PAGES]) || 0);
    });
    rawRevisionsForProj.forEach(function(r) {
      const pid = String(r[RC.PROJECT_ID] || "").trim();
      if (!pid) return;
      projPageMap[pid] = (projPageMap[pid] || 0) + (Number(r[RC.REV_PAGES]) || 0);
    });

    const activeProjectsList     = allProjects.filter(r => r[PC.STATUS] === "Active");
    const inProgressProjectsList = allProjects.filter(r => r[PC.STATUS] === "In Progress");
    const activeProjectPages     = activeProjectsList.reduce((s,r) => s + (projPageMap[String(r[PC.ID]).trim()] || 0), 0);
    const inProgressProjectPages = inProgressProjectsList.reduce((s,r) => s + (projPageMap[String(r[PC.ID]).trim()] || 0), 0);

    const summary = getMonthlySummary({
      year: year, month: month + 1,
      _tasks: rawTasks, _revisions: rawRevisions
    });

    // Payment-aware totals via centralized calculator
    var payTotals = calculateDashboardPaymentTotals(allTasks, allRevisions);
    var totalPaidAmt    = payTotals.totalPaidAmt;
    var totalUnpaidAmt  = payTotals.totalUnpaidAmt;
    var totalPartialAmt = payTotals.totalPartialAmt;

    return {
      summary,
      allTimeTotals: {
        totalProjects:     allProjects.length,
        activeProjects:    activeProjectsList.length,
        inProgressProjects: inProgressProjectsList.length,
        completedProjects: allProjects.filter(r => r[PC.STATUS] === "Completed").length,
        activeProjectPages,
        inProgressProjectPages,
        totalTasks:        allTasks.length,
        pendingTasks:      allTasks.filter(r => r[TC.STATUS] === "Pending").length,
        inProgressTasks:   allTasks.filter(r => r[TC.STATUS] === "In Progress").length,
        totalTaskPages:    allTasks.reduce((s,r) => s+(Number(r[TC.FINAL_PAGES])||0), 0),
        totalRevisions:    allRevisions.length,
        pendingRevisions:  allRevisions.filter(r => r[RC.STATUS] === "Pending" || r[RC.STATUS] === "In Progress").length,
        totalRevPages:     allRevisions.reduce((s,r) => s+(Number(r[RC.REV_PAGES])||0), 0),
        totalRevIHPages:   allRevisions.filter(r => r[RC.WORK_TYPE] === "In-House").reduce((s,r) => s+(Number(r[RC.REV_PAGES])||0), 0),
        totalRevVDPages:   allRevisions.filter(r => r[RC.WORK_TYPE] === "Vendor").reduce((s,r) => s+(Number(r[RC.REV_PAGES])||0), 0),
        inProgressPages:   allTasks.filter(r => r[TC.STATUS]==="In Progress").reduce((s,r) => s+(Number(r[TC.FINAL_PAGES])||0), 0),
        totalPaidAmt,
        totalUnpaidAmt,
        totalPartialAmt
      }
    };
  } catch(e) {
    console.error("getDashboardData:", e);
    return { summary:{}, allTimeTotals:{} };
  }
}

/* ================================================================
   PROJECT FULL SUMMARY
================================================================ */
function getProjectSummary(projectId) {
  try {
    const pid       = String(projectId).trim();
    const projFound = _findRow(_sh(SH_PROJECTS), pid);
    if (!projFound) return null;

    const tasks     = _sheetRows(SH_TASKS).filter(r => String(r[TC.PROJECT_ID]).trim() === pid).map(_fmtRow);
    const revisions = _sheetRows(SH_REVISIONS).filter(r => String(r[RC.PROJECT_ID]).trim() === pid).map(_fmtRow);

    const taskPages = tasks.reduce((s,r) => s+(Number(r[TC.FINAL_PAGES])||0), 0);
    const revPages  = revisions.reduce((s,r) => s+(Number(r[RC.REV_PAGES])||0), 0);

    var paymentSummary = calculateProjectPaymentSummary(tasks, revisions);

    return {
      project:    _fmtRow(projFound.row),
      tasks,
      revisions,
      taskPages,
      revPages,
      totalPages: taskPages + revPages,
      paidAmt:    paymentSummary.totalPaidAmt,
      unpaidAmt:  paymentSummary.totalUnpaidAmt,
      partialAmt: paymentSummary.totalPartialAmt,
      totalAmount:paymentSummary.totalAmount,
      paymentSummary
    };
  } catch (e) {
    console.error("getProjectSummary:", e);
    return null;
  }
}

/* ================================================================
   DEBUG — run once from Apps Script editor to check ID matching
================================================================ */
function debugTaskProjectMatch() {
  var rawProjects = _sheetRows(SH_PROJECTS);
  var rawTasks    = _sheetRows(SH_TASKS);
  var log = [];
  log.push("Projects count: " + rawProjects.length);
  log.push("Tasks count: " + rawTasks.length);
  if (rawProjects.length) log.push("First project ID: [" + String(rawProjects[0][PC.ID]) + "] col=" + PC.ID);
  if (rawTasks.length)    log.push("First task project_id: [" + String(rawTasks[0][TC.PROJECT_ID]) + "] col=" + TC.PROJECT_ID);
  /* Check first 5 tasks */
  rawTasks.slice(0, 5).forEach(function(t, i) {
    var pid = String(t[TC.PROJECT_ID] || "").trim();
    var match = rawProjects.find(function(p) { return String(p[PC.ID]||"").trim() === pid; });
    log.push("Task[" + i + "] pid=[" + pid + "] match=" + (match ? "YES" : "NO"));
  });
  Logger.log(log.join("\n"));
  SpreadsheetApp.getUi().alert(log.join("\n"));
}

/* ================================================================
   PROJECTS WITH TASK COUNT
================================================================ */
function getProjectsWithTaskCount() {
  try {
    const rawProjects = _sheetRows(SH_PROJECTS);
    const rawTasks    = _sheetRows(SH_TASKS);

    /* Build task map keyed by trimmed PROJECT_ID */
    const taskMap = {};
    rawTasks.forEach(function(t) {
      const pid = String(t[TC.PROJECT_ID] || "").trim();
      if (!pid) return;
      if (!taskMap[pid]) taskMap[pid] = { count: 0, pages: 0 };
      taskMap[pid].count++;
      taskMap[pid].pages += Number(t[TC.FINAL_PAGES]) || 0;
    });

    /* Build revision map keyed by PROJECT_ID */
    const rawRevisions = _sheetRows(SH_REVISIONS);
    const revMap = {};
    rawRevisions.forEach(function(r) {
      const pid = String(r[RC.PROJECT_ID] || "").trim();
      if (!pid) return;
      if (!revMap[pid]) revMap[pid] = { count: 0, pages: 0 };
      revMap[pid].count++;
      revMap[pid].pages += Number(r[RC.REV_PAGES]) || 0;
    });

    return rawProjects.map(function(p) {
      const pid    = String(p[PC.ID] || "").trim();
      const tInfo  = taskMap[pid] || { count: 0, pages: 0 };
      const rInfo  = revMap[pid]  || { count: 0, pages: 0 };
      const row    = _fmtRow(p).slice(0, 14);
      row[14] = tInfo.count;                /* task count       */
      row[15] = tInfo.pages;               /* task final pages  */
      row[16] = rInfo.count;               /* revision count    */
      row[17] = rInfo.pages;               /* revision pages    */
      row[18] = tInfo.pages + rInfo.pages; /* total pages       */
      return row;
    });
  } catch (e) {
    console.error("getProjectsWithTaskCount:", e);
    return [];
  }
}

/* ================================================================
   DROPDOWN DATA
================================================================ */
function getDropdownData() {
  try {
    var vendors = _sheetRows(SH_VENDORS).map(function(r) {
      return {
        id: String(r[VC.ID]||""), name: String(r[VC.NAME]||"").trim(),
        contactPerson: String(r[VC.CONTACT]||""), specialization: String(r[VC.SPECIALIZATION]||""),
        languages: String(r[VC.LANGUAGES]||""), ratePerPage: Number(r[VC.RATE_PER_PAGE])||0,
        currency: String(r[VC.CURRENCY]||""), status: String(r[VC.STATUS]||"Active").trim()
      };
    }).filter(function(v) { return v.name && v.status.toLowerCase() !== "inactive"; });

    var team = _sheetRows(SH_TEAM).map(function(r) {
      return {
        id: String(r[MC.ID]||""), name: String(r[MC.NAME]||"").trim(),
        role: String(r[MC.ROLE]||""), specialization: String(r[MC.SPECIALIZATION]||""),
        ratePerPage: Number(r[MC.RATE_PER_PAGE])||0, currency: String(r[MC.CURRENCY]||""),
        status: String(r[MC.STATUS]||"Active").trim()
      };
    }).filter(function(m) { return m.name && m.status.toLowerCase() !== "inactive"; });

    return { vendors, team };
  } catch (e) {
    console.error("getDropdownData:", e);
    return { vendors:[], team:[] };
  }
}

/* ================================================================
   TEAMS & VENDORS ANALYTICS
   Returns aggregated stats + monthly trends for the unified module.
================================================================ */
function getTeamsVendorsAnalytics() {
  try {
    var allTasks = _sheetRows(SH_TASKS).map(_fmtRow);
    var allRevs  = _sheetRows(SH_REVISIONS).map(_fmtRow);
    var now      = new Date();

    /* ── Team stats ── */
    var ihTasks = allTasks.filter(function(r){ return r[TC.WORK_TYPE] === 'In-House'; });
    var ihRevs  = allRevs.filter(function(r){  return r[RC.WORK_TYPE] === 'In-House'; });

    var ihTotalPages     = ihTasks.reduce(function(s,r){ return s+(Number(r[TC.FINAL_PAGES])||0); }, 0);
    var ihCompletedPages = ihTasks.filter(function(r){ return r[TC.STATUS]==='Completed'; })
                                  .reduce(function(s,r){ return s+(Number(r[TC.FINAL_PAGES])||0); }, 0);
    var ihPendingPages   = ihTotalPages - ihCompletedPages;
    var ihRevPages       = ihRevs.reduce(function(s,r){ return s+(Number(r[RC.REV_PAGES])||0); }, 0);
    var ihActiveTasks    = ihTasks.filter(function(r){ return r[TC.STATUS]==='In Progress'||r[TC.STATUS]==='Pending'; }).length;
    var ihCompletedTasks = ihTasks.filter(function(r){ return r[TC.STATUS]==='Completed'; }).length;

    /* Monthly team production — last 12 months */
    var teamMonthly = _buildMonthlyTrend(ihTasks, ihRevs, now, function(t){ return t[TC.FINAL_PAGES]; }, function(r){ return r[RC.REV_PAGES]; }, function(t){ return t[TC.START_DATE]||t[TC.CREATED_AT]; }, function(r){ return r[RC.REV_DATE]||r[RC.CREATED_AT]; });

    /* Per-member page distribution */
    var memberMap = {};
    ihTasks.forEach(function(r){
      var n = String(r[TC.ASSIGNED_TO]||'').trim(); if (!n) return;
      if (!memberMap[n]) memberMap[n] = { name:n, pages:0, tasks:0 };
      memberMap[n].pages += Number(r[TC.FINAL_PAGES])||0;
      memberMap[n].tasks++;
    });
    ihRevs.forEach(function(r){
      var n = String(r[RC.ASSIGNED_TO]||'').trim(); if (!n) return;
      if (!memberMap[n]) memberMap[n] = { name:n, pages:0, tasks:0 };
      memberMap[n].pages += Number(r[RC.REV_PAGES])||0;
      memberMap[n].tasks++;
    });
    var teamMembers = Object.values(memberMap).sort(function(a,b){ return b.pages-a.pages; });

    /* ── Vendor stats ── */
    var vdTasks = allTasks.filter(function(r){ return r[TC.WORK_TYPE] === 'Vendor'; });
    var vdRevs  = allRevs.filter(function(r){  return r[RC.WORK_TYPE] === 'Vendor'; });

    var vdTotalPages     = vdTasks.reduce(function(s,r){ return s+(Number(r[TC.FINAL_PAGES])||0); }, 0);
    var vdCompletedPages = vdTasks.filter(function(r){ return r[TC.STATUS]==='Completed'; })
                                  .reduce(function(s,r){ return s+(Number(r[TC.FINAL_PAGES])||0); }, 0);
    var vdPendingPages   = vdTotalPages - vdCompletedPages;
    var vdRevPages       = vdRevs.reduce(function(s,r){ return s+(Number(r[RC.REV_PAGES])||0); }, 0);
    var vdActiveTasks    = vdTasks.filter(function(r){ return r[TC.STATUS]==='In Progress'||r[TC.STATUS]==='Pending'; }).length;
    var vdCompletedTasks = vdTasks.filter(function(r){ return r[TC.STATUS]==='Completed'; }).length;

    /* Monthly vendor production — last 12 months */
    var vendorMonthly = _buildMonthlyTrend(vdTasks, vdRevs, now, function(t){ return t[TC.FINAL_PAGES]; }, function(r){ return r[RC.REV_PAGES]; }, function(t){ return t[TC.START_DATE]||t[TC.CREATED_AT]; }, function(r){ return r[RC.REV_DATE]||r[RC.CREATED_AT]; });

    /* Per-vendor page distribution */
    var vendorMap = {};
    vdTasks.forEach(function(r){
      var n = String(r[TC.VENDOR_NAME]||'').trim(); if (!n) return;
      if (!vendorMap[n]) vendorMap[n] = { name:n, pages:0, tasks:0 };
      vendorMap[n].pages += Number(r[TC.FINAL_PAGES])||0;
      vendorMap[n].tasks++;
    });
    vdRevs.forEach(function(r){
      var n = String(r[RC.VENDOR_NAME]||'').trim(); if (!n) return;
      if (!vendorMap[n]) vendorMap[n] = { name:n, pages:0, tasks:0 };
      vendorMap[n].pages += Number(r[RC.REV_PAGES])||0;
      vendorMap[n].tasks++;
    });
    var vendorMembers = Object.values(vendorMap).sort(function(a,b){ return b.pages-a.pages; });

    return {
      team:          { totalPages:ihTotalPages, completedPages:ihCompletedPages, pendingPages:ihPendingPages, revPages:ihRevPages, activeTasks:ihActiveTasks, completedTasks:ihCompletedTasks },
      vendor:        { totalPages:vdTotalPages, completedPages:vdCompletedPages, pendingPages:vdPendingPages, revPages:vdRevPages, activeTasks:vdActiveTasks, completedTasks:vdCompletedTasks },
      teamMonthly:   teamMonthly,
      vendorMonthly: vendorMonthly,
      teamMembers:   teamMembers,
      vendorMembers: vendorMembers
    };
  } catch(e) {
    console.error('getTeamsVendorsAnalytics:', e);
    return { team:{}, vendor:{}, teamMonthly:[], vendorMonthly:[], teamMembers:[], vendorMembers:[] };
  }
}

/* Build last-12-months trend array */
function _buildMonthlyTrend(tasks, revs, now, taskPages, revPages, taskDate, revDate) {
  var months = [];
  for (var i = 11; i >= 0; i--) {
    var d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    var m = d.getMonth(); var y = d.getFullYear();
    var tp = tasks.filter(function(r){ return _inMonth(taskDate(r), m, y); })
                  .reduce(function(s,r){ return s+(Number(taskPages(r))||0); }, 0);
    var rp = revs.filter(function(r){ return _inMonth(revDate(r), m, y); })
                 .reduce(function(s,r){ return s+(Number(revPages(r))||0); }, 0);
    months.push({ label: Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM yy'), pages: tp+rp });
  }
  return months;
}
