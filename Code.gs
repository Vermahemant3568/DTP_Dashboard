/* ================================================================
   Code.gs — DTP Project Tracker
   Entry point, shared helpers, batch load, monthly summary.
   All CRUD is in: Projects.gs | Tasks.gs | Revisions_Service.gs | Vendors_Service.gs
================================================================ */

const SS_ID           = "1-jcai8DeFYpHmfzWGjcYmA4Ehbt0ASO3Ue240P0HGM4";
const SH_PROJECTS     = "Projects";
const SH_TASKS        = "Tasks";
const SH_REVISIONS    = "Revisions";
const SH_VENDORS      = "Vendors";
const SH_TEAM         = "Team";
const SH_SNAPSHOT     = "MonthlySnapshot";

/* ── TASK column indices (0-based) ── */
const TC = {
  ID:0, PROJECT_ID:1, CLIENT:2, PROJECT_NAME:3,
  TASK_TYPE:4, WORK_TYPE:5, ASSIGNED_TO:6, VENDOR_NAME:7,
  LANGUAGE:8, SOURCE_PAGES:9, FINAL_PAGES:10, LANG_COUNT:11,
  STATUS:12, PRIORITY:13, START_DATE:14, DELIVERY_DATE:15,
  COMPLETED_DATE:16, SOURCE_LINK:17, DELIVERABLE_LINK:18, NOTES:19,
  CREATED_AT:20, UPDATED_AT:21
};

/* ── REVISION column indices (0-based) ── */
const RC = {
  ID:0, PROJECT_ID:1, TASK_ID:2, PROJECT_NAME:3,
  REV_NUMBER:4, REV_TYPE:5, LANGUAGE:6, REV_PAGES:7,
  WORK_TYPE:8, ASSIGNED_TO:9, VENDOR_NAME:10, STATUS:11,
  REV_DATE:12, DELIVERY_DATE:13, COMPLETED_DATE:14, NOTES:15,
  CREATED_AT:16, UPDATED_AT:17
};

/* ── PROJECT column indices (0-based) ── */
const PC = {
  ID:0, CLIENT:1, PROJECT_NAME:2, COORDINATOR:3,
  SOURCE_LANG:4, TARGET_LANGS:5, LANG_COUNT:6, SOURCE_PAGES:7,
  WORD_COUNT:8, PRIORITY:9, STATUS:10, RECEIVED_DATE:11,
  NOTES:12, CREATED_AT:13, UPDATED_AT:14
};

/* ── ENTRY POINT ── */
function doGet() {
  return HtmlService.createTemplateFromFile("Layout")
    .evaluate()
    .setTitle("DTP Project Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ── HTML INCLUDE ── */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function loadPageFile(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ── CORE HELPERS ── */
function _ss()   { return SpreadsheetApp.openById(SS_ID); }
function _sh(n)  { return _ss().getSheetByName(n); }
function _id(p)  { return p + "-" + Date.now(); }

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
  const d = v instanceof Date ? v : new Date(v);
  return !isNaN(d) && d.getMonth() === month && d.getFullYear() === year;
}

/* ================================================================
   BATCH LOAD — single round-trip for dashboard & pages
   Returns { projects, tasks, revisions, vendors, team }
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
   Accepts optional { year, month } — defaults to current month.
   Returns full breakdown by task type, work type, employee, vendor.
================================================================ */
function getMonthlySummary(params) {
  try {
    const now   = new Date();
    const year  = (params && params.year)  ? Number(params.year)  : now.getFullYear();
    const month = (params && params.month) ? Number(params.month) - 1 : now.getMonth(); // 0-based

    const tasks     = _sheetRows(SH_TASKS);
    const revisions = _sheetRows(SH_REVISIONS);

    /* filter tasks by start date falling in target month */
    const monthTasks = tasks.filter(r => _inMonth(r[TC.START_DATE], month, year));
    const monthRevs  = revisions.filter(r => _inMonth(r[RC.REV_DATE], month, year));

    /* ── by task type ── */
    const TASK_TYPES = ["Main DTP", "Content Extraction", "Corrections/Additions", "Bilingual Creation"];
    const TASK_KEYS  = { "Main DTP": "mainDTP", "Content Extraction": "extraction",
                         "Corrections/Additions": "corrections", "Bilingual Creation": "bilingual" };

    const byTaskType = {};
    TASK_TYPES.forEach(function(t) {
      const rows = monthTasks.filter(r => r[TC.TASK_TYPE] === t);
      const ih   = rows.filter(r => r[TC.WORK_TYPE] === "In-House").reduce((s, r) => s + (Number(r[TC.FINAL_PAGES]) || 0), 0);
      const vd   = rows.filter(r => r[TC.WORK_TYPE] === "Vendor").reduce((s, r)   => s + (Number(r[TC.FINAL_PAGES]) || 0), 0);
      byTaskType[TASK_KEYS[t]] = { label: t, total: ih + vd, inHouse: ih, vendor: vd, count: rows.length };
    });

    /* ── revisions ── */
    const revIH  = monthRevs.filter(r => r[RC.WORK_TYPE] === "In-House").reduce((s, r) => s + (Number(r[RC.REV_PAGES]) || 0), 0);
    const revVD  = monthRevs.filter(r => r[RC.WORK_TYPE] === "Vendor").reduce((s, r)   => s + (Number(r[RC.REV_PAGES]) || 0), 0);
    byTaskType.revisions = { label: "Revisions", total: revIH + revVD, inHouse: revIH, vendor: revVD, count: monthRevs.length };

    /* ── grand totals ── */
    const allTaskPages = monthTasks.reduce((s, r) => s + (Number(r[TC.FINAL_PAGES]) || 0), 0);
    const allRevPages  = monthRevs.reduce((s, r)  => s + (Number(r[RC.REV_PAGES])   || 0), 0);
    const grandTotal   = allTaskPages + allRevPages;
    const totalInHouse = monthTasks.filter(r => r[TC.WORK_TYPE] === "In-House").reduce((s, r) => s + (Number(r[TC.FINAL_PAGES]) || 0), 0) + revIH;
    const totalVendor  = monthTasks.filter(r => r[TC.WORK_TYPE] === "Vendor").reduce((s, r)   => s + (Number(r[TC.FINAL_PAGES]) || 0), 0) + revVD;

    /* ── by employee ── */
    const empMap = {};
    monthTasks.forEach(function(r) {
      const name = r[TC.ASSIGNED_TO] || "Unassigned";
      if (!empMap[name]) empMap[name] = { pages: 0, tasks: 0 };
      empMap[name].pages += Number(r[TC.FINAL_PAGES]) || 0;
      empMap[name].tasks++;
    });
    monthRevs.forEach(function(r) {
      const name = r[RC.ASSIGNED_TO] || "Unassigned";
      if (!empMap[name]) empMap[name] = { pages: 0, tasks: 0 };
      empMap[name].pages += Number(r[RC.REV_PAGES]) || 0;
      empMap[name].tasks++;
    });
    const byEmployee = Object.keys(empMap).map(n => ({ name: n, pages: empMap[n].pages, tasks: empMap[n].tasks }))
                             .sort((a, b) => b.pages - a.pages);

    /* ── by vendor ── */
    const vendMap = {};
    monthTasks.filter(r => r[TC.WORK_TYPE] === "Vendor").forEach(function(r) {
      const name = r[TC.VENDOR_NAME] || "Unknown Vendor";
      if (!vendMap[name]) vendMap[name] = { pages: 0, tasks: 0 };
      vendMap[name].pages += Number(r[TC.FINAL_PAGES]) || 0;
      vendMap[name].tasks++;
    });
    monthRevs.filter(r => r[RC.WORK_TYPE] === "Vendor").forEach(function(r) {
      const name = r[RC.VENDOR_NAME] || "Unknown Vendor";
      if (!vendMap[name]) vendMap[name] = { pages: 0, tasks: 0 };
      vendMap[name].pages += Number(r[RC.REV_PAGES]) || 0;
      vendMap[name].tasks++;
    });
    const byVendor = Object.keys(vendMap).map(n => ({ name: n, pages: vendMap[n].pages, tasks: vendMap[n].tasks }))
                           .sort((a, b) => b.pages - a.pages);

    /* ── by language ── */
    const langMap = {};
    monthTasks.forEach(function(r) {
      const lang = r[TC.LANGUAGE] || "Unknown";
      if (!langMap[lang]) langMap[lang] = 0;
      langMap[lang] += Number(r[TC.FINAL_PAGES]) || 0;
    });
    monthRevs.forEach(function(r) {
      const lang = r[RC.LANGUAGE] || "Unknown";
      if (!langMap[lang]) langMap[lang] = 0;
      langMap[lang] += Number(r[RC.REV_PAGES]) || 0;
    });
    const byLanguage = Object.keys(langMap).map(n => ({ language: n, pages: langMap[n] }))
                             .sort((a, b) => b.pages - a.pages);

    const monthLabel = Utilities.formatDate(
      new Date(year, month, 1), Session.getScriptTimeZone(), "MMMM yyyy"
    );

    return {
      month: monthLabel, year: year, monthIndex: month + 1,
      byTaskType, grandTotal, totalInHouse, totalVendor,
      byEmployee, byVendor, byLanguage,
      taskCount: monthTasks.length, revisionCount: monthRevs.length
    };
  } catch (e) {
    console.error("getMonthlySummary:", e);
    return { byTaskType: {}, grandTotal: 0, totalInHouse: 0, totalVendor: 0,
             byEmployee: [], byVendor: [], byLanguage: [], month: "", taskCount: 0, revisionCount: 0 };
  }
}

/* ================================================================
   PROJECT FULL SUMMARY — used by ViewProject page
   Returns project + all its tasks + all its revisions + page totals
================================================================ */
function getProjectSummary(projectId) {
  try {
    const pid      = String(projectId).trim();
    const projSh   = _sh(SH_PROJECTS);
    const projFound = _findRow(projSh, pid);
    if (!projFound) return null;

    const tasks     = _sheetRows(SH_TASKS).filter(r => String(r[TC.PROJECT_ID]).trim() === pid).map(_fmtRow);
    const revisions = _sheetRows(SH_REVISIONS).filter(r => String(r[RC.PROJECT_ID]).trim() === pid).map(_fmtRow);

    const taskPages = tasks.reduce((s, r) => s + (Number(r[TC.FINAL_PAGES]) || 0), 0);
    const revPages  = revisions.reduce((s, r) => s + (Number(r[RC.REV_PAGES]) || 0), 0);

    return {
      project:    _fmtRow(projFound.row),
      tasks:      tasks,
      revisions:  revisions,
      taskPages:  taskPages,
      revPages:   revPages,
      totalPages: taskPages + revPages
    };
  } catch (e) {
    console.error("getProjectSummary:", e);
    return null;
  }
}

/* ================================================================
   DROPDOWN DATA — vendors and team for form selects
================================================================ */
function getDropdownData() {
  try {
    const vendors = _sheetRows(SH_VENDORS).map(r => ({ id: r[0], name: r[1] || "", status: r[9] || "Active" }))
                                          .filter(v => v.status === "Active");
    const team    = _sheetRows(SH_TEAM).map(r => ({ id: r[0], name: r[1] || "", role: r[2] || "", status: r[6] || "Active" }))
                                       .filter(m => m.status === "Active");
    return { vendors, team };
  } catch (e) {
    console.error("getDropdownData:", e);
    return { vendors: [], team: [] };
  }
}
