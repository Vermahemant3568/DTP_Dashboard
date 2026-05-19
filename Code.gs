/* ================================================================
   Code.gs — DTP Project Tracker
   Entry point, shared helpers, batch load, monthly summary.
   All CRUD is in: Projects.gs | Tasks.gs | RevisionsService.gs | VendorsService.gs
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

/* Sequential short ID: PRJ-001, TSK-042, etc.
   Reads the sheet to find the highest existing number for that prefix,
   then increments. Falls back to timestamp if sheet is unreadable. */
function _id(prefix) {
  try {
    var sheetName = prefix === "PRJ" ? SH_PROJECTS
                  : prefix === "TSK" ? SH_TASKS
                  : prefix === "REV" ? SH_REVISIONS
                  : prefix === "VND" ? SH_VENDORS
                  : prefix === "MBR" ? SH_TEAM
                  : null;
    var max = 0;
    if (sheetName) {
      var rows = _sheetRows(sheetName);
      rows.forEach(function(r) {
        var m = String(r[0]).match(/^[A-Z]+-0*(\d+)$/);
        if (m) max = Math.max(max, parseInt(m[1], 10));
      });
    }
    var next = max + 1;
    var padded = next < 1000 ? String(next).padStart(3, "0") : String(next);
    return prefix + "-" + padded;
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
  /* Handle dd/MM/yyyy format produced by _fmt */
  if (typeof v === "string") {
    var dmy = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (dmy) v = new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]));
    else v = new Date(v);
  }
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
   Accepts optional { year, month, tasks, revisions } —
   tasks/revisions can be pre-loaded arrays to avoid double reads.
   Returns full breakdown by task type, work type, employee, vendor.
================================================================ */
function getMonthlySummary(params) {
  try {
    const now   = new Date();
    const year  = (params && params.year)  ? Number(params.year)  : now.getFullYear();
    const month = (params && params.month) ? Number(params.month) - 1 : now.getMonth(); // 0-based

    /* Accept pre-loaded rows to avoid double spreadsheet reads */
    const tasks     = (params && params._tasks)     ? params._tasks     : _sheetRows(SH_TASKS);
    const revisions = (params && params._revisions) ? params._revisions : _sheetRows(SH_REVISIONS);

    /* filter tasks by startDate — fall back to createdAt if startDate is blank,
       so tasks without a start date are still counted in the month they were created */
    const monthTasks = tasks.filter(r => {
      var dateVal = r[TC.START_DATE] || r[TC.CREATED_AT];
      return _inMonth(dateVal, month, year);
    });
    /* filter revisions by rev date — fall back to createdAt if rev date is blank */
    const monthRevs  = revisions.filter(r => {
      var dateVal = r[RC.REV_DATE] || r[RC.CREATED_AT];
      return _inMonth(dateVal, month, year);
    });

    /* ── by task type ── */
    const TASK_TYPES = ["Main DTP", "Pre-Engineering", "Bilingual Creation", "QC", "Others"];
    const TASK_KEYS  = { "Main DTP": "mainDTP", "Pre-Engineering": "preEng",
                         "Bilingual Creation": "bilingual", "QC": "qc", "Others": "others" };

    const byTaskType = {};
    TASK_TYPES.forEach(function(t) {
      const rows = monthTasks.filter(r => r[TC.TASK_TYPE] === t);
      const ih   = rows.filter(r => r[TC.WORK_TYPE] === "In-House").reduce((s, r) => s + (Number(r[TC.FINAL_PAGES]) || 0), 0);
      const vd   = rows.filter(r => r[TC.WORK_TYPE] === "Vendor").reduce((s, r)   => s + (Number(r[TC.FINAL_PAGES]) || 0), 0);
      byTaskType[TASK_KEYS[t]] = { label: t, total: ih + vd, inHouse: ih, vendor: vd, count: rows.length };
    });

    /* ── revisions pages (tracked separately, included in grand totals) ── */
    const revIH    = monthRevs.filter(r => r[RC.WORK_TYPE] === "In-House").reduce((s, r) => s + (Number(r[RC.REV_PAGES]) || 0), 0);
    const revVD    = monthRevs.filter(r => r[RC.WORK_TYPE] === "Vendor").reduce((s, r)   => s + (Number(r[RC.REV_PAGES]) || 0), 0);
    const revTotal = revIH + revVD;
    byTaskType.revisions = { label: "Revisions", total: revTotal, inHouse: revIH, vendor: revVD, count: monthRevs.length };

    /* ── grand totals (tasks + revisions) ── */
    const allTaskPages = monthTasks.reduce((s, r) => s + (Number(r[TC.FINAL_PAGES]) || 0), 0);
    const allRevPages  = monthRevs.reduce((s, r)  => s + (Number(r[RC.REV_PAGES])   || 0), 0);
    const grandTotal   = allTaskPages + allRevPages;
    const totalInHouse = monthTasks.filter(r => r[TC.WORK_TYPE] === "In-House").reduce((s, r) => s + (Number(r[TC.FINAL_PAGES]) || 0), 0) + revIH;
    const totalVendor  = monthTasks.filter(r => r[TC.WORK_TYPE] === "Vendor").reduce((s, r)   => s + (Number(r[TC.FINAL_PAGES]) || 0), 0) + revVD;

    const monthLabel = Utilities.formatDate(
      new Date(year, month, 1), Session.getScriptTimeZone(), "MMMM yyyy"
    );

    return {
      month: monthLabel, year: year, monthIndex: month + 1,
      byTaskType, grandTotal, totalInHouse, totalVendor,
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
   DASHBOARD DATA — single spreadsheet read for everything.
   Returns:
     summary      — monthly page/task breakdown (for ROW 2 cards + charts)
     monthTasks   — tasks whose startDate is in selected month
     monthRevs    — revisions whose revDate is in selected month
     monthProjects— projects received in selected month
     allProjects  — every project (for active pipeline, status-based)
     allTimeTotals— counts/pages across ALL time (for ROW 1 summary cards)
================================================================ */
function getDashboardData(params) {
  try {
    const now   = new Date();
    const year  = (params && params.year)  ? Number(params.year)  : now.getFullYear();
    const month = (params && params.month) ? Number(params.month) - 1 : now.getMonth();

    /* ── Single read of every sheet ── */
    const allTasks     = _sheetRows(SH_TASKS).map(_fmtRow);
    const allRevisions = _sheetRows(SH_REVISIONS).map(_fmtRow);
    const allProjects  = _sheetRows(SH_PROJECTS).map(_fmtRow);

    /* ── Monthly summary ── */
    const summary = getMonthlySummary({
      year:        year,
      month:       month + 1,
      _tasks:      _sheetRows(SH_TASKS),
      _revisions:  _sheetRows(SH_REVISIONS)
    });

    /* ── All-time totals for ROW 1 summary cards ── */
    const totalAllTasks     = allTasks.length;
    const totalAllRevisions = allRevisions.length;
    const totalAllProjects  = allProjects.length;
    const activeProjects    = allProjects.filter(r => r[PC.STATUS] === "Active").length;
    const completedProjects = allProjects.filter(r => r[PC.STATUS] === "Completed").length;
    const pendingTasks      = allTasks.filter(r => r[TC.STATUS] === "Pending").length;
    const inProgressTasks   = allTasks.filter(r => r[TC.STATUS] === "In Progress").length;
    const pendingRevisions  = allRevisions.filter(r => r[RC.STATUS] === "Pending" || r[RC.STATUS] === "In Progress").length;
    const totalRevPages     = allRevisions.reduce((s, r) => s + (Number(r[RC.REV_PAGES]) || 0), 0);
    const inProgressPages   = allTasks
      .filter(r => r[TC.STATUS] === "In Progress")
      .reduce((s, r) => s + (Number(r[TC.FINAL_PAGES]) || 0), 0);

    return {
      summary,
      allTimeTotals: {
        totalProjects:    totalAllProjects,
        activeProjects,
        completedProjects,
        totalTasks:       totalAllTasks,
        pendingTasks,
        inProgressTasks,
        totalRevisions:   totalAllRevisions,
        pendingRevisions,
        totalRevPages,
        inProgressPages
      }
    };
  } catch(e) {
    console.error("getDashboardData:", e);
    return { summary: {}, tasks: [], revisions: [], projects: [], allProjects: [], allTimeTotals: {} };
  }
}
function getProjectsWithTaskCount() {
  try {
    const projects = _sheetRows(SH_PROJECTS).map(_fmtRow);
    const tasks    = _sheetRows(SH_TASKS);

    /* Build a map: projectId -> { count, pages } */
    const taskMap = {};
    tasks.forEach(function(r) {
      const pid = String(r[TC.PROJECT_ID]).trim();
      if (!taskMap[pid]) taskMap[pid] = { count: 0, pages: 0 };
      taskMap[pid].count++;
      taskMap[pid].pages += Number(r[TC.FINAL_PAGES]) || 0;
    });

    return projects.map(function(p) {
      const pid  = String(p[0]).trim();
      const info = taskMap[pid] || { count: 0, pages: 0 };
      return p.concat([info.count, info.pages]); // append taskCount, taskPages at end
    });
  } catch (e) {
    console.error("getProjectsWithTaskCount:", e);
    return [];
  }
}

/* ================================================================
   DROPDOWN DATA — vendors and team for form selects
================================================================ */
function getDropdownData() {
  try {
    var vendorRows = _sheetRows(SH_VENDORS);
    var teamRows   = _sheetRows(SH_TEAM);

    var vendors = vendorRows
      .map(function(r) {
        return {
          id:             String(r[0] || ""),
          name:           String(r[1] || "").trim(),
          contactPerson:  String(r[2] || ""),
          specialization: String(r[5] || ""),
          languages:      String(r[6] || ""),
          ratePerPage:    Number(r[7]) || 0,
          currency:       String(r[8] || ""),
          status:         String(r[9] || "Active").trim()
        };
      })
      .filter(function(v) {
        return v.name !== "" && v.status.toLowerCase() !== "inactive";
      });

    var team = teamRows
      .map(function(r) {
        return {
          id:             String(r[0] || ""),
          name:           String(r[1] || "").trim(),
          role:           String(r[2] || ""),
          specialization: String(r[5] || ""),
          ratePerPage:    Number(r[7]) || 0,
          currency:       String(r[8] || ""),
          status:         String(r[6] || "Active").trim()
        };
      })
      .filter(function(m) {
        return m.name !== "" && m.status.toLowerCase() !== "inactive";
      });

    return { vendors: vendors, team: team };
  } catch (e) {
    console.error("getDropdownData:", e);
    return { vendors: [], team: [] };
  }
}
