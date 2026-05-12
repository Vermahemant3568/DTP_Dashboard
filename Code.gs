/* ================================================================
   Code.gs — DTP Project Tracker
   Single spreadsheet, two sheets: Projects + Tasks
   Tasks covers: DTP, Revision, Correction, Addition, Extraction, Bilingual
================================================================ */

const SS_ID        = "1-jcai8DeFYpHmfzWGjcYmA4Ehbt0ASO3Ue240P0HGM4";
const SH_PROJECTS  = "Projects";
const SH_TASKS     = "Tasks";

/* ----------------------------------------------------------------
   ENTRY POINT
---------------------------------------------------------------- */
function doGet() {
  return HtmlService.createTemplateFromFile("Layout")
    .evaluate()
    .setTitle("DTP Project Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ----------------------------------------------------------------
   HELPERS
---------------------------------------------------------------- */
function _ss() {
  return SpreadsheetApp.openById(SS_ID);
}

function _fmt(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy");
  return (v === null || v === undefined) ? "" : v;
}

function _fmtRow(row) { return row.map(_fmt); }

function _findRow(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) return { row: data[i], index: i + 1 };
  }
  return null;
}

/* ----------------------------------------------------------------
   BATCH LOAD — single server call for all data
   Returns projects + tasks. Dashboard computes summaries client-side.
---------------------------------------------------------------- */
function getAllData() {
  try {
    const ss = _ss();
    const pSheet = ss.getSheetByName(SH_PROJECTS);
    const tSheet = ss.getSheetByName(SH_TASKS);

    const pData = pSheet ? pSheet.getDataRange().getValues() : [];
    const tData = tSheet ? tSheet.getDataRange().getValues() : [];

    return {
      projects: pData.length > 1 ? pData.slice(1).filter(r => r[0] !== "").map(_fmtRow) : [],
      tasks:    tData.length > 1 ? tData.slice(1).filter(r => r[0] !== "").map(_fmtRow) : []
    };
  } catch (e) {
    console.error("getAllData:", e);
    return { projects: [], tasks: [] };
  }
}

/* ----------------------------------------------------------------
   PROJECTS
   Columns (0-based):
   0 projectId  1 clientName  2 projectName  3 coordinator
   4 startDate  5 dueDate     6 status       7 sourcePages
   8 sourceLang 9 notes       10 createdAt
---------------------------------------------------------------- */
function addProject(d) {
  const sheet = _ss().getSheetByName(SH_PROJECTS);
  const id    = "PRJ-" + Date.now();
  sheet.appendRow([
    id,
    d.clientName    || "",
    d.projectName   || "",
    d.coordinator   || "",
    d.startDate     || "",
    d.dueDate       || "",
    d.status        || "Pending",
    Number(d.sourcePages) || 0,
    d.sourceLang    || "English",
    d.notes         || "",
    new Date()
  ]);
  return { success: true, id };
}

function updateProject(id, d) {
  const sheet  = _ss().getSheetByName(SH_PROJECTS);
  const found  = _findRow(sheet, id);
  if (!found) throw new Error("Project not found: " + id);
  sheet.getRange(found.index, 2, 1, 9).setValues([[
    d.clientName    || "",
    d.projectName   || "",
    d.coordinator   || "",
    d.startDate     || "",
    d.dueDate       || "",
    d.status        || "Pending",
    Number(d.sourcePages) || 0,
    d.sourceLang    || "English",
    d.notes         || ""
  ]]);
  return { success: true };
}

function deleteProject(id) {
  const sheet = _ss().getSheetByName(SH_PROJECTS);
  const found = _findRow(sheet, id);
  if (!found) throw new Error("Project not found: " + id);
  sheet.deleteRow(found.index);
  return { success: true };
}

function getProjectById(id) {
  const sheet = _ss().getSheetByName(SH_PROJECTS);
  const found = _findRow(sheet, id);
  return found ? { details: _fmtRow(found.row) } : null;
}

/* ----------------------------------------------------------------
   TASKS
   Columns (0-based):
   0 taskId       1 projectId    2 projectName   3 clientName
   4 taskType     5 language     6 pages         7 assignedPerson
   8 workType     9 vendorName   10 status       11 startDate
   12 deliveryDate 13 notes      14 createdAt

   taskType values: DTP | Revision | Correction | Addition | Extraction | Bilingual
   workType values: In-House | Vendor
---------------------------------------------------------------- */
function addTask(d) {
  const sheet = _ss().getSheetByName(SH_TASKS);
  const id    = "TSK-" + Date.now();
  sheet.appendRow([
    id,
    d.projectId     || "",
    d.projectName   || "",
    d.clientName    || "",
    d.taskType      || "DTP",
    d.language      || "",
    Number(d.pages) || 0,
    d.assignedPerson|| "",
    d.workType      || "In-House",
    d.vendorName    || "",
    d.status        || "Pending",
    d.startDate     || "",
    d.deliveryDate  || "",
    d.notes         || "",
    new Date()
  ]);
  return { success: true, id };
}

function updateTask(id, d) {
  const sheet = _ss().getSheetByName(SH_TASKS);
  const found = _findRow(sheet, id);
  if (!found) throw new Error("Task not found: " + id);
  sheet.getRange(found.index, 2, 1, 13).setValues([[
    d.projectId     || "",
    d.projectName   || "",
    d.clientName    || "",
    d.taskType      || "DTP",
    d.language      || "",
    Number(d.pages) || 0,
    d.assignedPerson|| "",
    d.workType      || "In-House",
    d.vendorName    || "",
    d.status        || "Pending",
    d.startDate     || "",
    d.deliveryDate  || "",
    d.notes         || ""
  ]]);
  return { success: true };
}

function updateTaskStatus(id, status, deliveryDate) {
  const sheet = _ss().getSheetByName(SH_TASKS);
  const found = _findRow(sheet, id);
  if (!found) throw new Error("Task not found: " + id);
  sheet.getRange(found.index, 11).setValue(status);
  sheet.getRange(found.index, 13).setValue(deliveryDate);
  return { success: true };
}

function deleteTask(id) {
  const sheet = _ss().getSheetByName(SH_TASKS);
  const found = _findRow(sheet, id);
  if (!found) throw new Error("Task not found: " + id);
  sheet.deleteRow(found.index);
  return { success: true };
}

function getTasksByProject(projectId) {
  try {
    const sheet = _ss().getSheetByName(SH_TASKS);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const pid  = String(projectId).trim();
    return data.slice(1)
      .filter(r => r[0] !== "" && String(r[1]).trim() === pid)
      .map(_fmtRow);
  } catch (e) {
    console.error("getTasksByProject:", e);
    return [];
  }
}
