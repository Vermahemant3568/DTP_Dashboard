/* ================================================================
   Projects.gs — Master Project Registry CRUD
   Sheet: Projects (15 columns)
   Col map: PC constants defined in Code.gs
   0=ID, 1=Client, 2=ProjectName, 3=Coordinator, 4=SourceLang,
   5=TargetLangs, 6=LangCount, 7=SourcePages, 8=WordCount,
   9=Priority, 10=Status, 11=ReceivedDate, 12=Notes,
   13=CreatedAt, 14=UpdatedAt
================================================================ */

function getProjects() {
  try {
    return _sheetRows(SH_PROJECTS).map(_fmtRow);
  } catch (e) {
    console.error("getProjects:", e);
    return [];
  }
}

function getProjectById(id) {
  try {
    const sh    = _sh(SH_PROJECTS);
    const found = _findRow(sh, id);
    return found ? { project: _fmtRow(found.row) } : null;
  } catch (e) {
    console.error("getProjectById:", e);
    return null;
  }
}

function addProject(d) {
  try {
    const sh  = _sh(SH_PROJECTS);
    if (!sh) throw new Error("Projects sheet not found.");
    const id  = _id("PRJ");
    const now = new Date();
    sh.appendRow([
      id,
      d.clientName        || "",
      d.projectName       || "",
      d.coordinator       || "",
      d.sourceLanguage    || "English",
      d.targetLanguages   || "",
      Number(d.langCount) || 0,
      Number(d.sourcePages) || 0,
      Number(d.wordCount) || 0,
      d.priority          || "Medium",
      d.status            || "Active",
      d.receivedDate      || "",
      d.notes             || "",
      now,
      now
    ]);
    return { success: true, id: id };
  } catch (e) {
    console.error("addProject:", e);
    throw e;
  }
}

function updateProject(id, d) {
  try {
    const sh    = _sh(SH_PROJECTS);
    const found = _findRow(sh, id);
    if (!found) throw new Error("Project not found: " + id);
    const now = new Date();
    sh.getRange(found.index, 2, 1, 14).setValues([[
      d.clientName        || "",
      d.projectName       || "",
      d.coordinator       || "",
      d.sourceLanguage    || "English",
      d.targetLanguages   || "",
      Number(d.langCount) || 0,
      Number(d.sourcePages) || 0,
      Number(d.wordCount) || 0,
      d.priority          || "Medium",
      d.status            || "Active",
      d.receivedDate      || "",
      d.notes             || "",
      found.row[PC.CREATED_AT], // preserve createdAt
      now
    ]]);
    return { success: true };
  } catch (e) {
    console.error("updateProject:", e);
    throw e;
  }
}

function updateProjectStatus(id, status) {
  try {
    const sh    = _sh(SH_PROJECTS);
    const found = _findRow(sh, id);
    if (!found) throw new Error("Project not found: " + id);
    sh.getRange(found.index, PC.STATUS + 1).setValue(status);
    sh.getRange(found.index, PC.UPDATED_AT + 1).setValue(new Date());
    return { success: true };
  } catch (e) {
    console.error("updateProjectStatus:", e);
    throw e;
  }
}

function deleteProject(id) {
  try {
    const sh    = _sh(SH_PROJECTS);
    const found = _findRow(sh, id);
    if (!found) throw new Error("Project not found: " + id);
    sh.deleteRow(found.index);
    return { success: true };
  } catch (e) {
    console.error("deleteProject:", e);
    throw e;
  }
}

/* Search projects by name or client — used by task form lookup */
function searchProjects(query) {
  try {
    const q = String(query || "").toLowerCase().trim();
    if (!q) return [];
    return _sheetRows(SH_PROJECTS)
      .filter(r => String(r[PC.ID]).toLowerCase().includes(q) ||
                   String(r[PC.CLIENT]).toLowerCase().includes(q) ||
                   String(r[PC.PROJECT_NAME]).toLowerCase().includes(q))
      .slice(0, 10)
      .map(_fmtRow);
  } catch (e) {
    console.error("searchProjects:", e);
    return [];
  }
}
