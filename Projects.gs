/* ================================================================
   Projects.gs — Master Project Registry CRUD
   Sheet: Projects (14 columns)
   Col map: PC constants defined in Code.gs
   0=ID, 1=Client, 2=ProjectName, 3=Coordinator,
   4=SourceLang, 5=TargetLangs, 6=LangCount, 7=SourcePages,
   8=Priority, 9=Status, 10=ReceivedDate,
   11=Notes, 12=CreatedAt, 13=UpdatedAt
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

    /* ── Duplicate guard: block same client + project name (case-insensitive) ── */
    const clientNorm  = String(d.clientName  || "").toLowerCase().trim();
    const projectNorm = String(d.projectName || "").toLowerCase().trim();
    if (clientNorm && projectNorm) {
      const existing = _sheetRows(SH_PROJECTS).find(function(r) {
        return String(r[PC.CLIENT]       || "").toLowerCase().trim() === clientNorm &&
               String(r[PC.PROJECT_NAME] || "").toLowerCase().trim() === projectNorm;
      });
      if (existing) {
        throw new Error(
          'A project named "' + d.projectName + '" for client "' + d.clientName +
          '" already exists (ID: ' + existing[PC.ID] + '). Use that project or choose a different name.'
        );
      }
    }

    const id  = _id("PRJ");
    const now = new Date();
    sh.appendRow([
      id,
      d.clientName      || "",
      d.projectName     || "",
      d.coordinator     || "",
      d.sourceLanguage  || "",
      d.targetLanguages || "",
      Number(d.langCount)   || 0,
      Number(d.sourcePages) || 0,
      d.priority        || "Medium",
      d.status          || "Active",
      d.receivedDate    || "",
      d.notes           || "",
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

    /* ── Duplicate guard: block rename to an existing client+project combo ── */
    const clientNorm  = String(d.clientName  || "").toLowerCase().trim();
    const projectNorm = String(d.projectName || "").toLowerCase().trim();
    if (clientNorm && projectNorm) {
      const conflict = _sheetRows(SH_PROJECTS).find(function(r) {
        return String(r[PC.ID]).trim() !== String(id).trim() &&
               String(r[PC.CLIENT]       || "").toLowerCase().trim() === clientNorm &&
               String(r[PC.PROJECT_NAME] || "").toLowerCase().trim() === projectNorm;
      });
      if (conflict) {
        throw new Error(
          'Another project named "' + d.projectName + '" for client "' + d.clientName +
          '" already exists (ID: ' + conflict[PC.ID] + '). Choose a different project name.'
        );
      }
    }

    const now = new Date();
    const r   = found.row;
    sh.getRange(found.index, 2, 1, 13).setValues([[
      ("clientName"      in d) ? (d.clientName      || r[PC.CLIENT])        : r[PC.CLIENT],
      ("projectName"     in d) ? (d.projectName     || r[PC.PROJECT_NAME])  : r[PC.PROJECT_NAME],
      ("coordinator"     in d) ? d.coordinator                              : r[PC.COORDINATOR],
      ("sourceLanguage"  in d) ? d.sourceLanguage                           : r[PC.SOURCE_LANG],
      ("targetLanguages" in d) ? d.targetLanguages                          : r[PC.TARGET_LANGS],
      ("langCount"       in d) ? (Number(d.langCount)   || 0)               : (Number(r[PC.LANG_COUNT])   || 0),
      ("sourcePages"     in d) ? (Number(d.sourcePages) || 0)               : (Number(r[PC.SOURCE_PAGES]) || 0),
      ("priority"        in d) ? (d.priority        || r[PC.PRIORITY])      : r[PC.PRIORITY],
      ("status"          in d) ? (d.status          || r[PC.STATUS])        : r[PC.STATUS],
      ("receivedDate"    in d) ? d.receivedDate                             : r[PC.RECEIVED_DATE],
      ("notes"           in d) ? d.notes                                    : r[PC.NOTES],
      r[PC.CREATED_AT],
      now
    ]]);
    return { success: true, updated: _fmtRow(_findRow(sh, id).row) };
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
