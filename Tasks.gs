/* ================================================================
   Tasks.gs — Task CRUD (every unit of DTP work)
   Sheet: Tasks (25 columns)
   Col map: TC constants defined in Code.gs
   0=TaskID, 1=ProjectID, 2=Client, 3=ProjectName,
   4=TaskType, 5=WorkType, 6=AssignedTo, 7=VendorName,
   8=Language, 9=SourcePages, 10=FinalPages, 11=LangCount,
   12=Status, 13=Priority, 14=StartDate, 15=DeliveryDate,
   16=CompletedDate, 17=SourceLink, 18=DeliverableLink,
   19=Notes, 20=CreatedAt, 21=UpdatedAt,
   22=RatePerPage, 23=Currency, 24=PaymentStatus
================================================================ */

function getTasks() {
  try {
    return _sheetRows(SH_TASKS).map(_fmtRow);
  } catch (e) {
    console.error("getTasks:", e);
    return [];
  }
}

function getTaskById(id) {
  try {
    const sh    = _sh(SH_TASKS);
    const found = _findRow(sh, id);
    return found ? { task: _fmtRow(found.row) } : null;
  } catch (e) {
    console.error("getTaskById:", e);
    return null;
  }
}

function getTaskWithRevisions(taskId) {
  try {
    var sh    = _sh(SH_TASKS);
    var found = _findRow(sh, taskId);
    if (!found) return null;
    var task = _fmtRow(found.row);
    var revisions = _sheetRows(SH_REVISIONS)
      .filter(function(r) { return String(r[RC.TASK_ID]).trim() === String(taskId).trim(); })
      .map(_fmtRow);
    return { task: task, revisions: revisions };
  } catch (e) {
    console.error("getTaskWithRevisions:", e);
    return null;
  }
}

function getTasksByProjectId(projectId) {
  try {
    const pid = String(projectId).trim();
    return _sheetRows(SH_TASKS)
      .filter(r => String(r[TC.PROJECT_ID]).trim() === pid)
      .map(_fmtRow);
  } catch (e) {
    console.error("getTasksByProjectId:", e);
    return [];
  }
}

function addTask(d) {
  try {
    const sh  = _sh(SH_TASKS);
    if (!sh) throw new Error("Tasks sheet not found.");

    /* Resolve project details if projectId provided */
    let clientName   = d.clientName   || "";
    let projectName  = d.projectName  || "";

    if (d.projectId && (!clientName || !projectName)) {
      const projSh    = _sh(SH_PROJECTS);
      const projFound = _findRow(projSh, d.projectId);
      if (projFound) {
        clientName  = clientName  || projFound.row[PC.CLIENT]       || "";
        projectName = projectName || projFound.row[PC.PROJECT_NAME] || "";
      }
    }

    const id  = _id("TSK");
    const now = new Date();

    /* Auto-calculate finalPages = sourcePages × langCount if not manually overridden */
    var srcPages  = Number(d.sourcePages) || 0;
    var langCount = Number(d.langCount)   || 0;
    var finalPgs  = Number(d.finalPages)  || 0;
    /* For DTP/Extraction/QC/Bilingual: if langCount > 0 and finalPages not manually set, auto-calc */
    if (langCount > 0 && (!d.finalPages || Number(d.finalPages) === 0)) {
      finalPgs = srcPages * langCount;
    }

    sh.appendRow([
      id,
      d.projectId         || "",
      clientName,
      projectName,
      d.taskType          || "Main DTP",
      d.workType          || "In-House",
      d.assignedTo        || "",
      d.vendorName        || "",
      d.language          || "",
      srcPages,
      finalPgs,
      langCount,
      d.status            || "Pending",
      d.priority          || "Medium",
      d.startDate         || now,
      d.deliveryDate      || "",
      d.completedDate     || "",
      d.sourceLink        || "",
      d.deliverableLink   || "",
      d.notes             || "",
      now,
      now,
      Number(d.ratePerPage)  || 0,
      d.currency             || "",
      d.paymentStatus        || "Unpaid"
    ]);

    /* Update parent project status to Active if it was Pending */
    if (d.projectId) {
      const projSh    = _sh(SH_PROJECTS);
      const projFound = _findRow(projSh, d.projectId);
      if (projFound && projFound.row[PC.STATUS] === "Pending") {
        projSh.getRange(projFound.index, PC.STATUS + 1).setValue("Active");
        projSh.getRange(projFound.index, PC.UPDATED_AT + 1).setValue(now);
      }
    }

    return { success: true, id: id };
  } catch (e) {
    console.error("addTask:", e);
    throw e;
  }
}

function updateTask(id, d) {
  try {
    const sh    = _sh(SH_TASKS);
    const found = _findRow(sh, id);
    if (!found) throw new Error("Task not found: " + id);
    const now = new Date();
    const r   = found.row;

    /* Only recalculate finalPages if sourcePages or langCount was explicitly sent */
    var srcPages  = ("sourcePages" in d) ? (Number(d.sourcePages) || 0) : (Number(r[TC.SOURCE_PAGES]) || 0);
    var langCount = ("langCount"   in d) ? (Number(d.langCount)   || 0) : (Number(r[TC.LANG_COUNT])   || 0);
    var finalPgs;
    if ("finalPages" in d && d.finalPages !== "" && d.finalPages !== null) {
      finalPgs = Number(d.finalPages) || 0;
    } else if (("sourcePages" in d || "langCount" in d) && langCount > 0) {
      finalPgs = srcPages * langCount;
    } else {
      finalPgs = Number(r[TC.FINAL_PAGES]) || 0;
    }

    sh.getRange(found.index, 2, 1, 24).setValues([[
      r[TC.PROJECT_ID],
      r[TC.CLIENT],
      r[TC.PROJECT_NAME],
      ("taskType"        in d) ? (d.taskType        || r[TC.TASK_TYPE])        : r[TC.TASK_TYPE],
      ("workType"        in d) ? (d.workType        || r[TC.WORK_TYPE])        : r[TC.WORK_TYPE],
      ("assignedTo"      in d) ? d.assignedTo                                  : r[TC.ASSIGNED_TO],
      ("vendorName"      in d) ? d.vendorName                                  : r[TC.VENDOR_NAME],
      ("language"        in d) ? d.language                                    : r[TC.LANGUAGE],
      srcPages,
      finalPgs,
      langCount,
      ("status"          in d) ? (d.status          || r[TC.STATUS])           : r[TC.STATUS],
      ("priority"        in d) ? (d.priority        || r[TC.PRIORITY])         : r[TC.PRIORITY],
      ("startDate"       in d) ? (d.startDate       || r[TC.START_DATE])       : r[TC.START_DATE],
      ("deliveryDate"    in d) ? d.deliveryDate                                : r[TC.DELIVERY_DATE],
      ("completedDate"   in d) ? d.completedDate                               : r[TC.COMPLETED_DATE],
      ("sourceLink"      in d) ? d.sourceLink                                  : r[TC.SOURCE_LINK],
      ("deliverableLink" in d) ? d.deliverableLink                             : r[TC.DELIVERABLE_LINK],
      ("notes"           in d) ? d.notes                                       : r[TC.NOTES],
      r[TC.CREATED_AT],
      now,
      ("ratePerPage"     in d) ? (Number(d.ratePerPage) || r[TC.RATE_PER_PAGE] || 0) : (Number(r[TC.RATE_PER_PAGE]) || 0),
      ("currency"        in d) ? d.currency                                    : r[TC.CURRENCY],
      ("paymentStatus"   in d) ? (d.paymentStatus   || r[TC.PAYMENT_STATUS])   : r[TC.PAYMENT_STATUS]
    ]]);
    /* Return the full updated row so the frontend can patch in-memory cache */
    return { success: true, updated: _fmtRow(_findRow(sh, id).row) };
  } catch (e) {
    console.error("updateTask:", e);
    throw e;
  }
}

function updateTaskStatus(id, status, deliveryDate) {
  try {
    const sh    = _sh(SH_TASKS);
    const found = _findRow(sh, id);
    if (!found) throw new Error("Task not found: " + id);
    const now = new Date();
    sh.getRange(found.index, TC.STATUS + 1).setValue(status);
    if (deliveryDate) sh.getRange(found.index, TC.DELIVERY_DATE + 1).setValue(deliveryDate);
    if (status === "Completed") sh.getRange(found.index, TC.COMPLETED_DATE + 1).setValue(now);
    sh.getRange(found.index, TC.UPDATED_AT + 1).setValue(now);
    return { success: true };
  } catch (e) {
    console.error("updateTaskStatus:", e);
    throw e;
  }
}

function deleteTask(id) {
  try {
    const sh    = _sh(SH_TASKS);
    const found = _findRow(sh, id);
    if (!found) throw new Error("Task not found: " + id);
    sh.deleteRow(found.index);
    return { success: true };
  } catch (e) {
    console.error("deleteTask:", e);
    throw e;
  }
}

/* ================================================================
   ADD TASK WITH AUTO PROJECT CREATION
   Checks for an existing project with same clientName + projectName
   (case-insensitive) before creating. If a match is found, reuses it.
   This prevents duplicate projects from double-clicks or re-submissions.
================================================================ */
function addTaskWithProject(projectData, taskData) {
  try {
    const projSh = _sh(SH_PROJECTS);
    if (!projSh) throw new Error("Projects sheet not found.");
    const now = new Date();

    /* ── 1. Duplicate guard: find existing project by client + name ── */
    const clientNorm  = String(projectData.clientName  || "").toLowerCase().trim();
    const projectNorm = String(projectData.projectName || "").toLowerCase().trim();
    const allProjects = _sheetRows(SH_PROJECTS);
    const existing    = allProjects.find(function(r) {
      return String(r[PC.CLIENT]       || "").toLowerCase().trim() === clientNorm &&
             String(r[PC.PROJECT_NAME] || "").toLowerCase().trim() === projectNorm;
    });

    let projId;
    if (existing) {
      /* Reuse the existing project — do NOT create a duplicate */
      projId = String(existing[PC.ID]);
    } else {
      /* ── 2. Create the new project ── */
      projId = _id("PRJ");
      projSh.appendRow([
        projId,
        projectData.clientName      || "",
        projectData.projectName     || "",
        projectData.coordinator     || "",
        projectData.sourceLanguage  || "English",
        projectData.targetLanguages || "",
        Number(projectData.langCount)   || 0,
        Number(projectData.sourcePages) || 0,
        Number(projectData.wordCount)   || 0,
        projectData.priority        || "Medium",
        "Active",
        projectData.receivedDate    || "",
        projectData.notes           || "",
        now,
        now
      ]);
    }

    /* ── 3. Create the task linked to the project ── */
    taskData.projectId   = projId;
    taskData.clientName  = projectData.clientName  || "";
    taskData.projectName = projectData.projectName || "";
    const taskResult = addTask(taskData);

    return { success: true, projectId: projId, taskId: taskResult.id, reused: !!existing };
  } catch (e) {
    console.error("addTaskWithProject:", e);
    throw e;
  }
}
