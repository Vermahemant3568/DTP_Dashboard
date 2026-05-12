/* ================================================================
   Tasks.gs — Task CRUD (every unit of DTP work)
   Sheet: Tasks (22 columns)
   Col map: TC constants defined in Code.gs
   0=TaskID, 1=ProjectID, 2=Client, 3=ProjectName,
   4=TaskType, 5=WorkType, 6=AssignedTo, 7=VendorName,
   8=Language, 9=SourcePages, 10=FinalPages, 11=LangCount,
   12=Status, 13=Priority, 14=StartDate, 15=DeliveryDate,
   16=CompletedDate, 17=SourceLink, 18=DeliverableLink,
   19=Notes, 20=CreatedAt, 21=UpdatedAt
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
      Number(d.sourcePages)  || 0,
      Number(d.finalPages)   || 0,
      Number(d.langCount)    || 0,
      d.status            || "Pending",
      d.priority          || "Medium",
      d.startDate         || "",
      d.deliveryDate      || "",
      d.completedDate     || "",
      d.sourceLink        || "",
      d.deliverableLink   || "",
      d.notes             || "",
      now,
      now
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

    sh.getRange(found.index, 2, 1, 21).setValues([[
      found.row[TC.PROJECT_ID],          // preserve projectId
      d.clientName   || found.row[TC.CLIENT]        || "",
      d.projectName  || found.row[TC.PROJECT_NAME]  || "",
      d.taskType          || "Main DTP",
      d.workType          || "In-House",
      d.assignedTo        || "",
      d.vendorName        || "",
      d.language          || "",
      Number(d.sourcePages)  || 0,
      Number(d.finalPages)   || 0,
      Number(d.langCount)    || 0,
      d.status            || "Pending",
      d.priority          || "Medium",
      d.startDate         || "",
      d.deliveryDate      || "",
      d.completedDate     || "",
      d.sourceLink        || "",
      d.deliverableLink   || "",
      d.notes             || "",
      found.row[TC.CREATED_AT],          // preserve createdAt
      now
    ]]);
    return { success: true };
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
