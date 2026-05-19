/* ================================================================
   RevisionsService.gs — Revision CRUD
   Sheet: Revisions (21 columns)
   Col map: RC constants defined in Code.gs
   0=RevID, 1=ProjectID, 2=TaskID, 3=ProjectName,
   4=RevNumber, 5=RevType, 6=Language, 7=RevPages,
   8=WorkType, 9=AssignedTo, 10=VendorName, 11=Status,
   12=RevDate, 13=DeliveryDate, 14=CompletedDate, 15=Notes,
   16=CreatedAt, 17=UpdatedAt,
   18=RatePerPage, 19=Currency, 20=PaymentStatus
================================================================ */

function getRevisions() {
  try {
    return _sheetRows(SH_REVISIONS).map(_fmtRow);
  } catch (e) {
    console.error("getRevisions:", e);
    return [];
  }
}

function getRevisionById(id) {
  try {
    var sh    = _sh(SH_REVISIONS);
    var found = _findRow(sh, id);
    return found ? { revision: _fmtRow(found.row) } : null;
  } catch (e) {
    console.error("getRevisionById:", e);
    return null;
  }
}

function getRevisionsByProjectId(projectId) {
  try {
    var pid = String(projectId).trim();
    return _sheetRows(SH_REVISIONS)
      .filter(function(r) { return String(r[RC.PROJECT_ID]).trim() === pid; })
      .map(_fmtRow);
  } catch (e) {
    console.error("getRevisionsByProjectId:", e);
    return [];
  }
}

function addRevision(d) {
  try {
    var sh = _sh(SH_REVISIONS);
    if (!sh) throw new Error("Revisions sheet not found.");

    /* Resolve project name */
    var projectName = d.projectName || "";
    if (d.projectId && !projectName) {
      var projSh    = _sh(SH_PROJECTS);
      var projFound = _findRow(projSh, d.projectId);
      if (projFound) projectName = projFound.row[PC.PROJECT_NAME] || "";
    }

    /* If a taskId is provided, inherit any missing fields from the task */
    if (d.taskId) {
      var taskSh    = _sh(SH_TASKS);
      var taskFound = _findRow(taskSh, d.taskId);
      if (taskFound) {
        var tr = taskFound.row;
        if (!d.language   || d.language   === "") d.language   = tr[TC.LANGUAGE]    || "";
        if (!d.workType   || d.workType   === "") d.workType   = tr[TC.WORK_TYPE]   || "In-House";
        if (!d.assignedTo || d.assignedTo === "") d.assignedTo = tr[TC.ASSIGNED_TO] || "";
        if (!d.vendorName || d.vendorName === "") d.vendorName = tr[TC.VENDOR_NAME] || "";
        if ((!d.ratePerPage || Number(d.ratePerPage) === 0) && tr[TC.RATE_PER_PAGE]) {
          d.ratePerPage = tr[TC.RATE_PER_PAGE];
          d.currency    = d.currency || tr[TC.CURRENCY] || "";
        }
      }
    }

    /* Auto-assign revision number scoped to project+task */
    var revNumber = d.revisionNumber || "";
    if (!revNumber && d.projectId) {
      var existing = _sheetRows(SH_REVISIONS).filter(function(r) {
        var sameProj = String(r[RC.PROJECT_ID]).trim() === String(d.projectId).trim();
        var sameTask = d.taskId ? String(r[RC.TASK_ID]).trim() === String(d.taskId).trim() : true;
        return sameProj && sameTask;
      });
      revNumber = "R" + (existing.length + 1);
    }

    var id  = _id("REV");
    var now = new Date();

    sh.appendRow([
      id,
      d.projectId       || "",
      d.taskId          || "",
      projectName,
      revNumber,
      d.revisionType    || "Client Feedback",
      d.language        || "",
      Number(d.revisionPages) || 0,
      d.workType        || "In-House",
      d.assignedTo      || "",
      d.vendorName      || "",
      d.status          || "Pending",
      d.revisionDate    || "",
      d.deliveryDate    || "",
      d.completedDate   || "",
      d.notes           || "",
      now,
      now,
      Number(d.ratePerPage)  || 0,
      d.currency             || "",
      d.paymentStatus        || "Unpaid"
    ]);
    return { success: true, id: id };
  } catch (e) {
    console.error("addRevision:", e);
    throw e;
  }
}

function updateRevision(id, d) {
  try {
    var sh    = _sh(SH_REVISIONS);
    var found = _findRow(sh, id);
    if (!found) throw new Error("Revision not found: " + id);
    var now = new Date();
    var r   = found.row;

    sh.getRange(found.index, 2, 1, 20).setValues([[
      r[RC.PROJECT_ID],
      r[RC.TASK_ID],
      r[RC.PROJECT_NAME],
      ("revisionNumber" in d) ? (d.revisionNumber || r[RC.REV_NUMBER])      : r[RC.REV_NUMBER],
      ("revisionType"   in d) ? (d.revisionType   || r[RC.REV_TYPE])        : r[RC.REV_TYPE],
      ("language"       in d) ? d.language                                  : r[RC.LANGUAGE],
      ("revisionPages"  in d) ? (Number(d.revisionPages) || 0)              : (Number(r[RC.REV_PAGES]) || 0),
      ("workType"       in d) ? (d.workType        || r[RC.WORK_TYPE])      : r[RC.WORK_TYPE],
      ("assignedTo"     in d) ? d.assignedTo                                : r[RC.ASSIGNED_TO],
      ("vendorName"     in d) ? d.vendorName                                : r[RC.VENDOR_NAME],
      ("status"         in d) ? (d.status          || r[RC.STATUS])         : r[RC.STATUS],
      ("revisionDate"   in d) ? d.revisionDate                              : r[RC.REV_DATE],
      ("deliveryDate"   in d) ? d.deliveryDate                              : r[RC.DELIVERY_DATE],
      ("completedDate"  in d) ? d.completedDate                             : r[RC.COMPLETED_DATE],
      ("notes"          in d) ? d.notes                                     : r[RC.NOTES],
      r[RC.CREATED_AT],
      now,
      ("ratePerPage"    in d) ? (Number(d.ratePerPage) || r[RC.RATE_PER_PAGE] || 0) : (Number(r[RC.RATE_PER_PAGE]) || 0),
      ("currency"       in d) ? d.currency                                  : r[RC.CURRENCY],
      ("paymentStatus"  in d) ? (d.paymentStatus   || r[RC.PAYMENT_STATUS]) : r[RC.PAYMENT_STATUS]
    ]]);
    /* Return the full updated row so the frontend can patch in-memory cache */
    return { success: true, updated: _fmtRow(_findRow(sh, id).row) };
  } catch (e) {
    console.error("updateRevision:", e);
    throw e;
  }
}

function updateRevisionStatus(id, status, deliveryDate) {
  try {
    var sh    = _sh(SH_REVISIONS);
    var found = _findRow(sh, id);
    if (!found) throw new Error("Revision not found: " + id);
    var now = new Date();
    sh.getRange(found.index, RC.STATUS + 1).setValue(status);
    if (deliveryDate) sh.getRange(found.index, RC.DELIVERY_DATE + 1).setValue(deliveryDate);
    if (status === "Completed") sh.getRange(found.index, RC.COMPLETED_DATE + 1).setValue(now);
    sh.getRange(found.index, RC.UPDATED_AT + 1).setValue(now);
    return { success: true };
  } catch (e) {
    console.error("updateRevisionStatus:", e);
    throw e;
  }
}

function deleteRevision(id) {
  try {
    var sh    = _sh(SH_REVISIONS);
    var found = _findRow(sh, id);
    if (!found) throw new Error("Revision not found: " + id);
    sh.deleteRow(found.index);
    return { success: true };
  } catch (e) {
    console.error("deleteRevision:", e);
    throw e;
  }
}
