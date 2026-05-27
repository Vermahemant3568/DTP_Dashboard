/* ================================================================
   VendorsService.gs — Vendor & Team CRUD + Performance Analytics
   Sheets: Vendors (13 cols) | Team (10 cols)
================================================================ */

/* ── VENDORS ── */

function getVendors() {
  try {
    return _sheetRows(SH_VENDORS).map(_fmtRow);
  } catch (e) {
    console.error("getVendors:", e);
    return [];
  }
}

function addVendor(d) {
  try {
    var sh  = _sh(SH_VENDORS);
    if (!sh) throw new Error("Vendors sheet not found.");
    var id  = _id("VND");
    var now = new Date();
    sh.appendRow([
      id,
      d.vendorName       || "",
      d.contactPerson    || "",
      d.email            || "",
      d.phone            || "",
      d.specialization   || "",
      d.languages        || "",
      Number(d.ratePerPage) || 0,
      d.currency         || "USD",
      d.status           || "Active",
      d.notes            || "",
      now,
      now
    ]);
    return { success: true, id: id };
  } catch (e) {
    console.error("addVendor:", e);
    throw e;
  }
}

function updateVendor(id, d) {
  try {
    var sh    = _sh(SH_VENDORS);
    var found = _findRow(sh, id);
    if (!found) throw new Error("Vendor not found: " + id);
    var now = new Date();
    sh.getRange(found.index, 2, 1, 12).setValues([[
      d.vendorName       || "",
      d.contactPerson    || "",
      d.email            || "",
      d.phone            || "",
      d.specialization   || "",
      d.languages        || "",
      Number(d.ratePerPage) || 0,
      d.currency         || "USD",
      d.status           || "Active",
      d.notes            || "",
      found.row[VC.CREATED_AT],
      now
    ]]);
    return { success: true };
  } catch (e) {
    console.error("updateVendor:", e);
    throw e;
  }
}

function deleteVendor(id) {
  try {
    var sh    = _sh(SH_VENDORS);
    var found = _findRow(sh, id);
    if (!found) throw new Error("Vendor not found: " + id);
    sh.deleteRow(found.index);
    return { success: true };
  } catch (e) {
    console.error("deleteVendor:", e);
    throw e;
  }
}

/* ── TEAM ── */

function getTeam() {
  try {
    return _sheetRows(SH_TEAM).map(_fmtRow);
  } catch (e) {
    console.error("getTeam:", e);
    return [];
  }
}

function addTeamMember(d) {
  try {
    var sh  = _sh(SH_TEAM);
    if (!sh) throw new Error("Team sheet not found.");
    var id  = _id("MBR");
    var now = new Date();
    sh.appendRow([
      id,
      d.name             || "",
      d.role             || "",
      d.email            || "",
      d.phone            || "",
      d.specialization   || "",
      d.status           || "Active",
      Number(d.ratePerPage) || 0,
      d.currency         || "",
      now
    ]);
    return { success: true, id: id };
  } catch (e) {
    console.error("addTeamMember:", e);
    throw e;
  }
}

function updateTeamMember(id, d) {
  try {
    var sh    = _sh(SH_TEAM);
    var found = _findRow(sh, id);
    if (!found) throw new Error("Team member not found: " + id);
    sh.getRange(found.index, 2, 1, 9).setValues([[
      d.name             || "",
      d.role             || "",
      d.email            || "",
      d.phone            || "",
      d.specialization   || "",
      d.status           || "Active",
      Number(d.ratePerPage) || 0,
      d.currency         || "",
      found.row[MC.CREATED_AT]
    ]]);
    return { success: true };
  } catch (e) {
    console.error("updateTeamMember:", e);
    throw e;
  }
}

function deleteTeamMember(id) {
  try {
    var sh    = _sh(SH_TEAM);
    var found = _findRow(sh, id);
    if (!found) throw new Error("Team member not found: " + id);
    sh.deleteRow(found.index);
    return { success: true };
  } catch (e) {
    console.error("deleteTeamMember:", e);
    throw e;
  }
}

/* ================================================================
   SHARED HELPER — parse a date string or Date into a "YYYY-MM" key.
   Used by both performance functions to build monthly charts.
================================================================ */
function _toMonthKey(v) {
  if (!v) return null;
  if (typeof v === "string") {
    var dmy = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    v = dmy ? new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1])) : new Date(v);
  }
  if (!(v instanceof Date) || isNaN(v)) return null;
  return v.getFullYear() + "-" + String(v.getMonth() + 1).padStart(2, "0");
}

/* ================================================================
   SHARED HELPER — build breakdown maps used by both performance fns.
   Returns { monthly, byProject, byTaskType, byLanguage, rows,
             projectOptions, langOptions }
   allItems  = unfiltered rows for this person (for monthly + dropdowns)
   fItems    = filtered rows (for breakdowns + detail table)
   Each item: { pages, date, projectId, projectName, language,
                taskType, status, paymentStatus, rate, currency,
                id, workStatus, completedDate, startDate, type }
================================================================ */
function _buildPerformanceData(allItems, fItems, defaultCurrency) {

  /* Monthly chart — always unfiltered, last 12 months */
  var monthlyMap = {};
  allItems.forEach(function(item) {
    var key = _toMonthKey(item.date);
    if (!key) return;
    if (!monthlyMap[key]) monthlyMap[key] = { pages: 0, count: 0 };
    monthlyMap[key].pages += item.pages;
    monthlyMap[key].count++;
  });
  var monthly = Object.keys(monthlyMap).sort().slice(-12).map(function(k) {
    var parts = k.split("-");
    var label = Utilities.formatDate(
      new Date(Number(parts[0]), Number(parts[1]) - 1, 1),
      Session.getScriptTimeZone(), "MMM yyyy"
    );
    return { key: k, label: label, pages: monthlyMap[k].pages, count: monthlyMap[k].count };
  });

  /* Project breakdown with payment totals — filtered */
  var projMap = {};
  fItems.forEach(function(item) {
    var pid = item.projectId;
    if (!projMap[pid]) projMap[pid] = {
      projectId: pid, projectName: item.projectName,
      pages: 0, count: 0, completed: 0,
      paidAmt: 0, partialAmt: 0, unpaidAmt: 0
    };
    var amt = item.pages * item.rate;
    projMap[pid].pages += item.pages;
    projMap[pid].count++;
    if (item.workStatus === "Completed") projMap[pid].completed++;
    if (item.paymentStatus === "Paid")         projMap[pid].paidAmt    += amt;
    else if (item.paymentStatus === "Partial") projMap[pid].partialAmt += amt;
    else if (item.paymentStatus === "Unpaid")  projMap[pid].unpaidAmt  += amt;
  });
  var byProject = Object.values(projMap).sort(function(a, b) { return b.pages - a.pages; });

  /* Task-type breakdown — filtered */
  var typeMap = {};
  fItems.forEach(function(item) {
    var t = item.taskType || "Other";
    if (!typeMap[t]) typeMap[t] = { pages: 0, count: 0 };
    typeMap[t].pages += item.pages;
    typeMap[t].count++;
  });
  var byTaskType = Object.keys(typeMap).map(function(k) {
    return { type: k, pages: typeMap[k].pages, count: typeMap[k].count };
  }).sort(function(a, b) { return b.pages - a.pages; });

  /* Language breakdown — filtered */
  var langMap = {};
  fItems.forEach(function(item) {
    var l = item.language || "Unknown";
    if (!langMap[l]) langMap[l] = 0;
    langMap[l] += item.pages;
  });
  var byLanguage = Object.keys(langMap).map(function(l) {
    return { language: l, pages: langMap[l] };
  }).sort(function(a, b) { return b.pages - a.pages; });

  /* Detail rows — filtered, sorted newest first */
  var rows = fItems.map(function(item) {
    return {
      id:            item.id,
      projectId:     item.projectId,
      projectName:   item.projectName,
      taskType:      item.taskType,
      language:      item.language,
      pages:         item.pages,
      status:        item.workStatus,
      startDate:     item.startDate,
      completedDate: item.completedDate,
      ratePerPage:   item.rate,
      currency:      item.currency || defaultCurrency,
      amount:        item.pages * item.rate,
      paymentStatus: item.paymentStatus,
      type:          item.type
    };
  }).sort(function(a, b) {
    return String(b.startDate).localeCompare(String(a.startDate));
  });

  /* Filter dropdown options — from unfiltered data */
  var seenPids = {}, projectOptions = [];
  allItems.forEach(function(item) {
    if (item.projectId && !seenPids[item.projectId]) {
      seenPids[item.projectId] = true;
      projectOptions.push({ id: item.projectId, name: item.projectName });
    }
  });
  var seenLangs = {}, langOptions = [];
  allItems.forEach(function(item) {
    if (item.language && !seenLangs[item.language]) {
      seenLangs[item.language] = true;
      langOptions.push(item.language);
    }
  });

  return { monthly, byProject, byTaskType, byLanguage, rows, projectOptions, langOptions };
}

/* ================================================================
   VENDOR PERFORMANCE
   Accepts: vendorName (string), filters { month, year, projectId, language, status }
================================================================ */
function getVendorPerformance(vendorName, filters) {
  try {
    filters = filters || {};
    var name      = String(vendorName || "").trim();
    var nameLower = name.toLowerCase();

    /* Single batch load — passed into stats calculator, no double reads */
    var allTasks     = _sheetRows(SH_TASKS).map(_fmtRow);
    var allRevisions = _sheetRows(SH_REVISIONS).map(_fmtRow);
    var allVendors   = _sheetRows(SH_VENDORS).map(_fmtRow);

    var vendorRow = allVendors.find(function(r) {
      return String(r[VC.NAME] || "").trim().toLowerCase() === nameLower;
    });
    var defaultCurrency = vendorRow ? (vendorRow[VC.CURRENCY] || "") : "";

    /* Filter to this vendor */
    var vTasks = allTasks.filter(function(r) {
      return r[TC.WORK_TYPE] === "Vendor" &&
             String(r[TC.VENDOR_NAME] || "").trim().toLowerCase() === nameLower;
    });
    var vRevs = allRevisions.filter(function(r) {
      return r[RC.WORK_TYPE] === "Vendor" &&
             String(r[RC.VENDOR_NAME] || "").trim().toLowerCase() === nameLower;
    });

    /* Stats via centralized calculator — pass preloaded data, zero extra reads */
    var stats = calculateVendorStats(name, filters, {
      tasks: allTasks, revisions: allRevisions, vendorRow: vendorRow
    });
    var cur = stats.currency || defaultCurrency;

    /* Apply filters for breakdowns */
    var filtered = applyFilters(vTasks, vRevs, filters);
    var fTasks   = filtered.tasks;
    var fRevs    = filtered.revs;

    /* Normalise tasks + revisions into a flat item list for shared helpers */
    function _taskToItem(r) {
      return {
        id: r[TC.ID], projectId: String(r[TC.PROJECT_ID]).trim(),
        projectName: r[TC.PROJECT_NAME] || String(r[TC.PROJECT_ID]).trim(),
        taskType: r[TC.TASK_TYPE] || "Other",
        language: r[TC.LANGUAGE] || "",
        pages: Number(r[TC.FINAL_PAGES]) || 0,
        rate: Number(r[TC.RATE_PER_PAGE]) || 0,
        currency: r[TC.CURRENCY] || cur,
        workStatus: r[TC.STATUS] || "Pending",
        paymentStatus: String(r[TC.PAYMENT_STATUS] || "Unpaid").trim(),
        startDate: r[TC.START_DATE],
        completedDate: r[TC.COMPLETED_DATE] || r[TC.DELIVERY_DATE],
        date: r[TC.START_DATE] || r[TC.CREATED_AT],
        type: "task"
      };
    }
    function _revToItem(r) {
      return {
        id: r[RC.ID], projectId: String(r[RC.PROJECT_ID]).trim(),
        projectName: r[RC.PROJECT_NAME] || String(r[RC.PROJECT_ID]).trim(),
        taskType: "Revision",
        language: r[RC.LANGUAGE] || "",
        pages: Number(r[RC.REV_PAGES]) || 0,
        rate: Number(r[RC.RATE_PER_PAGE]) || 0,
        currency: r[RC.CURRENCY] || cur,
        workStatus: r[RC.STATUS] || "Pending",
        paymentStatus: String(r[RC.PAYMENT_STATUS] || "Unpaid").trim(),
        startDate: r[RC.REV_DATE] || r[RC.CREATED_AT],
        completedDate: r[RC.COMPLETED_DATE] || r[RC.DELIVERY_DATE],
        date: r[RC.REV_DATE] || r[RC.CREATED_AT],
        type: "revision"
      };
    }

    var allItems = vTasks.map(_taskToItem).concat(vRevs.map(_revToItem));
    var fItems   = fTasks.map(_taskToItem).concat(fRevs.map(_revToItem));

    var breakdown = _buildPerformanceData(allItems, fItems, cur);

    return {
      vendor: vendorRow ? {
        id: vendorRow[VC.ID], name: vendorRow[VC.NAME],
        contactPerson: vendorRow[VC.CONTACT], email: vendorRow[VC.EMAIL],
        phone: vendorRow[VC.PHONE], specialization: vendorRow[VC.SPECIALIZATION],
        languages: vendorRow[VC.LANGUAGES], ratePerPage: stats.ratePerPage,
        currency: cur, status: vendorRow[VC.STATUS], notes: vendorRow[VC.NOTES]
      } : { name: name, ratePerPage: 0, currency: "" },
      stats: stats,
      monthly:       breakdown.monthly,
      byProject:     breakdown.byProject,
      byTaskType:    breakdown.byTaskType,
      byLanguage:    breakdown.byLanguage,
      rows:          breakdown.rows,
      projectOptions:breakdown.projectOptions,
      langOptions:   breakdown.langOptions
    };
  } catch (e) {
    console.error("getVendorPerformance:", e);
    throw e;
  }
}

/* ================================================================
   TEAM MEMBER PERFORMANCE
   Accepts: memberName (string), filters { month, year, projectId, language, taskType }
================================================================ */
function getTeamMemberPerformance(memberName, filters) {
  try {
    filters = filters || {};
    var name      = String(memberName || "").trim();
    var nameLower = name.toLowerCase();

    /* Single batch load — passed into stats calculator, no double reads */
    var allTasks     = _sheetRows(SH_TASKS).map(_fmtRow);
    var allRevisions = _sheetRows(SH_REVISIONS).map(_fmtRow);
    var allTeam      = _sheetRows(SH_TEAM).map(_fmtRow);

    var memberRow = allTeam.find(function(r) {
      return String(r[MC.NAME] || "").trim().toLowerCase() === nameLower;
    });
    var defaultCurrency = memberRow ? (memberRow[MC.CURRENCY] || "") : "";

    /* Filter to this member */
    var mTasks = allTasks.filter(function(r) {
      return r[TC.WORK_TYPE] === "In-House" &&
             String(r[TC.ASSIGNED_TO] || "").trim().toLowerCase() === nameLower;
    });
    var mRevs = allRevisions.filter(function(r) {
      return r[RC.WORK_TYPE] === "In-House" &&
             String(r[RC.ASSIGNED_TO] || "").trim().toLowerCase() === nameLower;
    });

    /* Stats via centralized calculator — pass preloaded data, zero extra reads */
    var stats = calculateTeamMemberStats(name, filters, {
      tasks: allTasks, revisions: allRevisions, memberRow: memberRow
    });
    var cur = stats.currency || defaultCurrency;

    /* Apply filters for breakdowns */
    var filtered = applyFilters(mTasks, mRevs, filters);
    var fTasks   = filtered.tasks;
    var fRevs    = filtered.revs;

    /* Normalise into flat item list */
    function _taskToItem(r) {
      return {
        id: r[TC.ID], projectId: String(r[TC.PROJECT_ID]).trim(),
        projectName: r[TC.PROJECT_NAME] || String(r[TC.PROJECT_ID]).trim(),
        taskType: r[TC.TASK_TYPE] || "Other",
        language: r[TC.LANGUAGE] || "",
        pages: Number(r[TC.FINAL_PAGES]) || 0,
        rate: Number(r[TC.RATE_PER_PAGE]) || 0,
        currency: r[TC.CURRENCY] || cur,
        workStatus: r[TC.STATUS] || "Pending",
        paymentStatus: String(r[TC.PAYMENT_STATUS] || "Unpaid").trim(),
        startDate: r[TC.START_DATE],
        completedDate: r[TC.COMPLETED_DATE] || r[TC.DELIVERY_DATE],
        date: r[TC.START_DATE] || r[TC.CREATED_AT],
        type: "task"
      };
    }
    function _revToItem(r) {
      return {
        id: r[RC.ID], projectId: String(r[RC.PROJECT_ID]).trim(),
        projectName: r[RC.PROJECT_NAME] || String(r[RC.PROJECT_ID]).trim(),
        taskType: "Revision",
        language: r[RC.LANGUAGE] || "",
        pages: Number(r[RC.REV_PAGES]) || 0,
        rate: Number(r[RC.RATE_PER_PAGE]) || 0,
        currency: r[RC.CURRENCY] || cur,
        workStatus: r[RC.STATUS] || "Pending",
        paymentStatus: String(r[RC.PAYMENT_STATUS] || "Unpaid").trim(),
        startDate: r[RC.REV_DATE] || r[RC.CREATED_AT],
        completedDate: r[RC.COMPLETED_DATE] || r[RC.DELIVERY_DATE],
        date: r[RC.REV_DATE] || r[RC.CREATED_AT],
        type: "revision"
      };
    }

    var allItems = mTasks.map(_taskToItem).concat(mRevs.map(_revToItem));
    var fItems   = fTasks.map(_taskToItem).concat(fRevs.map(_revToItem));

    var breakdown = _buildPerformanceData(allItems, fItems, cur);

    return {
      member: memberRow ? {
        id: memberRow[MC.ID], name: memberRow[MC.NAME], role: memberRow[MC.ROLE],
        email: memberRow[MC.EMAIL], phone: memberRow[MC.PHONE],
        specialization: memberRow[MC.SPECIALIZATION], status: memberRow[MC.STATUS],
        ratePerPage: Number(memberRow[MC.RATE_PER_PAGE]) || 0,
        currency: memberRow[MC.CURRENCY] || ""
      } : { name: name },
      stats: stats,
      monthly:       breakdown.monthly,
      byProject:     breakdown.byProject,
      byTaskType:    breakdown.byTaskType,
      byLanguage:    breakdown.byLanguage,
      rows:          breakdown.rows,
      projectOptions:breakdown.projectOptions,
      langOptions:   breakdown.langOptions
    };
  } catch (e) {
    console.error("getTeamMemberPerformance:", e);
    throw e;
  }
}
