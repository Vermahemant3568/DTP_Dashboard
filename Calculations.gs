/* ================================================================
   Calculations.gs — Centralized Payment & Stats Calculations
   Single source of truth for all amount and page calculations.

   Payment Status Rules:
     "Paid"    → included in paidAmount + totalAmount
     "Partial" → included in partialAmount + totalAmount
     "Unpaid"  → included in unpaidAmount only (NOT in totalAmount)
     "N/A"     → excluded from all payment totals
================================================================ */

/**
 * Core payment-aware stats calculator.
 * Accepts pre-filtered task and revision rows (already _fmtRow'd).
 * Separates Task and Revision calculations, then combines.
 *
 * @param {Array}  tasks           - formatted task rows for one person/project
 * @param {Array}  revs            - formatted revision rows for one person/project
 * @param {string} defaultCurrency
 * @returns {Object} stats
 */
function _calcPaymentStats(tasks, revs, defaultCurrency) {
  tasks = tasks || [];
  revs  = revs  || [];
  var cur = defaultCurrency || "";

  // Task totals
  var taskPages          = 0;
  var taskCompletedPages = 0;
  var taskPendingPages   = 0;
  var taskPaidAmt        = 0;
  var taskUnpaidAmt      = 0;
  var taskPartialAmt     = 0;
  var taskCompleted      = 0;

  tasks.forEach(function(r) {
    var pages   = Number(r[TC.FINAL_PAGES])   || 0;
    var rate    = Number(r[TC.RATE_PER_PAGE])  || 0;
    var pStatus = String(r[TC.PAYMENT_STATUS]  || "Unpaid").trim();
    var wStatus = String(r[TC.STATUS]          || "Pending").trim();
    var amt     = pages * rate;

    taskPages += pages;
    if (wStatus === "Completed") { taskCompletedPages += pages; taskCompleted++; }
    else taskPendingPages += pages;

    if (pStatus === "Paid")    taskPaidAmt    += amt;
    else if (pStatus === "Partial") taskPartialAmt += amt;
    else if (pStatus === "Unpaid")  taskUnpaidAmt  += amt;
    // "N/A" → excluded from all

    if (!cur && r[TC.CURRENCY]) cur = r[TC.CURRENCY];
  });

  // Revision totals
  var revPages          = 0;
  var revCompletedPages = 0;
  var revPendingPages   = 0;
  var revPaidAmt        = 0;
  var revUnpaidAmt      = 0;
  var revPartialAmt     = 0;
  var revCompleted      = 0;

  revs.forEach(function(r) {
    var pages   = Number(r[RC.REV_PAGES])      || 0;
    var rate    = Number(r[RC.RATE_PER_PAGE])   || 0;
    var pStatus = String(r[RC.PAYMENT_STATUS]   || "Unpaid").trim();
    var wStatus = String(r[RC.STATUS]           || "Pending").trim();
    var amt     = pages * rate;

    revPages += pages;
    if (wStatus === "Completed") { revCompletedPages += pages; revCompleted++; }
    else revPendingPages += pages;

    if (pStatus === "Paid")    revPaidAmt    += amt;
    else if (pStatus === "Partial") revPartialAmt += amt;
    else if (pStatus === "Unpaid")  revUnpaidAmt  += amt;

    if (!cur && r[RC.CURRENCY]) cur = r[RC.CURRENCY];
  });

  var totalPages      = taskPages + revPages;
  var completedPages  = taskCompletedPages + revCompletedPages;
  var pendingPages    = taskPendingPages + revPendingPages;
  var paidAmount      = taskPaidAmt + revPaidAmt;
  var partialAmount   = taskPartialAmt + revPartialAmt;
  var unpaidAmount    = taskUnpaidAmt + revUnpaidAmt;
  var totalAmount     = paidAmount + partialAmount;  // Unpaid excluded
  var totalAssigned   = tasks.length + revs.length;
  var totalCompleted  = taskCompleted + revCompleted;
  var avgPages        = totalAssigned > 0 ? Math.round(totalPages / totalAssigned) : 0;

  return {
    totalAssigned,
    totalCompleted,
    totalPages,
    completedPages,
    pendingPages,
    avgPages,
    // Task-specific
    taskPages,
    taskCompletedPages,
    taskPaidAmt,
    taskUnpaidAmt,
    taskPartialAmt,
    // Revision-specific
    revPages,
    revCompletedPages,
    revPaidAmt,
    revUnpaidAmt,
    revPartialAmt,
    // Combined payment totals
    paidAmount,
    partialAmount,
    unpaidAmount,
    totalAmount,   // paid + partial only (Unpaid excluded)
    currency: cur
  };
}

/**
 * Apply date/project/language/status filters to tasks and revisions.
 */
function applyFilters(tasks, revs, filters) {
  filters = filters || {};
  var month     = filters.month     ? Number(filters.month) - 1 : null;
  var year      = filters.year      ? Number(filters.year)      : null;
  var projectId = filters.projectId ? String(filters.projectId).trim() : "";
  var language  = filters.language  ? String(filters.language).trim().toLowerCase() : "";
  var status    = filters.status    ? String(filters.status).trim() : "";

  function matchDate(dateVal, fallback) {
    if (month === null || year === null) return true;
    return _inMonth(dateVal || fallback, month, year);
  }

  var fTasks = tasks.filter(function(r) {
    if (projectId && String(r[TC.PROJECT_ID]).trim() !== projectId) return false;
    if (language  && String(r[TC.LANGUAGE] || "").toLowerCase() !== language) return false;
    if (status    && r[TC.STATUS] !== status) return false;
    if (month !== null) return matchDate(r[TC.START_DATE], r[TC.CREATED_AT]);
    return true;
  });

  var fRevs = revs.filter(function(r) {
    if (projectId && String(r[RC.PROJECT_ID]).trim() !== projectId) return false;
    if (language  && String(r[RC.LANGUAGE] || "").toLowerCase() !== language) return false;
    if (status    && r[RC.STATUS] !== status) return false;
    if (month !== null) return matchDate(r[RC.REV_DATE], r[RC.CREATED_AT]);
    return true;
  });

  return { tasks: fTasks, revs: fRevs };
}

/**
 * Calculate vendor stats — payment-aware, with optional filters.
 * Accepts pre-loaded allTasks/allRevs/vendorRow to avoid duplicate sheet reads
 * when called from getVendorPerformance. Falls back to reading sheets if not provided.
 */
function calculateVendorStats(vendorName, filters, preloaded) {
  var name     = String(vendorName || "").trim().toLowerCase();
  var allTasks = (preloaded && preloaded.tasks)     || _sheetRows(SH_TASKS).map(_fmtRow);
  var allRevs  = (preloaded && preloaded.revisions) || _sheetRows(SH_REVISIONS).map(_fmtRow);

  var vTasks = allTasks.filter(function(r) {
    return r[TC.WORK_TYPE] === "Vendor" &&
           String(r[TC.VENDOR_NAME] || "").trim().toLowerCase() === name;
  });
  var vRevs = allRevs.filter(function(r) {
    return r[RC.WORK_TYPE] === "Vendor" &&
           String(r[RC.VENDOR_NAME] || "").trim().toLowerCase() === name;
  });

  var filtered = applyFilters(vTasks, vRevs, filters || {});

  var vendorRow = (preloaded && preloaded.vendorRow) ||
    _sheetRows(SH_VENDORS).map(_fmtRow).find(function(r) {
      return String(r[VC.NAME] || "").trim().toLowerCase() === name;
    });
  var defaultCurrency = vendorRow ? (vendorRow[VC.CURRENCY] || "") : "";

  var stats = _calcPaymentStats(filtered.tasks, filtered.revs, defaultCurrency);
  stats.ratePerPage = vendorRow ? (Number(vendorRow[VC.RATE_PER_PAGE]) || 0) : 0;
  stats.currency    = stats.currency || defaultCurrency;
  return stats;
}

/**
 * Calculate team member stats — payment-aware, with optional filters.
 * Accepts pre-loaded allTasks/allRevs/memberRow to avoid duplicate sheet reads
 * when called from getTeamMemberPerformance. Falls back to reading sheets if not provided.
 */
function calculateTeamMemberStats(memberName, filters, preloaded) {
  var name     = String(memberName || "").trim().toLowerCase();
  var allTasks = (preloaded && preloaded.tasks)     || _sheetRows(SH_TASKS).map(_fmtRow);
  var allRevs  = (preloaded && preloaded.revisions) || _sheetRows(SH_REVISIONS).map(_fmtRow);

  var mTasks = allTasks.filter(function(r) {
    return r[TC.WORK_TYPE] === "In-House" &&
           String(r[TC.ASSIGNED_TO] || "").trim().toLowerCase() === name;
  });
  var mRevs = allRevs.filter(function(r) {
    return r[RC.WORK_TYPE] === "In-House" &&
           String(r[RC.ASSIGNED_TO] || "").trim().toLowerCase() === name;
  });

  var filtered = applyFilters(mTasks, mRevs, filters || {});

  var memberRow = (preloaded && preloaded.memberRow) ||
    _sheetRows(SH_TEAM).map(_fmtRow).find(function(r) {
      return String(r[MC.NAME] || "").trim().toLowerCase() === name;
    });
  var defaultCurrency = memberRow ? (memberRow[MC.CURRENCY] || "") : "";

  var stats = _calcPaymentStats(filtered.tasks, filtered.revs, defaultCurrency);
  stats.ratePerPage = memberRow ? (Number(memberRow[MC.RATE_PER_PAGE]) || 0) : 0;
  stats.currency    = stats.currency || defaultCurrency;
  return stats;
}

/**
 * Calculate payment summary for a single project.
 * Returns separate task and revision breakdowns.
 */
function calculateProjectPaymentSummary(tasks, revisions) {
  tasks     = tasks     || [];
  revisions = revisions || [];
  var taskPaid = 0, taskUnpaid = 0, taskPartial = 0;
  var revPaid  = 0, revUnpaid  = 0, revPartial  = 0;

  tasks.forEach(function(r) {
    var amt     = (Number(r[TC.FINAL_PAGES]) || 0) * (Number(r[TC.RATE_PER_PAGE]) || 0);
    var pStatus = String(r[TC.PAYMENT_STATUS] || "Unpaid").trim();
    if (pStatus === "Paid")    taskPaid    += amt;
    else if (pStatus === "Partial") taskPartial += amt;
    else if (pStatus === "Unpaid")  taskUnpaid  += amt;
  });

  revisions.forEach(function(r) {
    var amt     = (Number(r[RC.REV_PAGES]) || 0) * (Number(r[RC.RATE_PER_PAGE]) || 0);
    var pStatus = String(r[RC.PAYMENT_STATUS] || "Unpaid").trim();
    if (pStatus === "Paid")    revPaid    += amt;
    else if (pStatus === "Partial") revPartial += amt;
    else if (pStatus === "Unpaid")  revUnpaid  += amt;
  });

  return {
    taskPaidAmt:    taskPaid,
    taskUnpaidAmt:  taskUnpaid,
    taskPartialAmt: taskPartial,
    revPaidAmt:     revPaid,
    revUnpaidAmt:   revUnpaid,
    revPartialAmt:  revPartial,
    totalPaidAmt:   taskPaid    + revPaid,
    totalUnpaidAmt: taskUnpaid  + revUnpaid,
    totalPartialAmt:taskPartial + revPartial,
    totalAmount:    taskPaid + taskPartial + revPaid + revPartial  // excludes Unpaid
  };
}

/**
 * Calculate dashboard-level payment totals across all tasks and revisions.
 * Unpaid items are tracked but NOT included in totalAmount.
 */
function calculateDashboardPaymentTotals(allTasks, allRevisions) {
  allTasks     = allTasks     || [];
  allRevisions = allRevisions || [];
  var totalPaidAmt    = 0;
  var totalUnpaidAmt  = 0;
  var totalPartialAmt = 0;

  allTasks.forEach(function(r) {
    var amt = (Number(r[TC.FINAL_PAGES]) || 0) * (Number(r[TC.RATE_PER_PAGE]) || 0);
    var ps  = String(r[TC.PAYMENT_STATUS] || "Unpaid").trim();
    if (ps === "Paid")    totalPaidAmt    += amt;
    else if (ps === "Partial") totalPartialAmt += amt;
    else if (ps === "Unpaid")  totalUnpaidAmt  += amt;
  });

  allRevisions.forEach(function(r) {
    var amt = (Number(r[RC.REV_PAGES]) || 0) * (Number(r[RC.RATE_PER_PAGE]) || 0);
    var ps  = String(r[RC.PAYMENT_STATUS] || "Unpaid").trim();
    if (ps === "Paid")    totalPaidAmt    += amt;
    else if (ps === "Partial") totalPartialAmt += amt;
    else if (ps === "Unpaid")  totalUnpaidAmt  += amt;
  });

  return {
    totalPaidAmt,
    totalUnpaidAmt,
    totalPartialAmt,
    totalAmount: totalPaidAmt + totalPartialAmt  // Unpaid excluded
  };
}
