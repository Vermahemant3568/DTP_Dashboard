/* ================================================================
   VendorsService.gs — Vendor & Team CRUD
   Sheets: Vendors (12 cols) | Team (8 cols)
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
    sh.getRange(found.index, 2, 1, 11).setValues([[
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
      found.row[11]
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
    sh.getRange(found.index, 2, 1, 7).setValues([[
      d.name             || "",
      d.role             || "",
      d.email            || "",
      d.phone            || "",
      d.specialization   || "",
      d.status           || "Active",
      found.row[7]
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
   VENDOR PERFORMANCE — full analytics for a single vendor
   Accepts: vendorName (string), filters { month, year, projectId, language, status }
   Returns: stats, monthly breakdown, project/task/language contributions
================================================================ */
function getVendorPerformance(vendorName, filters) {
  try {
    filters = filters || {};
    var name = String(vendorName || "").trim();

    /* ── Batch load all data in one pass ── */
    var allTasks     = _sheetRows(SH_TASKS).map(_fmtRow);
    var allRevisions = _sheetRows(SH_REVISIONS).map(_fmtRow);
    var allVendors   = _sheetRows(SH_VENDORS).map(_fmtRow);
    var allProjects  = _sheetRows(SH_PROJECTS).map(_fmtRow);

    /* ── Find vendor record ── */
    var vendorRow = allVendors.find(function(r) {
      return String(r[1] || "").trim().toLowerCase() === name.toLowerCase();
    });

    /* ── Filter tasks/revisions belonging to this vendor ── */
    var vTasks = allTasks.filter(function(r) {
      return r[TC.WORK_TYPE] === "Vendor" &&
             String(r[TC.VENDOR_NAME] || "").trim().toLowerCase() === name.toLowerCase();
    });
    var vRevs = allRevisions.filter(function(r) {
      return r[RC.WORK_TYPE] === "Vendor" &&
             String(r[RC.VENDOR_NAME] || "").trim().toLowerCase() === name.toLowerCase();
    });

    /* ── Apply optional filters ── */
    function _applyFilters(tasks, revs) {
      var ft = tasks, fr = revs;
      if (filters.month && filters.year) {
        var m = Number(filters.month) - 1, y = Number(filters.year);
        ft = ft.filter(function(r) { return _inMonth(r[TC.START_DATE], m, y); });
        fr = fr.filter(function(r) { return _inMonth(r[RC.REV_DATE] || r[RC.CREATED_AT], m, y); });
      }
      if (filters.projectId) {
        ft = ft.filter(function(r) { return String(r[TC.PROJECT_ID]).trim() === String(filters.projectId).trim(); });
        fr = fr.filter(function(r) { return String(r[RC.PROJECT_ID]).trim() === String(filters.projectId).trim(); });
      }
      if (filters.language) {
        ft = ft.filter(function(r) { return String(r[TC.LANGUAGE] || "").toLowerCase() === filters.language.toLowerCase(); });
        fr = fr.filter(function(r) { return String(r[RC.LANGUAGE] || "").toLowerCase() === filters.language.toLowerCase(); });
      }
      if (filters.status) {
        ft = ft.filter(function(r) { return r[TC.STATUS] === filters.status; });
        fr = fr.filter(function(r) { return r[RC.STATUS] === filters.status; });
      }
      return { tasks: ft, revs: fr };
    }

    var filtered = _applyFilters(vTasks, vRevs);
    var fTasks = filtered.tasks, fRevs = filtered.revs;

    /* ── Overall stats (unfiltered) ── */
    var totalAssigned   = vTasks.length + vRevs.length;
    var totalCompleted  = vTasks.filter(function(r){ return r[TC.STATUS] === "Completed"; }).length +
                         vRevs.filter(function(r){ return r[RC.STATUS] === "Completed"; }).length;
    var totalPages      = vTasks.reduce(function(s,r){ return s+(Number(r[TC.FINAL_PAGES])||0); }, 0) +
                         vRevs.reduce(function(s,r){ return s+(Number(r[RC.REV_PAGES])||0); }, 0);
    var completedPages  = vTasks.filter(function(r){ return r[TC.STATUS]==="Completed"; })
                               .reduce(function(s,r){ return s+(Number(r[TC.FINAL_PAGES])||0); }, 0) +
                         vRevs.filter(function(r){ return r[RC.STATUS]==="Completed"; })
                              .reduce(function(s,r){ return s+(Number(r[RC.REV_PAGES])||0); }, 0);
    var pendingPages    = totalPages - completedPages;
    var ratePerPage     = vendorRow ? (Number(vendorRow[7]) || 0) : 0;
    var currency        = vendorRow ? (vendorRow[8] || "INR") : "INR";
    var estimatedAmount = completedPages * ratePerPage;

    /* ── Filtered stats ── */
    var filtPages = fTasks.reduce(function(s,r){ return s+(Number(r[TC.FINAL_PAGES])||0); }, 0) +
                   fRevs.reduce(function(s,r){ return s+(Number(r[RC.REV_PAGES])||0); }, 0);

    /* ── Monthly breakdown (last 12 months, unfiltered) ── */
    var monthlyMap = {};
    function _addToMonth(dateVal, pages) {
      var d = dateVal;
      if (typeof d === "string") {
        var dmy = d.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (dmy) d = new Date(Number(dmy[3]), Number(dmy[2])-1, Number(dmy[1]));
        else d = new Date(d);
      }
      if (!(d instanceof Date) || isNaN(d)) return;
      var key = d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0");
      if (!monthlyMap[key]) monthlyMap[key] = { pages: 0, tasks: 0 };
      monthlyMap[key].pages += pages;
      monthlyMap[key].tasks++;
    }
    vTasks.forEach(function(r) { _addToMonth(r[TC.START_DATE], Number(r[TC.FINAL_PAGES])||0); });
    vRevs.forEach(function(r)  { _addToMonth(r[RC.REV_DATE]||r[RC.CREATED_AT], Number(r[RC.REV_PAGES])||0); });
    var monthly = Object.keys(monthlyMap).sort().slice(-12).map(function(k) {
      var parts = k.split("-");
      var label = Utilities.formatDate(new Date(Number(parts[0]), Number(parts[1])-1, 1), Session.getScriptTimeZone(), "MMM yyyy");
      return { key: k, label: label, pages: monthlyMap[k].pages, tasks: monthlyMap[k].tasks };
    });

    /* ── Project-wise contribution (filtered) ── */
    var projMap = {};
    fTasks.forEach(function(r) {
      var pid = String(r[TC.PROJECT_ID]).trim();
      var pname = r[TC.PROJECT_NAME] || pid;
      if (!projMap[pid]) projMap[pid] = { projectId: pid, projectName: pname, pages: 0, tasks: 0, completed: 0 };
      projMap[pid].pages += Number(r[TC.FINAL_PAGES]) || 0;
      projMap[pid].tasks++;
      if (r[TC.STATUS] === "Completed") projMap[pid].completed++;
    });
    fRevs.forEach(function(r) {
      var pid = String(r[RC.PROJECT_ID]).trim();
      var pname = r[RC.PROJECT_NAME] || pid;
      if (!projMap[pid]) projMap[pid] = { projectId: pid, projectName: pname, pages: 0, tasks: 0, completed: 0 };
      projMap[pid].pages += Number(r[RC.REV_PAGES]) || 0;
      projMap[pid].tasks++;
      if (r[RC.STATUS] === "Completed") projMap[pid].completed++;
    });
    var byProject = Object.values(projMap).sort(function(a,b){ return b.pages - a.pages; });

    /* ── Task-type contribution (filtered) ── */
    var taskTypeMap = {};
    fTasks.forEach(function(r) {
      var t = r[TC.TASK_TYPE] || "Other";
      if (!taskTypeMap[t]) taskTypeMap[t] = { pages: 0, count: 0 };
      taskTypeMap[t].pages += Number(r[TC.FINAL_PAGES]) || 0;
      taskTypeMap[t].count++;
    });
    if (fRevs.length) {
      if (!taskTypeMap["Revisions"]) taskTypeMap["Revisions"] = { pages: 0, count: 0 };
      fRevs.forEach(function(r) {
        taskTypeMap["Revisions"].pages += Number(r[RC.REV_PAGES]) || 0;
        taskTypeMap["Revisions"].count++;
      });
    }
    var byTaskType = Object.keys(taskTypeMap).map(function(k) {
      return { type: k, pages: taskTypeMap[k].pages, count: taskTypeMap[k].count };
    }).sort(function(a,b){ return b.pages - a.pages; });

    /* ── Language-wise work (filtered) ── */
    var langMap = {};
    fTasks.forEach(function(r) {
      var l = r[TC.LANGUAGE] || "Unknown";
      if (!langMap[l]) langMap[l] = 0;
      langMap[l] += Number(r[TC.FINAL_PAGES]) || 0;
    });
    fRevs.forEach(function(r) {
      var l = r[RC.LANGUAGE] || "Unknown";
      if (!langMap[l]) langMap[l] = 0;
      langMap[l] += Number(r[RC.REV_PAGES]) || 0;
    });
    var byLanguage = Object.keys(langMap).map(function(l) {
      return { language: l, pages: langMap[l] };
    }).sort(function(a,b){ return b.pages - a.pages; });

    /* ── Detailed task rows (filtered, for table) ── */
    var taskRows = fTasks.map(function(r) {
      return {
        id: r[TC.ID], projectId: r[TC.PROJECT_ID], projectName: r[TC.PROJECT_NAME],
        taskType: r[TC.TASK_TYPE], language: r[TC.LANGUAGE],
        pages: Number(r[TC.FINAL_PAGES]) || 0, status: r[TC.STATUS],
        startDate: r[TC.START_DATE], completedDate: r[TC.COMPLETED_DATE],
        type: "task"
      };
    });
    var revRows = fRevs.map(function(r) {
      return {
        id: r[RC.ID], projectId: r[RC.PROJECT_ID], projectName: r[RC.PROJECT_NAME],
        taskType: "Revision", language: r[RC.LANGUAGE],
        pages: Number(r[RC.REV_PAGES]) || 0, status: r[RC.STATUS],
        startDate: r[RC.REV_DATE] || r[RC.CREATED_AT], completedDate: r[RC.COMPLETED_DATE],
        type: "revision"
      };
    });
    var allRows = taskRows.concat(revRows).sort(function(a,b) {
      return String(b.startDate).localeCompare(String(a.startDate));
    });

    /* ── Unique project list for filter dropdown ── */
    var projectOptions = [];
    var seenPids = {};
    vTasks.concat(vRevs).forEach(function(r) {
      var pid = String(r[TC.PROJECT_ID] || r[RC.PROJECT_ID] || "").trim();
      var pname = r[TC.PROJECT_NAME] || r[RC.PROJECT_NAME] || pid;
      if (pid && !seenPids[pid]) { seenPids[pid] = true; projectOptions.push({ id: pid, name: pname }); }
    });

    /* ── Unique language list for filter dropdown ── */
    var langOptions = [];
    var seenLangs = {};
    vTasks.concat(vRevs).forEach(function(r) {
      var l = r[TC.LANGUAGE] || r[RC.LANGUAGE] || "";
      if (l && !seenLangs[l]) { seenLangs[l] = true; langOptions.push(l); }
    });

    return {
      vendor: vendorRow ? {
        id: vendorRow[0], name: vendorRow[1], contactPerson: vendorRow[2],
        email: vendorRow[3], phone: vendorRow[4], specialization: vendorRow[5],
        languages: vendorRow[6], ratePerPage: ratePerPage, currency: currency,
        status: vendorRow[9], notes: vendorRow[10]
      } : { name: name, ratePerPage: 0, currency: "INR" },
      stats: {
        totalAssigned, totalCompleted, totalPages, completedPages,
        pendingPages, ratePerPage, currency, estimatedAmount, filtPages
      },
      monthly, byProject, byTaskType, byLanguage,
      rows: allRows,
      projectOptions, langOptions
    };
  } catch (e) {
    console.error("getVendorPerformance:", e);
    throw e;
  }
}

/* ================================================================
   TEAM MEMBER PERFORMANCE — full analytics for a single team member
   Accepts: memberName (string), filters { month, year, projectId, language, taskType }
   Returns: stats, monthly breakdown, project/task-type/language contributions
================================================================ */
function getTeamMemberPerformance(memberName, filters) {
  try {
    filters = filters || {};
    var name = String(memberName || "").trim();

    /* ── Batch load ── */
    var allTasks     = _sheetRows(SH_TASKS).map(_fmtRow);
    var allRevisions = _sheetRows(SH_REVISIONS).map(_fmtRow);
    var allTeam      = _sheetRows(SH_TEAM).map(_fmtRow);

    /* ── Find member record ── */
    var memberRow = allTeam.find(function(r) {
      return String(r[1] || "").trim().toLowerCase() === name.toLowerCase();
    });

    /* ── Filter tasks/revisions for this member ── */
    var mTasks = allTasks.filter(function(r) {
      return r[TC.WORK_TYPE] === "In-House" &&
             String(r[TC.ASSIGNED_TO] || "").trim().toLowerCase() === name.toLowerCase();
    });
    var mRevs = allRevisions.filter(function(r) {
      return r[RC.WORK_TYPE] === "In-House" &&
             String(r[RC.ASSIGNED_TO] || "").trim().toLowerCase() === name.toLowerCase();
    });

    /* ── Apply optional filters ── */
    function _applyFilters(tasks, revs) {
      var ft = tasks, fr = revs;
      if (filters.month && filters.year) {
        var m = Number(filters.month) - 1, y = Number(filters.year);
        ft = ft.filter(function(r) { return _inMonth(r[TC.START_DATE], m, y); });
        fr = fr.filter(function(r) { return _inMonth(r[RC.REV_DATE] || r[RC.CREATED_AT], m, y); });
      }
      if (filters.projectId) {
        ft = ft.filter(function(r) { return String(r[TC.PROJECT_ID]).trim() === String(filters.projectId).trim(); });
        fr = fr.filter(function(r) { return String(r[RC.PROJECT_ID]).trim() === String(filters.projectId).trim(); });
      }
      if (filters.language) {
        ft = ft.filter(function(r) { return String(r[TC.LANGUAGE]||""  ).toLowerCase() === filters.language.toLowerCase(); });
        fr = fr.filter(function(r) { return String(r[RC.LANGUAGE]||""  ).toLowerCase() === filters.language.toLowerCase(); });
      }
      if (filters.taskType && filters.taskType !== "Revision") {
        ft = ft.filter(function(r) { return r[TC.TASK_TYPE] === filters.taskType; });
      } else if (filters.taskType === "Revision") {
        ft = []; /* show only revisions */
      }
      return { tasks: ft, revs: fr };
    }

    var filtered = _applyFilters(mTasks, mRevs);
    var fTasks = filtered.tasks, fRevs = filtered.revs;

    /* ── Overall stats (unfiltered) ── */
    var totalAssigned  = mTasks.length + mRevs.length;
    var totalCompleted = mTasks.filter(function(r){ return r[TC.STATUS]==="Completed"; }).length +
                        mRevs.filter(function(r){ return r[RC.STATUS]==="Completed"; }).length;
    var totalPages     = mTasks.reduce(function(s,r){ return s+(Number(r[TC.FINAL_PAGES])||0); }, 0) +
                        mRevs.reduce(function(s,r){ return s+(Number(r[RC.REV_PAGES])||0); }, 0);
    var completedPages = mTasks.filter(function(r){ return r[TC.STATUS]==="Completed"; })
                               .reduce(function(s,r){ return s+(Number(r[TC.FINAL_PAGES])||0); }, 0) +
                        mRevs.filter(function(r){ return r[RC.STATUS]==="Completed"; })
                             .reduce(function(s,r){ return s+(Number(r[RC.REV_PAGES])||0); }, 0);
    var inProgressPages = mTasks.filter(function(r){ return r[TC.STATUS]==="In Progress"; })
                                .reduce(function(s,r){ return s+(Number(r[TC.FINAL_PAGES])||0); }, 0) +
                         mRevs.filter(function(r){ return r[RC.STATUS]==="In Progress"; })
                              .reduce(function(s,r){ return s+(Number(r[RC.REV_PAGES])||0); }, 0);
    var avgPages = totalAssigned > 0 ? Math.round(totalPages / totalAssigned) : 0;

    /* ── Monthly breakdown (last 12 months, unfiltered) ── */
    var monthlyMap = {};
    function _addToMonth(dateVal, pages) {
      var d = dateVal;
      if (typeof d === "string") {
        var dmy = d.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (dmy) d = new Date(Number(dmy[3]), Number(dmy[2])-1, Number(dmy[1]));
        else d = new Date(d);
      }
      if (!(d instanceof Date) || isNaN(d)) return;
      var key = d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0");
      if (!monthlyMap[key]) monthlyMap[key] = { pages: 0, tasks: 0 };
      monthlyMap[key].pages += pages;
      monthlyMap[key].tasks++;
    }
    mTasks.forEach(function(r) { _addToMonth(r[TC.START_DATE], Number(r[TC.FINAL_PAGES])||0); });
    mRevs.forEach(function(r)  { _addToMonth(r[RC.REV_DATE]||r[RC.CREATED_AT], Number(r[RC.REV_PAGES])||0); });
    var monthly = Object.keys(monthlyMap).sort().slice(-12).map(function(k) {
      var parts = k.split("-");
      var label = Utilities.formatDate(new Date(Number(parts[0]), Number(parts[1])-1, 1), Session.getScriptTimeZone(), "MMM yyyy");
      return { key: k, label: label, pages: monthlyMap[k].pages, tasks: monthlyMap[k].tasks };
    });

    /* ── Project contribution (filtered) ── */
    var projMap = {};
    fTasks.forEach(function(r) {
      var pid = String(r[TC.PROJECT_ID]).trim();
      var pname = r[TC.PROJECT_NAME] || pid;
      if (!projMap[pid]) projMap[pid] = { projectId: pid, projectName: pname, pages: 0, tasks: 0, completed: 0 };
      projMap[pid].pages += Number(r[TC.FINAL_PAGES]) || 0;
      projMap[pid].tasks++;
      if (r[TC.STATUS] === "Completed") projMap[pid].completed++;
    });
    fRevs.forEach(function(r) {
      var pid = String(r[RC.PROJECT_ID]).trim();
      var pname = r[RC.PROJECT_NAME] || pid;
      if (!projMap[pid]) projMap[pid] = { projectId: pid, projectName: pname, pages: 0, tasks: 0, completed: 0 };
      projMap[pid].pages += Number(r[RC.REV_PAGES]) || 0;
      projMap[pid].tasks++;
      if (r[RC.STATUS] === "Completed") projMap[pid].completed++;
    });
    var byProject = Object.values(projMap).sort(function(a,b){ return b.pages - a.pages; });

    /* ── Task-type breakdown (filtered) ── */
    var taskTypeMap = {};
    fTasks.forEach(function(r) {
      var t = r[TC.TASK_TYPE] || "Other";
      if (!taskTypeMap[t]) taskTypeMap[t] = { pages: 0, count: 0 };
      taskTypeMap[t].pages += Number(r[TC.FINAL_PAGES]) || 0;
      taskTypeMap[t].count++;
    });
    if (fRevs.length) {
      if (!taskTypeMap["Revisions"]) taskTypeMap["Revisions"] = { pages: 0, count: 0 };
      fRevs.forEach(function(r) {
        taskTypeMap["Revisions"].pages += Number(r[RC.REV_PAGES]) || 0;
        taskTypeMap["Revisions"].count++;
      });
    }
    var byTaskType = Object.keys(taskTypeMap).map(function(k) {
      return { type: k, pages: taskTypeMap[k].pages, count: taskTypeMap[k].count };
    }).sort(function(a,b){ return b.pages - a.pages; });

    /* ── Language contribution (filtered) ── */
    var langMap = {};
    fTasks.forEach(function(r) {
      var l = r[TC.LANGUAGE] || "Unknown";
      if (!langMap[l]) langMap[l] = 0;
      langMap[l] += Number(r[TC.FINAL_PAGES]) || 0;
    });
    fRevs.forEach(function(r) {
      var l = r[RC.LANGUAGE] || "Unknown";
      if (!langMap[l]) langMap[l] = 0;
      langMap[l] += Number(r[RC.REV_PAGES]) || 0;
    });
    var byLanguage = Object.keys(langMap).map(function(l) {
      return { language: l, pages: langMap[l] };
    }).sort(function(a,b){ return b.pages - a.pages; });

    /* ── Detailed rows (filtered) ── */
    var taskRows = fTasks.map(function(r) {
      return {
        id: r[TC.ID], projectId: r[TC.PROJECT_ID], projectName: r[TC.PROJECT_NAME],
        taskType: r[TC.TASK_TYPE], language: r[TC.LANGUAGE],
        pages: Number(r[TC.FINAL_PAGES]) || 0, status: r[TC.STATUS],
        startDate: r[TC.START_DATE], completedDate: r[TC.COMPLETED_DATE],
        type: "task"
      };
    });
    var revRows = fRevs.map(function(r) {
      return {
        id: r[RC.ID], projectId: r[RC.PROJECT_ID], projectName: r[RC.PROJECT_NAME],
        taskType: "Revision", language: r[RC.LANGUAGE],
        pages: Number(r[RC.REV_PAGES]) || 0, status: r[RC.STATUS],
        startDate: r[RC.REV_DATE] || r[RC.CREATED_AT], completedDate: r[RC.COMPLETED_DATE],
        type: "revision"
      };
    });
    var allRows = taskRows.concat(revRows).sort(function(a,b) {
      return String(b.startDate).localeCompare(String(a.startDate));
    });

    /* ── Dropdown options ── */
    var projectOptions = [], seenPids = {};
    mTasks.concat(mRevs).forEach(function(r) {
      var pid = String(r[TC.PROJECT_ID] || r[RC.PROJECT_ID] || "").trim();
      var pname = r[TC.PROJECT_NAME] || r[RC.PROJECT_NAME] || pid;
      if (pid && !seenPids[pid]) { seenPids[pid] = true; projectOptions.push({ id: pid, name: pname }); }
    });
    var langOptions = [], seenLangs = {};
    mTasks.concat(mRevs).forEach(function(r) {
      var l = r[TC.LANGUAGE] || r[RC.LANGUAGE] || "";
      if (l && !seenLangs[l]) { seenLangs[l] = true; langOptions.push(l); }
    });

    return {
      member: memberRow ? {
        id: memberRow[0], name: memberRow[1], role: memberRow[2],
        email: memberRow[3], phone: memberRow[4], specialization: memberRow[5],
        status: memberRow[6]
      } : { name: name },
      stats: {
        totalAssigned, totalCompleted, totalPages, completedPages,
        inProgressPages, avgPages
      },
      monthly, byProject, byTaskType, byLanguage,
      rows: allRows,
      projectOptions, langOptions
    };
  } catch (e) {
    console.error("getTeamMemberPerformance:", e);
    throw e;
  }
}
