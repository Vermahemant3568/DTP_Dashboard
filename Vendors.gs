/* ================================================================
   Vendors.gs — Vendor & Team CRUD
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
    const sh  = _sh(SH_VENDORS);
    if (!sh) throw new Error("Vendors sheet not found.");
    const id  = _id("VND");
    const now = new Date();
    sh.appendRow([
      id,
      d.vendorName      || "",
      d.contactPerson   || "",
      d.email           || "",
      d.phone           || "",
      d.specialization  || "",
      d.languages       || "",
      Number(d.ratePerPage) || 0,
      d.currency        || "USD",
      d.status          || "Active",
      d.notes           || "",
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
    const sh    = _sh(SH_VENDORS);
    const found = _findRow(sh, id);
    if (!found) throw new Error("Vendor not found: " + id);
    sh.getRange(found.index, 2, 1, 11).setValues([[
      d.vendorName      || "",
      d.contactPerson   || "",
      d.email           || "",
      d.phone           || "",
      d.specialization  || "",
      d.languages       || "",
      Number(d.ratePerPage) || 0,
      d.currency        || "USD",
      d.status          || "Active",
      d.notes           || "",
      found.row[11]     // preserve createdAt
    ]]);
    return { success: true };
  } catch (e) {
    console.error("updateVendor:", e);
    throw e;
  }
}

function deleteVendor(id) {
  try {
    const sh    = _sh(SH_VENDORS);
    const found = _findRow(sh, id);
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
    const sh  = _sh(SH_TEAM);
    if (!sh) throw new Error("Team sheet not found.");
    const id  = _id("MBR");
    const now = new Date();
    sh.appendRow([
      id,
      d.name            || "",
      d.role            || "",
      d.email           || "",
      d.phone           || "",
      d.specialization  || "",
      d.status          || "Active",
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
    const sh    = _sh(SH_TEAM);
    const found = _findRow(sh, id);
    if (!found) throw new Error("Team member not found: " + id);
    sh.getRange(found.index, 2, 1, 7).setValues([[
      d.name            || "",
      d.role            || "",
      d.email           || "",
      d.phone           || "",
      d.specialization  || "",
      d.status          || "Active",
      found.row[7]      // preserve createdAt
    ]]);
    return { success: true };
  } catch (e) {
    console.error("updateTeamMember:", e);
    throw e;
  }
}

function deleteTeamMember(id) {
  try {
    const sh    = _sh(SH_TEAM);
    const found = _findRow(sh, id);
    if (!found) throw new Error("Team member not found: " + id);
    sh.deleteRow(found.index);
    return { success: true };
  } catch (e) {
    console.error("deleteTeamMember:", e);
    throw e;
  }
}
