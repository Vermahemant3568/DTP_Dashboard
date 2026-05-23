/* ================================================================
   Setup.gs — Run once to create all 6 sheets

   SCHEMA & RELATIONSHIPS
   ─────────────────────────────────────────────────────────────────
   Projects  (1) ──< Tasks     (many)  via Tasks.Project ID
   Projects  (1) ──< Revisions (many)  via Revisions.Project ID
   Tasks     (1) ──< Revisions (many)  via Revisions.Task ID (optional)
   Tasks.Assigned To  → Team.Name    (Work Type = In-House)
   Tasks.Vendor Name  → Vendors.Name (Work Type = Vendor)
   Revisions.Assigned To → Team.Name
   Revisions.Vendor Name → Vendors.Name

   DENORMALIZATION POLICY
   ─────────────────────────────────────────────────────────────────
   Tasks and Revisions store Client Name and Project Name as
   denormalized copies from Projects. This avoids expensive lookups
   on every read in Google Sheets. These are written once on creation
   and treated as read-only snapshots thereafter.

   PAYMENT STATUS VALUES
   ─────────────────────────────────────────────────────────────────
   Paid    → included in paid totals
   Partial → included in paid totals (partially settled)
   Unpaid  → tracked separately, excluded from paid totals
   N/A     → excluded from all payment calculations

   Run setupDatabase()    — create all sheets (safe, never overwrites)
   Run addMissingColumns() — add new columns to existing sheets safely
================================================================ */

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SHEETS = {

    /* ── 1. PROJECTS — master project registry ── */
    /* 14 columns. Word Count removed (never used). */
    Projects: [
      "Project ID", "Client Name", "Project Name", "Project Coordinator",
      "Source Language", "Target Languages", "Target Lang Count", "Source Pages",
      "Priority", "Status", "Received Date",
      "Notes", "Created At", "Updated At"
    ],

    /* ── 2. TASKS — every unit of work ── */
    /* 25 columns. Client Name + Project Name are denormalized from Projects. */
    /* Assigned To (col 6) used when Work Type = In-House. */
    /* Vendor Name (col 7) used when Work Type = Vendor. */
    Tasks: [
      "Task ID", "Project ID", "Client Name", "Project Name",
      "Task Type", "Work Type", "Assigned To", "Vendor Name",
      "Language", "Source Pages", "Final Pages", "Lang Count",
      "Status", "Priority", "Start Date", "Delivery Date",
      "Completed Date", "Source Link", "Deliverable Link", "Notes",
      "Created At", "Updated At",
      "Rate Per Page", "Currency", "Payment Status"
    ],

    /* ── 3. REVISIONS — rework rounds ── */
    /* 21 columns. Project Name is denormalized from Projects. */
    /* Task ID (col 2) is optional — links revision to a specific task. */
    Revisions: [
      "Revision ID", "Project ID", "Task ID", "Project Name",
      "Revision Number", "Revision Type", "Language", "Revision Pages",
      "Work Type", "Assigned To", "Vendor Name", "Status",
      "Revision Date", "Delivery Date", "Completed Date", "Notes",
      "Created At", "Updated At",
      "Rate Per Page", "Currency", "Payment Status"
    ],

    /* ── 4. VENDORS — vendor master list ── */
    /* 13 columns. Added Updated At for consistency with all other sheets. */
    Vendors: [
      "Vendor ID", "Vendor Name", "Contact Person", "Email",
      "Phone", "Specialization", "Languages", "Rate Per Page",
      "Currency", "Status", "Notes", "Created At", "Updated At"
    ],

    /* ── 5. TEAM — in-house team member list ── */
    /* 10 columns. */
    Team: [
      "Member ID", "Name", "Role", "Email",
      "Phone", "Specialization", "Status", "Rate Per Page",
      "Currency", "Created At"
    ],

    /* ── 6. MONTHLY SNAPSHOT — pre-computed summaries ── */
    MonthlySnapshot: [
      "Snapshot ID", "Year Month", "Task Type", "Work Type",
      "Total Pages", "Total Tasks", "In House Pages", "Vendor Pages",
      "Generated At"
    ]
  };

  const COLORS = {
    Projects:        "#0f172a",
    Tasks:           "#1e3a5f",
    Revisions:       "#3b0764",
    Vendors:         "#064e3b",
    Team:            "#1c1917",
    MonthlySnapshot: "#1e1b4b"
  };

  Object.keys(SHEETS).forEach(function(name) {
    let sheet = ss.getSheetByName(name);
    // Only create if it doesn't exist — never delete existing data
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    const headers = SHEETS[name];
    const color   = COLORS[name] || "#0f172a";

    // Write headers only if row 1 is empty
    const firstCell = sheet.getRange(1, 1).getValue();
    if (!firstCell) {
      sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight("bold")
        .setBackground(color)
        .setFontColor("#ffffff");
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, headers.length);
    }
  });

  SpreadsheetApp.getUi().alert(
    "✅ Database setup complete.\n\n" +
    "Sheets created:\n" +
    "• Projects (14 cols)\n" +
    "• Tasks (25 cols)\n" +
    "• Revisions (21 cols)\n" +
    "• Vendors (13 cols)\n" +
    "• Team (10 cols)\n" +
    "• MonthlySnapshot (9 cols)\n\n" +
    "If you have existing data, run migrateOldProjectsToTasks() next."
  );
}

/* ================================================================
   ADD MISSING COLUMNS
   Safe to run on a live sheet with existing data.
   Checks the current header row of each sheet and appends only the
   columns that are not already present. Never deletes or moves data.
   Run this after any schema update (e.g. adding Rate Per Page etc.).
================================================================ */
function addMissingColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  /* Full expected header list per sheet — order matters, new cols at end */
  const EXPECTED = {
    Tasks: [
      "Task ID", "Project ID", "Client Name", "Project Name",
      "Task Type", "Work Type", "Assigned To", "Vendor Name",
      "Language", "Source Pages", "Final Pages", "Lang Count",
      "Status", "Priority", "Start Date", "Delivery Date",
      "Completed Date", "Source Link", "Deliverable Link", "Notes",
      "Created At", "Updated At",
      "Rate Per Page", "Currency", "Payment Status"
    ],
    Projects: [
      "Project ID", "Client Name", "Project Name", "Project Coordinator",
      "Source Language", "Target Languages", "Target Lang Count", "Source Pages",
      "Priority", "Status", "Received Date",
      "Notes", "Created At", "Updated At"
    ],
    Revisions: [
      "Revision ID", "Project ID", "Task ID", "Project Name",
      "Revision Number", "Revision Type", "Language", "Revision Pages",
      "Work Type", "Assigned To", "Vendor Name", "Status",
      "Revision Date", "Delivery Date", "Completed Date", "Notes",
      "Created At", "Updated At",
      "Rate Per Page", "Currency", "Payment Status"
    ],
    Vendors: [
      "Vendor ID", "Vendor Name", "Contact Person", "Email",
      "Phone", "Specialization", "Languages", "Rate Per Page",
      "Currency", "Status", "Notes", "Created At", "Updated At"
    ],
    Team: [
      "Member ID", "Name", "Role", "Email",
      "Phone", "Specialization", "Status", "Rate Per Page",
      "Currency", "Created At"
    ]
  };

  const COLORS = {
    Tasks:     "#1e3a5f",
    Projects:  "#0f172a",
    Revisions: "#3b0764",
    Vendors:   "#064e3b",
    Team:      "#1c1917"
  };

  var report = [];

  Object.keys(EXPECTED).forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      report.push("⚠️  " + sheetName + ": sheet not found — skipped.");
      return;
    }

    var lastCol      = sheet.getLastColumn();
    var existingHdrs = lastCol > 0
      ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(v) { return String(v).trim(); })
      : [];

    var expected  = EXPECTED[sheetName];
    var color     = COLORS[sheetName] || "#0f172a";
    var added     = [];

    expected.forEach(function(hdr) {
      if (!existingHdrs.includes(hdr)) {
        /* Append this header in the next available column */
        var newCol = sheet.getLastColumn() + 1;
        var cell   = sheet.getRange(1, newCol);
        cell.setValue(hdr)
            .setFontWeight("bold")
            .setBackground(color)
            .setFontColor("#ffffff");
        sheet.autoResizeColumn(newCol);
        added.push(hdr);
      }
    });

    if (added.length) {
      report.push("✅ " + sheetName + ": added " + added.length + " column(s) — " + added.join(", "));
    } else {
      report.push("✔️  " + sheetName + ": all columns already present.");
    }
  });

  SpreadsheetApp.getUi().alert(
    "Column Migration Complete\n\n" + report.join("\n")
  );
}

/* ================================================================
   MIGRATION — copies old Projects sheet data into new Tasks sheet
   Safe: reads old data, writes to Tasks, does NOT delete old sheet.
   Run once after setupDatabase().
================================================================ */
function migrateOldProjectsToTasks() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const oldSheet  = ss.getSheetByName("Projects");
  const taskSheet = ss.getSheetByName("Tasks");
  const projSheet = ss.getSheetByName("Projects");

  if (!oldSheet) {
    SpreadsheetApp.getUi().alert("❌ Old Projects sheet not found.");
    return;
  }
  if (!taskSheet) {
    SpreadsheetApp.getUi().alert("❌ Tasks sheet not found. Run setupDatabase() first.");
    return;
  }

  const data = oldSheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("ℹ️ No data rows found in Projects sheet.");
    return;
  }

  // Old column map (0-based):
  // 0=projectId, 1=customerName, 2=projectName, 3=projectCoordinator
  // 4=modeOfTask, 5=assignedPerson, 6=vendorTeam, 7=priority
  // 8=startDate, 9=deliveryDate, 10=status, 11=sourcePages
  // 12=sourceLanguage, 13=targetLangCount, 14=targetLangNames
  // 15=finalPages, 16=notes, 17=sourceLink, 18=deliverableLink
  // 19=createdAt, 20=updatedAt, 21=taskType, 22=workType

  const rows    = data.slice(1).filter(r => r[0] !== "");
  const now     = new Date();
  let   migrated = 0;

  rows.forEach(function(r) {
    const oldId    = String(r[0]).trim();
    const taskType = r[21] || "Main DTP";
    const workType = r[22] || "In-House";

    // Generate new Task ID preserving timestamp from old PRJ- id
    const taskId = "TSK-" + Date.now() + "-" + migrated;

    taskSheet.appendRow([
      taskId,                          // Task ID
      oldId,                           // Project ID (keep old PRJ- id as project ref)
      r[1]  || "",                     // Client Name
      r[2]  || "",                     // Project Name
      taskType,                        // Task Type
      workType,                        // Work Type
      r[5]  || "",                     // Assigned To
      r[6]  || "",                     // Vendor Name
      r[14] || "",                     // Language (target lang names)
      Number(r[11]) || 0,              // Source Pages
      Number(r[15]) || 0,              // Final Pages
      Number(r[13]) || 0,              // Lang Count
      r[10] || "Pending",              // Status
      r[7]  || "Medium",               // Priority
      r[8]  || "",                     // Start Date
      r[9]  || "",                     // Delivery Date
      r[10] === "Completed" ? r[20] || "" : "", // Completed Date
      r[17] || "",                     // Source Link
      r[18] || "",                     // Deliverable Link
      r[16] || "",                     // Notes
      r[19] || now,                    // Created At
      r[20] || now                     // Updated At
    ]);

    // Also ensure a master Projects record exists for this project
    // (we use the old PRJ- id as the ProjectID in the new Projects sheet)
    const existingProj = _findRowInSheet(projSheet, oldId);
    if (!existingProj) {
      projSheet.appendRow([
        oldId,                         // Project ID
        r[1]  || "",                   // Client Name
        r[2]  || "",                   // Project Name
        r[3]  || "",                   // Project Coordinator
        r[12] || "English",            // Source Language
        r[14] || "",                   // Target Languages
        Number(r[13]) || 0,            // Target Lang Count
        Number(r[11]) || 0,            // Source Pages
        r[7]  || "Medium",             // Priority
        r[10] || "Pending",            // Status
        r[8]  || "",                   // Received Date
        r[16] || "",                   // Notes
        r[19] || now,                  // Created At
        r[20] || now                   // Updated At
      ]);
    }

    migrated++;
    Utilities.sleep(50); // avoid quota issues
  });

  SpreadsheetApp.getUi().alert(
    "✅ Migration complete.\n" +
    migrated + " project rows migrated to Tasks sheet.\n\n" +
    "The old Projects sheet has NOT been deleted.\n" +
    "Verify the data, then you can rename it to 'Projects_OLD'."
  );
}

/* helper used only inside Setup.gs */
function _findRowInSheet(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) return { row: data[i], index: i + 1 };
  }
  return null;
}

/* ================================================================
   REMOVE WORD COUNT COLUMN FROM PROJECTS
   Run once on a live sheet that still has the old 15-column schema.
   Finds the "Word Count" header and deletes that column.
   Safe: checks header name before deleting, never touches other cols.
================================================================ */
function removeWordCountColumn() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Projects");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ Projects sheet not found.");
    return;
  }
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colIdx  = headers.indexOf("Word Count");
  if (colIdx === -1) {
    SpreadsheetApp.getUi().alert("✔️  Word Count column not found — nothing to remove.");
    return;
  }
  sheet.deleteColumn(colIdx + 1);  // deleteColumn is 1-based
  SpreadsheetApp.getUi().alert("✅ Word Count column removed from Projects sheet.");
}
