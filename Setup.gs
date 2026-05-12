/* ================================================================
   Setup.gs — Run once to create/reset the database sheets
================================================================ */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SHEETS = {
    Projects: [
      "Project ID", "Client Name", "Project Name", "Coordinator",
      "Start Date", "Due Date", "Status", "Source Pages",
      "Source Language", "Notes", "Created At"
    ],
    Tasks: [
      "Task ID", "Project ID", "Project Name", "Client Name",
      "Task Type", "Language", "Pages", "Assigned Person",
      "Work Type", "Vendor Name", "Status", "Start Date",
      "Delivery Date", "Notes", "Created At"
    ]
  };

  Object.keys(SHEETS).forEach(function(name) {
    const existing = ss.getSheetByName(name);
    if (existing) ss.deleteSheet(existing);

    const sheet   = ss.insertSheet(name);
    const headers = SHEETS[name];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold")
      .setBackground("#0f172a")
      .setFontColor("#ffffff");

    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).createFilter();
    sheet.autoResizeColumns(1, headers.length);
  });

  SpreadsheetApp.getUi().alert("✅ Database setup complete. Sheets: Projects + Tasks");
}
