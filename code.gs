/**
 * THE MULTI-TAB DOCUMENT ENGINE v4.6
 * Fixes: DriveApp Permissions & Search Logic
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 DOCUMENT ENGINE')
      .addItem('Generate & Send PDF Now', 'runEngine')
      .addSeparator()
      .addItem('Initial Setup Check', 'checkConnections')
      .addToUi();
}

function runEngine() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // 1. GET SEARCH TERM
  const response = ui.prompt('Generate Document', 'Enter the SITE NAME (e.g., "Edinburgh Leisure"):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const searchTerm = response.getResponseText().trim();
  
  if (!searchTerm) {
    ui.alert('❌ Error: Site Name cannot be empty.');
    return;
  }

  // 2. DEFENSIVE TABS
  const settings = ss.getSheetByName('⚙️ Settings');
  const quoteSheet = ss.getSheetByName('Quote gen');
  const boqSheet = ss.getSheetByName('Bill of quantities');
  const groundworksSheet = ss.getSheetByName('Ground works B&Q');
  
  if (!settings) { ui.alert('❌ Tab "⚙️ Settings" not found.'); return; }

  // 3. GET SETTINGS & CLEAN DATA
  // Based on your uploaded CSV: Folder ID is near the bottom, Template ID is "Costing Template ID"
  // We search the settings sheet for these labels to be safe
  const settingsData = settings.getDataRange().getValues();
  let templateId = "";
  let folderId = "";
  let bizName = "Your Company";

  settingsData.forEach(row => {
    if (row[0].toString().includes("Costing Template ID")) templateId = row[1].toString().trim();
    if (row[0].toString().includes("Folder ID")) folderId = row[1].toString().trim();
  });

  if (!templateId || !folderId) {
    ui.alert('❌ Setup Error: Could not find Template ID or Folder ID in Settings.');
    return;
  }

  // 4. CONSOLIDATE DATA BY SEARCHING ROWS
  let masterData = {};
  let matchFound = false;

  // Define which column in which sheet contains the "Site Name"
  const searchConfigs = [
    { sheet: quoteSheet, colName: "Site" },
    { sheet: boqSheet, colName: "Project Name" },
    { sheet: groundworksSheet, colName: "Site Name" }
  ];

  searchConfigs.forEach(config => {
    if (!config.sheet) return;
    const data = config.sheet.getDataRange().getValues();
    const headers = data[0];
    const siteColIdx = headers.indexOf(config.colName);

    if (siteColIdx === -1) return;

    for (let i = 1; i < data.length; i++) {
      if (data[i][siteColIdx].toString().toLowerCase() === searchTerm.toLowerCase()) {
        matchFound = true;
        headers.forEach((header, idx) => {
          if (header) masterData[header] = data[i][idx];
        });
      }
    }
  });

  if (!matchFound) {
    ui.alert(`❌ No data found for: "${searchTerm}"`);
    return;
  }

  // 5. DRIVE OPERATIONS
  try {
    // Check Folder Access
    const folder = DriveApp.getFolderById(folderId);
    const templateFile = DriveApp.getFileById(templateId);
    
    const copy = templateFile.makeCopy(`Quote - ${searchTerm}`, folder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    // REPLACEMENT LOGIC
    for (let key in masterData) {
      let tag = `{{${key}}}`;
      let value = masterData[key] || "";
      
      // Auto-format currency
      if (typeof value === 'number' || (!isNaN(value) && value !== "")) {
         if (key.toLowerCase().includes('cost') || key.toLowerCase().includes('total') || key.toLowerCase().includes('price')) {
           value = "£" + Number(value).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
         }
      }
      body.replaceText(tag, value.toString());
    }
    
    body.replaceText('{{Date}}', Utilities.formatDate(new Date(), "GMT", "dd/MM/yyyy"));
    doc.saveAndClose();

    // CONVERT TO PDF
    const pdfBlob = copy.getAs(MimeType.PDF);
    const finalFile = folder.createFile(pdfBlob);
    finalFile.setName(`${searchTerm}_Quote.pdf`);
    
    // Trash temp doc
    copy.setTrashed(true);

    ss.toast(`Successfully created PDF for ${searchTerm}`, '✅ Done');

  } catch (err) {
    ui.alert('Drive Error: ' + err.message + "\n\nCheck if your Folder ID is correct: " + folderId);
  }
}
