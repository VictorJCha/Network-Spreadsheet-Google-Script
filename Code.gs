// ================== SETTINGS ==================
const PASSWORD_SHEET_NAME = 'Credentials'; // The name of the sheet/tab
const PASSWORD_CELL = 'B2';               // The specific cell to check
const PASSWORD_URL = 'https://www.xkpasswd.net/?c=ZW5nbGlzaCwyLDQsNixSQU5ET00sUkFORE9NLEl3LExUb3VJVDlBSmcsMCwyLEZJWEVELDgsUkFORE9NLCxMVG91SVQ5QUpnLDAsMA'; // URL to generate the password

// DHCP CSV template headers (MUST match uploaded template, in order)
const DHCP_TEMPLATE_HEADERS = [
  'MAC Address',
  'IP Address',
  'Hostname',
  'Local DNS Record',
  'Lease Type',
  'Name',
  'Expiration Time'
];

// Optional default expiration to match your template
// (1969-12-31T16:00:00.000-08:00)
const DHCP_DEFAULT_EXPIRATION = '';

// ================== ON OPEN / EDIT ==================

// Triggered when the sheet is opened
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  // ----- Custom Tools menu -----
ui.createMenu('Cantara Tools')
  .addItem('Create Password', 'openPasswordGenerator')
  .addSeparator()
  .addSubMenu(
    ui.createMenu('Export DHCP Reservations')
      .addItem('Export Management VLAN', 'exportDhcpManagementVlan')
      .addItem('Export AV VLAN', 'exportDhcpAvVlan')
      .addItem('Export Surveillance VLAN', 'exportDhcpSurveillanceVlan')
      .addSeparator()
      .addItem('Export All VLANs', 'exportDhcpAllVlans')
  )
  .addToUi();

// Opens prompt if page flip to Credentials tab hasn't been activated

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PASSWORD_SHEET_NAME);
  if (!sheet) return;

  const passwordValue = sheet.getRange(PASSWORD_CELL).getValue();

  if (!passwordValue) {
    const html = HtmlService.createHtmlOutputFromFile('PasswordPrompt')
      .setWidth(360)     // Wider to fit the logo + text nicely
      .setHeight(300);   // Taller to avoid scrolling
    SpreadsheetApp.getUi().showModalDialog(html, 'Site Admin Password Required');
  }
}

// Opens Password Generator in a new tab
function openPasswordGenerator() {
  const template = HtmlService.createTemplate(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
      </head>
      <body>
        <script>
          window.open('<?= url ?>', '_blank');
          google.script.host.close();
        </script>
      </body>
    </html>
  `);

  template.url = PASSWORD_URL;

  const html = template.evaluate()
    .setWidth(10)
    .setHeight(10);

  SpreadsheetApp.getUi().showModalDialog(html, 'Opening password generator...');
}

// Triggered whenever a cell is edited
function onEdit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const passwordSheet = ss.getSheetByName(PASSWORD_SHEET_NAME);
  if (!passwordSheet) return;

  const passwordValue = passwordSheet.getRange(PASSWORD_CELL).getValue();

  const editedRange = e.range;
  const editedSheet = editedRange.getSheet();
  const editedCell = editedRange.getA1Notation();

  // Password gate: if password is empty and we're not editing the password cell, block the edit
  if (!passwordValue && !(editedSheet.getName() === PASSWORD_SHEET_NAME && editedCell === PASSWORD_CELL)) {
    const html = HtmlService.createHtmlOutputFromFile('PasswordPrompt')
      .setWidth(360)
      .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, 'Site Admin Password Required');
    editedRange.setValue(''); // Clear unauthorized input
    return;
  }

  // If we get here, password is present OR we're editing the password cell itself.
  // Apply MAC/UID normalization where appropriate.
  handleMacAddressEdit_(e);
  handleSavantUidEdit_(e);
}

// Switch the active sheet to "Credentials"
function switchToCredentialsSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Credentials");
  if (sheet) {
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  }
}

// ================== MAC / UID HANDLING ==================

/**
 * MAC / UID handling:
 * - Applies to ANY sheet that has a header "MAC Address" in row 1
 * - Only runs on cells in that column, below the header row
 * - Accepts common formats, e.g.:
 *   - AA:BB:CC:DD:EE:FF
 *   - AA-BB-CC-DD-EE-FF
 *   - AABBCCDDEEFF
 *   - AABB.CCDD.EEFF (Cisco-style)
 *   - Any of the above with 4 extra hex chars on the end (UID): ...YYYY
 * - Always rewrites to: AA:BB:CC:DD:EE:FF (uppercase, colons)
 * - Sets font to Black
 * - Copies background color from cell to the left, includes checks in case its in column A and there are no columns to the left
 * - On invalid input:
 *   - Revert to previous value
 *   - Highlight cell red
 *   - Show toast
 */
function handleMacAddressEdit_(e) {
  const range = e.range;
  const sheet = range.getSheet();

  if (range.getRow() <= 2) return;

  // Skip sheets whose tab name contains "csv"
  if (sheet.getName().toLowerCase().includes("csv")) return;

  const HEADER_ROW = 1;
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;

  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const macColIndex = headers.indexOf("MAC Address"); // exact header match
  if (macColIndex === -1) return;

  const macCol = macColIndex + 1; // convert 0-based index to column number

  // Only process edits in the MAC Address column
  if (range.getColumn() !== macCol) return;

  // Get all values in the edited range (2D array)
  const values = range.getValues();
  const oldValues = range.getOldValues ? range.getOldValues() : values; // optional fallback

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = range.getCell(r + 1, c + 1);
      let val = values[r][c];

      // Empty cell: reset background and continue
      if (!val) {
        clearCellAndMatchRow_(cell);

        // Clear UID partner cell
        const uidCol = headers.findIndex(h =>
          h.toString().toLowerCase().includes("savant uid")
        ) + 1;

        if (uidCol > 0) {
          clearCellAndMatchRow_(sheet.getRange(cell.getRow(), uidCol));
        }

        continue;
      }

      // Normalize: uppercase and remove all separators/spaces
      val = val.toString().toUpperCase().replace(/[:\-\.\s]/g, "");

      // Validate: 12 hex (MAC) or 16 hex (UID)
      const isMac12 = /^[0-9A-F]{12}$/.test(val);
      const isUid16 = /^[0-9A-F]{16}$/.test(val);

      if (!isMac12 && !isUid16) {
        cell.setValue(null);
        cell.setBackground('#ffcccc');
        SpreadsheetApp.getActive().toast(
          "Invalid MAC/UID in '" + sheet.getName() +
          "' — expected 12 or 16 hex characters.",
          "MAC Address Validation",
          5
        );
        continue;
      }

      // Take first 12 chars for MAC portion
      const macHex = val.substring(0, 12);
      const octets = macHex.match(/.{2}/g);
      if (!octets) continue;

      const formattedMac = octets.join(":");

      // Write normalized MAC back, set font to black, and set and copy background from left cell if not column A
      cell.setFontColor("#000000");
      cell.setValue(formattedMac);
      cell.getColumn() > 1 && cell.setBackground(cell.offset(0, -1).getBackground());
    }
  }
}

/**
 * Handles "Savant UID" normalization:
 * - Validates 16-char hex UIDs
 * - Converts to MAC (first 12 chars)
 * - Formats MAC as AA:BB:CC:DD:EE:FF
 * - Writes result to the MAC Address column (same row)
 * - Sets font to Black for UID & MAC
 * - Copies background color from cell to the left, includes checks in case its in column A and there are no columns to the left (for UID & MAC)
 * - On invalid input:
 *   - Revert to previous value
 *   - Highlight cell red
 *   - Show toast
 */
function handleSavantUidEdit_(e) {
  const range = e.range;
  const sheet = range.getSheet();

  if (range.getRow() <= 2) return;

  // Skip sheets with "csv" in the tab name
  if (sheet.getName().toLowerCase().includes("csv")) return;

  const HEADER_ROW = 1;

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];

  const uidCol = headers.findIndex(h => h.toString().toLowerCase().includes("savant uid")) + 1;
  const macCol = headers.indexOf("MAC Address") + 1;

  if (uidCol < 1 || macCol < 1) return; // required columns missing

  // Only process edits in the Savant UID column
  if (range.getColumn() !== uidCol) return;

  // Get all values in the edited range (handles multi-cell pastes)
  const values = range.getValues(); // 2D array

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {

      let uid = values[r][c];
      const cell = range.getCell(r + 1, c + 1);

      //if (!uid) {
      //  cell.setBackground(null); // reset background if empty
      //  continue;
      //}

      if (!uid) {
        clearCellAndMatchRow_(cell);

        // Clear MAC partner cell
        clearCellAndMatchRow_(sheet.getRange(cell.getRow(), macCol));

        continue;
      }

      // Auto-capitalize and remove invalid characters
      uid = uid.toString().toUpperCase().replace(/[^0-9A-F]/g, "");
      cell.setFontColor("#000000"); // set font to black
      cell.setValue(uid); // overwrite with cleaned UID
      cell.getColumn() > 1 && cell.setBackground(cell.offset(0, -1).getBackground()); // Update cell color to match row

      // Validate 16-character hex UID
      if (!/^[0-9A-F]{16}$/.test(uid)) {
        // Invalid: clear value, highlight red and show toast
        range.setValue(null);
        cell.setBackground('#ffcccc');
        SpreadsheetApp.getActive().toast(
          "Invalid Savant UID in '" + sheet.getName() +
          "' — expected exactly 16 hexadecimal characters.",
          "Savant UID Validation",
          5
        );
        continue;
      }

      // cell.setBackground(null);

      // Take first 12 chars for MAC
      const macHex = uid.substring(0, 12);
      const octets = macHex.match(/.{2}/g);
      if (!octets) continue;

      const formattedMac = octets.join(":");

      // Write formatted MAC to MAC Address column in same row and change color of the cell
      const macCell = sheet.getRange(range.getRow() + r, macCol);

      macCell.setFontColor("#000000");
      macCell.setValue(formattedMac);
      macCell.getColumn() > 1 &&
      macCell.setBackground(macCell.offset(0, -1).getBackground());

    }
  }
}

// ================== CELL CLEANER AND COLOR MATCH HELPER ==================

// Used by MAC and UID function to Clear partner cell and set background color
function clearCellAndMatchRow_(cell) {
  cell.setValue("");
  cell.getColumn() > 1 &&
    cell.setBackground(cell.offset(0, -1).getBackground());
}

// ================== DHCP EXPORT MENU HANDLERS ==================

function exportDhcpManagementVlan() {
  exportDhcpForVlan_('Management');
}

function exportDhcpAvVlan() {
  exportDhcpForVlan_('AV');
}

function exportDhcpSurveillanceVlan() {
  exportDhcpForVlan_('Surveillance');
}

function exportDhcpAllVlans() {
  exportDhcpForVlan_('ALL');
}

/**
 * Build filename based on VLAN key.
 * Patterns:
 *  Management   -> Unifi_Mgmt_VLAN_DHCP_<timestamp>.csv
 *  AV           -> Unifi_AV_VLAN_DHCP_<timestamp>.csv
 *  Surveillance -> Unifi_Surveillance_VLAN_DHCP_<timestamp>.csv
 *  ALL          -> Unifi_All_VLAN_DHCP_<timestamp>.csv
 */
function buildDhcpFilename_(vlanKey) {
  let prefix;
  switch (vlanKey) {
    case 'Management':
      prefix = 'Unifi_Mgmt_VLAN_DHCP_';
      break;
    case 'AV':
      prefix = 'Unifi_AV_VLAN_DHCP_';
      break;
    case 'Surveillance':
      prefix = 'Unifi_Surveillance_VLAN_DHCP_';
      break;
    case 'ALL':
    default:
      prefix = 'Unifi_All_VLAN_DHCP_';
      break;
  }

  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMdd_HHmmss'
  );

  return prefix + timestamp + '.csv';
}

/**
 * Core export function.
 * vlanKey is one of: 'Management', 'AV', 'Surveillance', 'ALL'
 */
function exportDhcpForVlan_(vlanKey) {
  const ui = SpreadsheetApp.getUi();

  // 1) Get raw records from the sheet(s) for this VLAN
  const records = getDhcpRecordsForVlan_(vlanKey);

  if (!records || records.length === 0) {
    ui.alert('No DHCP reservations found for ' + vlanKey + (vlanKey === 'ALL' ? '' : ' VLAN') + '.');
    return;
  }

  // 2) Convert records into CSV rows following the template column order:
  //    [MAC Address, IP Address, Hostname, Local DNS Record, Lease Type, Name, Expiration Time]
  const dataRows = records.map(rec => {
    const mac = rec.mac || '';
    const ip = rec.ip || '';
    const name = rec.name || '';

    return [
      mac,                    // MAC Address
      ip,                     // IP Address
      '',                     // Hostname (left blank)
      '',                     // Local DNS Record (left blank)
      'Fixed',                // Lease Type (always "Fixed")
      name,                   // Name
      DHCP_DEFAULT_EXPIRATION // or '' if you prefer blank
    ];
  });

  // 3) Build CSV string
  const allRows = [DHCP_TEMPLATE_HEADERS].concat(dataRows);
  const csvContent = allRows
    .map(row => row.map(value => {
      const str = value == null ? '' : String(value);
      // Escape double quotes and wrap in quotes if needed
      if (str.includes('"') || str.includes(',') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }).join(','))
    .join('\r\n');

  // 4) Trigger a direct download via HTML dialog
  const filename = buildDhcpFilename_(vlanKey);
  showCsvDownloadDialog_(filename, csvContent);
}

// ================== DHCP DATA GATHERING ==================

function getSheetNameForVlanKey_(vlanKey) {
  switch (vlanKey) {
    case 'Management':
      return 'Management VLAN 1';
    case 'AV':
      return 'AV VLAN 5';
    case 'Surveillance':
      return 'Surveillance VLAN 7';
    default:
      return null;
  }
}


/**
 * Fetch DHCP data for a given VLAN or all VLANs.
 *
 * Returns an array of objects:
 *   { mac: string, ip: string, name: string, vlan: string }
 *
 * Mapping from sheet:
 *   MAC Address (sheet)  -> mac
 *   OVRC Name (sheet)    -> name
 *   IP Address (sheet)   -> ip
 *
 * Rules:
 *   - Skip row if MAC is missing
 *   - Skip row if IP is missing
 *   - Skip row if OVRC Name is missing
 *   - Toast a warning summarizing skipped rows + reasons
 */
function getDhcpRecordsForVlan_(vlanKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const keys = (vlanKey === 'ALL')
    ? ['Management', 'AV', 'Surveillance']
    : [vlanKey];

  const records = [];
  const skipped = []; // collect skipped rows + reasons

  keys.forEach(function (key) {
    const sheetName = getSheetNameForVlanKey_(key);
    if (!sheetName) return;

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol === 0) return; // no data

    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = values[0];

    // EXACT header matching
    const macColIdx  = headers.indexOf('MAC Address');
    const nameColIdx = headers.indexOf('OVRC Name');
    const ipColIdx   = headers.indexOf('IP');

    if (macColIdx === -1 || nameColIdx === -1 || ipColIdx === -1) {
      ui.alert(
        'DHCP Export',
        'Sheet "' + sheetName + '" is missing one of the required headers:\n\n' +
        ' • MAC Address\n' +
        ' • OVRC Name\n' +
        ' • IP\n\n' +
        'Please correct the header row.',
        ui.ButtonSet.OK
      );
      return;
    }

    // Optional Manufacturer column (Management VLAN 1 only)
    let manufacturerColIdx = -1;
    let manufacturerFilterEnabled = false;

    if (sheetName === 'Management VLAN 1') {
      manufacturerColIdx = headers.indexOf('Manufacturer');
      if (manufacturerColIdx === -1) {
        // Warn once per export if Manufacturer column is missing
        SpreadsheetApp.getActive().toast(
          'Management VLAN 1 is missing "Manufacturer" column. Ubiquiti devices will NOT be auto-skipped.',
          'DHCP Export Warning',
          10
        );
      } else {
        manufacturerFilterEnabled = true;
      }
    }

    // Iterate through rows
    for (let r = 1; r < values.length; r++) {
      const row  = values[r];
      const mac  = row[macColIdx];
      const name = row[nameColIdx];
      const ip   = row[ipColIdx];
      let reason = null;

      // Extra rule: on Management VLAN 1, skip Ubiquiti-manufactured devices
      if (manufacturerFilterEnabled) {
        const manufacturer = row[manufacturerColIdx];
        if (manufacturer === 'Ubiquiti') {
          reason = 'Ubiquiti manufacturer (handled separately)';
        }
      }

      // Base validation: require MAC, IP, OVRC Name
      if (!reason) {
        if (!mac)      reason = 'Missing MAC Address';
        else if (!ip)  reason = 'Missing IP Address';
        else if (!name) reason = 'Missing OVRC Name';
      }

      if (reason) {
        skipped.push({
          sheet: sheetName,
          row:   r + 1, // 1-based row number for the UI
          reason: reason
        });
        continue;
      }

      // Row validated
      records.push({
        mac:  mac,
        ip:   ip,
        name: name,
        vlan: key
      });
    }
  });

  // Toast summarizing skipped rows
  if (skipped.length > 0) {
    const summary = skipped
      .slice(0, 5)
      .map(s => `${s.sheet} R${s.row}: ${s.reason}`)
      .join(' | ');

    const moreText = skipped.length > 5
      ? ` ...and ${skipped.length - 5} more`
      : '';

    SpreadsheetApp.getActive().toast(
      `Some rows were skipped: ${summary}${moreText}`,
      'DHCP Export Warning',
      10
    );
  }

  return records;
}

// ================== HTML DOWNLOAD BRIDGE ==================

/**
 * Opens a tiny HTML dialog that immediately triggers a CSV download
 * on the user's machine using a Blob + a[n] <a download> click.
 */
function showCsvDownloadDialog_(filename, csvContent) {
  const template = HtmlService.createTemplateFromFile('CsvDownload');
  template.filename = filename;
  template.csvBase64 = Utilities.base64Encode(csvContent);

  const html = template.evaluate()
    .setWidth(10)
    .setHeight(10);

  SpreadsheetApp.getUi().showModalDialog(html, 'Preparing download...');
}
