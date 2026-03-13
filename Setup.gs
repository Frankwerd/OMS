/********************************
 * Setup.gs (Standalone, No Triggers)
 * - NO ScriptApp.newTrigger
 * - NO onOpen
 * - Dashboard merges are filter-safe
 ********************************/

function omsSetupSheet() {
  const ss = OMS_Utils.ss();

  const inbound = getOrCreateSheet_(ss, OMS_CONFIG.TABS.INBOUND);
  const outbound = getOrCreateSheet_(ss, OMS_CONFIG.TABS.OUTBOUND);
  const dashboard = getOrCreateSheet_(ss, OMS_CONFIG.TABS.DASHBOARD);
  const masterTable = getOrCreateSheet_(ss, OMS_CONFIG.TABS.MASTER_TABLE);
  const meta = getOrCreateSheet_(ss, OMS_CONFIG.TABS.META);

  // IMPORTANT:
  // Inbound/Outbound get filters (table sheets)
  // Dashboard gets NO filter (merge-safe)
  // clearAll = false to preserve manual data while adding missing schema columns.
  applyHeaderRow_(inbound, OMS_SCHEMA_INBOUND_(), { createFilter: true, clearAll: false });
  applyHeaderRow_(outbound, OMS_SCHEMA_OUTBOUND_(), { createFilter: true, clearAll: false });
  applyHeaderRow_(dashboard, OMS_SCHEMA_DASHBOARD_(), { createFilter: false, clearAll: true });
  applyHeaderRow_(masterTable, OMS_SCHEMA_MASTER_TABLE_(), { createFilter: true, clearAll: true });

  updateMetaSheet_(ss, meta, {
    [OMS_CONFIG.TABS.INBOUND]: OMS_SCHEMA_INBOUND_(),
    [OMS_CONFIG.TABS.OUTBOUND]: OMS_SCHEMA_OUTBOUND_(),
    [OMS_CONFIG.TABS.MASTER_TABLE]: OMS_SCHEMA_MASTER_TABLE_(),
  });
  meta.hideSheet();

  styleInbound_(inbound);
  styleOutbound_(outbound);

  buildDashboard_(dashboard); // does merges safely because dashboard has no filter
  refreshMasterOmsTable(); // Generate the master table content

  inbound.setTabColor('#1E3A8A');
  outbound.setTabColor('#047857');
  dashboard.setTabColor('#111827');
  masterTable.setTabColor('#374151');

  // Validate after setup/repair
  validateSchema(ss);

  SpreadsheetApp.flush();
}

function omsRefreshDashboard() {
  const ss = OMS_Utils.ss();
  const dashboard = ss.getSheetByName(OMS_CONFIG.TABS.DASHBOARD);
  if (!dashboard) throw new Error('Dashboard sheet not found. Run omsSetupSheet() first.');

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Refresh Dashboard', 'Enter date range (e.g. 2024-01-01 to 2024-01-31) or leave blank to keep current:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  let startDate = dashboard.getRange('B3').getValue();
  let endDate = dashboard.getRange('D3').getValue();

  if (input) {
    const parts = input.split(/ to | - |,/i);
    startDate = new Date(parts[0].trim());
    if (parts.length > 1) {
      endDate = new Date(parts[1].trim());
    } else {
      endDate = new Date(startDate);
    }

    if (!isNaN(startDate.getTime()) && !isNaN(endDate.getTime())) {
      dashboard.getRange('B3').setValue(startDate);
      dashboard.getRange('D3').setValue(endDate);
    } else {
      ui.alert('Invalid date format. Proceeding with current dates.');
      startDate = dashboard.getRange('B3').getValue();
      endDate = dashboard.getRange('D3').getValue();
    }
  }

  // Ensure Master Table is refreshed with the same range
  refreshMasterOmsTable(startDate, endDate);

  buildDashboard_(dashboard);
  SpreadsheetApp.flush();
}

/** ---------------- SCHEMAS ---------------- **/

function OMS_SCHEMA_INBOUND_() {
  return [
    'merchant-order-id','merchant-order-item-id','line-item-index','purchase-date','purchase-time','order-created-at','buyer-email','buyer-name',
    'buyer-phone-number','customer-id','system-gmail-id','order-source-email','sales-channel','customer-classification','is-business-order',
    'sku','product-name','mag-safe-stand','model','club-type','product-category','hand','flex','shaft-length-option','grip-size','head-material',
    'shaft-material','loft','lie-angle','offset','quantity-purchased','currency','item-price','item-tax','shipping-price',
    'total-amount','coupon-code','discount-amount','refund-amount','refund-date','return-reason-code','recipient-name','ship-address-1',
    'ship-city','ship-state','ship-postal-code','ship-country','ship-service-level','serial-number-allocated','notes',
    'automation-notes','item-life-cycle','order-life-cycle','replacement-for-order-id','replacement-for-order-item-id',
    'replacement-type','replacement-group-id','system-created-at','system-updated-at','parse-status',
    'source-system','source-order-id','source-order-item-id','oms-order-id','oms-order-item-id','buyer-email-hash'
  ];
}

function OMS_SCHEMA_OUTBOUND_() {
  // Added: us-ship-date, delivered-date (needed for velocity metrics)
  // Added: package weight/dims for outbound stub creation
  return [
    'merchant-order-id','merchant-order-item-id','sku','customer-id','outbound-workflow-type',
    'original-merchant-order-id','original-merchant-order-item-id',
    'order-created-at','delivery-country','package-type',
    'domestic-tracking-kr','hub-received-date','hub-location','international-tracking-us','carrier-us',
    'us-ship-date','delivered-date',
    'outbound-status','serial-number-scanned','sn-verify','customer-email-status','last-email-at','system-updated-at','notes','stage-timeline',
    'oms-order-id','oms-order-item-id','shipment-id',
    'actual-weight-kg','package-length-cm','package-width-cm','package-height-cm'
  ];
}

function OMS_SCHEMA_DASHBOARD_() {
  // Dashboard sheet doesn’t need a big table header; it’s a layout sheet.
  return ['metric','value','notes'];
}

function OMS_SCHEMA_MASTER_TABLE_() {
  return [
    'oms-order-id','oms-order-item-id','source-system','source-order-id','source-order-item-id',
    'merchant-order-id','merchant-order-item-id','sku','customer-id','buyer-email-hash','buyer-email','buyer-name',
    'purchase-date','sales-channel','item-life-cycle','order-life-cycle','replacement-type','replacement-group-id',
    'quantity-purchased','currency','item-price','item-tax','shipping-price','total-amount','refund-amount','refund-date',
    'serial-number-allocated','serial-number-scanned','sn-verify',
    'domestic-tracking-kr','hub-received-date','international-tracking-us','carrier-us',
    'us-ship-date','delivered-date',
    'outbound-status','customer-email-status','last-email-at','shipment-id','notes'
  ];
}

/** ---------------- SHEET HELPERS ---------------- **/

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/**
 * Schema Validation
 * Throws error if required headers are missing or version mismatch.
 */
function validateSchema(ss) {
  const meta = ss.getSheetByName(OMS_CONFIG.TABS.META);
  if (!meta) return; // Skip if meta doesn't exist yet (first run)

  const lr = meta.getLastRow();
  if (lr < 2) return;

  // Check version from a hidden cell or property
  const props = PropertiesService.getScriptProperties();
  const currentVersion = props.getProperty('OMS_SCHEMA_VERSION');
  if (currentVersion && currentVersion !== OMS_CONFIG.SCHEMA_VERSION) {
    OMS_Utils.opsAlert_(`Schema Version Mismatch! Code expects ${OMS_CONFIG.SCHEMA_VERSION}, found ${currentVersion}. Running setup might be required.`);
  }
  props.setProperty('OMS_SCHEMA_VERSION', OMS_CONFIG.SCHEMA_VERSION);

  // Validate Inbound
  const inbound = ss.getSheetByName(OMS_CONFIG.TABS.INBOUND);
  if (inbound) {
    const required = [
      'merchant-order-id','buyer-email','sku','oms-order-item-id',
      'line-item-index','order-created-at','order-source-email','product-category',
      'shaft-length-option','discount-amount','buyer-email-hash'
    ];
    try {
      OMS_Utils.requireCols_(inbound, required);
    } catch (e) {
      throw new Error(`CRITICAL: Inbound sheet schema validation failed. ${e.message}`);
    }
  }

  // Validate Outbound
  const outbound = ss.getSheetByName(OMS_CONFIG.TABS.OUTBOUND);
  if (outbound) {
    const required = [
      'oms-order-item-id','shipment-id','outbound-status','stage-timeline',
      'order-created-at','delivery-country','package-type',
      'actual-weight-kg','package-length-cm','package-width-cm','package-height-cm'
    ];
    try {
      OMS_Utils.requireCols_(outbound, required);
    } catch (e) {
      throw new Error(`CRITICAL: Outbound sheet schema validation failed. ${e.message}`);
    }
  }
}

/**
 * Update _Meta sheet with header-to-column mappings
 */
function updateMetaSheet_(ss, meta, schemas) {
  meta.clear();
  const rows = [['sheet', 'header', 'column', 'letter']];

  for (const sheetName in schemas) {
    const headers = schemas[sheetName];
    headers.forEach((h, i) => {
      const col = i + 1;
      rows.push([sheetName, h, col, OMS_Utils.columnLetter_(col)]);
    });
  }

  meta.getRange(1, 1, rows.length, 4).setValues(rows);
  meta.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#efefef');
}

/**
 * ✅ FIX: Always remove an existing filter BEFORE doing any merges later.
 * ✅ Master sheet: createFilter=false so dashboard merges are always allowed.
 * Logic: If columns missing, append. Never overwrite existing columns.
 */
function applyHeaderRow_(sheet, headers, options) {
  const opts = options || {};
  const createFilter = (opts.createFilter !== false);
  const clearAll = (opts.clearAll !== false);

  // Remove filter FIRST (prevents merge conflicts)
  const existing = sheet.getFilter();
  if (existing) existing.remove();

  if (clearAll) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    // Check which headers are missing
    const currentMap = OMS_Utils.getHeadersMap_(sheet);
    const missing = headers.filter(h => !currentMap[h.toLowerCase()]);
    if (missing.length) {
      sheet.getRange(1, sheet.getLastColumn() + 1, 1, missing.length).setValues([missing]);
    }
  }

  sheet.setFrozenRows(1);

  const lastCol = sheet.getLastColumn();
  if (lastCol > 0) {
    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight('bold')
      .setBackground('#111827')
      .setFontColor('#FFFFFF')
      .setWrap(true)
      .setVerticalAlignment('middle');

    sheet.setRowHeight(1, 36);

    if (createFilter) {
      sheet.getRange(1, 1, 1, lastCol).createFilter();
    }
  }
}

/** ---------------- DROPDOWNS + COLORS ---------------- **/

function styleInbound_(sheet) {
  sheet.setHiddenGridlines(false);
  sheet.getDataRange().setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  sheet.setFrozenColumns(6);

  const map = OMS_Utils.getHeadersMap_(sheet);

  // widths
  setColWidth_(sheet, map, 'oms-order-id', 220);
  setColWidth_(sheet, map, 'oms-order-item-id', 260);
  setColWidth_(sheet, map, 'buyer-email', 220);
  setColWidth_(sheet, map, 'buyer-email-hash', 250);
  setColWidth_(sheet, map, 'ship-address-1', 300);
  setColWidth_(sheet, map, 'sku', 160);
  setColWidth_(sheet, map, 'product-name', 220);
  setColWidth_(sheet, map, 'order-created-at', 150);
  setColWidth_(sheet, map, 'order-source-email', 220);
  setColWidth_(sheet, map, 'product-category', 140);
  setColWidth_(sheet, map, 'discount-amount', 120);
  setColWidth_(sheet, map, 'notes', 260);
  setColWidth_(sheet, map, 'automation-notes', 260);

  // formats
  setNumberFormat_(sheet, map, 'purchase-date', 'yyyy-mm-dd');
  setNumberFormat_(sheet, map, 'order-created-at', 'yyyy-mm-dd hh:mm:ss');
  setNumberFormat_(sheet, map, 'refund-date', 'yyyy-mm-dd');
  setNumberFormat_(sheet, map, 'item-price', '$0.00');
  setNumberFormat_(sheet, map, 'item-tax', '$0.00');
  setNumberFormat_(sheet, map, 'shipping-price', '$0.00');
  setNumberFormat_(sheet, map, 'total-amount', '$0.00');
  setNumberFormat_(sheet, map, 'discount-amount', '$0.00');
  setNumberFormat_(sheet, map, 'refund-amount', '$0.00');
  setNumberFormat_(sheet, map, 'quantity-purchased', '0');
  setNumberFormat_(sheet, map, 'system-created-at', 'yyyy-mm-dd hh:mm:ss');
  setNumberFormat_(sheet, map, 'system-updated-at', 'yyyy-mm-dd hh:mm:ss');

  // dropdowns (based on your prior sheet + Pro/Basic model)
  applyDropdown_(sheet, map, 'customer-classification', ['Active','Replaced','Refunded','']);
  applyDropdown_(sheet, map, 'model', ['Pro','Basic','']);
  applyDropdown_(sheet, map, 'grip-size', ['Standard','Mid','']);
  applyDropdown_(sheet, map, 'shaft-length-option', ['Standard','Longer','']);
  applyDropdown_(sheet, map, 'hand', ['Right','Left','']);
  applyDropdown_(sheet, map, 'flex', ['L','R','S','X','']);
  applyDropdown_(sheet, map, 'mag-safe-stand', ['Yes','0','']);
  applyDropdown_(sheet, map, 'item-life-cycle', ['ACTIVE','REFUNDED','RETURNED','REPLACED','CANCELLED','']);
  applyDropdown_(sheet, map, 'order-life-cycle', ['ACTIVE','PARTIAL_REFUND','FULL_REFUND','CANCELLED','']);
  applyDropdown_(sheet, map, 'parse-status', ['OK','ERROR','RES_AUTO','MANUAL_EDIT','']);
  applyDropdown_(sheet, map, 'source-system', ['amazon_fba','shopify','samcart','imweb','manual','unknown','']);

  // UI: Section Shading
  const sections = [
    { color: '#EEF2FF', headers: ['merchant-order-id','merchant-order-item-id','line-item-index','purchase-date','purchase-time','order-created-at','source-*','oms-*','system-gmail-id','order-source-email','buyer-email-hash','customer-id'] },
    { color: '#ECFEFF', headers: ['buyer-email','buyer-name','buyer-phone-number','recipient-name'] },
    { color: '#ECFDF5', headers: ['sku','product-name','product-category','model','club-type','hand','flex','shaft-length-option','grip-size','mag-safe-stand'] },
    { color: '#FFFBEB', headers: ['currency','item-price','item-tax','shipping-price','discount-amount','refund-amount','total-amount', 'coupon-code'] },
    { color: '#F5F3FF', headers: ['ship-address-1','ship-city','ship-state','ship-postal-code','ship-country','ship-service-level'] },
    { color: '#FFF1F2', headers: ['serial-number-allocated','item-life-cycle','order-life-cycle','replacement-*','parse-status','notes','automation-notes'] }
  ];
  OMS_Utils.applySectionShading_(sheet, map, sections);

  // Restore dark header readability
  const lc = sheet.getLastColumn();
  if (lc > 0) {
    sheet.getRange(1, 1, 1, lc)
      .setFontWeight('bold')
      .setBackground('#111827')
      .setFontColor('#FFFFFF');
  }

  // UI: Banding
  OMS_Utils.applyBanding_(sheet);

  // conditional formatting to mimic “chips”
  sheet.setConditionalFormatRules([]);
  const rules = [];
  const maxRows = sheet.getMaxRows();
  const lastCol = sheet.getLastColumn();

  // A) Parse failed (row highlight)
  if (map['parse-status']) {
    const full = sheet.getRange(2, 1, maxRows - 1, lastCol);
    rules.push(cfRowEquals_(full, map['parse-status'], 'ERROR', '#FEE2E2'));
  }

  // B) Missing shipping address (cell highlight)
  const addrFields = ['ship-address-1', 'ship-city', 'ship-state', 'ship-postal-code'];
  const addrRanges = [];
  addrFields.forEach(f => {
    if (map[f]) addrRanges.push(sheet.getRange(2, map[f], maxRows - 1, 1));
  });
  if (addrRanges.length) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenCellEmpty()
      .setBackground('#FCE7F3')
      .setRanges(addrRanges).build());
  }

  // C) Missing customer-id
  if (map['customer-id']) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenCellEmpty()
      .setBackground('#FCA5A5')
      .setRanges([sheet.getRange(2, map['customer-id'], maxRows - 1, 1)]).build());
  }

  // D) Refund present (row highlight)
  if (map['refund-amount']) {
    const full = sheet.getRange(2, 1, maxRows - 1, lastCol);
    const c = OMS_Utils.columnLetter_(map['refund-amount']);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=N($${c}2)>0`)
      .setBackground('#FEF3C7')
      .setRanges([full]).build());
  }

  // E) Stand order alert
  if (map['mag-safe-stand']) {
    const c = OMS_Utils.columnLetter_(map['mag-safe-stand']);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${c}2="Yes"`)
      .setBackground('#FED7AA')
      .setRanges([sheet.getRange(2, map['mag-safe-stand'], maxRows - 1, 1)]).build());
  }

  // F) Duplicate system-gmail-id
  if (map['system-gmail-id']) {
    const c = OMS_Utils.columnLetter_(map['system-gmail-id']);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=COUNTIF($${c}:$${c}, $${c}2)>1`)
      .setBackground('#FED7AA')
      .setRanges([sheet.getRange(2, map['system-gmail-id'], maxRows - 1, 1)]).build());
  }

  // Existing: customer classification colors
  addChipRules_(rules, sheet, map, 'customer-classification', {
    'Active':   { bg:'#DBEAFE', fg:'#1D4ED8' },
    'Replaced': { bg:'#0F766E', fg:'#FFFFFF' },
    'Refunded': { bg:'#B91C1C', fg:'#FFFFFF' },
  });

  // Existing: grip-size / length / hand / model
  addChipRules_(rules, sheet, map, 'grip-size', {
    'Standard': { bg:'#D1E7F0', fg:'#0F4C5C' },
    'Mid':      { bg:'#FDE68A', fg:'#92400E' },
  });
  addChipRules_(rules, sheet, map, 'shaft-length-option', {
    'Standard': { bg:'#D1E7F0', fg:'#0F4C5C' },
    'Longer':   { bg:'#D9F99D', fg:'#166534' },
  });
  addChipRules_(rules, sheet, map, 'hand', {
    'Right': { bg:'#FDE68A', fg:'#92400E' },
    'Left':  { bg:'#D1E7F0', fg:'#0F4C5C' },
  });
  addChipRules_(rules, sheet, map, 'model', {
    'Pro':   { bg:'#E9D5FF', fg:'#6B21A8' },
    'Basic': { bg:'#E5E7EB', fg:'#111827' },
  });

  // Existing: missing serial allocation cue (ACTIVE + empty)
  if (map['serial-number-allocated'] && map['item-life-cycle']) {
    const snCol = OMS_Utils.columnLetter_(map['serial-number-allocated']);
    const lifeCol = OMS_Utils.columnLetter_(map['item-life-cycle']);
    const r = sheet.getRange(2, map['serial-number-allocated'], maxRows - 1, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($${lifeCol}2="ACTIVE",LEN($${snCol}2)=0)`)
      .setBackground('#FCE7F3')
      .setRanges([r]).build());
  }

  sheet.setConditionalFormatRules(rules);
}

function styleOutbound_(sheet) {
  sheet.setHiddenGridlines(false);
  sheet.getDataRange().setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  sheet.setFrozenColumns(6);

  const map = OMS_Utils.getHeadersMap_(sheet);

  setColWidth_(sheet, map, 'oms-order-id', 220);
  setColWidth_(sheet, map, 'oms-order-item-id', 260);
  setColWidth_(sheet, map, 'shipment-id', 320);
  setColWidth_(sheet, map, 'order-created-at', 180);
  setColWidth_(sheet, map, 'domestic-tracking-kr', 180);
  setColWidth_(sheet, map, 'international-tracking-us', 200);
  setColWidth_(sheet, map, 'notes', 260);
  setColWidth_(sheet, map, 'stage-timeline', 450);
  setColWidth_(sheet, map, 'package-type', 140);
  setColWidth_(sheet, map, 'actual-weight-kg', 120);
  setColWidth_(sheet, map, 'package-length-cm', 140);
  setColWidth_(sheet, map, 'package-width-cm', 140);
  setColWidth_(sheet, map, 'package-height-cm', 140);

  setNumberFormat_(sheet, map, 'order-created-at', 'yyyy-mm-dd hh:mm:ss');
  setNumberFormat_(sheet, map, 'hub-received-date', 'yyyy-mm-dd');
  setNumberFormat_(sheet, map, 'us-ship-date', 'yyyy-mm-dd');
  setNumberFormat_(sheet, map, 'delivered-date', 'yyyy-mm-dd');
  setNumberFormat_(sheet, map, 'last-email-at', 'yyyy-mm-dd hh:mm:ss');
  setNumberFormat_(sheet, map, 'system-updated-at', 'yyyy-mm-dd hh:mm:ss');

  setNumberFormat_(sheet, map, 'actual-weight-kg', '0.00');
  setNumberFormat_(sheet, map, 'package-length-cm', '0.00');
  setNumberFormat_(sheet, map, 'package-width-cm', '0.00');
  setNumberFormat_(sheet, map, 'package-height-cm', '0.00');

  applyDropdown_(sheet, map, 'carrier-us', ['FEDEX','UPS','USPS','DHL','OTHER','']);
  applyDropdown_(sheet, map, 'hub-location', ['Seoul','Busan','Los Angeles','']);
  applyDropdown_(sheet, map, 'package-type', ['standard-club','club-with-stand','']);
  applyDropdown_(sheet, map, 'outbound-status', ['CREATED','KR_SHIPPED','HUB_RECEIVED','US_SHIPPED','DELIVERED','EXCEPTION','HOLD','CANCELLED','']);
  applyDropdown_(sheet, map, 'sn-verify', ['OK','MISMATCH','ERROR: No allocated S/N','']);
  applyDropdown_(sheet, map, 'customer-email-status', ['Sent: Final Delivery','Error','SKIP','']);

  // UI: Section Shading
  const sections = [
    { color: '#EEF2FF', headers: ['merchant-order-id','merchant-order-item-id','sku','customer-id','oms-*','shipment-id'] },
    { color: '#ECFEFF', headers: ['outbound-workflow-type','original-merchant-order-id','original-merchant-order-item-id'] },
    { color: '#ECFDF5', headers: ['order-created-at','hub-received-date','us-ship-date','delivered-date','hub-location','delivery-country'] },
    { color: '#FFFBEB', headers: ['domestic-tracking-kr','international-tracking-us','carrier-us'] },
    { color: '#FFF1F2', headers: ['outbound-status','sn-verify','customer-email-status','stage-timeline','notes'] },
    { color: '#F5F3FF', headers: ['package-type','actual-weight-kg','package-length-cm','package-width-cm','package-height-cm'] }
  ];
  OMS_Utils.applySectionShading_(sheet, map, sections);

  // Restore dark header readability
  const lc = sheet.getLastColumn();
  if (lc > 0) {
    sheet.getRange(1, 1, 1, lc)
      .setFontWeight('bold')
      .setBackground('#111827')
      .setFontColor('#FFFFFF');
  }

  // UI: Banding
  OMS_Utils.applyBanding_(sheet);

  sheet.setConditionalFormatRules([]);
  const rules = [];
  const maxRows = sheet.getMaxRows();
  const lastCol = sheet.getLastColumn();

  // A) Hub backlog (row orange)
  if (map['domestic-tracking-kr'] && map['international-tracking-us']) {
    const full = sheet.getRange(2, 1, maxRows - 1, lastCol);
    const krCol = OMS_Utils.columnLetter_(map['domestic-tracking-kr']);
    const usCol = OMS_Utils.columnLetter_(map['international-tracking-us']);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(LEN($${krCol}2)>0, LEN($${usCol}2)=0)`)
      .setBackground('#FFEDD5')
      .setRanges([full]).build());
  }

  // B) International tracking set but no us-ship-date
  if (map['international-tracking-us'] && map['us-ship-date']) {
    const usTrk = OMS_Utils.columnLetter_(map['international-tracking-us']);
    const usDate = OMS_Utils.columnLetter_(map['us-ship-date']);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(LEN($${usTrk}2)>0, LEN($${usDate}2)=0)`)
      .setBackground('#FED7AA')
      .setRanges([sheet.getRange(2, map['us-ship-date'], maxRows - 1, 1)]).build());
  }

  // C) Delivered but missing date
  if (map['outbound-status'] && map['delivered-date']) {
    const full = sheet.getRange(2, 1, maxRows - 1, lastCol);
    const statusCol = OMS_Utils.columnLetter_(map['outbound-status']);
    const dateCol = OMS_Utils.columnLetter_(map['delivered-date']);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($${statusCol}2="DELIVERED", LEN($${dateCol}2)=0)`)
      .setBackground('#FEE2E2')
      .setRanges([full]).build());
  }

  // D) S/N mismatch (critical)
  if (map['sn-verify']) {
    const full = sheet.getRange(2, 1, maxRows - 1, lastCol);
    const c = OMS_Utils.columnLetter_(map['sn-verify']);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=OR($${c}2="MISMATCH", LEFT($${c}2, 5)="ERROR")`)
      .setBackground('#FEE2E2')
      .setRanges([full]).build());
  }

  // E) Email gate
  if (map['customer-email-status']) {
    const c = OMS_Utils.columnLetter_(map['customer-email-status']);
    const range = sheet.getRange(2, map['customer-email-status'], maxRows - 1, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=SEARCH("Sent", $${c}2)`)
      .setBackground('#DCFCE7')
      .setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(LEN($${c}2)=0, LEN(INDIRECT("R"&ROW()&"C"&${map['international-tracking-us']}, FALSE))>0)`)
      .setBackground('#FEF3C7')
      .setRanges([range]).build());
  }

  // F) Stage timeline missing
  if (map['outbound-status'] && map['stage-timeline']) {
    const sCol = OMS_Utils.columnLetter_(map['outbound-status']);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(LEN($${sCol}2)>0, LEN(INDIRECT("R"&ROW()&"C"&${map['stage-timeline']}, FALSE))=0)`)
      .setBackground('#FCE7F3')
      .setRanges([sheet.getRange(2, map['stage-timeline'], maxRows - 1, 1)]).build());
  }

  // Existing: status chips
  addChipRules_(rules, sheet, map, 'outbound-status', {
    'CREATED':      { bg:'#E5E7EB', fg:'#111827' },
    'KR_SHIPPED':   { bg:'#DBEAFE', fg:'#1D4ED8' },
    'HUB_RECEIVED': { bg:'#FEF3C7', fg:'#92400E' },
    'US_SHIPPED':   { bg:'#E9D5FF', fg:'#6B21A8' },
    'DELIVERED':    { bg:'#DCFCE7', fg:'#166534' },
    'EXCEPTION':    { bg:'#FCA5A5', fg:'#7F1D1D' },
    'HOLD':         { bg:'#FCA5A5', fg:'#7F1D1D' },
    'CANCELLED':    { bg:'#9CA3AF', fg:'#111827' },
  });

  sheet.setConditionalFormatRules(rules);
}

/** ---------------- DASHBOARD (MERGE SAFE) ---------------- **/

function buildDashboard_(sheet) {
  const ss = OMS_Utils.ss();
  // ✅ FIX: remove any existing filter before merges
  const f = sheet.getFilter();
  if (f) f.remove();

  // Preserve existing dates if they exist
  const existingStart = sheet.getRange('B3').getValue();
  const existingEnd = sheet.getRange('D3').getValue();

  sheet.clear();
  sheet.setHiddenGridlines(true);

  // Title
  sheet.getRange(1, 1).setValue('G·GRIP OMS Dashboard');
  sheet.getRange(1, 1, 1, 6).merge()
    .setBackground('#111827')
    .setFontColor('#FFFFFF')
    .setFontSize(16)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);

  sheet.getRange(2, 1).setValue('Manual dashboard. No triggers are created automatically.');
  sheet.getRange(2, 1, 1, 6).merge().setFontColor('#374151');
  sheet.setRowHeight(2, 22);

  // Date Range Controls
  sheet.getRange('A3').setValue('Start Date:').setFontWeight('bold');
  sheet.getRange('C3').setValue('End Date:').setFontWeight('bold');

  const startCell = sheet.getRange('B3');
  const endCell = sheet.getRange('D3');

  startCell.setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build()).setNumberFormat('yyyy-mm-dd');
  endCell.setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build()).setNumberFormat('yyyy-mm-dd');

  if (existingStart instanceof Date) startCell.setValue(existingStart);
  else {
    const d = new Date(); d.setDate(1); startCell.setValue(d);
  }
  if (existingEnd instanceof Date) endCell.setValue(existingEnd);
  else endCell.setValue(new Date());

  // Simple metric table layout
  sheet.getRange(5, 1).setValue('Metric').setFontWeight('bold');
  sheet.getRange(5, 2).setValue('Value').setFontWeight('bold');
  sheet.getRange(5, 3).setValue('Notes').setFontWeight('bold');
  sheet.getRange(5, 1, 1, 3).setBackground('#EFEFEF');

  const MST_NAME = OMS_CONFIG.TABS.MASTER_TABLE;
  const mstSheet = ss.getSheetByName(MST_NAME);
  if (!mstSheet) throw new Error('Master Table not found.');
  
  const mstMap = OMS_Utils.getHeadersMap_(mstSheet);
  const mstCol = (h) => {
    const c = mstMap[h.toLowerCase()];
    if (!c) return `'${MST_NAME}'!$A2:$A`; // Fallback
    const letter = OMS_Utils.columnLetter_(c);
    return `'${MST_NAME}'!$${letter}2:$${letter}`;
  };

  let r = 6;

  // Period Metrics (from Master Table for consistency)
  const mstPD = mstCol('purchase-date');
  const dateFilter = `${mstPD},">="&$B$3,${mstPD},"<="&$D$3`;

  putMetric_(sheet, r++, 'Inbound Items (In Range)', `=IFERROR(COUNTIFS(${dateFilter}), 0)`, '');
  putMetric_(sheet, r++, 'Outbound US Shipped (In Range)', `=IFERROR(COUNTIFS(${mstCol('us-ship-date')},">="&$B$3,${mstCol('us-ship-date')},"<="&$D$3), 0)`, '');
  putMetric_(sheet, r++, 'Delivered (In Range)', `=IFERROR(COUNTIFS(${mstCol('delivered-date')},">="&$B$3,${mstCol('delivered-date')},"<="&$D$3), 0)`, '');

  r++;

  // Logistics Velocity (Calculated using Master Table joins)
  const mstHD = mstCol('hub-received-date');
  const mstUSD = mstCol('us-ship-date');
  const mstDD = mstCol('delivered-date');
  const velFilter = `(${mstPD}<>"")*(${mstPD}>=$B$3)*(${mstPD}<=$D$3)`;

  putMetric_(sheet, r++, 'Avg Time to Hub (days)', `=IFERROR(AVERAGE(FILTER(${mstHD}-${mstPD}, (${mstHD}<>"")*${velFilter})),"")`, 'purchase → hub-received');
  putMetric_(sheet, r++, 'Avg Customs Clearance (days)', `=IFERROR(AVERAGE(FILTER(${mstUSD}-${mstHD}, (${mstUSD}<>"")*(${mstHD}<>"")*${velFilter})),"")`, 'hub → us-ship');
  putMetric_(sheet, r++, 'Avg Last Mile (days)', `=IFERROR(AVERAGE(FILTER(${mstDD}-${mstUSD}, (${mstDD}<>"")*(${mstUSD}<>"")*${velFilter})),"")`, 'us-ship → delivered');
  putMetric_(sheet, r++, 'Avg Click-to-Door (days)', `=IFERROR(AVERAGE(FILTER(${mstDD}-${mstPD}, (${mstDD}<>"")*(${mstPD}<>"")*${velFilter})),"")`, 'purchase → delivered');

  r++;

  // Backlog Monitor
  putMetric_(sheet, r++, 'Hub Backlog', `=IFERROR(COUNTIFS(${mstCol('domestic-tracking-kr')},"<>",${mstCol('international-tracking-us')},""), 0)`, 'KR tracking present, US tracking empty');
  putMetric_(sheet, r++, 'S/N Mismatch Count', `=IFERROR(COUNTIF(${mstCol('sn-verify')},"MISMATCH"), 0)`, '');

  r++;

  // Financials & Health
  const reshipPattern = `"*${OMS_CONFIG.RESHIP_SUFFIX}*"`;
  putMetric_(sheet, r++, 'Reshipment Rate (%)', `=IFERROR(COUNTIFS(${mstCol('merchant-order-id')}, ${reshipPattern}, ${mstPD}, ">="&$B$3, ${mstPD}, "<="&$D$3)/MAX(1,COUNTIFS(${dateFilter})),"")`, '');
  putMetric_(sheet, r++, 'Total Lost Revenue (In Range)', `=IFERROR(SUMIFS(${mstCol('refund-amount')}, ${mstPD}, ">="&$B$3, ${mstPD}, "<="&$D$3),0)`, 'Sum refund-amount');
  putMetric_(sheet, r++, 'Total LTV (In Range)', `=IFERROR(SUMIFS(${mstCol('total-amount')}, ${mstPD}, ">="&$B$3, ${mstPD}, "<="&$D$3),0)`, 'Sum of total-amount');

  r++;

  // Customer Loyalty
  const mstEmail = mstCol('buyer-email');
  putMetric_(sheet, r++, 'Repeat Purchase Rate (%)', `=IFERROR((COUNTIFS(${dateFilter})-COUNTUNIQUE(FILTER(${mstEmail}, ${velFilter})))/MAX(1, COUNTIFS(${dateFilter})),"")`, 'Total items - unique emails / total items');

  r++;

  // Charts Integration
  const dataSheet = getOrCreateSheet_(ss, OMS_CONFIG.TABS.DASHBOARD_DATA);
  dataSheet.hideSheet();
  buildDashboardData_(dataSheet, mstCol);

  addDashboardCharts_(sheet, dataSheet);

  // Cosmetics
  sheet.setColumnWidth(1, 320);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 420);
  sheet.getRange(5, 1, 60, 3).setFontFamily('Arial').setFontSize(10);
}

function putMetric_(sheet, row, label, formula, notes) {
  sheet.getRange(row, 1).setValue(label).setFontWeight('bold');
  sheet.getRange(row, 2).setFormula(formula);
  sheet.getRange(row, 3).setValue(notes || '');
}

/**
 * Builds source data for charts in a hidden sheet.
 */
function buildDashboardData_(sheet, mstCol) {
  sheet.clear();
  const MST_DASH = OMS_CONFIG.TABS.DASHBOARD;

  // 1. Inbound Items by Day (Dynamic Range)
  sheet.getRange('A1').setValue('Inbound Trend (In Range)');
  sheet.getRange('A2').setValue('Date');
  sheet.getRange('B2').setValue('Items');
  
  // Dynamic sequence of dates between Start (B3) and End (D3) on Dashboard
  const seqFormula = `=IFERROR(SEQUENCE('${MST_DASH}'!$D$3 - '${MST_DASH}'!$B$3 + 1, 1, '${MST_DASH}'!$B$3), TODAY())`;
  sheet.getRange('A3').setFormula(seqFormula);
  
  // Adjusted to use indirect to prevent range shifting
  const countFormula = `=IFERROR(ARRAYFORMULA(COUNTIFS(${mstCol('purchase-date')}, A3:INDEX(A:A, MATCH(9^9, A:A)))), 0)`;
  sheet.getRange('B3').setFormula(countFormula);

  // 2. Outbound Status Distribution (Filtered by Range)
  sheet.getRange('D1').setValue('Outbound Status (In Range)');
  const statusCol = mstCol('outbound-status');
  const dateCol = mstCol('purchase-date');
  
  // Wrapped in IFERROR and ensured query is robust. Used data range starting from row 2.
  const query = `=IFERROR(QUERY({${statusCol}, ${dateCol}}, "SELECT Col1, COUNT(Col1) WHERE Col1 IS NOT NULL AND Col2 >= DATE '"&TEXT('${MST_DASH}'!$B$3,"yyyy-mm-dd")&"' AND Col2 <= DATE '"&TEXT('${MST_DASH}'!$D$3,"yyyy-mm-dd")&"' GROUP BY Col1 LABEL COUNT(Col1) 'Count'", 0), "No Data")`;
  sheet.getRange('D2').setFormula(query);
}

/**
 * Adds charts to the Dashboard sheet.
 */
function addDashboardCharts_(sheet, dataSheet) {
  // Remove existing charts
  const charts = sheet.getCharts();
  charts.forEach(c => sheet.removeChart(c));

  // 1. Inbound Line Chart
  const inboundChart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dataSheet.getRange('A2:B102')) // support up to 100 days
    .setPosition(24, 1, 0, 0)
    .setOption('title', 'Inbound Velocity')
    .setOption('legend', { position: 'none' })
    .setOption('vAxis', { title: 'Items' })
    .setOption('hAxis', { title: 'Date' })
    .setOption('width', 600)
    .setOption('height', 300)
    .setOption('backgroundColor.stroke', 'none')
    .setOption('backgroundColor.strokeWidth', 0)
    .build();

  // 2. Status Pie Chart
  // Repositioned to start halfway through column C (Column 3, offset 210)
  const statusChart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(dataSheet.getRange('D2:E10'))
    .setPosition(24, 3, 210, 0)
    .setOption('title', 'Outbound Pipeline Status')
    .setOption('width', 400)
    .setOption('height', 300)
    .setOption('backgroundColor.stroke', 'none')
    .setOption('backgroundColor.strokeWidth', 0)
    .build();

  sheet.insertChart(inboundChart);
  sheet.insertChart(statusChart);
}

/** ---------------- Formatting Helpers ---------------- **/

function setColWidth_(sheet, map, header, width) {
  const col = map[String(header).toLowerCase()];
  if (col) sheet.setColumnWidth(col, width);
}

function setNumberFormat_(sheet, map, header, fmt) {
  const col = map[String(header).toLowerCase()];
  if (!col) return;
  sheet.getRange(2, col, sheet.getMaxRows()-1, 1).setNumberFormat(fmt);
}

function applyDropdown_(sheet, map, header, values) {
  const col = map[String(header).toLowerCase()];
  if (!col) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(true)
    .build();

  sheet.getRange(2, col, sheet.getMaxRows()-1, 1).setDataValidation(rule);
}

function cfRowEquals_(fullRowRange, colIndex, value, bg) {
  const letter = OMS_Utils.columnLetter_(colIndex);
  return SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${letter}2="${value}"`)
    .setBackground(bg)
    .setRanges([fullRowRange])
    .build();
}

function addChipRules_(rules, sheet, map, header, valueToStyle) {
  const col = map[String(header).toLowerCase()];
  if (!col) return;

  const letter = OMS_Utils.columnLetter_(col);
  const range = sheet.getRange(2, col, sheet.getMaxRows()-1, 1);

  Object.keys(valueToStyle).forEach(val => {
    const s = valueToStyle[val];
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=$${letter}2="${val}"`)
        .setBackground(s.bg)
        .setFontColor(s.fg)
        .setRanges([range])
        .build()
    );
  });
}
