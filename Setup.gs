/********************************
 * Setup.gs (Standalone, No Triggers)
 * - NO ScriptApp.newTrigger
 * - NO onOpen
 * - Dashboard merges are filter-safe
 ********************************/

function omsSetupSheet() {
  const ss = OMS_Utils.ss();

  // Validate on startup
  validateSchema(ss);

  const inbound = getOrCreateSheet_(ss, OMS_CONFIG.TABS.INBOUND);
  const outbound = getOrCreateSheet_(ss, OMS_CONFIG.TABS.OUTBOUND);
  const master = getOrCreateSheet_(ss, OMS_CONFIG.TABS.MASTER);
  const meta = getOrCreateSheet_(ss, OMS_CONFIG.TABS.META);

  // IMPORTANT:
  // Inbound/Outbound get filters (table sheets)
  // Master dashboard gets NO filter (merge-safe)
  applyHeaderRow_(inbound, OMS_SCHEMA_INBOUND_(), { createFilter: true, clearAll: true });
  applyHeaderRow_(outbound, OMS_SCHEMA_OUTBOUND_(), { createFilter: true, clearAll: true });
  applyHeaderRow_(master, OMS_SCHEMA_MASTER_(), { createFilter: false, clearAll: true }); // ✅ FIX

  updateMetaSheet_(ss, meta, {
    [OMS_CONFIG.TABS.INBOUND]: OMS_SCHEMA_INBOUND_(),
    [OMS_CONFIG.TABS.OUTBOUND]: OMS_SCHEMA_OUTBOUND_(),
  });
  meta.hideSheet();

  styleInbound_(inbound);
  styleOutbound_(outbound);

  buildDashboard_(master); // does merges safely because master has no filter

  inbound.setTabColor('#1E3A8A');
  outbound.setTabColor('#047857');
  master.setTabColor('#111827');

  SpreadsheetApp.flush();
}

function omsRefreshDashboard() {
  const ss = OMS_Utils.ss();
  const master = ss.getSheetByName(OMS_CONFIG.TABS.MASTER);
  if (!master) throw new Error('Master_OMS_View not found. Run omsSetupSheet() first.');
  buildDashboard_(master);
  SpreadsheetApp.flush();
}

/** ---------------- SCHEMAS ---------------- **/

function OMS_SCHEMA_INBOUND_() {
  return [
    'merchant-order-id','merchant-order-item-id','purchase-date','purchase-time','buyer-email','buyer-name',
    'buyer-phone-number','customer-id','system-gmail-id','sales-channel','customer-classification','is-business-order',
    'sku','product-name','mag-safe-stand','model','club-type','hand','flex','length','grip-size','head-material',
    'shaft-material','loft','lie-angle','offset','quantity-purchased','currency','item-price','item-tax','shipping-price',
    'total-amount','coupon-code','refund-amount','refund-date','return-reason-code','recipient-name','ship-address-1',
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
    'domestic-tracking-kr','hub-received-date','international-tracking-us','carrier-us',
    'us-ship-date','delivered-date',
    'outbound-status','serial-number-scanned','sn-verify','customer-email-status','last-email-at','system-updated-at','notes',
    'oms-order-id','oms-order-item-id','shipment-id',
    'actual-weight-kg','package-length-cm','package-width-cm','package-height-cm'
  ];
}

function OMS_SCHEMA_MASTER_() {
  // Dashboard sheet doesn’t need a big table header; it’s a layout sheet.
  return ['metric','value','notes'];
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
    const required = ['merchant-order-id','buyer-email','sku','oms-order-item-id'];
    try {
      OMS_Utils.requireCols_(inbound, required);
    } catch (e) {
      throw new Error(`CRITICAL: Inbound sheet schema validation failed. ${e.message}`);
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
 */
function applyHeaderRow_(sheet, headers, options) {
  const opts = options || {};
  const createFilter = (opts.createFilter !== false);
  const clearAll = (opts.clearAll !== false);

  // Remove filter FIRST (prevents merge conflicts)
  const existing = sheet.getFilter();
  if (existing) existing.remove();

  if (clearAll) sheet.clear();

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#111827')
    .setFontColor('#FFFFFF')
    .setWrap(true)
    .setVerticalAlignment('middle');

  sheet.setRowHeight(1, 36);

  if (createFilter) {
    sheet.getRange(1, 1, 1, headers.length).createFilter();
  }
}

/** ---------------- DROPDOWNS + COLORS ---------------- **/

function styleInbound_(sheet) {
  sheet.setHiddenGridlines(false);
  sheet.getDataRange().setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');

  const map = OMS_Utils.getHeadersMap_(sheet);

  // widths
  setColWidth_(sheet, map, 'oms-order-id', 220);
  setColWidth_(sheet, map, 'oms-order-item-id', 260);
  setColWidth_(sheet, map, 'buyer-email', 220);
  setColWidth_(sheet, map, 'buyer-email-hash', 240);
  setColWidth_(sheet, map, 'ship-address-1', 280);
  setColWidth_(sheet, map, 'sku', 160);
  setColWidth_(sheet, map, 'product-name', 220);
  setColWidth_(sheet, map, 'notes', 260);
  setColWidth_(sheet, map, 'automation-notes', 260);

  // formats
  setNumberFormat_(sheet, map, 'purchase-date', 'yyyy-mm-dd');
  setNumberFormat_(sheet, map, 'refund-date', 'yyyy-mm-dd');
  setNumberFormat_(sheet, map, 'item-price', '$0.00');
  setNumberFormat_(sheet, map, 'item-tax', '$0.00');
  setNumberFormat_(sheet, map, 'shipping-price', '$0.00');
  setNumberFormat_(sheet, map, 'total-amount', '$0.00');
  setNumberFormat_(sheet, map, 'refund-amount', '$0.00');
  setNumberFormat_(sheet, map, 'quantity-purchased', '0');
  setNumberFormat_(sheet, map, 'system-created-at', 'yyyy-mm-dd hh:mm:ss');
  setNumberFormat_(sheet, map, 'system-updated-at', 'yyyy-mm-dd hh:mm:ss');

  // dropdowns (based on your prior sheet + Pro/Basic model)
  applyDropdown_(sheet, map, 'customer-classification', ['Active','Replaced','Refunded','']);
  applyDropdown_(sheet, map, 'model', ['Pro','Basic','']);
  applyDropdown_(sheet, map, 'grip-size', ['Standard','Mid','']);
  applyDropdown_(sheet, map, 'length', ['Standard','Longer','']);
  applyDropdown_(sheet, map, 'hand', ['Right','Left','']);
  applyDropdown_(sheet, map, 'flex', ['L','R','S','X','']);
  applyDropdown_(sheet, map, 'mag-safe-stand', ['Yes','0','']);
  applyDropdown_(sheet, map, 'item-life-cycle', ['ACTIVE','REFUNDED','RETURNED','REPLACED','CANCELLED','']);
  applyDropdown_(sheet, map, 'order-life-cycle', ['ACTIVE','PARTIAL_REFUND','FULL_REFUND','CANCELLED','']);
  applyDropdown_(sheet, map, 'parse-status', ['OK','ERROR','RES_AUTO','MANUAL_EDIT','']);
  applyDropdown_(sheet, map, 'source-system', ['amazon_fba','shopify','samcart','imweb','manual','unknown','']);

  // conditional formatting to mimic “chips”
  sheet.setConditionalFormatRules([]);
  const rules = [];

  // parse error row tint
  if (map['parse-status']) {
    const full = sheet.getRange(2, 1, sheet.getMaxRows()-1, sheet.getLastColumn());
    rules.push(cfRowEquals_(full, map['parse-status'], 'ERROR', '#FDE2E2'));
  }

  // refund tint
  if (map['refund-amount']) {
    const full = sheet.getRange(2, 1, sheet.getMaxRows()-1, sheet.getLastColumn());
    const c = OMS_Utils.columnLetter_(map['refund-amount']);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=N($${c}2)>0`)
      .setBackground('#FEF3C7')
      .setRanges([full]).build());
  }

  // customer classification colors (like your screenshot)
  addChipRules_(rules, sheet, map, 'customer-classification', {
    'Active':   { bg:'#DBEAFE', fg:'#1D4ED8' },
    'Replaced': { bg:'#0F766E', fg:'#FFFFFF' },
    'Refunded': { bg:'#B91C1C', fg:'#FFFFFF' },
  });

  // grip-size / length / hand (your screenshot vibe)
  addChipRules_(rules, sheet, map, 'grip-size', {
    'Standard': { bg:'#D1E7F0', fg:'#0F4C5C' },
    'Mid':      { bg:'#FDE68A', fg:'#92400E' },
  });

  addChipRules_(rules, sheet, map, 'length', {
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

  // missing serial allocation cue (ACTIVE + empty)
  if (map['serial-number-allocated'] && map['item-life-cycle']) {
    const snCol = OMS_Utils.columnLetter_(map['serial-number-allocated']);
    const lifeCol = OMS_Utils.columnLetter_(map['item-life-cycle']);
    const r = sheet.getRange(2, map['serial-number-allocated'], sheet.getMaxRows()-1, 1);
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

  const map = OMS_Utils.getHeadersMap_(sheet);

  setColWidth_(sheet, map, 'oms-order-id', 220);
  setColWidth_(sheet, map, 'oms-order-item-id', 260);
  setColWidth_(sheet, map, 'domestic-tracking-kr', 180);
  setColWidth_(sheet, map, 'international-tracking-us', 200);
  setColWidth_(sheet, map, 'notes', 260);
  setColWidth_(sheet, map, 'actual-weight-kg', 120);
  setColWidth_(sheet, map, 'package-length-cm', 140);
  setColWidth_(sheet, map, 'package-width-cm', 140);
  setColWidth_(sheet, map, 'package-height-cm', 140);

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
  applyDropdown_(sheet, map, 'outbound-status', ['CREATED','KR_SHIPPED','HUB_RECEIVED','US_SHIPPED','DELIVERED','EXCEPTION','HOLD','CANCELLED','']);
  applyDropdown_(sheet, map, 'sn-verify', ['OK','MISMATCH','ERROR: No allocated S/N','']);
  applyDropdown_(sheet, map, 'customer-email-status', ['Sent: Final Delivery','Error','SKIP','']);

  sheet.setConditionalFormatRules([]);
  const rules = [];

  // status chips
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

  // SN mismatch row tint
  if (map['sn-verify']) {
    const full = sheet.getRange(2, 1, sheet.getMaxRows()-1, sheet.getLastColumn());
    const c = OMS_Utils.columnLetter_(map['sn-verify']);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=OR($${c}2="MISMATCH",LEFT($${c}2,5)="ERROR")`)
      .setBackground('#FDE2E2')
      .setRanges([full]).build());
  }

  sheet.setConditionalFormatRules(rules);
}

/** ---------------- DASHBOARD (MERGE SAFE) ---------------- **/

function buildDashboard_(sheet) {
  // ✅ FIX: remove any existing filter before merges (even though we don’t create one)
  const f = sheet.getFilter();
  if (f) f.remove();

  sheet.clear();
  sheet.setHiddenGridlines(true);

  // Title row merge (safe now)
  sheet.getRange(1,1).setValue('G·GRIP OMS Dashboard');
  sheet.getRange(1,1,1,6).merge()
    .setBackground('#111827')
    .setFontColor('#FFFFFF')
    .setFontSize(16)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);

  sheet.getRange(2,1).setValue('Manual dashboard. No triggers are created automatically.');
  sheet.getRange(2,1,1,6).merge().setFontColor('#374151');
  sheet.setRowHeight(2, 22);

  // simple metric table layout (you asked to plan + build dashboard now)
  sheet.getRange(4,1).setValue('Metric').setFontWeight('bold');
  sheet.getRange(4,2).setValue('Value').setFontWeight('bold');
  sheet.getRange(4,3).setValue('Notes').setFontWeight('bold');
  sheet.getRange(4,1,1,3).setBackground('#EFEFEF');

  const IN = OMS_CONFIG.TABS.INBOUND;
  const OUT = OMS_CONFIG.TABS.OUTBOUND;
  const META = OMS_CONFIG.TABS.META;

  // Helper for meta-based range
  const metaRng = (s, h) => `INDIRECT("'"&${IN}!A1&"'!"&VLOOKUP("${s}|${h}",{${META}!A:A&"|"&${META}!B:B,${META}!D:D},2,FALSE)&":"&VLOOKUP("${s}|${h}",{${META}!A:A&"|"&${META}!B:B,${META}!D:D},2,FALSE))`.replace(`${IN}!A1`, `"${s}"`);

  // Actually, Apps Script string templates are easier. Let's use a function.
  // We use VLOOKUP against the _Meta sheet which contains sheet|header in Col A (if concatenated)
  // or we can use a simpler lookup since Col A=Sheet, Col B=Header.
  // Using INDEX(MATCH(..., INDEX(JOINED_RANGE, 0), 0)) to handle array context in standard formulas.
  const lookup_ = (s, h) => `INDEX('${META}'!$D$1:$D, MATCH("${s}|${h}", INDEX('${META}'!$A$1:$A & "|" & '${META}'!$B$1:$B, 0), 0))`;
  const col_ = (s, h) => `INDIRECT("'"&"${s}"&"'!" & ${lookup_(s,h)} & ":" & ${lookup_(s,h)})`;

  // Week / month anchors
  sheet.getRange('E4').setValue('WeekStart');
  sheet.getRange('F4').setFormula('=TODAY()-WEEKDAY(TODAY(),2)+1');
  sheet.getRange('E5').setValue('MonthStart');
  sheet.getRange('F5').setFormula('=EOMONTH(TODAY(),-1)+1');

  let r = 6;

  // Weekly & Monthly inbound/outbound
  putMetric_(sheet, r++, 'Inbound Items (This Week)', `=COUNTIFS(${col_(IN, 'purchase-date')},">="&$F$4,${col_(IN, 'purchase-date')},"<"&$F$4+7)`, '');
  putMetric_(sheet, r++, 'Inbound Items (This Month)', `=COUNTIFS(${col_(IN, 'purchase-date')},">="&$F$5,${col_(IN, 'purchase-date')},"<"&EOMONTH(TODAY(),0)+1)`, '');

  putMetric_(sheet, r++, 'Outbound US Shipped (This Week)', `=COUNTIFS(${col_(OUT, 'us-ship-date')},">="&$F$4,${col_(OUT, 'us-ship-date')},"<"&$F$4+7)`, '');

  putMetric_(sheet, r++, 'Delivered (This Week)', `=COUNTIFS(${col_(OUT, 'delivered-date')},">="&$F$4,${col_(OUT, 'delivered-date')},"<"&$F$4+7)`, '');

  r++;

  // Logistics Velocity
  const inPD = col_(IN, 'purchase-date');
  const outHD = col_(OUT, 'hub-received-date');
  const outUSD = col_(OUT, 'us-ship-date');
  const outDD = col_(OUT, 'delivered-date');

  putMetric_(sheet, r++, 'Avg Time to Hub (days)', `=IFERROR(AVERAGE(FILTER(${outHD}-${inPD}, (${outHD}<>"")*(${inPD}<>""))),"")`, 'purchase-date → hub-received-date');
  putMetric_(sheet, r++, 'Avg Customs Clearance (days)', `=IFERROR(AVERAGE(FILTER(${outUSD}-${outHD}, (${outUSD}<>"")*(${outHD}<>""))),"")`, 'hub → us-ship');
  putMetric_(sheet, r++, 'Avg Last Mile (days)', `=IFERROR(AVERAGE(FILTER(${outDD}-${outUSD}, (${outDD}<>"")*(${outUSD}<>""))),"")`, 'us-ship → delivered');
  putMetric_(sheet, r++, 'Avg Click-to-Door (days)', `=IFERROR(AVERAGE(FILTER(${outDD}-${inPD}, (${outDD}<>"")*(${inPD}<>""))),"")`, 'purchase → delivered');

  r++;

  // Backlog Monitor
  putMetric_(sheet, r++, 'Hub Backlog', `=COUNTIFS(${col_(OUT, 'domestic-tracking-kr')},"<>",${col_(OUT, 'international-tracking-us')},"")`, 'KR tracking present, US tracking empty');

  putMetric_(sheet, r++, 'S/N Mismatch Count', `=COUNTIF(${col_(OUT, 'sn-verify')},"MISMATCH")`, '');

  r++;

  // -RES / Refunds
  const inOID = col_(IN, 'merchant-order-id');
  const reshipPattern = `"*${OMS_CONFIG.RESHIP_SUFFIX}*"`;
  putMetric_(sheet, r++, 'Reshipment Rate (%)', `=IFERROR(COUNTIF(${inOID}, ${reshipPattern})/MAX(1,COUNTA(${inOID})-1),"")`, '');

  putMetric_(sheet, r++, 'Total Lost Revenue', `=IFERROR(SUM(${col_(IN, 'refund-amount')}),0)`, 'Sum refund-amount');

  r++;

  // More detailed metrics
  const inEmail = col_(IN, 'buyer-email');
  putMetric_(sheet, r++, 'Repeat Purchase Rate (%)', `=IFERROR((COUNTA(${inEmail})-1-COUNTUNIQUE(${inEmail}))/(COUNTA(${inEmail})-1),"")`, 'Total items - unique emails / total items');

  putMetric_(sheet, r++, 'Total LTV', `=SUM(${col_(IN, 'total-amount')})`, 'Sum of total-amount across all orders');

  const inRR = col_(IN, 'return-reason-code');
  putMetric_(sheet, r++, 'Top Return Reason', `=IFERROR(INDEX(${inRR}, MATCH(MAX(COUNTIF(${inRR}, ${inRR})), COUNTIF(${inRR}, ${inRR}), 0)), "")`, 'Most frequent return reason');

  // Cosmetics
  sheet.setColumnWidth(1, 320);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 420);
  sheet.getRange(4,1,60,3).setFontFamily('Arial').setFontSize(10);
}

function putMetric_(sheet, row, label, formula, notes) {
  sheet.getRange(row, 1).setValue(label).setFontWeight('bold');
  sheet.getRange(row, 2).setFormula(formula);
  sheet.getRange(row, 3).setValue(notes || '');
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
