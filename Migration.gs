/********************************
 * Migration.gs
 * Legacy Data Migration Logic
 ********************************/

/**
 * Strips $ and commas and converts to Number.
 */
function migrate_parseCurrency_(val) {
  if (val === null || val === undefined || val === '') return 0;
  const cleaned = String(val).replace(/[$,]/g, '');
  return Number(cleaned) || 0;
}

/**
 * Strips leading "=" and normalizes phone.
 */
function migrate_parsePhone_(val) {
  let s = String(val || '').trim();
  if (s.startsWith('=')) s = s.substring(1).replace(/["']/g, '');
  return OMS_Utils.normalizePhone(s);
}

/**
 * Parses MM/DD/YYYY into YYYY-MM-DD string.
 */
function migrate_parseDate_(val) {
  const s = String(val || '').trim();
  if (!s) return '';
  const d = new Date(s);
  if (isNaN(d.getTime())) return '';
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Extracts date from strings like "Wednesday, 2/25/26 at 11:55 AM"
 */
function migrate_parseCompletionDate_(val) {
  const s = String(val || '').trim();
  if (!s) return null;
  // Try to match M/D/YY or M/D/YYYY
  const match = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (!match) return null;
  let m = parseInt(match[1]) - 1;
  let d = parseInt(match[2]);
  let y = parseInt(match[3]);
  if (y < 100) y += 2000;
  const date = new Date(y, m, d);
  return isNaN(date.getTime()) ? null : date;
}

/**
 * migrateInboundLegacyToNew
 * Manually triggered.
 */
function migrateInboundLegacyToNew() {
  const ss = OMS_Utils.ss();
  const rawSheet = ss.getSheetByName('Inbound_Legacy_Raw');
  if (!rawSheet) {
    SpreadsheetApp.getUi().alert('Staging sheet "Inbound_Legacy_Raw" not found.');
    return;
  }

  const inbound = OMS_Utils.sheet_(OMS_CONFIG.TABS.INBOUND);
  const inCols = OMS_Utils.getHeadersMap_(inbound);
  const lc = inbound.getLastColumn();

  const rawData = rawSheet.getDataRange().getValues();
  if (rawData.length < 1) return; // No data

  // Existing order IDs to avoid duplicates
  const existingOrderIds = new Set();
  const inLr = inbound.getLastRow();
  if (inLr >= 2) {
    const ids = inbound.getRange(2, inCols['merchant-order-id'], inLr - 1, 1).getValues();
    ids.forEach(r => {
      const id = String(r[0] || '').trim();
      if (id) existingOrderIds.add(id);
    });
  }

  const now = new Date();
  const stamp = Utilities.formatDate(now, OMS_CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');
  const migratedRows = [];
  const stubData = [];

  // Legacy Column mapping (0-indexed)
  // 0 Order ID, 1 Customer ID, 2 S/N, 3 Notes, 4 First name, 5 Surname, 6 Email, 7 Customer Classification
  // 8 Order Platform, 9 Phone Number, 10 Shipping Address, 11 City, 12 State, 13 Zip, 14 Country
  // 15 Order Date (MM/DD/YYYY), 16 Stand (1 or 0), 17 Order Option / Flex, 18 Grip Size, 19 Length
  // 20 Hand, 21 Quantity, 22 Order Subtotal, 23 Order Shipping, 24 Order Tax, 25 Discount Amount
  // 26 Refund Amount, 27 Total Payment Amount, 28 Coupon Applied, 29 System_Gmail_ID, 30 Full Name

  rawData.forEach((row, idx) => {
    // Skip header if it looks like one (usually row[0] is Order ID)
    if (idx === 0 && (String(row[0]).toLowerCase().includes('order') || String(row[0]).toLowerCase().includes('id'))) return;

    const merchantOrderId = String(row[0] || '').trim();
    if (!merchantOrderId || existingOrderIds.has(merchantOrderId)) return;

    const sourceSystem = 'legacy';
    const sourceOrderId = merchantOrderId;
    const lineItemIndex = 1;
    const merchantOrderItemId = OMS_Utils.generateLineItemId_(lineItemIndex);
    const sourceOrderItemId = merchantOrderItemId;

    const omsOrderId = OMS_Utils.buildOmsOrderId_(sourceSystem, sourceOrderId);
    const omsOrderItemId = OMS_Utils.buildOmsOrderItemId_(omsOrderId, merchantOrderItemId);

    const purchaseDate = migrate_parseDate_(row[15]);
    const buyerEmail = OMS_Utils.normalizeEmail_(row[6]);
    const customerId = String(row[1] || '').trim() || OMS_Utils.lookupOrCreateCustomerId_(buyerEmail);
    const emailHash = OMS_Utils.emailHash_(buyerEmail);

    const model = String(row[17] || '').toUpperCase().includes('S') ? 'Pro' : 'Basic';
    const hand = String(row[20] || '').trim();
    const flex = String(row[17] || '').trim();
    const length = String(row[19] || '').trim();
    const gripSize = String(row[18] || '').trim();
    const magSafeStand = (String(row[16]) === '1') ? 'Yes' : '0';

    const sku = OMS_Utils.deriveSku({
      model,
      clubType: 'Iron',
      hand,
      flex,
      length,
      gripSize,
      magSafeStand
    });

    const v = {
      merchantOrderId,
      merchantOrderItemId,
      lineItemIndex,
      purchaseDate,
      purchaseTime: '',
      orderCreatedAt: purchaseDate ? `${purchaseDate}T00:00:00` : '',

      buyerEmail,
      buyerName: String(row[30] || '').trim() || (String(row[4] || '') + ' ' + String(row[5] || '')).trim(),
      buyerPhone: migrate_parsePhone_(row[9]),

      customerId,
      systemGmailId: String(row[29] || '').trim(),
      orderSourceEmail: buyerEmail,

      salesChannel: String(row[8] || '').trim(),
      customerClassification: String(row[7] || '').trim() || 'Active',
      isBusinessOrder: 'false',

      sku,
      productName: 'G-GRIP',
      magSafeStand,
      model,
      clubType: 'Iron',
      productCategory: 'Golf Club',
      hand,
      flex,
      shaftLengthOption: length,
      gripSize,

      qty: Number(row[21]) || 1,
      currency: 'USD',
      itemPrice: migrate_parseCurrency_(row[22]),
      shippingPrice: migrate_parseCurrency_(row[23]),
      itemTax: migrate_parseCurrency_(row[24]),
      discountAmount: migrate_parseCurrency_(row[25]),
      refundAmount: migrate_parseCurrency_(row[26]),
      totalAmount: migrate_parseCurrency_(row[27]),
      couponCode: String(row[28] || '').trim(),

      recipientName: String(row[30] || '').trim() || (String(row[4] || '') + ' ' + String(row[5] || '')).trim(),
      shipAddr1: String(row[10] || '').trim(),
      shipCity: String(row[11] || '').trim(),
      shipState: String(row[12] || '').trim(),
      shipPostal: String(row[13] || '').trim(),
      shipCountry: String(row[14] || '').trim() || 'United States',

      serialAllocated: String(row[2] || '').trim(),
      notes: String(row[3] || '').trim(),
      automationNotes: '',
      itemLifeCycle: 'ACTIVE',
      orderLifeCycle: 'ACTIVE',
      createdAt: stamp,
      updatedAt: stamp,
      parseStatus: 'LEGACY_IMPORT',

      sourceSystem,
      sourceOrderId,
      sourceOrderItemId,
      omsOrderId,
      omsOrderItemId,
      buyerEmailHash: emailHash
    };

    migratedRows.push(inbound_buildRowFromHeaders_(inCols, lc, v));

    stubData.push({
      merchantOrderId,
      merchantOrderItemId,
      sku,
      customerId,
      omsOrderId,
      omsOrderItemId,
      magSafeStand,
      orderCreatedAt: v.orderCreatedAt,
      deliveryCountry: v.shipCountry
    });
  });

  if (migratedRows.length > 0) {
    inbound.getRange(inbound.getLastRow() + 1, 1, migratedRows.length, lc).setValues(migratedRows);
    outbound_createStubs_(stubData);
    SpreadsheetApp.getUi().alert(`Successfully migrated ${migratedRows.length} inbound rows.`);
  } else {
    SpreadsheetApp.getUi().alert('No new inbound rows to migrate.');
  }
}

/**
 * migrateOutboundLegacyToNew
 * Manually triggered.
 */
function migrateOutboundLegacyToNew() {
  const ss = OMS_Utils.ss();
  const rawSheet = ss.getSheetByName('Outbound_Legacy_Raw');
  if (!rawSheet) {
    SpreadsheetApp.getUi().alert('Staging sheet "Outbound_Legacy_Raw" not found.');
    return;
  }

  const outbound = OMS_Utils.sheet_(OMS_CONFIG.TABS.OUTBOUND);
  const outCols = OMS_Utils.getHeadersMap_(outbound);
  const rawData = rawSheet.getDataRange().getValues();
  if (rawData.length < 1) return;

  const outLr = outbound.getLastRow();
  if (outLr < 2) {
    SpreadsheetApp.getUi().alert('No stubs found in Outbound_Logistics. Run Inbound migration first.');
    return;
  }

  // Load Outbound sheet into memory for fast matching
  const outData = outbound.getRange(1, 1, outLr, outbound.getLastColumn()).getValues();
  const midIdx = outCols['merchant-order-id'] - 1;
  const outMap = {};
  outData.forEach((r, i) => {
    const mid = String(r[midIdx] || '').trim();
    if (mid && i > 0) outMap[mid] = i + 1; // row number
  });

  const now = new Date();
  const stamp = Utilities.formatDate(now, OMS_CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');
  let updateCount = 0;

  // Legacy Column mapping:
  // 0 Order ID, 1 Customer ID, 2 Email, 3 Full Name, 4 Customer Classification, 5 Shipping Status, 6 Tracking Number (Logen), 7 Provider, 8 TrackingNumber, 9 Completion Date, 10 Destination

  rawData.forEach((row, idx) => {
    if (idx === 0 && (String(row[0]).toLowerCase().includes('order') || String(row[0]).toLowerCase().includes('id'))) return;

    const merchantOrderId = String(row[0] || '').trim();
    const rowNum = outMap[merchantOrderId];
    if (!rowNum) return;

    const domesticTracking = String(row[6] || '').trim();
    const internationalTracking = String(row[8] || '').trim();
    const provider = String(row[7] || '').trim().toUpperCase();
    const destination = String(row[10] || '').trim();
    const completionDate = migrate_parseCompletionDate_(row[9]);

    let carrier = '';
    if (provider.includes('FEDEX')) carrier = 'FEDEX';
    else if (provider.includes('UPS')) carrier = 'UPS';
    else if (provider.includes('DHL')) carrier = 'DHL';
    else if (provider.includes('USPS')) carrier = 'USPS';

    let status = 'CREATED';
    if (completionDate) status = 'DELIVERED';
    else if (internationalTracking) status = 'US_SHIPPED';
    else if (domesticTracking) status = 'KR_SHIPPED';

    const legacyMeta = `Legacy Info: ${row[3]} (${row[2]}), Class: ${row[4]}, Prov: ${row[7]}`;
    const rIdx = rowNum - 1;

    // Updates in memory array
    if (outCols['domestic-tracking-kr']) outData[rIdx][outCols['domestic-tracking-kr'] - 1] = domesticTracking;
    if (outCols['international-tracking-us']) outData[rIdx][outCols['international-tracking-us'] - 1] = internationalTracking;
    if (outCols['carrier-us']) outData[rIdx][outCols['carrier-us'] - 1] = carrier;
    if (outCols['delivery-country']) outData[rIdx][outCols['delivery-country'] - 1] = destination;
    if (outCols['delivered-date'] && completionDate) outData[rIdx][outCols['delivered-date'] - 1] = completionDate;
    if (outCols['outbound-status']) outData[rIdx][outCols['outbound-status'] - 1] = status;

    if (outCols['notes']) {
      const currentNote = String(outData[rIdx][outCols['notes'] - 1] || '').trim();
      outData[rIdx][outCols['notes'] - 1] = currentNote ? currentNote + '\n' + legacyMeta : legacyMeta;
    }

    if (outCols['system-updated-at']) outData[rIdx][outCols['system-updated-at'] - 1] = stamp;

    updateCount++;
  });

  if (updateCount > 0) {
    // Batch write updated data
    outbound.getRange(1, 1, outData.length, outData[0].length).setValues(outData);

    // Timeline and Linkify must happen after values are written for display/rich-text purposes
    rawData.forEach((row, idx) => {
      if (idx === 0 && (String(row[0]).toLowerCase().includes('order') || String(row[0]).toLowerCase().includes('id'))) return;
      const merchantOrderId = String(row[0] || '').trim();
      const rowNum = outMap[merchantOrderId];
      if (!rowNum) return;

      outbound_updateStageTimeline_(outbound, rowNum, outCols, outCols['international-tracking-us']);
      outbound_linkifyRow_(outbound, rowNum, outCols);
    });

    SpreadsheetApp.getUi().alert(`Successfully updated ${updateCount} outbound rows.`);
  } else {
    SpreadsheetApp.getUi().alert('No matching outbound rows to update.');
  }
}
