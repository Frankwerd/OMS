/********************************
 * Outbound.gs
 * Trigger-ready but NOT installed automatically
 ********************************/

function outbound_onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== OMS_CONFIG.TABS.OUTBOUND) return;
  if (e.range.getRow() === 1) return;

  const cols = OMS_Utils.getHeadersMap_(sheet);
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // stamp updated
  const updatedCol = OMS_Utils.col_(cols, 'system-updated-at');
  if (updatedCol) sheet.getRange(row, updatedCol).setValue(Utilities.formatDate(new Date(), OMS_CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss'));

  // linkify on tracking changes
  const domCol = OMS_Utils.col_(cols, 'domestic-tracking-kr');
  const intlCol = OMS_Utils.col_(cols, 'international-tracking-us');
  if (col === domCol || col === intlCol) {
    outbound_linkifyRow_(sheet, row, cols);
  }

  // SN verify on scan
  const scanCol = OMS_Utils.col_(cols, 'serial-number-scanned');
  if (col === scanCol) outbound_verifySerial_(row);

  // Hub trigger gate: ONLY when international-tracking-us updated
  if (col === intlCol) outbound_sendFinalDeliveryEmail_(row);

  // Velocity Automation: set status based on dates
  outbound_updateStatusFromDates_(sheet, row, cols, col);

  // Timeline tracking
  outbound_updateStageTimeline_(sheet, row, cols, col);
}

/**
 * Velocity Metrics Automation
 * When hub-received-date filled → outbound-status = HUB_RECEIVED
 * When us-ship-date filled → outbound-status = US_SHIPPED
 * When delivered-date filled → outbound-status = DELIVERED
 */
function outbound_updateStatusFromDates_(sheet, row, cols, editedCol) {
  const hubDateCol = OMS_Utils.col_(cols, 'hub-received-date');
  const usDateCol = OMS_Utils.col_(cols, 'us-ship-date');
  const delDateCol = OMS_Utils.col_(cols, 'delivered-date');
  const statusCol = OMS_Utils.col_(cols, 'outbound-status');

  if (!statusCol) return;
  if (editedCol !== hubDateCol && editedCol !== usDateCol && editedCol !== delDateCol) return;

  const hubDate = sheet.getRange(row, hubDateCol).getValue();
  const usDate = sheet.getRange(row, usDateCol).getValue();
  const delDate = sheet.getRange(row, delDateCol).getValue();

  let newStatus = '';
  if (delDate) {
    newStatus = 'DELIVERED';
  } else if (usDate) {
    newStatus = 'US_SHIPPED';
  } else if (hubDate) {
    newStatus = 'HUB_RECEIVED';
  }

  if (newStatus) {
    sheet.getRange(row, statusCol).setValue(newStatus);
  }
}

/**
 * Create Outbound stubs for Inbound items
 */
function outbound_createStubs_(items) {
  if (!items || !items.length) return;

  const out = OMS_Utils.sheet_(OMS_CONFIG.TABS.OUTBOUND);
  const cols = OMS_Utils.getHeadersMap_(out);
  const omsIidCol = cols['oms-order-item-id'];
  if (!omsIidCol) {
    OMS_Utils.opsAlert_('Outbound stub creation failed: oms-order-item-id column not found.');
    return;
  }

  const lr = out.getLastRow();
  const existingMap = {};
  if (lr >= 2) {
    const ids = out.getRange(2, omsIidCol, lr - 1, 1).getValues();
    ids.forEach((r, i) => {
      const id = String(r[0] || '').trim();
      if (id) existingMap[id] = i + 2;
    });
  }

  const stamp = Utilities.formatDate(new Date(), OMS_CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');

  items.forEach(it => {
    const row = new Array(out.getLastColumn()).fill('');

    OMS_Utils.setByHeader_(row, cols, 'merchant-order-id', it.merchantOrderId);
    OMS_Utils.setByHeader_(row, cols, 'merchant-order-item-id', it.merchantOrderItemId);
    OMS_Utils.setByHeader_(row, cols, 'sku', it.sku);
    OMS_Utils.setByHeader_(row, cols, 'customer-id', it.customerId);
    OMS_Utils.setByHeader_(row, cols, 'order-created-at', it.orderCreatedAt);
    OMS_Utils.setByHeader_(row, cols, 'delivery-country', it.deliveryCountry);
    OMS_Utils.setByHeader_(row, cols, 'outbound-workflow-type', 'DIRECT_SHIP');
    OMS_Utils.setByHeader_(row, cols, 'outbound-status', 'CREATED');
    OMS_Utils.setByHeader_(row, cols, 'oms-order-id', it.omsOrderId);
    OMS_Utils.setByHeader_(row, cols, 'oms-order-item-id', it.omsOrderItemId);
    OMS_Utils.setByHeader_(row, cols, 'shipment-id', `${it.omsOrderItemId}:DIRECT_SHIP:001`);
    OMS_Utils.setByHeader_(row, cols, 'stage-timeline', `CREATED — ${it.orderCreatedAt || stamp}`);

    // Package defaults based on stand
    if (it.magSafeStand === 'Yes' || it.magSafeStand === '1') {
      OMS_Utils.setByHeader_(row, cols, 'package-type', 'club-with-stand');
      OMS_Utils.setByHeader_(row, cols, 'actual-weight-kg', OMS_CONFIG.STAND_DEFAULTS.WEIGHT_KG);
      OMS_Utils.setByHeader_(row, cols, 'package-length-cm', OMS_CONFIG.STAND_DEFAULTS.LENGTH_CM);
      OMS_Utils.setByHeader_(row, cols, 'package-width-cm', OMS_CONFIG.STAND_DEFAULTS.WIDTH_CM);
      OMS_Utils.setByHeader_(row, cols, 'package-height-cm', OMS_CONFIG.STAND_DEFAULTS.HEIGHT_CM);
      OMS_Utils.setByHeader_(row, cols, 'notes', 'Stand included.');
    } else {
      OMS_Utils.setByHeader_(row, cols, 'package-type', 'standard-club');
      OMS_Utils.setByHeader_(row, cols, 'actual-weight-kg', OMS_CONFIG.PACKAGE_DEFAULTS.WEIGHT_KG);
      OMS_Utils.setByHeader_(row, cols, 'package-length-cm', OMS_CONFIG.PACKAGE_DEFAULTS.LENGTH_CM);
      OMS_Utils.setByHeader_(row, cols, 'package-width-cm', OMS_CONFIG.PACKAGE_DEFAULTS.WIDTH_CM);
      OMS_Utils.setByHeader_(row, cols, 'package-height-cm', OMS_CONFIG.PACKAGE_DEFAULTS.HEIGHT_CM);
    }

    OMS_Utils.setByHeader_(row, cols, 'system-updated-at', stamp);

    const existingRow = existingMap[it.omsOrderItemId];
    if (existingRow) {
      // Upsert: Only update core stub fields to avoid wiping out tracking/dates
      const stubHeaders = [
        'merchant-order-id','merchant-order-item-id','sku','customer-id',
        'order-created-at','delivery-country',
        'outbound-workflow-type','oms-order-id','oms-order-item-id',
        'system-updated-at'
      ];
      // Also update package defaults if they are blank in the existing row
      // but for simplicity, we'll just update these core ones and leave the rest.
      stubHeaders.forEach(h => {
        const c = cols[h];
        if (c) out.getRange(existingRow, c).setValue(row[c - 1]);
      });

      // Update dimensions if currently empty
      ['actual-weight-kg','package-length-cm','package-width-cm','package-height-cm','notes'].forEach(h => {
        const c = cols[h];
        if (c) {
          const current = out.getRange(existingRow, c).getValue();
          if (!current) out.getRange(existingRow, c).setValue(row[c - 1]);
        }
      });
    } else {
      out.getRange(out.getLastRow() + 1, 1, 1, row.length)
        .setValues([row]);
    }
  });
  SpreadsheetApp.flush();
}

function outbound_linkifyRow_(sheet, row, cols) {
  const domCol = OMS_Utils.col_(cols, 'domestic-tracking-kr');
  if (domCol) {
    const v = sheet.getRange(row, domCol).getDisplayValue();
    if (v) sheet.getRange(row, domCol).setRichTextValue(OMS_Utils.buildTrackingRichText_('LOGEN', v));
  }

  const intlCol = OMS_Utils.col_(cols, 'international-tracking-us');
  if (intlCol) {
    const carrierCol = OMS_Utils.col_(cols, 'carrier-us');
    const carrier = carrierCol ? sheet.getRange(row, carrierCol).getValue() : 'FEDEX';
    const v = sheet.getRange(row, intlCol).getDisplayValue();
    if (v) sheet.getRange(row, intlCol).setRichTextValue(OMS_Utils.buildTrackingRichText_(String(carrier || 'FEDEX'), v));
  }
}

function outbound_verifySerial_(outboundRow) {
  const out = OMS_Utils.sheet_(OMS_CONFIG.TABS.OUTBOUND);
  const outCols = OMS_Utils.requireCols_(out, ['oms-order-item-id','serial-number-scanned','sn-verify','notes','sku']);

  const omsItem = String(out.getRange(outboundRow, outCols['oms-order-item-id']).getValue() || '').trim();
  const scanned = String(out.getRange(outboundRow, outCols['serial-number-scanned']).getValue() || '').trim();
  const sku = String(out.getRange(outboundRow, outCols['sku']).getValue() || '').trim();

  if (!omsItem || !scanned) return;

  const inbound = OMS_Utils.sheet_(OMS_CONFIG.TABS.INBOUND);
  const inCols = OMS_Utils.requireCols_(inbound, ['oms-order-item-id','serial-number-allocated','sku']);

  const allocated = findInboundSerial_(inbound, inCols, omsItem, sku);

  if (!allocated) {
    out.getRange(outboundRow, outCols['sn-verify']).setValue('ERROR: No allocated S/N');
    OMS_Utils.opsAlert_(`S/N verify ERROR (no allocated)\noms-order-item-id: ${omsItem}\nSKU: ${sku}`);
    return;
  }

  if (allocated === scanned) {
    out.getRange(outboundRow, outCols['sn-verify']).setValue('OK');
  } else {
    out.getRange(outboundRow, outCols['sn-verify']).setValue('MISMATCH');
    out.getRange(outboundRow, outCols['notes']).setValue(`S/N mismatch. Allocated=${allocated}, Scanned=${scanned}`.slice(0, 220));
    OMS_Utils.opsAlert_(`S/N MISMATCH\noms-order-item-id: ${omsItem}\nSKU: ${sku}\nAllocated: ${allocated}\nScanned: ${scanned}`);
  }
}

function findInboundSerial_(sheet, cols, omsItem, sku) {
  const lr = sheet.getLastRow();
  if (lr < 2) return '';
  const data = sheet.getRange(2, 1, lr - 1, sheet.getLastColumn()).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    const oi = String(data[i][cols['oms-order-item-id'] - 1] || '').trim();
    const s = String(data[i][cols['sku'] - 1] || '').trim();
    if (oi === omsItem && (!sku || s === sku)) {
      return String(data[i][cols['serial-number-allocated'] - 1] || '').trim();
    }
  }
  return '';
}

function outbound_sendFinalDeliveryEmail_(outboundRow) {
  const out = OMS_Utils.sheet_(OMS_CONFIG.TABS.OUTBOUND);
  const outCols = OMS_Utils.requireCols_(out, [
    'oms-order-item-id','international-tracking-us','carrier-us','customer-email-status','last-email-at','notes'
  ]);

  const omsItem = String(out.getRange(outboundRow, outCols['oms-order-item-id']).getValue() || '').trim();
  const tracking = String(out.getRange(outboundRow, outCols['international-tracking-us']).getDisplayValue() || '').trim();
  const carrier = String(out.getRange(outboundRow, outCols['carrier-us']).getValue() || '').trim() || 'FEDEX';
  const status = String(out.getRange(outboundRow, outCols['customer-email-status']).getValue() || '').trim();

  if (!tracking) return;
  if (/sent/i.test(status)) return;

  const inbound = OMS_Utils.sheet_(OMS_CONFIG.TABS.INBOUND);
  const inCols = OMS_Utils.requireCols_(inbound, ['oms-order-item-id','buyer-email','buyer-name','merchant-order-id']);

  const buyer = findBuyer_(inbound, inCols, omsItem);
  if (!buyer.email) {
    out.getRange(outboundRow, outCols['notes']).setValue('Cannot send email: buyer-email not found in inbound.');
    OMS_Utils.opsAlert_(`Final email blocked (no buyer-email)\noms-order-item-id: ${omsItem}`);
    return;
  }

  // linkify the tracking cell too
  out.getRange(outboundRow, outCols['international-tracking-us'])
    .setRichTextValue(OMS_Utils.buildTrackingRichText_(carrier, tracking));

  const subject = `Your G·GRIP order #${buyer.orderId} is on final delivery`;
  const html = `
    <div style="font-family:Arial,sans-serif;line-height:1.6;">
      <p>Hello ${buyer.name || ''},</p>
      <p>Your package is now on its final delivery leg.</p>
      <p><b>Order:</b> ${buyer.orderId}<br/>
         <b>Carrier:</b> ${carrier}<br/>
         <b>Tracking:</b> ${tracking}</p>
      <p>Thank you,<br/>The G·GRIP Team</p>
    </div>
  `;

  try {
    OMS_Utils.sendCustomerEmail_(buyer.email, subject, html);
    out.getRange(outboundRow, outCols['customer-email-status']).setValue('Sent: Final Delivery');
    out.getRange(outboundRow, outCols['last-email-at']).setValue(Utilities.formatDate(new Date(), OMS_CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss'));
  } catch (err) {
    out.getRange(outboundRow, outCols['customer-email-status']).setValue('Error');
    out.getRange(outboundRow, outCols['notes']).setValue(`Email error: ${err.message}`.slice(0, 200));
    OMS_Utils.opsAlert_(`Final email FAILED\noms-order-item-id: ${omsItem}\nError: ${err.message}`);
  }
}

/**
 * Stage Timeline Tracking
 * Appends human-readable stage updates to stage-timeline column
 */
function outbound_updateStageTimeline_(sheet, row, cols, editedCol) {
  const timelineCol = OMS_Utils.col_(cols, 'stage-timeline');
  if (!timelineCol) return;

  const fields = {
    'hub-received-date': 'Hub received',
    'us-ship-date': 'US shipped',
    'delivered-date': 'Delivered',
    'domestic-tracking-kr': 'KR Tracking',
    'international-tracking-us': 'US Tracking'
  };

  const fieldKeys = Object.keys(fields);
  const relevantCols = fieldKeys.map(k => OMS_Utils.col_(cols, k));

  if (!relevantCols.includes(editedCol)) return;

  const timelineRange = sheet.getRange(row, timelineCol);

  // Rebuild timeline deterministically based on current values of all tracked fields.
  const events = [];

  // Initial event
  const createdAtCol = OMS_Utils.col_(cols, 'order-created-at');
  if (createdAtCol) {
    const createdVal = sheet.getRange(row, createdAtCol).getDisplayValue();
    if (createdVal) events.push(`CREATED — ${createdVal}`);
  }

  fieldKeys.forEach(k => {
    const c = OMS_Utils.col_(cols, k);
    if (!c) return;
    const val = sheet.getRange(row, c).getDisplayValue();
    if (val) {
      events.push(`${fields[k]} — ${val}`);
    }
  });

  timelineRange.setValue(events.join('\n'));
}

function findBuyer_(sheet, cols, omsItem) {
  const lr = sheet.getLastRow();
  if (lr < 2) return { email:'', name:'', orderId:'' };
  const data = sheet.getRange(2, 1, lr - 1, sheet.getLastColumn()).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    const oi = String(data[i][cols['oms-order-item-id'] - 1] || '').trim();
    if (oi === omsItem) {
      return {
        email: String(data[i][cols['buyer-email'] - 1] || '').trim(),
        name: String(data[i][cols['buyer-name'] - 1] || '').trim(),
        orderId: String(data[i][cols['merchant-order-id'] - 1] || '').trim(),
      };
    }
  }
  return { email:'', name:'', orderId:'' };
}
