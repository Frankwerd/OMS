/********************************
 * Inbound.gs
 * Restored SamCart ingestion (NOT Same Cart)
 * NO triggers created automatically
 ********************************/

function inbound_runSamCart() {
  const inbound = OMS_Utils.sheet_(OMS_CONFIG.TABS.INBOUND);

  const required = [
    'merchant-order-id','merchant-order-item-id','purchase-date','buyer-email','customer-id','system-gmail-id',
    'sales-channel','sku','quantity-purchased',
    'source-system','source-order-id','source-order-item-id','oms-order-id','oms-order-item-id','buyer-email-hash',
    'system-created-at','system-updated-at','parse-status'
  ];
  const cols = OMS_Utils.requireCols_(inbound, required);

  const labels = {
    toProcess: OMS_Utils.getOrCreateLabel_(OMS_CONFIG.GMAIL.SAMCART_TO_PROCESS),
    processed: OMS_Utils.getOrCreateLabel_(OMS_CONFIG.GMAIL.SAMCART_PROCESSED),
    error: OMS_Utils.getOrCreateLabel_(OMS_CONFIG.GMAIL.SAMCART_ERROR),
  };

  const threads = labels.toProcess.getThreads();
  if (!threads.length) return;

  const existing = inbound_collectExistingMessageIds_(inbound, cols['system-gmail-id']);

  threads.reverse().forEach(thread => {
    let okThread = true;

    thread.getMessages().forEach(msg => {
      const msgId = String(msg.getId() || '').trim().toLowerCase();
      if (!msgId || existing.has(msgId)) return;

      try {
        const clean = OMS_Utils.ultraCleanText_(msg.getBody());
        const parsed = inbound_parseSamCartInvoice_(clean);

        const customerId = OMS_Utils.lookupOrCreateCustomerId_(parsed.buyerEmail);
        const emailHash = OMS_Utils.emailHash_(parsed.buyerEmail);

        const sourceSystem = OMS_CONFIG.SOURCE_SYSTEMS.SAMCART;
        const sourceOrderId = parsed.orderId;
        const omsOrderId = OMS_Utils.buildOmsOrderId_(sourceSystem, sourceOrderId);

        const now = new Date();
        const stamp = Utilities.formatDate(now, OMS_CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');

        const rows = [];
        parsed.items.forEach((it, idx) => {
          const sourceOrderItemId = OMS_Utils.generateLineItemId_(idx + 1);
          const omsOrderItemId = OMS_Utils.buildOmsOrderItemId_(omsOrderId, sourceOrderItemId);

          rows.push(inbound_buildRowFromHeaders_(cols, inbound.getLastColumn(), {
            // core ids
            merchantOrderId: sourceOrderId,
            merchantOrderItemId: sourceOrderItemId,

            purchaseDate: parsed.purchaseDate,
            purchaseTime: parsed.purchaseTime,

            buyerEmail: parsed.buyerEmail,
            buyerName: parsed.buyerName,
            buyerPhone: parsed.buyerPhone,

            customerId,
            systemGmailId: msgId,

            salesChannel: 'SamCart',
            customerClassification: 'Active',
            isBusinessOrder: 'false',

            sku: it.sku,
            productName: it.productName || parsed.productName,
            magSafeStand: it.magSafeStand,
            model: it.model || parsed.model,
            clubType: it.clubType || parsed.clubType,
            hand: it.hand || parsed.hand,
            flex: it.flex || parsed.flex,
            length: it.length || parsed.length,
            gripSize: it.gripSize || parsed.gripSize,

            qty: it.quantity || parsed.quantity || 1,
            currency: 'USD',
            itemPrice: parsed.subtotal,
            itemTax: parsed.tax,
            shippingPrice: parsed.shipping,
            totalAmount: parsed.totalAmount,
            couponCode: parsed.couponCode,
            refundAmount: parsed.refundAmount,
            refundDate: parsed.refundDate,
            returnReasonCode: parsed.returnReasonCode,

            recipientName: parsed.recipientName || parsed.buyerName,
            shipAddr1: parsed.shipAddr1,
            shipCity: parsed.shipCity,
            shipState: parsed.shipState,
            shipPostal: parsed.shipPostal,
            shipCountry: parsed.shipCountry || 'United States',
            shipServiceLevel: parsed.shipServiceLevel,

            // ops fields
            serialAllocated: '',
            notes: '',
            automationNotes: '',
            itemLifeCycle: 'ACTIVE',
            orderLifeCycle: (Number(parsed.refundAmount || 0) > 0 ? 'PARTIAL_REFUND' : 'ACTIVE'),
            parseStatus: 'OK',

            createdAt: stamp,
            updatedAt: stamp,

            // canonical
            sourceSystem,
            sourceOrderId,
            sourceOrderItemId,
            omsOrderId,
            omsOrderItemId,
            buyerEmailHash: emailHash,
          }));
        });

        if (!rows.length) throw new Error('No items parsed from SamCart invoice.');

        inbound.getRange(inbound.getLastRow() + 1, 1, rows.length, rows[0].length)
          .setNumberFormat('@')
          .setValues(rows);

        existing.add(msgId);

      } catch (err) {
        okThread = false;
        OMS_Utils.opsAlert_(
          `SamCart parse/write failed.\nThread: ${thread.getId()}\nMsg: ${msg.getId()}\nError: ${err.message}`
        );
      }
    });

    if (okThread) {
      thread.removeLabel(labels.toProcess).addLabel(labels.processed).moveToArchive();
    } else {
      thread.removeLabel(labels.toProcess).addLabel(labels.error);
    }
  });

  SpreadsheetApp.flush();
}

function inbound_collectExistingMessageIds_(sheet, msgCol) {
  const lr = sheet.getLastRow();
  const set = new Set();
  if (lr < 2) return set;
  sheet.getRange(2, msgCol, lr - 1, 1).getValues().forEach(r => {
    const v = String(r[0] || '').trim().toLowerCase();
    if (v) set.add(v);
  });
  return set;
}

/**
 * SamCart parsing: based on your legacy SamCartEngine.parse
 * Returns normalized object.
 */
function inbound_parseSamCartInvoice_(text) {
  const t = String(text || '');
  const u = t.toUpperCase();

  // Order ID
  const idMatch = t.match(/Order\s*ID:\s*#?\s*([A-Za-z0-9\-]+)/i) || t.match(/Order\s*#\s*([A-Za-z0-9\-]+)/i);
  const orderId = idMatch ? String(idMatch[1]).trim() : '';
  if (!orderId) throw new Error('Order ID not found');

  // Date
  const dateMatch =
    t.match(/(?:Original Order Date|Receipt Date|Date):\s*([A-Z][a-z]+ \d{1,2}(?:st|nd|rd|th)?,? \d{4})/i) ||
    t.match(/Date:\s*([A-Z][a-z]+ \d{1,2}, \d{4})/i);

  const purchaseDate = dateMatch ? normalizeDateYYYYMMDD_(dateMatch[1]) : '';
  const purchaseTime = ''; // SamCart emails usually don’t include a stable time; leave blank.

  // Customer name/email
  const cust = t.match(/Customer\s*\n([^\n]+)\n([^\n]+@\S+)/i);
  let buyerName = cust ? String(cust[1] || '').trim() : '';
  let buyerEmail = cust ? String(cust[2] || '').trim() : '';

  if (!buyerEmail) {
    const emails = [...t.matchAll(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g)].map(m => m[1]);
    buyerEmail = emails[0] ? String(emails[0]).trim() : '';
  }
  if (!buyerEmail) throw new Error('buyer-email not found');

  // Phone
  const phoneMatch =
    t.match(/(?:Phone|delivery\.):\s*([\d\+\-\s\(\)]{7,20})/i) ||
    t.match(/Customer\s*\n.*\n.*\n(\+?[\d\s-]{7,20})/i);
  const buyerPhone = phoneMatch ? String(phoneMatch[1]).trim().replace(/\s/g, '') : '';

  // Ship To block
  let shipAddr1 = '', shipCity = '', shipState = '', shipPostal = '', shipCountry = 'United States';
  const shipBlockMatch = t.match(/Ship To\n([\s\S]*?)\n(?:Flex|Hand|Length|Grip|Order ID|Original|Customer|Order Items)/i);
  if (shipBlockMatch) {
    const lines = shipBlockMatch[1].split('\n').map(x => x.trim()).filter(Boolean);
    const addr = OMS_Utils.parseGlobalAddress(lines);
    shipAddr1 = addr.addr1;
    shipCity = addr.city;
    shipState = addr.state;
    shipPostal = addr.zip;
    shipCountry = addr.country || shipCountry;
  }

  // Specs (same as legacy)
  const flex = (u.includes('L-FLEX') || u.includes('LADIE')) ? 'L' : (u.includes('R-FLEX') || u.includes('REGULAR')) ? 'R' : 'S';
  const gripSize = (u.includes('MIDSIZE') || u.includes('MID SIZE')) ? 'Mid' : 'Standard';
  const length = u.includes('LONGER') ? 'Longer' : 'Standard';
  const hand = (u.includes('LEFT HAND') || u.includes('HAND: LEFT')) ? 'Left' : 'Right';
  const magSafeStand = (u.includes('TRIPOD') || u.includes('BUNDLE')) ? 'Yes' : '0';

  // Money
  const money = (label) => {
    const re = new RegExp(label + "\\s*\\$?([\\d,.]+)", "gi");
    const matches = [...t.matchAll(re)];
    if (!matches.length) return 0;
    return Number(String(matches[matches.length - 1][1]).replace(/,/g, '')) || 0;
  };

  const subtotal = money('Subtotal');
  const shipping = money('Shipping');
  const tax = money('Tax');
  const discount = money('Discount');
  const refundAmount = money('Refund');
  const totalAmount = Number((subtotal + shipping + tax - discount - refundAmount).toFixed(2));

  const couponCode = (t.match(/Coupon:\s*([^\n\r]+)/i) || ['', ''])[1].trim();
  const quantity = Number((t.match(/Qty:\s*(\d+)/i) || ['', '1'])[1]) || 1;

  // SKU (SamCart often does not have it) → safe placeholder
  let sku = '';
  const skuMatch = t.match(/\bSKU\b\s*[:#]?\s*([A-Za-z0-9\-_.]{4,})/i);
  if (skuMatch) sku = String(skuMatch[1]).trim();
  if (!sku) sku = 'SAMCART-UNMAPPED';

  return {
    orderId,
    purchaseDate,
    purchaseTime,
    buyerName,
    buyerEmail,
    buyerPhone,

    shipAddr1,
    shipCity,
    shipState,
    shipPostal,
    shipCountry,
    shipServiceLevel: '',

    productName: '',
    model: '',        // your Pro/Basic can be inferred later if text contains
    clubType: '',

    hand,
    flex,
    length,
    gripSize,
    magSafeStand,

    quantity,
    subtotal,
    shipping,
    tax,
    totalAmount,
    couponCode,

    refundAmount,
    refundDate: '',
    returnReasonCode: '',
    recipientName: '',

    items: [{
      sku,
      quantity,
      productName: '',
      model: '',
      clubType: '',
      hand,
      flex,
      length,
      gripSize,
      magSafeStand,
    }],
  };
}

function inbound_buildRowFromHeaders_(headersMap, lastCol, v) {
  const row = new Array(lastCol).fill('');

  OMS_Utils.setByHeader_(row, headersMap, 'merchant-order-id', v.merchantOrderId);
  OMS_Utils.setByHeader_(row, headersMap, 'merchant-order-item-id', v.merchantOrderItemId);
  OMS_Utils.setByHeader_(row, headersMap, 'purchase-date', v.purchaseDate);
  OMS_Utils.setByHeader_(row, headersMap, 'purchase-time', v.purchaseTime);

  OMS_Utils.setByHeader_(row, headersMap, 'buyer-email', v.buyerEmail);
  OMS_Utils.setByHeader_(row, headersMap, 'buyer-name', v.buyerName);
  OMS_Utils.setByHeader_(row, headersMap, 'buyer-phone-number', v.buyerPhone);

  OMS_Utils.setByHeader_(row, headersMap, 'customer-id', v.customerId);
  OMS_Utils.setByHeader_(row, headersMap, 'system-gmail-id', v.systemGmailId);

  OMS_Utils.setByHeader_(row, headersMap, 'sales-channel', v.salesChannel);
  OMS_Utils.setByHeader_(row, headersMap, 'customer-classification', v.customerClassification);
  OMS_Utils.setByHeader_(row, headersMap, 'is-business-order', v.isBusinessOrder);

  OMS_Utils.setByHeader_(row, headersMap, 'sku', v.sku);
  OMS_Utils.setByHeader_(row, headersMap, 'product-name', v.productName);
  OMS_Utils.setByHeader_(row, headersMap, 'mag-safe-stand', v.magSafeStand);
  OMS_Utils.setByHeader_(row, headersMap, 'model', v.model);
  OMS_Utils.setByHeader_(row, headersMap, 'club-type', v.clubType);
  OMS_Utils.setByHeader_(row, headersMap, 'hand', v.hand);
  OMS_Utils.setByHeader_(row, headersMap, 'flex', v.flex);
  OMS_Utils.setByHeader_(row, headersMap, 'length', v.length);
  OMS_Utils.setByHeader_(row, headersMap, 'grip-size', v.gripSize);

  OMS_Utils.setByHeader_(row, headersMap, 'quantity-purchased', v.qty);
  OMS_Utils.setByHeader_(row, headersMap, 'currency', v.currency);
  OMS_Utils.setByHeader_(row, headersMap, 'item-price', v.itemPrice);
  OMS_Utils.setByHeader_(row, headersMap, 'item-tax', v.itemTax);
  OMS_Utils.setByHeader_(row, headersMap, 'shipping-price', v.shippingPrice);
  OMS_Utils.setByHeader_(row, headersMap, 'total-amount', v.totalAmount);
  OMS_Utils.setByHeader_(row, headersMap, 'coupon-code', v.couponCode);

  OMS_Utils.setByHeader_(row, headersMap, 'refund-amount', v.refundAmount);
  OMS_Utils.setByHeader_(row, headersMap, 'refund-date', v.refundDate);
  OMS_Utils.setByHeader_(row, headersMap, 'return-reason-code', v.returnReasonCode);

  OMS_Utils.setByHeader_(row, headersMap, 'recipient-name', v.recipientName);
  OMS_Utils.setByHeader_(row, headersMap, 'ship-address-1', v.shipAddr1);
  OMS_Utils.setByHeader_(row, headersMap, 'ship-city', v.shipCity);
  OMS_Utils.setByHeader_(row, headersMap, 'ship-state', v.shipState);
  OMS_Utils.setByHeader_(row, headersMap, 'ship-postal-code', v.shipPostal);
  OMS_Utils.setByHeader_(row, headersMap, 'ship-country', v.shipCountry);
  OMS_Utils.setByHeader_(row, headersMap, 'ship-service-level', v.shipServiceLevel);

  OMS_Utils.setByHeader_(row, headersMap, 'serial-number-allocated', v.serialAllocated);

  OMS_Utils.setByHeader_(row, headersMap, 'notes', v.notes);
  OMS_Utils.setByHeader_(row, headersMap, 'automation-notes', v.automationNotes);

  OMS_Utils.setByHeader_(row, headersMap, 'item-life-cycle', v.itemLifeCycle);
  OMS_Utils.setByHeader_(row, headersMap, 'order-life-cycle', v.orderLifeCycle);

  OMS_Utils.setByHeader_(row, headersMap, 'system-created-at', v.createdAt);
  OMS_Utils.setByHeader_(row, headersMap, 'system-updated-at', v.updatedAt);
  OMS_Utils.setByHeader_(row, headersMap, 'parse-status', v.parseStatus);

  OMS_Utils.setByHeader_(row, headersMap, 'source-system', v.sourceSystem);
  OMS_Utils.setByHeader_(row, headersMap, 'source-order-id', v.sourceOrderId);
  OMS_Utils.setByHeader_(row, headersMap, 'source-order-item-id', v.sourceOrderItemId);
  OMS_Utils.setByHeader_(row, headersMap, 'oms-order-id', v.omsOrderId);
  OMS_Utils.setByHeader_(row, headersMap, 'oms-order-item-id', v.omsOrderItemId);
  OMS_Utils.setByHeader_(row, headersMap, 'buyer-email-hash', v.buyerEmailHash);

  return row;
}

function normalizeDateYYYYMMDD_(dateStr) {
  const s = String(dateStr || '').replace(/(st|nd|rd|th)/g, '').replace(/,/g, '').trim();
  const d = new Date(s);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
