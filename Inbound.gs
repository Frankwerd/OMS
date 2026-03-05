/********************************
 * Inbound.gs
 * Restored SamCart ingestion (NOT Same Cart)
 * NO triggers created automatically
 ********************************/

/**
 * Inbound Shopify ingestion
 */
function inbound_runShopify() {
  const ss = OMS_Utils.ss();
  validateSchema(ss);
  const inbound = OMS_Utils.sheet_(OMS_CONFIG.TABS.INBOUND);

  const required = [
    'merchant-order-id','merchant-order-item-id','purchase-date','buyer-email','customer-id','system-gmail-id',
    'sales-channel','sku','quantity-purchased',
    'source-system','source-order-id','source-order-item-id','oms-order-id','oms-order-item-id','buyer-email-hash',
    'system-created-at','system-updated-at','parse-status'
  ];
  const cols = OMS_Utils.requireCols_(inbound, required);

  const labels = {
    toProcess: OMS_Utils.getOrCreateLabel_(OMS_CONFIG.GMAIL.SHOPIFY_TO_PROCESS),
    processed: OMS_Utils.getOrCreateLabel_(OMS_CONFIG.GMAIL.SHOPIFY_PROCESSED),
    error: OMS_Utils.getOrCreateLabel_(OMS_CONFIG.GMAIL.SHOPIFY_ERROR),
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
        const parsed = inbound_parseShopifyOrder_(clean);

        // Standardize reshipments as Outbound-only
        if (String(parsed.orderId).toUpperCase().endsWith(OMS_CONFIG.RESHIP_SUFFIX)) return;

        const buyerEmail = OMS_Utils.normalizeEmail_(parsed.buyerEmail);
        const customerId = buyerEmail ? OMS_Utils.lookupOrCreateCustomerId_(buyerEmail) : 'C-UNKNOWN';
        const emailHash = buyerEmail ? OMS_Utils.emailHash_(buyerEmail) : '';

        const sourceSystem = OMS_CONFIG.SOURCE_SYSTEMS.SHOPIFY;
        const sourceOrderId = parsed.orderId;
        const omsOrderId = OMS_Utils.buildOmsOrderId_(sourceSystem, sourceOrderId);

        const orderCreatedAt = (parsed.purchaseDate && parsed.purchaseTime)
          ? `${parsed.purchaseDate}T${parsed.purchaseTime}:00`
          : Utilities.formatDate(msg.getDate(), OMS_CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss");

        const orderSourceEmail = String(msg.getFrom() || '').match(/<([^>]+)>/)?.[1] || msg.getFrom();

        const now = new Date();
        const stamp = Utilities.formatDate(now, OMS_CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');

        const rows = [];
        const stubData = [];
        parsed.items.forEach((it, idx) => {
          const lineItemIndex = idx + 1;
          const sourceOrderItemId = it.itemId || OMS_Utils.generateLineItemId_(lineItemIndex);
          const omsOrderItemId = OMS_Utils.buildOmsOrderItemId_(omsOrderId, sourceOrderItemId);

          rows.push(inbound_buildRowFromHeaders_(cols, inbound.getLastColumn(), {
            // core ids
            merchantOrderId: sourceOrderId,
            merchantOrderItemId: sourceOrderItemId,
            lineItemIndex,

            purchaseDate: parsed.purchaseDate,
            purchaseTime: parsed.purchaseTime,
            orderCreatedAt,

            buyerEmail,
            buyerName: parsed.buyerName,
            buyerPhone: parsed.buyerPhone,

            customerId,
            systemGmailId: msgId,
            orderSourceEmail,

            salesChannel: 'Shopify',
            customerClassification: 'Active',
            isBusinessOrder: 'false',

            sku: it.sku || 'SHOPIFY-UNMAPPED',
            productName: OMS_Utils.decodeHtmlEntities_(it.productName),
            magSafeStand: it.magSafeStand,
            model: it.model,
            clubType: it.clubType,
            productCategory: 'Golf Club',
            hand: it.hand,
            flex: it.flex,
            shaftLengthOption: it.length,
            gripSize: it.gripSize,

            qty: it.quantity || 1,
            currency: parsed.currency || 'USD',
            itemPrice: it.price,
            itemTax: (parsed.tax / parsed.items.length), // simple split
            shippingPrice: (parsed.shipping / parsed.items.length),
            totalAmount: parsed.totalAmount,
            couponCode: parsed.couponCode,

            recipientName: parsed.recipientName || parsed.buyerName,
            shipAddr1: parsed.shipAddr1,
            shipCity: parsed.shipCity,
            shipState: parsed.shipState,
            shipPostal: parsed.shipPostal,
            shipCountry: parsed.shipCountry,

            // ops fields
            serialAllocated: '',
            notes: parsed.parseNotes || '',
            automationNotes: '',
            itemLifeCycle: 'ACTIVE',
            orderLifeCycle: 'ACTIVE',
            parseStatus: parsed.parseStatus || 'OK',

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

          stubData.push({
            merchantOrderId: sourceOrderId,
            merchantOrderItemId: sourceOrderItemId,
            sku: it.sku || 'SHOPIFY-UNMAPPED',
            customerId,
            omsOrderId,
            omsOrderItemId,
            magSafeStand: it.magSafeStand,
            orderCreatedAt,
            deliveryCountry: parsed.shipCountry,
          });
        });

        if (!rows.length) throw new Error('No items parsed from Shopify order.');

        inbound.getRange(inbound.getLastRow() + 1, 1, rows.length, rows[0].length)
          .setNumberFormat('@')
          .setValues(rows);

        // Create Outbound stubs
        outbound_createStubs_(stubData);

        existing.add(msgId);

      } catch (err) {
        okThread = false;
        OMS_Utils.opsAlert_(
          `Shopify parse/write failed.\nThread: ${thread.getId()}\nMsg: ${msg.getId()}\nError: ${err.message}`
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

/**
 * Shopify parsing logic
 */
function inbound_parseShopifyOrder_(text) {
  const t = String(text || '');
  const u = t.toUpperCase();
  let parseStatus = 'OK', parseNotes = '';

  // Order ID: Order #1234
  const idMatch = t.match(/Order\s*#?\s*(\d{4,10})/i);
  let orderId = idMatch ? idMatch[1] : '';
  if (!orderId) {
    orderId = 'UNKNOWN';
    parseStatus = OMS_CONFIG.ERRORS.MISSING_ORDER_ID;
  }

  // Date: Feb 15, 2024
  const dateMatch = t.match(/([A-Z][a-z]{2}\s\d{1,2},\s\d{4})/);
  const purchaseDate = dateMatch ? normalizeDateYYYYMMDD_(dateMatch[1]) : '';

  // Email
  const emailMatch = t.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
  let buyerEmail = emailMatch ? emailMatch[1].trim() : '';
  if (!buyerEmail && parseStatus === 'OK') {
    parseStatus = OMS_CONFIG.ERRORS.MISSING_EMAIL;
  }

  // Shipping
  let shipAddr1 = '', shipCity = '', shipState = '', shipPostal = '', shipCountry = 'United States';
  const shipMatch = t.match(/Shipping address\n([\s\S]*?)\n(?:Billing address|Payment method|Shipping method)/i);
  if (shipMatch) {
    const lines = shipMatch[1].split('\n').map(x => x.trim()).filter(Boolean);
    const addr = OMS_Utils.parseGlobalAddress(lines);
    shipAddr1 = addr.addr1;
    shipCity = addr.city;
    shipState = addr.state;
    shipPostal = addr.zip;
    shipCountry = addr.country || shipCountry;
    if (!addr.success) {
      parseStatus = OMS_CONFIG.ERRORS.MISSING_ADDRESS;
      parseNotes = 'Address parsing fallback to raw block.';
    }
  } else {
    parseStatus = OMS_CONFIG.ERRORS.MISSING_ADDRESS;
  }

  // Items
  const items = [];
  // Basic Shopify email item pattern: [Product Name] - [Variant] x [Qty]
  // This is highly variable, but we'll attempt a common one.
  const itemBlocks = [...t.matchAll(/(.*)\s+×\s+(\d+)\n\$([\d,.]+)/g)];
  itemBlocks.forEach(m => {
    const rawName = m[1].trim();
    const qty = parseInt(m[2]);
    const price = parseFloat(m[3].replace(/,/g, ''));

    // Specs from variant string? e.g. "G-GRIP Pro - Wood / Right / Regular / Standard / Standard / Yes"
    const specs = {
      productName: rawName,
      quantity: qty,
      price: price,
      model: rawName.toUpperCase().includes('PRO') ? 'Pro' : 'Basic',
      clubType: rawName.toUpperCase().includes('WOOD') ? 'Wood' : (rawName.toUpperCase().includes('IRON') ? 'Iron' : '7-iron'),
      hand: rawName.toUpperCase().includes('LEFT') ? 'Left' : 'Right',
      flex: rawName.toUpperCase().includes('L-FLEX') ? 'L' : (rawName.toUpperCase().includes('S-FLEX') ? 'S' : 'R'),
      length: rawName.toUpperCase().includes('LONGER') ? 'Longer' : 'Standard',
      gripSize: rawName.toUpperCase().includes('MID') ? 'Mid' : 'Standard',
      magSafeStand: rawName.toUpperCase().includes('YES') ? 'Yes' : '0'
    };
    specs.sku = OMS_Utils.deriveSku(specs);
    items.push(specs);
  });

  return {
    orderId,
    purchaseDate,
    buyerEmail,
    shipAddr1,
    shipCity,
    shipState,
    shipPostal,
    shipCountry,
    items,
    totalAmount: 0, // could sum from text if needed
    currency: 'USD',
    parseStatus,
    parseNotes
  };
}

/**
 * Inbound Imweb ingestion (KR localized)
 */
function inbound_runImweb() {
  const ss = OMS_Utils.ss();
  validateSchema(ss);
  const inbound = OMS_Utils.sheet_(OMS_CONFIG.TABS.INBOUND);

  const required = [
    'merchant-order-id','merchant-order-item-id','purchase-date','buyer-email','customer-id','system-gmail-id',
    'sales-channel','sku','quantity-purchased',
    'source-system','source-order-id','source-order-item-id','oms-order-id','oms-order-item-id','buyer-email-hash',
    'system-created-at','system-updated-at','parse-status'
  ];
  const cols = OMS_Utils.requireCols_(inbound, required);

  const labels = {
    toProcess: OMS_Utils.getOrCreateLabel_(OMS_CONFIG.GMAIL.IMWEB_TO_PROCESS),
    processed: OMS_Utils.getOrCreateLabel_(OMS_CONFIG.GMAIL.IMWEB_PROCESSED),
    error: OMS_Utils.getOrCreateLabel_(OMS_CONFIG.GMAIL.IMWEB_ERROR),
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
        const parsed = inbound_parseImwebOrder_(clean);

        // Standardize reshipments as Outbound-only
        if (String(parsed.orderId).toUpperCase().endsWith(OMS_CONFIG.RESHIP_SUFFIX)) return;

        const buyerEmail = OMS_Utils.normalizeEmail_(parsed.buyerEmail);
        const customerId = buyerEmail ? OMS_Utils.lookupOrCreateCustomerId_(buyerEmail) : 'C-UNKNOWN';
        const emailHash = buyerEmail ? OMS_Utils.emailHash_(buyerEmail) : '';

        const sourceSystem = OMS_CONFIG.SOURCE_SYSTEMS.IMWEB;
        const sourceOrderId = parsed.orderId;
        const omsOrderId = OMS_Utils.buildOmsOrderId_(sourceSystem, sourceOrderId);

        const orderCreatedAt = (parsed.purchaseDate && parsed.purchaseTime)
          ? `${parsed.purchaseDate}T${parsed.purchaseTime}:00`
          : Utilities.formatDate(msg.getDate(), OMS_CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss");

        const orderSourceEmail = String(msg.getFrom() || '').match(/<([^>]+)>/)?.[1] || msg.getFrom();

        const now = new Date();
        const stamp = Utilities.formatDate(now, OMS_CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');

        const rows = [];
        const stubData = [];
        parsed.items.forEach((it, idx) => {
          const lineItemIndex = idx + 1;
          const sourceOrderItemId = it.itemId || OMS_Utils.generateLineItemId_(lineItemIndex);
          const omsOrderItemId = OMS_Utils.buildOmsOrderItemId_(omsOrderId, sourceOrderItemId);

          rows.push(inbound_buildRowFromHeaders_(cols, inbound.getLastColumn(), {
            merchantOrderId: sourceOrderId,
            merchantOrderItemId: sourceOrderItemId,
            lineItemIndex,
            purchaseDate: parsed.purchaseDate,
            purchaseTime: parsed.purchaseTime,
            orderCreatedAt,
            buyerEmail,
            buyerName: parsed.buyerName,
            buyerPhone: parsed.buyerPhone,
            customerId,
            systemGmailId: msgId,
            orderSourceEmail,
            salesChannel: 'Imweb',
            customerClassification: 'Active',
            isBusinessOrder: 'false',
            sku: it.sku || 'IMWEB-UNMAPPED',
            productName: OMS_Utils.decodeHtmlEntities_(it.productName),
            magSafeStand: it.magSafeStand,
            model: it.model,
            clubType: it.clubType,
            productCategory: 'Golf Club',
            hand: it.hand,
            flex: it.flex,
            shaftLengthOption: it.length,
            gripSize: it.gripSize,
            qty: it.quantity || 1,
            currency: 'KRW',
            itemPrice: it.price,
            itemTax: 0,
            shippingPrice: (parsed.shipping || 0) / parsed.items.length,
            totalAmount: parsed.totalAmount,
            recipientName: parsed.recipientName || parsed.buyerName,
            shipAddr1: parsed.shipAddr1,
            shipCity: parsed.shipCity,
            shipState: parsed.shipState,
            shipPostal: parsed.shipPostal,
            shipCountry: parsed.shipCountry || 'South Korea',
            serialAllocated: '',
            notes: parsed.parseNotes || '',
            automationNotes: '',
            itemLifeCycle: 'ACTIVE',
            orderLifeCycle: 'ACTIVE',
            parseStatus: parsed.parseStatus || 'OK',
            createdAt: stamp,
            updatedAt: stamp,
            sourceSystem,
            sourceOrderId,
            sourceOrderItemId,
            omsOrderId,
            omsOrderItemId,
            buyerEmailHash: emailHash,
          }));

          stubData.push({
            merchantOrderId: sourceOrderId,
            merchantOrderItemId: sourceOrderItemId,
            sku: it.sku || 'IMWEB-UNMAPPED',
            customerId,
            omsOrderId,
            omsOrderItemId,
            magSafeStand: it.magSafeStand,
            orderCreatedAt,
            deliveryCountry: parsed.shipCountry || 'South Korea',
          });
        });

        if (!rows.length) throw new Error('No items parsed from Imweb order.');

        inbound.getRange(inbound.getLastRow() + 1, 1, rows.length, rows[0].length)
          .setNumberFormat('@')
          .setValues(rows);

        // Create Outbound stubs
        outbound_createStubs_(stubData);

        existing.add(msgId);
      } catch (err) {
        okThread = false;
        OMS_Utils.opsAlert_(`Imweb parse failed.\nMsg: ${msg.getId()}\nError: ${err.message}`);
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

/**
 * Imweb parsing logic (simplified example)
 */
function inbound_parseImwebOrder_(text) {
  const t = String(text || '');
  let parseStatus = 'OK', parseNotes = '';

  const idMatch = t.match(/주문번호\s*[:]\s*([A-Z0-9\-]+)/) || t.match(/Order\s*No\.\s*([A-Z0-9\-]+)/i);
  let orderId = idMatch ? idMatch[1] : '';
  if (!orderId) {
    orderId = 'UNKNOWN';
    parseStatus = OMS_CONFIG.ERRORS.MISSING_ORDER_ID;
  }

  const dateMatch = t.match(/주문일시\s*[:]\s*(\d{4}-\d{2}-\d{2})/);
  const purchaseDate = dateMatch ? dateMatch[1] : '';

  const emailMatch = t.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
  let buyerEmail = emailMatch ? emailMatch[1] : '';
  if (!buyerEmail && parseStatus === 'OK') {
    parseStatus = OMS_CONFIG.ERRORS.MISSING_EMAIL;
  }

  // Address
  let shipAddr1 = '', shipCity = '', shipState = '', shipPostal = '', shipCountry = 'South Korea';
  const addrMatch = t.match(/배송지\s*주소\s*[:]\s*(.*)/);
  if (addrMatch) {
    const raw = addrMatch[1].trim();
    // basic KR parse: [12345] Seoul...
    const krZip = raw.match(/^\[(\d{5})\]\s*(.*)$/);
    if (krZip) {
      shipPostal = krZip[1];
      shipAddr1 = krZip[2];
      // simplistic city extraction for KR
      shipCity = shipAddr1.split(' ')[0];
    } else {
      shipAddr1 = raw;
      parseStatus = OMS_CONFIG.ERRORS.MISSING_ADDRESS;
      parseNotes = 'KR address parsing fallback.';
    }
  } else {
    parseStatus = OMS_CONFIG.ERRORS.MISSING_ADDRESS;
  }

  // Items
  const items = [];
  // Sample: 상품명 : G-GRIP Pro (Option: Right/Regular) x 1
  const itemMatch = [...t.matchAll(/상품명\s*[:]\s*([^\(]+)(?:\(옵션\s*[:]\s*([^\)]+)\))?\s*x\s*(\d+)/g)];
  itemMatch.forEach(m => {
    const prodName = m[1].trim();
    const optStr = (m[2] || '').trim().toUpperCase();
    const qty = parseInt(m[3]);

    const specs = {
      productName: prodName,
      quantity: qty,
      model: prodName.toUpperCase().includes('PRO') ? 'Pro' : 'Basic',
      clubType: optStr.includes('WOOD') ? 'Wood' : (optStr.includes('IRON') ? 'Iron' : '7-iron'),
      hand: optStr.includes('LEFT') ? 'Left' : 'Right',
      flex: optStr.includes('L-FLEX') ? 'L' : (optStr.includes('S-FLEX') ? 'S' : 'R'),
      length: optStr.includes('LONGER') ? 'Longer' : 'Standard',
      gripSize: optStr.includes('MID') ? 'Mid' : 'Standard',
      magSafeStand: optStr.includes('YES') ? 'Yes' : '0'
    };
    specs.sku = OMS_Utils.deriveSku(specs);
    items.push(specs);
  });

  return {
    orderId,
    purchaseDate,
    buyerEmail,
    shipAddr1,
    shipCity,
    shipState,
    shipPostal,
    shipCountry,
    items,
    totalAmount: 0,
    currency: 'KRW',
    parseStatus,
    parseNotes
  };
}

function inbound_runSamCart() {
  const ss = OMS_Utils.ss();
  validateSchema(ss);
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
        const msgDate = msg.getDate();
        const purchaseTime = Utilities.formatDate(msgDate, OMS_CONFIG.TZ, 'HH:mm');
        const parsed = inbound_parseSamCartInvoice_(clean, purchaseTime);

        // Standardize reshipments as Outbound-only
        if (String(parsed.orderId).toUpperCase().endsWith(OMS_CONFIG.RESHIP_SUFFIX)) return;

        const buyerEmail = OMS_Utils.normalizeEmail_(parsed.buyerEmail);
        const customerId = buyerEmail ? OMS_Utils.lookupOrCreateCustomerId_(buyerEmail) : 'C-UNKNOWN';
        const emailHash = buyerEmail ? OMS_Utils.emailHash_(buyerEmail) : '';

        const sourceSystem = OMS_CONFIG.SOURCE_SYSTEMS.SAMCART;
        const sourceOrderId = parsed.orderId;
        const omsOrderId = OMS_Utils.buildOmsOrderId_(sourceSystem, sourceOrderId);

        const orderCreatedAt = (parsed.purchaseDate && parsed.purchaseTime)
          ? `${parsed.purchaseDate}T${parsed.purchaseTime}:00`
          : Utilities.formatDate(msgDate, OMS_CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss");

        const orderSourceEmail = String(msg.getFrom() || '').match(/<([^>]+)>/)?.[1] || msg.getFrom();

        const now = new Date();
        const stamp = Utilities.formatDate(now, OMS_CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');

        const rows = [];
        const stubData = [];
        parsed.items.forEach((it, idx) => {
          const lineItemIndex = idx + 1;
          const sourceOrderItemId = OMS_Utils.generateLineItemId_(lineItemIndex);
          const omsOrderItemId = OMS_Utils.buildOmsOrderItemId_(omsOrderId, sourceOrderItemId);

          // Derive SKU if not present or placeholder
          if (!it.sku || it.sku === 'SAMCART-UNMAPPED') {
            it.sku = OMS_Utils.deriveSku(it);
          }

          rows.push(inbound_buildRowFromHeaders_(cols, inbound.getLastColumn(), {
            // core ids
            merchantOrderId: sourceOrderId,
            merchantOrderItemId: sourceOrderItemId,
            lineItemIndex,

            purchaseDate: parsed.purchaseDate,
            purchaseTime: parsed.purchaseTime,
            orderCreatedAt,

            buyerEmail,
            buyerName: parsed.buyerName,
            buyerPhone: parsed.buyerPhone,

            customerId,
            systemGmailId: msgId,
            orderSourceEmail,

            salesChannel: 'SamCart',
            customerClassification: 'Active',
            isBusinessOrder: 'false',

            sku: it.sku,
            productName: OMS_Utils.decodeHtmlEntities_(it.productName || parsed.productName),
            magSafeStand: it.magSafeStand,
            model: it.model || parsed.model,
            clubType: it.clubType || parsed.clubType,
            productCategory: 'Golf Club',
            hand: it.hand || parsed.hand,
            flex: it.flex || parsed.flex,
            shaftLengthOption: it.length || parsed.length,
            gripSize: it.gripSize || parsed.gripSize,

            qty: it.quantity || parsed.quantity || 1,
            currency: 'USD',
            itemPrice: it.price || parsed.subtotal,
            itemTax: parsed.tax,
            shippingPrice: parsed.shipping,
            totalAmount: parsed.totalAmount,
            couponCode: parsed.couponCode,
            discountAmount: parsed.discount,
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
            notes: parsed.parseNotes || '',
            automationNotes: '',
            itemLifeCycle: 'ACTIVE',
            orderLifeCycle: (Number(parsed.refundAmount || 0) > 0 ? 'PARTIAL_REFUND' : 'ACTIVE'),
            parseStatus: parsed.parseStatus || 'OK',

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

          stubData.push({
            merchantOrderId: sourceOrderId,
            merchantOrderItemId: sourceOrderItemId,
            sku: it.sku,
            customerId,
            omsOrderId,
            omsOrderItemId,
            magSafeStand: it.magSafeStand,
            orderCreatedAt,
            deliveryCountry: parsed.shipCountry || 'United States',
          });
        });

        if (!rows.length) throw new Error('No items parsed from SamCart invoice.');

        inbound.getRange(inbound.getLastRow() + 1, 1, rows.length, rows[0].length)
          .setNumberFormat('@')
          .setValues(rows);

        // Create Outbound stubs
        outbound_createStubs_(stubData);

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
function inbound_parseSamCartInvoice_(text, purchaseTime) {
  const t = String(text || '');
  const u = t.toUpperCase();
  let parseStatus = 'OK', parseNotes = '';

  // Order ID
  const idMatch = t.match(/Order\s*ID:\s*#?\s*([A-Za-z0-9\-]+)/i) || t.match(/Order\s*#\s*([A-Za-z0-9\-]+)/i);
  let orderId = idMatch ? String(idMatch[1]).trim() : '';
  if (!orderId) {
    orderId = 'UNKNOWN';
    parseStatus = OMS_CONFIG.ERRORS.MISSING_ORDER_ID;
  }

  // Date
  const dateMatch =
    t.match(/(?:Original Order Date|Receipt Date|Date):\s*([A-Z][a-z]+ \d{1,2}(?:st|nd|rd|th)?,? \d{4})/i) ||
    t.match(/Date:\s*([A-Z][a-z]+ \d{1,2}, \d{4})/i);

  const purchaseDate = dateMatch ? normalizeDateYYYYMMDD_(dateMatch[1]) : '';

  // Parse Time from: Mar 2, 2026 at 6:03 AM
  let pTime = purchaseTime || '';
  const timeMatch = t.match(/(?:at\s*)(\d{1,2}:\d{2}\s*(?:AM|PM))/i);
  if (timeMatch) {
    const raw = timeMatch[1].toUpperCase();
    const [h, m] = raw.replace(/\s*(?:AM|PM)/, '').split(':');
    let hour = parseInt(h);
    if (raw.includes('PM') && hour < 12) hour += 12;
    if (raw.includes('AM') && hour === 12) hour = 0;
    pTime = `${String(hour).padStart(2, '0')}:${m}`;
  }

  // Customer name/email
  const cust = t.match(/Customer\s*\n([^\n]+)\n([^\n]+@\S+)/i);
  let buyerName = cust ? String(cust[1] || '').trim() : '';
  let buyerEmail = cust ? String(cust[2] || '').trim() : '';

  if (!buyerEmail) {
    const emails = [...t.matchAll(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g)].map(m => m[1]);
    buyerEmail = emails[0] ? String(emails[0]).trim() : '';
  }
  if (!buyerEmail && parseStatus === 'OK') {
    parseStatus = OMS_CONFIG.ERRORS.MISSING_EMAIL;
  }

  // Phone
  const phoneMatch =
    t.match(/(?:Phone|delivery\.):\s*(\+?\d[\d \-\(\)]{8,15})/i) ||
    t.match(/Customer\s*\n.*\n.*\n(\+?[\d \-\(\)]{8,15})/i);
  const buyerPhone = phoneMatch ? OMS_Utils.normalizePhone(phoneMatch[1]) : '';

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
    if (!addr.success) {
      parseStatus = OMS_CONFIG.ERRORS.MISSING_ADDRESS;
      parseNotes = 'Address parsing fallback to raw block.';
    }
  } else {
    parseStatus = OMS_CONFIG.ERRORS.MISSING_ADDRESS;
  }

  // Specs (same as legacy)
  let model = 'Basic';
  if (/\bPRO\b/i.test(t)) model = 'Pro';
  else if (/\bBASIC\b/i.test(t)) model = 'Basic';

  let clubType = '7-iron';
  if (/\bWOOD\b/i.test(t)) clubType = 'Wood';
  else if (/\bIRON\b/i.test(t)) clubType = 'Iron';
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

  // Parse product line: G-GRIP - Men's 7-Iron Qty: 1 $220.00
  let productName = '', quantity = 1, itemPrice = subtotal;
  const itemMatch = t.match(/(G-GRIP\s*-\s*[^Q\n]+)\s*Qty:\s*(\d+)\s*\$([\d,.]+)/i);
  if (itemMatch) {
    productName = itemMatch[1].trim();
    quantity = parseInt(itemMatch[2]);
    itemPrice = parseFloat(itemMatch[3].replace(/,/g, ''));
  } else {
    quantity = Number((t.match(/Qty:\s*(\d+)/i) || ['', '1'])[1]) || 1;
  }

  // SKU (SamCart often does not have it) → safe placeholder
  let sku = '';
  const skuMatch = t.match(/\bSKU\b\s*[:#]?\s*([A-Za-z0-9\-_.]{4,})/i);
  if (skuMatch) sku = String(skuMatch[1]).trim();
  if (!sku) sku = 'SAMCART-UNMAPPED';

  return {
    orderId,
    purchaseDate,
    purchaseTime: pTime,
    buyerName,
    buyerEmail,
    buyerPhone,

    shipAddr1,
    shipCity,
    shipState,
    shipPostal,
    shipCountry,
    shipServiceLevel: '',

    productName,
    model,
    clubType,

    hand,
    flex,
    length,
    gripSize,
    magSafeStand,

    quantity,
    subtotal,
    shipping,
    tax,
    discount,
    totalAmount,
    couponCode,

    refundAmount,
    refundDate: '',
    returnReasonCode: '',
    recipientName: '',

    items: [{
      sku,
      quantity,
      price: itemPrice,
      productName,
      model,
      clubType,
      hand,
      flex,
      length,
      gripSize,
      magSafeStand,
    }],
    parseStatus,
    parseNotes
  };
}

function inbound_buildRowFromHeaders_(headersMap, lastCol, v) {
  const row = new Array(lastCol).fill('');

  OMS_Utils.setByHeader_(row, headersMap, 'merchant-order-id', v.merchantOrderId);
  OMS_Utils.setByHeader_(row, headersMap, 'merchant-order-item-id', v.merchantOrderItemId);
  OMS_Utils.setByHeader_(row, headersMap, 'line-item-index', v.lineItemIndex);
  OMS_Utils.setByHeader_(row, headersMap, 'purchase-date', v.purchaseDate);
  OMS_Utils.setByHeader_(row, headersMap, 'purchase-time', v.purchaseTime);
  OMS_Utils.setByHeader_(row, headersMap, 'order-created-at', v.orderCreatedAt);

  OMS_Utils.setByHeader_(row, headersMap, 'buyer-email', v.buyerEmail);
  OMS_Utils.setByHeader_(row, headersMap, 'buyer-name', v.buyerName);
  OMS_Utils.setByHeader_(row, headersMap, 'buyer-phone-number', v.buyerPhone);

  OMS_Utils.setByHeader_(row, headersMap, 'customer-id', v.customerId);
  OMS_Utils.setByHeader_(row, headersMap, 'system-gmail-id', v.systemGmailId);
  OMS_Utils.setByHeader_(row, headersMap, 'order-source-email', v.orderSourceEmail);

  OMS_Utils.setByHeader_(row, headersMap, 'sales-channel', v.salesChannel);
  OMS_Utils.setByHeader_(row, headersMap, 'customer-classification', v.customerClassification);
  OMS_Utils.setByHeader_(row, headersMap, 'is-business-order', v.isBusinessOrder);

  OMS_Utils.setByHeader_(row, headersMap, 'sku', v.sku);
  OMS_Utils.setByHeader_(row, headersMap, 'product-name', v.productName);
  OMS_Utils.setByHeader_(row, headersMap, 'mag-safe-stand', v.magSafeStand);
  OMS_Utils.setByHeader_(row, headersMap, 'model', v.model);
  OMS_Utils.setByHeader_(row, headersMap, 'club-type', v.clubType);
  OMS_Utils.setByHeader_(row, headersMap, 'product-category', v.productCategory);
  OMS_Utils.setByHeader_(row, headersMap, 'hand', v.hand);
  OMS_Utils.setByHeader_(row, headersMap, 'flex', v.flex);
  OMS_Utils.setByHeader_(row, headersMap, 'shaft-length-option', v.shaftLengthOption);
  OMS_Utils.setByHeader_(row, headersMap, 'grip-size', v.gripSize);

  OMS_Utils.setByHeader_(row, headersMap, 'quantity-purchased', v.qty);
  OMS_Utils.setByHeader_(row, headersMap, 'currency', v.currency);
  OMS_Utils.setByHeader_(row, headersMap, 'item-price', v.itemPrice);
  OMS_Utils.setByHeader_(row, headersMap, 'item-tax', v.itemTax);
  OMS_Utils.setByHeader_(row, headersMap, 'shipping-price', v.shippingPrice);
  OMS_Utils.setByHeader_(row, headersMap, 'total-amount', v.totalAmount);
  OMS_Utils.setByHeader_(row, headersMap, 'coupon-code', v.couponCode);
  OMS_Utils.setByHeader_(row, headersMap, 'discount-amount', v.discountAmount);

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
  return OMS_Utils.normalizeDateYYYYMMDD(dateStr);
}
