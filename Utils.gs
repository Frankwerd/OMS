/********************************
 * Utils.gs
 ********************************/

var OMS_Utils = {

  /********************************
   * Spreadsheet Access
   ********************************/
  ss() {
    // Prefer active spreadsheet (bound scripts)
    try {
      const active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) return active;
    } catch (e) {}

    // Fallback for standalone
    const id = String(OMS_CONFIG.SPREADSHEET_ID || '').trim();
    if (!id) {
      throw new Error('No active spreadsheet and OMS_CONFIG.SPREADSHEET_ID is empty.');
    }

    try {
      return SpreadsheetApp.openById(id);
    } catch (err) {
      throw new Error(
        'Failed to open spreadsheet by ID.\n' +
        'Check Spreadsheet ID + permissions.\n' +
        'ID: ' + id + '\n' +
        'Error: ' + err.message
      );
    }
  },

  sheet_(name) {
    const s = this.ss().getSheetByName(name);
    if (!s) throw new Error(`Sheet not found: ${name}`);
    return s;
  },

  /********************************
   * Header Utilities
   ********************************/
  getHeadersMap_(sheet) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h || '').trim().toLowerCase());

    const map = {};
    headers.forEach((h, i) => {
      if (h) map[h] = i + 1;
    });
    return map;
  },

  requireCols_(sheet, requiredHeadersLower) {
    const map = this.getHeadersMap_(sheet);
    requiredHeadersLower.forEach(h => {
      if (!map[h]) {
        throw new Error(`Missing required column "${h}" on "${sheet.getName()}"`);
      }
    });
    return map;
  },

  col_(headersMap, headerLower) {
    return headersMap[String(headerLower).toLowerCase()] || 0;
  },

  setByHeader_(rowArray, headersMap, headerLower, value) {
    const idx = this.col_(headersMap, headerLower);
    if (!idx) return; // optional header missing: ignore
    rowArray[idx - 1] = (value === null || value === undefined) ? '' : value;
  },

  columnLetter_(col) {
    let temp = col;
    let letter = '';
    while (temp > 0) {
      let mod = (temp - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      temp = Math.floor((temp - 1) / 26);
    }
    return letter;
  },

  /********************************
   * Text Utilities
   ********************************/
  ultraCleanText_(html) {
    return String(html || '')
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<head[\s\S]*?<\/head>/gi, '')
      .replace(/<\/p>|<\/div>|<\/tr>|<li>|<br\s*\/?>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/[ \t]+/g, ' ')
      .split('\n')
      .map(l => l.trim())
      .filter(Boolean)
      .join('\n');
  },

  /********************************
   * Email Hashing
   ********************************/
  normalizeEmail_(email) {
    return String(email || '').trim().toLowerCase();
  },

  sha256Hex_(text) {
    const bytes = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      String(text || ''),
      Utilities.Charset.UTF_8
    );
    return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  },

  emailHash_(email) {
    const n = this.normalizeEmail_(email);
    return n ? this.sha256Hex_(n) : '';
  },

  /********************************
   * Customer ID: CYYYYMMDD-###
   ********************************/
  lookupOrCreateCustomerId_(buyerEmail) {
    const email = this.normalizeEmail_(buyerEmail);
    if (!email) throw new Error('Missing buyer-email for customer-id.');

    const inbound = this.sheet_(OMS_CONFIG.TABS.INBOUND);
    const cols = this.requireCols_(inbound, ['buyer-email', 'customer-id']);

    const lr = inbound.getLastRow();
    if (lr >= 2) {
      const emails = inbound.getRange(2, cols['buyer-email'], lr - 1, 1).getValues();
      const cids = inbound.getRange(2, cols['customer-id'], lr - 1, 1).getValues();

      for (let i = emails.length - 1; i >= 0; i--) {
        if (this.normalizeEmail_(emails[i][0]) === email) {
          const existing = String(cids[i][0] || '').trim();
          if (existing) return existing;
        }
      }
    }

    const todayKey = Utilities.formatDate(new Date(), OMS_CONFIG.TZ, 'yyyyMMdd');
    const props = PropertiesService.getScriptProperties();
    const k = `CID_COUNTER_${todayKey}`;
    const n = Number(props.getProperty(k) || '0') + 1;
    props.setProperty(k, String(n));

    return `${OMS_CONFIG.CUSTOMER_ID_PREFIX}${todayKey}-${String(n).padStart(3, '0')}`;
  },

  /********************************
   * Canonical OMS IDs
   ********************************/
  buildOmsOrderId_(sourceSystem, sourceOrderId) {
    const sys = String(sourceSystem || '').trim().toLowerCase() || 'unknown';
    const oid = String(sourceOrderId || '').trim();
    return oid ? `${sys}:${oid}` : '';
  },

  buildOmsOrderItemId_(omsOrderId, sourceOrderItemId) {
    const iid = String(sourceOrderItemId || '').trim();
    return (omsOrderId && iid) ? `${omsOrderId}:${iid}` : '';
  },

  generateLineItemId_(n) {
    return `line-${String(Number(n || 1)).padStart(3, '0')}`;
  },

  /********************************
   * Slack + Gmail
   ********************************/
  slack_(text) {
    if (!OMS_CONFIG.SLACK.ENABLED || !OMS_CONFIG.SLACK.WEBHOOK_URL) return;
    UrlFetchApp.fetch(OMS_CONFIG.SLACK.WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text }),
      muteHttpExceptions: true,
    });
  },

  opsAlert_(text) {
    this.slack_(`🚨 ${OMS_CONFIG.SLACK.OPS_ALERTS_TAG}\n${text}`);
  },

  getOrCreateLabel_(name) {
    return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
  },

  sendCustomerEmail_(to, subject, htmlBody) {
    if (!OMS_CONFIG.EMAIL.ENABLED) return;
    const finalTo = OMS_CONFIG.EMAIL.GLOBAL_LIVE ? to : OMS_CONFIG.EMAIL.TEST_RECIPIENT;
    GmailApp.sendEmail(finalTo, subject, 'Please view as HTML.', {
      htmlBody,
      name: OMS_CONFIG.EMAIL.SENDER_NAME,
    });
  },

  /********************************
   * Tracking Rich Text
   ********************************/
  buildTrackingRichText_(provider, tracking) {
    const cleaned = String(tracking || '').trim().replace(/\s|-/g, '');
    if (!cleaned) return SpreadsheetApp.newRichTextValue().setText('').build();

    const p = String(provider || '').trim().toUpperCase();
    const pattern = OMS_CONFIG.TRACKING_URLS[p] || OMS_CONFIG.TRACKING_URLS.OTHER;

    const b = SpreadsheetApp.newRichTextValue().setText(cleaned);

    if (pattern) {
      b.setLinkUrl(
        pattern.includes('{{TRACKING_NUMBER}}')
          ? pattern.replace('{{TRACKING_NUMBER}}', encodeURIComponent(cleaned))
          : pattern
      );
    }

    return b.build();
  },

  /********************************
   * Address Parsing (Robust)
   ********************************/
  parseGlobalAddress(lines) {
    let d = { addr1: "", city: "", state: "", zip: "", country: "United States" };
    if (!lines || !lines.length) return d;

    let working = lines.map(l => String(l || '').trim()).filter(Boolean);
    if (!working.length) return d;

    // Last line = geo
    const geo = working.pop();

    const usMatch = geo.match(/^(.*?),?\s*([A-Z]{2})\s*(\d{5}(?:-\d{4})?)$/);
    if (usMatch) {
      d.city = usMatch[1].trim();
      d.state = usMatch[2].trim();
      d.zip = usMatch[3].trim();
    } else {
      d.city = geo;
    }

    if (working.length) {
      d.addr1 = working.join(', ');
    }

    return d;
  },
};
