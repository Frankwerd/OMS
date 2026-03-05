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

  /**
   * Derive SKU from product specs
   * Logic: GG-{Model}-{ClubType}-{Hand}{Flex}-{Length}-{Grip}-{MagSafe}
   */
  deriveSku(specs) {
    const p = specs || {};
    const model = (String(p.model || '').toUpperCase().includes('PRO')) ? 'PRO' : 'BAS';
    const club = (String(p.clubType || '').toUpperCase().includes('WOOD')) ? 'WD' : 'IR';
    const hand = (String(p.hand || '').toUpperCase().startsWith('L')) ? 'L' : 'R';

    // Normalize flex: R, S, L, X (default R)
    let flex = String(p.flex || '').toUpperCase().charAt(0);
    if (!['R','S','L','X'].includes(flex)) flex = 'R';

    const length = (String(p.length || '').toUpperCase().includes('LONG')) ? 'LG' : 'ST';
    const grip = (String(p.gripSize || '').toUpperCase().includes('MID')) ? 'MS' : 'ST';
    const mag = (String(p.magSafeStand || '') === 'Yes' || String(p.magSafeStand || '') === '1' || String(p.magSafeStand || '').toUpperCase() === 'TRUE') ? 'M' : '0';

    return `GG-${model}-${club}-${hand}${flex}-${length}-${grip}-${mag}`;
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
   * Supports US/CA/EU/UK/JP/KR
   ********************************/
  parseGlobalAddress(lines) {
    let d = { addr1: "", city: "", state: "", zip: "", country: "United States", success: false };
    if (!lines || !lines.length) return d;

    const originalBlock = lines.join(', ');
    let working = lines.map(l => String(l || '').trim()).filter(Boolean);
    if (!working.length) return d;

    // 1. Identify country from last line
    const lastLine = working[working.length - 1].toUpperCase();
    const countryMap = {
      'UNITED STATES': 'United States', 'USA': 'United States',
      'CANADA': 'Canada',
      'UNITED KINGDOM': 'United Kingdom', 'UK': 'United Kingdom', 'GREAT BRITAIN': 'United Kingdom',
      'JAPAN': 'Japan',
      'KOREA': 'South Korea', 'SOUTH KOREA': 'South Korea', 'REPUBLIC OF KOREA': 'South Korea', 'ROK': 'South Korea',
      'GERMANY': 'Germany', 'DEUTSCHLAND': 'Germany',
      'FRANCE': 'France',
      'ITALY': 'Italy',
      'SPAIN': 'Spain',
      'AUSTRALIA': 'Australia'
    };

    for (let key in countryMap) {
      if (lastLine === key || lastLine.endsWith(' ' + key)) {
        d.country = countryMap[key];
        working.pop();
        break;
      }
    }

    if (!working.length) {
      d.addr1 = originalBlock;
      return d;
    }

    // 2. Parse Geo (now last line)
    const geo = working.pop();

    // Robust US/Canada Pattern
    // 1. "City, State Zip" or "City State Zip" (2-letter state)
    let usCaMatch = geo.match(/^(.*?)[,\s]+([A-Z]{2})\s+(\d{5}(?:-\d{4})?|[A-Z]\d[A-Z]\s*\d[A-Z]\d)$/i);
    if (!usCaMatch) {
      // 2. Full State Name: "City, StateName Zip"
      usCaMatch = geo.match(/^(.*?)[,\s]+([A-Za-z\s]{3,20})\s+(\d{5}(?:-\d{4})?)$/i);
    }
    if (!usCaMatch) {
      // 3. Reversed Zip: "Zip City State"
      usCaMatch = geo.match(/^(\d{5}(?:-\d{4})?)\s+(.*?)[,\s]+([A-Z]{2})$/i);
      if (usCaMatch) {
        d.zip = usCaMatch[1].trim().toUpperCase();
        d.city = usCaMatch[2].trim();
        d.state = usCaMatch[3].trim().toUpperCase();
        d.success = true;
      }
    }

    if (!d.success && usCaMatch) {
      d.city = usCaMatch[1].trim();
      d.state = this.mapStateToAbbr_(usCaMatch[2].trim());
      d.zip = usCaMatch[3].trim().toUpperCase();
      d.success = true;
    }
    // UK Pattern: "City Postcode"
    else if (d.country === 'United Kingdom') {
      const ukMatch = geo.match(/^(.*?)\s+([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})$/i);
      if (ukMatch) {
        d.city = ukMatch[1].trim();
        d.zip = ukMatch[2].trim().toUpperCase();
        d.success = true;
      } else {
        d.city = geo;
      }
    }
    // JP/KR Pattern: "[Zip] City" or "City Zip" or "Zip City"
    else if (d.country === 'Japan' || d.country === 'South Korea') {
      const eastMatch = geo.match(/^\[?(\d{3,7}[-\s]?\d{0,4})\]?\s*(.*)$/) || geo.match(/^(.*?)\s+(\d{3,7}[-\s]?\d{0,4})$/);
      if (eastMatch) {
        if (isNaN(parseInt(eastMatch[1].charAt(0)))) { // Zip at end
          d.city = eastMatch[1].trim();
          d.zip = eastMatch[2].trim();
        } else { // Zip at start
          d.zip = eastMatch[1].trim();
          d.city = eastMatch[2].trim();
        }
        d.success = (!!d.zip && !!d.city);
      } else {
        d.city = geo;
      }
    }
    // EU Pattern: "Zip City"
    else {
      const euMatch = geo.match(/^(\d{3,7})\s+(.*)$/) || geo.match(/^(.*?)\s+(\d{3,7})$/);
      if (euMatch) {
        if (isNaN(parseInt(euMatch[1].charAt(0)))) { // Zip at end
          d.city = euMatch[1].trim();
          d.zip = euMatch[2].trim();
        } else { // Zip at start
          d.zip = euMatch[1].trim();
          d.city = euMatch[2].trim();
        }
        d.success = true;
      } else {
        d.city = geo;
      }
    }

    // 3. Addr1 is whatever is left
    if (working.length) {
      d.addr1 = working.join(', ');
    } else {
      d.addr1 = originalBlock;
    }

    return d;
  },

  mapStateToAbbr_(state) {
    const s = String(state || '').trim().toUpperCase();
    if (s.length === 2) return s;
    const map = {
      'ALABAMA': 'AL', 'ALASKA': 'AK', 'ARIZONA': 'AZ', 'ARKANSAS': 'AR', 'CALIFORNIA': 'CA', 'COLORADO': 'CO', 'CONNECTICUT': 'CT', 'DELAWARE': 'DE', 'FLORIDA': 'FL', 'GEORGIA': 'GA', 'HAWAII': 'HI', 'IDAHO': 'ID', 'ILLINOIS': 'IL', 'INDIANA': 'IN', 'IOWA': 'IA', 'KANSAS': 'KS', 'KENTUCKY': 'KY', 'LOUISIANA': 'LA', 'MAINE': 'ME', 'MARYLAND': 'MD', 'MASSACHUSETTS': 'MA', 'MICHIGAN': 'MI', 'MINNESOTA': 'MN', 'MISSISSIPPI': 'MS', 'MISSOURI': 'MO', 'MONTANA': 'MT', 'NEBRASKA': 'NE', 'NEVADA': 'NV', 'NEW HAMPSHIRE': 'NH', 'NEW JERSEY': 'NJ', 'NEW MEXICO': 'NM', 'NEW YORK': 'NY', 'NORTH CAROLINA': 'NC', 'NORTH DAKOTA': 'ND', 'OHIO': 'OH', 'OKLAHOMA': 'OK', 'OREGON': 'OR', 'PENNSYLVANIA': 'PA', 'RHODE ISLAND': 'RI', 'SOUTH CAROLINA': 'SC', 'SOUTH DAKOTA': 'SD', 'TENNESSEE': 'TN', 'TEXAS': 'TX', 'UTAH': 'UT', 'VERMONT': 'VT', 'VIRGINIA': 'VA', 'WASHINGTON': 'WA', 'WEST VIRGINIA': 'WV', 'WISCONSIN': 'WI', 'WYOMING': 'WY'
    };
    return map[s] || s;
  },

  normalizeDateYYYYMMDD(dateStr) {
    const s = String(dateStr || '').replace(/(st|nd|rd|th)/g, '').replace(/,/g, '').trim();
    const d = new Date(s);
    if (isNaN(d.getTime())) return '';
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  },
};
