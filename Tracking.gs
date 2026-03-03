/********************************
 * Tracking.gs
 ********************************/

function tracking_linkifyOutboundSelection() {
  const ss = OMS_Utils.ss();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== OMS_CONFIG.TABS.OUTBOUND) {
    throw new Error(`Active sheet must be "${OMS_CONFIG.TABS.OUTBOUND}"`);
  }

  const cols = OMS_Utils.getHeadersMap_(sheet);
  const range = sheet.getActiveRange();
  if (!range || range.getRow() === 1) return;

  const start = range.getRow();
  const end = range.getLastRow();

  for (let r = start; r <= end; r++) {
    // domestic = LOGEN
    const dCol = OMS_Utils.col_(cols, 'domestic-tracking-kr');
    if (dCol) {
      const v = sheet.getRange(r, dCol).getDisplayValue();
      if (v) sheet.getRange(r, dCol).setRichTextValue(OMS_Utils.buildTrackingRichText_('LOGEN', v));
    }

    // international = carrier-us
    const iCol = OMS_Utils.col_(cols, 'international-tracking-us');
    if (iCol) {
      const carrierCol = OMS_Utils.col_(cols, 'carrier-us');
      const carrier = carrierCol ? sheet.getRange(r, carrierCol).getValue() : 'FEDEX';
      const v = sheet.getRange(r, iCol).getDisplayValue();
      if (v) sheet.getRange(r, iCol).setRichTextValue(OMS_Utils.buildTrackingRichText_(String(carrier || 'FEDEX'), v));
    }
  }
}
