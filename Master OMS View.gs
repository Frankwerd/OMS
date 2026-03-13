/*******************************
 * Master OMS View (place anywhere)
 *******************************/

function refreshMasterOmsView() {
  const ss = OMS_Utils.ss();
  const inbound = OMS_Utils.sheet_(OMS_CONFIG.TABS.INBOUND);
  const outbound = OMS_Utils.sheet_(OMS_CONFIG.TABS.OUTBOUND);

  let master = ss.getSheetByName(OMS_CONFIG.TABS.MASTER);
  if (!master) master = ss.insertSheet(OMS_CONFIG.TABS.MASTER);
  master.clearContents();

  const inCols = OMS_Utils.getHeadersMap_(inbound);
  const outCols = OMS_Utils.getHeadersMap_(outbound);

  const inData = inbound.getLastRow() >= 2
    ? inbound.getRange(1, 1, inbound.getLastRow(), inbound.getLastColumn()).getValues()
    : [];
  const outData = outbound.getLastRow() >= 2
    ? outbound.getRange(1, 1, outbound.getLastRow(), outbound.getLastColumn()).getValues()
    : [];

  const headers = [
    'merchant-order-id','sku','customer-id','buyer-email','full-name',
    'mag-safe-stand','flex','hand','quantity','order-value-usd','order-date',
    'serial-number-allocated','serial-number-scanned','sn-verify',
    'domestic-tracking-kr','hub-received-date','international-tracking-us','carrier-us','status-email','last-email-at',
    'inbound-parse-status','inbound-notes','outbound-status','outbound-notes',
  ];
  master.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  master.setFrozenRows(1);

  if (!inData.length) return;

  // Build outbound map by (orderId|sku)
  const outMap = new Map();
  if (outData.length >= 2) {
    const oh = outData[0].map(h => String(h || '').toLowerCase().trim());
    const idx = {};
    oh.forEach((h, i) => idx[h] = i);

    for (let r = 1; r < outData.length; r++) {
      const row = outData[r];
      const oid = String(row[idx['merchant-order-id']] || '').trim();
      const sku = String(row[idx['sku']] || '').trim();
      if (!oid || !sku) continue;
      outMap.set(`${oid}|${sku}`, row);
    }
  }

  // Inbound rows → join outbound
  const ih = inData[0].map(h => String(h || '').toLowerCase().trim());
  const iidx = {};
  ih.forEach((h, i) => iidx[h] = i);

  const outHeadersLower = outData.length ? outData[0].map(h => String(h || '').toLowerCase().trim()) : [];
  const oidx = {};
  outHeadersLower.forEach((h, i) => oidx[h] = i);

  const outRows = [];
  for (let r = 1; r < inData.length; r++) {
    const row = inData[r];
    const oid = String(row[iidx['merchant-order-id']] || '').trim();
    const sku = String(row[iidx['sku']] || '').trim();
    if (!oid || !sku) continue;

    const joined = outMap.get(`${oid}|${sku}`);

    const serialAllocated = String(row[iidx['serial-number-allocated']] || '').trim();
    const serialScanned = joined ? String(joined[oidx['serial-number-scanned']] || '').trim() : '';
    const snVerify = joined ? String(joined[oidx['sn-verify']] || '').trim() : (serialAllocated && serialScanned ? (serialAllocated === serialScanned ? 'OK' : 'MISMATCH') : '');

    outRows.push([
      oid,
      sku,
      row[iidx['customer-id']] || '',
      row[iidx['buyer-email']] || '',
      row[iidx['full-name']] || '',
      row[iidx['mag-safe-stand']] || '',
      row[iidx['flex']] || '',
      row[iidx['hand']] || '',
      row[iidx['quantity']] || '',
      row[iidx['order-value-usd']] || '',
      row[iidx['order-date']] || '',
      serialAllocated,
      serialScanned,
      snVerify,
      joined ? joined[oidx['domestic-tracking-kr']] || '' : '',
      joined ? joined[oidx['hub-received-date']] || '' : '',
      joined ? joined[oidx['international-tracking-us']] || '' : '',
      joined ? joined[oidx['carrier-us']] || '' : '',
      joined ? joined[oidx['status-email']] || '' : '',
      joined ? joined[oidx['last-email-at']] || '' : '',
      row[iidx['parse-status']] || '',
      row[iidx['notes']] || '',
      joined ? joined[oidx['outbound-status']] || '' : '',
      joined ? joined[oidx['notes']] || '' : '',
    ]);
  }

  if (outRows.length) {
    master.getRange(2, 1, outRows.length, headers.length).setValues(outRows);
  }

  // Optional visual flagging for mismatches
  const snVerifyCol = headers.indexOf('sn-verify') + 1;
  if (outRows.length) {
    const rng = master.getRange(2, snVerifyCol, outRows.length, 1);
    const vals = rng.getValues();
    const bgs = vals.map(v => {
      const s = String(v[0] || '').toUpperCase();
      if (s === 'MISMATCH' || s.startsWith('ERROR')) return ['#f4cccc'];
      if (s === 'OK') return ['#d9ead3'];
      return [null];
    });
    rng.setBackgrounds(bgs);
  }
}
