/********************************
 * Master.gs
 ********************************/

function refreshMasterOmsTable() {
  const ss = OMS_Utils.ss();
  const inbound = OMS_Utils.sheet_(OMS_CONFIG.TABS.INBOUND);
  const outbound = OMS_Utils.sheet_(OMS_CONFIG.TABS.OUTBOUND);
  const masterTable = OMS_Utils.sheet_(OMS_CONFIG.TABS.MASTER_TABLE);

  const inCols = OMS_Utils.requireCols_(inbound, [
    'oms-order-item-id','oms-order-id','source-system','source-order-id','source-order-item-id',
    'merchant-order-id','merchant-order-item-id','sku','customer-id','buyer-email-hash','buyer-email','buyer-name',
    'purchase-date','sales-channel','item-life-cycle','order-life-cycle','replacement-type','replacement-group-id',
    'quantity-purchased','currency','item-price','item-tax','shipping-price','total-amount','refund-amount','refund-date',
    'serial-number-allocated','notes'
  ]);

  const outCols = OMS_Utils.requireCols_(outbound, [
    'oms-order-item-id','domestic-tracking-kr','hub-received-date','international-tracking-us','carrier-us',
    'outbound-status','serial-number-scanned','sn-verify','customer-email-status','last-email-at','shipment-id',
    'us-ship-date','delivered-date'
  ]);

  // Build outbound map by oms-order-item-id
  const outMap = new Map();
  const outLR = outbound.getLastRow();
  if (outLR >= 2) {
    const outData = outbound.getRange(2, 1, outLR - 1, outbound.getLastColumn()).getValues();
    outData.forEach(row => {
      const key = String(row[outCols['oms-order-item-id'] - 1] || '').trim();
      if (!key) return;
      outMap.set(key, row);
    });
  }

  const headers = OMS_SCHEMA_MASTER_TABLE_();

  masterTable.clear();
  masterTable.getRange(1,1,1,headers.length).setValues([headers]);

  // Style header
  masterTable.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#111827')
    .setFontColor('#FFFFFF')
    .setWrap(true)
    .setVerticalAlignment('middle');

  masterTable.setRowHeight(1, 36);
  masterTable.setFrozenRows(1);
  masterTable.getRange(1, 1, 1, headers.length).createFilter();

  const inLR = inbound.getLastRow();
  if (inLR < 2) return;

  const inData = inbound.getRange(2, 1, inLR - 1, inbound.getLastColumn()).getValues();
  const outRows = [];

  inData.forEach(inRow => {
    const omsItem = String(inRow[inCols['oms-order-item-id'] - 1] || '').trim();
    if (!omsItem) return;

    const o = outMap.get(omsItem);
    const serialScanned = o ? String(o[outCols['serial-number-scanned'] - 1] || '').trim() : '';
    const snVerify = o ? String(o[outCols['sn-verify'] - 1] || '').trim() : '';

    outRows.push([
      inRow[inCols['oms-order-id'] - 1],
      omsItem,
      inRow[inCols['source-system'] - 1],
      inRow[inCols['source-order-id'] - 1],
      inRow[inCols['source-order-item-id'] - 1],
      inRow[inCols['merchant-order-id'] - 1],
      inRow[inCols['merchant-order-item-id'] - 1],
      inRow[inCols['sku'] - 1],
      inRow[inCols['customer-id'] - 1],
      inRow[inCols['buyer-email-hash'] - 1],
      inRow[inCols['buyer-email'] - 1],
      inRow[inCols['buyer-name'] - 1],
      inRow[inCols['purchase-date'] - 1],
      inRow[inCols['sales-channel'] - 1],
      inRow[inCols['item-life-cycle'] - 1],
      inRow[inCols['order-life-cycle'] - 1],
      inRow[inCols['replacement-type'] - 1],
      inRow[inCols['replacement-group-id'] - 1],
      inRow[inCols['quantity-purchased'] - 1],
      inRow[inCols['currency'] - 1],
      inRow[inCols['item-price'] - 1],
      inRow[inCols['item-tax'] - 1],
      inRow[inCols['shipping-price'] - 1],
      inRow[inCols['total-amount'] - 1],
      inRow[inCols['refund-amount'] - 1],
      inRow[inCols['refund-date'] - 1],
      inRow[inCols['serial-number-allocated'] - 1],

      serialScanned,
      snVerify,

      o ? o[outCols['domestic-tracking-kr'] - 1] : '',
      o ? o[outCols['hub-received-date'] - 1] : '',
      o ? o[outCols['international-tracking-us'] - 1] : '',
      o ? o[outCols['carrier-us'] - 1] : '',
      o ? o[outCols['us-ship-date'] - 1] : '',
      o ? o[outCols['delivered-date'] - 1] : '',
      o ? o[outCols['outbound-status'] - 1] : '',
      o ? o[outCols['customer-email-status'] - 1] : '',
      o ? o[outCols['last-email-at'] - 1] : '',
      o ? o[outCols['shipment-id'] - 1] : '',

      inRow[inCols['notes'] - 1],
    ]);
  });

  if (outRows.length) {
    masterTable.getRange(2, 1, outRows.length, headers.length).setValues(outRows);
  }
}
