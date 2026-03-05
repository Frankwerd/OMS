/********************************
 * Config.gs
 ********************************/
const OMS_CONFIG = {
  // Standalone script: REQUIRED
  SPREADSHEET_ID: '1I7l8LrFjdNQePw5gcVBiSoKAm3RdX9wqmUbHVFV4tuA',

  SCHEMA_VERSION: '1.4.0',

  TZ: Session.getScriptTimeZone(),

  PACKAGE_DEFAULTS: {
    WEIGHT_KG: 1.5,
    WIDTH_CM: 106,
    LENGTH_CM: 16.5,
    HEIGHT_CM: 8.5,
  },

  STAND_DEFAULTS: {
    WEIGHT_KG: 2.2,
    WIDTH_CM: 108,
    LENGTH_CM: 20,
    HEIGHT_CM: 15,
  },

  ERRORS: {
    MISSING_ORDER_ID: 'ERR_NO_ORDER_ID',
    MISSING_EMAIL: 'ERR_NO_EMAIL',
    MISSING_ADDRESS: 'ERR_NO_ADDRESS',
    AMBIGUOUS_ITEMS: 'ERR_AMBIGUOUS_ITEMS',
    MISSING_SKU_MAPPING: 'ERR_MISSING_SKU',
  },

  TABS: {
    INBOUND: 'Inbound_Orders',
    OUTBOUND: 'Outbound_Logistics',
    DASHBOARD: 'Master_OMS_Dashboard',
    MASTER_TABLE: 'Master_OMS_Table',
    META: '_Meta',
  },

  CUSTOMER_ID_PREFIX: 'C',
  RESHIP_SUFFIX: '-RES',

  // Gmail labels
  GMAIL: {
    SAMCART_TO_PROCESS: 'OMS/Inbound/SamCart/To Process',
    SAMCART_PROCESSED: 'OMS/Inbound/SamCart/Processed',
    SAMCART_ERROR: 'OMS/Inbound/SamCart/Error',

    SHOPIFY_TO_PROCESS: 'OMS/Inbound/Shopify/To Process',
    SHOPIFY_PROCESSED: 'OMS/Inbound/Shopify/Processed',
    SHOPIFY_ERROR: 'OMS/Inbound/Shopify/Error',

    IMWEB_TO_PROCESS: 'OMS/Inbound/Imweb/To Process',
    IMWEB_PROCESSED: 'OMS/Inbound/Imweb/Processed',
    IMWEB_ERROR: 'OMS/Inbound/Imweb/Error',
  },

  SLACK: {
    ENABLED: true,
    WEBHOOK_URL: '',        // paste webhook
    OPS_ALERTS_TAG: '#ops-alerts',
  },

  EMAIL: {
    ENABLED: true,
    SENDER_NAME: 'The G·GRIP Team',
    GLOBAL_LIVE: true,
    TEST_RECIPIENT: 'test@example.com',
  },

  // Canonical IDs
  SOURCE_SYSTEMS: {
    SAMCART: 'samcart',
    SHOPIFY: 'shopify',
    IMWEB: 'imweb',
  },

  // Provider URLs (rich links)
  TRACKING_URLS: {
    FEDEX: 'https://www.fedex.com/fedextrack/?trknbr={{TRACKING_NUMBER}}',
    UPS: 'https://www.ups.com/track?tracknum={{TRACKING_NUMBER}}',
    USPS: 'https://tools.usps.com/go/TrackConfirmAction?tLabels={{TRACKING_NUMBER}}',
    DHL: 'https://www.dhl.com/global-en/home/tracking/tracking-express.html?submit=1&tracking-id={{TRACKING_NUMBER}}',
    LOGEN: 'https://www.ilogen.com/web/personal/tkSearch.jsp?slipno={{TRACKING_NUMBER}}',
    ECMS: 'https://www.ecmsglobal.com/track?tracking_number={{TRACKING_NUMBER}}',
    EMS: 'http://www.emspremium.com/tracking/',
    OTHER: '',
  },
};
