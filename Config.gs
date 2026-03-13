/********************************
 * Config.gs
 ********************************/
const OMS_CONFIG = {
  // Standalone script: REQUIRED
  SPREADSHEET_ID: '1I7l8LrFjdNQePw5gcVBiSoKAm3RdX9wqmUbHVFV4tuA',

  TZ: Session.getScriptTimeZone(),

  TABS: {
    INBOUND: 'Inbound_Orders',
    OUTBOUND: 'Outbound_Logistics',
    MASTER: 'Master_OMS_View',
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
