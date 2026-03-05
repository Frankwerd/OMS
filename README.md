# OMS Operations Manual
## G·GRIP Order Management System

This system tracks every order from sale → shipment → delivery using three main sheets.

### Workflow Diagram

```mermaid
graph TD
    A[Customer Purchase] --> B[Inbound_Orders]
    B --> C[Outbound_Logistics]
    C --> D[Master_OMS_View]
```

### Process Diagram

```text
Customer Purchase
      ↓
Inbound_Orders (sales record)
      ↓
Outbound_Logistics (shipment tracking)
      ↓
Master_OMS_View / Dashboard (analytics + monitoring)
```

## 1. Inbound_Orders (Sales & Product Information)

### Purpose
This sheet records every item sold. Each row represents one product in an order. If a customer buys 3 clubs in one cart, there will be 3 rows.

### What gets filled automatically
These fields are created automatically by the system:

| Field | Meaning |
| :--- | :--- |
| merchant-order-id | Order number from source system |
| merchant-order-item-id | Unique item ID within the order |
| line-item-index | Position of item in cart |
| purchase-date | Date order was created |
| purchase-time | Time of purchase |
| order-created-at | Exact timestamp (ISO 8601) |
| customer-id | Internal customer identifier (CYYYYMMDD-###) |
| system-gmail-id | Email ID used for deduplication |
| source-system | Platform (samcart, shopify, or imweb) |
| oms-order-id | Internal canonical order ID |
| oms-order-item-id | Internal canonical item ID |
| buyer-email-hash | Privacy-safe SHA-256 email hash |
| parse-status | Parsing result (OK or ERROR) |

### Customer Information
These fields identify the buyer.

| Field | Meaning |
| :--- | :--- |
| buyer-email | Customer email |
| buyer-name | Customer full name |
| buyer-phone-number | Normalized phone number |
| recipient-name | Shipping recipient name |

### Product Details
These describe the club purchased.

| Field | Meaning |
| :--- | :--- |
| sku | Internal product SKU (GG-Model-Club-HandFlex-Length-Grip-Mag) |
| product-name | Product name |
| model | Basic / Pro |
| club-type | Club type (Wood / Iron / 7-iron) |
| product-category | Product classification (default: Golf Club) |
| hand | Right / Left |
| flex | Shaft flex (L, R, S, X) |
| shaft-length-option | Standard / Longer |
| grip-size | Standard / Mid |
| mag-safe-stand | Whether stand is included (Yes / 0) |

### Financial Fields
These represent the transaction values.

| Field | Meaning |
| :--- | :--- |
| currency | USD / KRW |
| item-price | Club price |
| item-tax | Sales tax |
| shipping-price | Shipping charge |
| discount-amount | Discount applied |
| refund-amount | Refunds issued |
| total-amount | Total payment |

### Shipping Address
Customer shipping destination.

| Field | Meaning |
| :--- | :--- |
| ship-address-1 | Street address |
| ship-city | City |
| ship-state | State |
| ship-postal-code | ZIP / Postal Code |
| ship-country | Country |
| ship-service-level | Shipping service level |

### Operational Fields
Used internally for order lifecycle.

| Field | Meaning |
| :--- | :--- |
| serial-number-allocated | Assigned club serial number |
| item-life-cycle | ACTIVE / REFUNDED / RETURNED / REPLACED / CANCELLED |
| order-life-cycle | ACTIVE / PARTIAL_REFUND / FULL_REFUND / CANCELLED |
| replacement-* | Replacement tracking (for ID, item ID, type, group ID) |
| notes | Human notes |
| automation-notes | Script notes |

## 2. Outbound_Logistics (Shipment Tracking)

### Purpose
Tracks how the order moves through the shipping pipeline. Each row represents one shipment for one order item. Outbound rows are automatically created when inbound rows appear.

### Identity Fields
These connect the shipment to the original order.

| Field | Meaning |
| :--- | :--- |
| merchant-order-id | Source system order ID |
| merchant-order-item-id | Line item ID |
| oms-order-id | Internal order ID |
| oms-order-item-id | Internal item ID |
| shipment-id | Unique shipment identifier |

### Workflow Information
| Field | Meaning |
| :--- | :--- |
| outbound-workflow-type | Shipment type (DIRECT_SHIP, RESHIP, etc.) |
| outbound-status | Current stage of shipment |

**Typical lifecycle:**
`CREATED` → `KR_SHIPPED` → `HUB_RECEIVED` → `US_SHIPPED` → `DELIVERED`

### Shipping Stages
These fields record when the package reaches each step.

| Field | Meaning |
| :--- | :--- |
| domestic-tracking-kr | Korean courier tracking |
| hub-received-date | Arrival at export hub |
| hub-location | Hub city (Seoul / Busan / Los Angeles) |
| international-tracking-us | US carrier tracking (FedEx / UPS / USPS / DHL) |
| carrier-us | Carrier name |
| us-ship-date | Shipment departure date |
| delivered-date | Final delivery date |

### Shipment Timeline
| Field | Meaning |
| :--- | :--- |
| stage-timeline | Human-readable event history |

*Example:*
`CREATED — 2026-03-01`
`Hub Received — 2026-03-03`
`US Shipped — 2026-03-04`
`Delivered — 2026-03-06`

### Serial Number Verification
| Field | Meaning |
| :--- | :--- |
| serial-number-scanned | Serial scanned in warehouse |
| sn-verify | OK / MISMATCH / ERROR |
| notes | Issue notes |

### Customer Communication
| Field | Meaning |
| :--- | :--- |
| customer-email-status | Email status (Sent: Final Delivery / Error / SKIP) |
| last-email-at | When customer was notified |

*Note: Emails only trigger once international tracking is available.*

### Package Information
Used for logistics and shipping labels.

| Field | Meaning |
| :--- | :--- |
| package-type | standard-club / club-with-stand |
| actual-weight-kg | Shipping weight |
| package-length-cm | Length |
| package-width-cm | Width |
| package-height-cm | Height |
| delivery-country | Destination |

## 3. Meta / Dashboard (Master_OMS_View)

### Purpose
This sheet is the control center for operations and analytics. It combines information from Inbound and Outbound to show metrics and backlogs.

### Key Operational Metrics

#### Logistics Velocity
Measures shipping speed.

| Metric | Meaning |
| :--- | :--- |
| Avg Time to Hub | Purchase → hub arrival |
| Avg Customs Clearance | Hub → US ship |
| Avg Last Mile | US ship → delivery |
| Avg Click-to-Door | Purchase → delivery |

#### Operational Efficiency
Identifies bottlenecks.

| Metric | Meaning |
| :--- | :--- |
| Hub Backlog | Packages stuck at hub (KR tracking present, US tracking empty) |
| S/N Mismatch Count | Warehouse errors |
| Reshipment Rate | Orders ending in -RES |

#### Product Intelligence (Planned)
| Metric | Meaning |
| :--- | :--- |
| Mag-Safe Attachment Rate | Stand upsell rate |
| Flex Split | R vs S |
| Hand Split | Right vs Left |
| Model Popularity | Basic vs Pro |

#### Customer Metrics
| Metric | Meaning |
| :--- | :--- |
| Repeat Purchase Rate | Returning customers |
| Total LTV | Total spend per customer |
| Top Return Reason | Most frequent return reason |

## Daily Workflow for the Team

### Step 1 — Sales ingestion
Email arrives (SamCart, Shopify, or Imweb). The system automatically adds rows to `Inbound_Orders` and creates stubs in `Outbound_Logistics`. No manual work required.

### Step 2 — Warehouse fulfillment
Warehouse staff allocate a serial number, scan it during packing, and add outbound tracking. This updates `Outbound_Logistics`.

### Step 3 — Shipping updates
Logistics team updates `hub-received-date`, `international-tracking-us`, and `delivered-date`. The `stage-timeline` automatically updates.

### Step 4 — Dashboard monitoring
Operations monitors `Master_OMS_View` to check for backlogs, shipping delays, errors, and product demand.

## Important Rules

### Never edit these fields manually
- `oms-order-id`
- `oms-order-item-id`
- `buyer-email-hash`
- `system-gmail-id`
*They are generated by the system to ensure data integrity.*

### Always check for these alerts
- **S/N Mismatch:** Highlighted in Red.
- **Hub Backlog:** Highlighted in Orange.
- **Missing Data:** Missing addresses or customer IDs are highlighted.
- **Parse Errors:** Rows highlighted in Light Red.

## Summary
The OMS separates responsibilities into three layers:

**Inbound_Orders** (Sales Data)
↓
**Outbound_Logistics** (Shipment Operations)
↓
**Master_OMS_View** (Operations + Analytics)

This structure ensures clean sales data, reliable shipment tracking, and clear operational visibility.
