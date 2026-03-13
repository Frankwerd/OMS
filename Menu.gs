/********************************
 * Menu.gs
 ********************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ OMS')
    .addItem('Setup Sheet', 'omsSetupSheet')
    .addSeparator()
    .addItem('Refresh Dashboard', 'omsRefreshDashboard')
    .addItem('Refresh Master Table', 'refreshMasterOmsTable')
    .addSeparator()
    .addItem('Migrate Legacy Inbound', 'migrateInboundLegacyToNew')
    .addItem('Migrate Legacy Outbound', 'migrateOutboundLegacyToNew')
    .addToUi();
}
