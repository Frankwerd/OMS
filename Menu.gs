/********************************
 * Menu.gs
 ********************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ OMS')
    .addItem('Setup Sheet', 'omsSetupSheet')
    .addSeparator()
    .addItem('Refresh Dashboard', 'omsRefreshDashboard')
    .addItem('Refresh Master OMS View', 'refreshMasterOmsView')
    .addToUi();
}
