/********************************
 * Menu.gs
 ********************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ OMS')
    .addItem('Setup Sheet', 'omsSetupSheet')
    .addSeparator()
    .addItem('Refresh Master OMS View', 'refreshMasterOmsView')
    .addToUi();
}
