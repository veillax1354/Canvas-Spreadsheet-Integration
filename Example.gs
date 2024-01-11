function refresh() {
  var accessToken = "token"
  var namespace = "namespace"
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  CSI.run(accessToken, spreadsheet, namespace);
}
