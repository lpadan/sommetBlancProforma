function assumptions() {
    //var data = {};
    //html.data = data;
    var html = HtmlService.createTemplateFromFile('SB.assumptions');
    html = html.evaluate().setWidth(1000).setHeight(750);
    SpreadsheetApp.getUi().showModalDialog(html, 'Assumptions');
}

function include(filename) {

   return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function saveExpenses(expenses) {
   var documentProperties = PropertiesService.getDocumentProperties();
   documentProperties.setProperty('expenses', JSON.stringify(expenses));
}