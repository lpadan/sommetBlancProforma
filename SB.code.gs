function assumptions() {
    //var data = {};
    //html.data = data;
    var html = HtmlService.createTemplateFromFile('SB.assumptions');
    html = html.evaluate().setWidth(800).setHeight(750);
    SpreadsheetApp.getUi().showModalDialog(html, 'Assumptions');
}

function include(filename) {

   return HtmlService.createHtmlOutputFromFile(filename).getContent();
}