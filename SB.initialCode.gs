function onOpen(e) {
	//container bound code
	ui = SpreadsheetApp.getUi();
	ui.createMenu('Schedule Maker')
	.addItem("Open Sidebar",'openSidebar')
	.addToUi();
     ui.createMenu('Proforma')
     .addItem('Assumptions','assumptions')
    .addItem('Update Proforma', 'updateProforma')
    .addToUi();
}