function assumptions() {
    //var data = {};
    //html.data = data;
    var html = HtmlService.createTemplateFromFile('SB.assumptions');
    html = html.evaluate().setWidth(1000).setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(html, 'Assumptions');
}

function include(filename) {

   return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function saveExpenses(expenses) {

   var land = expenses.land
   var hardCosts = expenses.hardCosts
   var softCosts = expenses.softCosts

   land = ArrayLib.sort(land,0,true);
   hardCosts = ArrayLib.sort(hardCosts,0,true);
   softCosts = ArrayLib.sort(softCosts,0,true);

   expenses.land = land;
   expenses.hardCosts = hardCosts;
   expenses.softCosts = softCosts;

   var documentProperties = PropertiesService.getDocumentProperties();
   documentProperties.setProperty('expenses', JSON.stringify(expenses));
}

function saveRevenue(revenue) {
   var documentProperties = PropertiesService.getDocumentProperties();
   documentProperties.setProperty('revenue', JSON.stringify(revenue));
}

function saveConstruction(construction) {
   var documentProperties = PropertiesService.getDocumentProperties();
   documentProperties.setProperty('construction', JSON.stringify(construction));
}

function saveLoans(loans) {
   var documentProperties = PropertiesService.getDocumentProperties();
   documentProperties.setProperty('loans', JSON.stringify(loans));
}

function addExpense(formData) {
   // formData = {category, description, duration, futureCosts, milestone, paidCosts, startDate, startFinish, type}

   // ADD
   // consider saving properties as objects instead of arrays. Allows to change order or add to easily in the future.

   var documentProperties = PropertiesService.getDocumentProperties();
   var expenses = JSON.parse(documentProperties.getProperty('expenses'));
   var land = expenses.land;
   var hardCosts = expenses.hardCosts;
   var softCosts = expenses.softCosts;

   var category = formData.category;
   if (!category) return;

   var expenseArray = [
      formData.description,
      formData.paidCosts,
      formData.futureCosts,
      formData.type,
      formData.startDate,
      formData.milestone,
      formData.startFinish,
      formData.duration
   ]

   switch(category) {

    case 'land':
        land.push(expenseArray);
        land = ArrayLib.sort(land,0,true);
        expenses.land = land;
        documentProperties.setProperty('expenses',JSON.stringify(expenses));
        return land;
        break;

    case 'hardCosts':
        hardCosts.push(expenseArray);
        hardCosts = ArrayLib.sort(hardCosts,0,true);
        expenses.hardCosts = hardCosts;
        documentProperties.setProperty('expenses',JSON.stringify(expenses));
        return hardCosts;
        break;

    case 'softCosts':
        softCosts.push(expenseArray);
        softCosts = ArrayLib.sort(softCosts,0,true);
        expenses.softCosts = softCosts;
        documentProperties.setProperty('expenses',JSON.stringify(expenses));
        return softCosts;
        break;
    }

    return
}

function getExpenseItems(category) {
    var documentProperties = PropertiesService.getDocumentProperties();
    var expenses = JSON.parse(documentProperties.getProperty('expenses'));
    var data = expenses[category]; // 2D array of expense line items
    var itemDescriptions = data.map(function(elem) {
        return elem[0];
    });
    return itemDescriptions;
}

function deleteExpense(category,item) {
    var documentProperties = PropertiesService.getDocumentProperties();
    var expenses = JSON.parse(documentProperties.getProperty('expenses'));
    var data = expenses[category]; // 2D array of expense line items

    index = ArrayLib.indexOf(data,0,item,false);
    if (index > -1) {
        data.splice(index, 1);
    }
    expenses[category] = data;
    documentProperties.setProperty('expenses',JSON.stringify(expenses));
    return data;
}

var ArrayLib = (function() {
   var arrayLib = {};
    arrayLib.sort = function(data, columnIndex, ascOrDesc) { // true for asc, false for desc
        if (data.length > 0) {
        if (typeof columnIndex != "number" || columnIndex > data[0].length) {
          throw 'Choose a valid column index';
        }
        var r = new Array();
        var areDates = true;
        for (var i = 0; i < data.length; i++) {
          if(data[i] != null){ //
            // this code attempts to convert strings to dates and then compare them
            // not working properly.  "All Tasks 2" is considered a date, which is wrong.
            // var value = data[i][columnIndex];
            // if(value && typeof value == 'string') {
            //   var date = new Date(value);
            //   if (isNaN(date.getFullYear())) areDates = false;
            //   else data[i][columnIndex] = date;
            // }
            r.push(data[i]);
          }
        }
        return r.sort(function (a, b) {
          if (ascOrDesc) return ((a[columnIndex] < b[columnIndex]) ? -1 : ((a[columnIndex] > b[columnIndex]) ? 1 : 0));
          return ((a[columnIndex] > b[columnIndex]) ? -1 : ((a[columnIndex] < b[columnIndex]) ? 1 : 0));
        });
      }
      else {
        return data;
      }
    }

    /* case insensitive

        arrayLib.indexOf = function(data, columnIndex, value) {
            if (data.length > 0) {
                if (typeof columnIndex != "number" || columnIndex > data[0].length) {
                  throw 'Choose a valid column index';
                }
                var r = -1;
                var reg = new RegExp(escape(value).toUpperCase());
                for (var i = 0; i < data.length; i++) {
                  if (data[0][0] == undefined) {
                    if (escape(data[i].toString()).toUpperCase().search(reg) != -1) return i;
                  }
                  else {
                    if (columnIndex < 0 && escape(data[i].toString()).toUpperCase().search(reg) != -1 || columnIndex >= 0 && escape(data[i][columnIndex].toString()).toUpperCase().search(reg) != -1) return i;
                  }
                }
                return r;
            }
            else {
            return data;
          }
        }
    */



    /** case sensitive true or false
      * @param caseSensitive - boolean
    */

    // problem when the "value" is an integrer.  its being converted to a string for comparison
    // test for typeof and add code for integer comparisions

    arrayLib.indexOf = function(data, columnIndex, value, caseSensitive) {

        caseSensitive = (typeof caseSensitive !== 'undefined') ?  caseSensitive : false;
        if (data.length > 0) {

            if (typeof columnIndex != "number" || columnIndex > data[0].length) {
              throw 'Choose a valid column index';
            }
            var r = -1;

            if (caseSensitive) {
                var reg = new RegExp(escape(value));
                for (var i = 0; i < data.length; i++) {
                  if (data[0][0] == undefined) {
                    if (escape(data[i].toString()).search(reg) != -1) return i;
                  }
                  else {
                    if (columnIndex < 0 && escape(data[i].toString()).search(reg) != -1 || columnIndex >= 0 && escape(data[i][columnIndex].toString()).search(reg) != -1) return i;
                  }
                }
            }

            else if (typeof value === 'number') {
                for (var i = 0; i < data.length; i++) {
                    if (data[i][columnIndex] == value) {
                        return i;
                    }
                }
            }

            else {
                var reg = new RegExp(escape(value).toUpperCase());
                for (var i = 0; i < data.length; i++) {
                  if (data[0][0] == undefined) {
                    if (escape(data[i].toString()).toUpperCase().search(reg) != -1) return i;
                  }
                  else {
                    if (columnIndex < 0 && escape(data[i].toString()).toUpperCase().search(reg) != -1 || columnIndex >= 0 && escape(data[i][columnIndex].toString()).toUpperCase().search(reg) != -1) return i;
                  }
                }
            }
            return r;
        }
        else {
        return data;
      }
    }
    return arrayLib;
}) ();
