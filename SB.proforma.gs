function updateProforma() {

    var arrayTemplate=[],headers,i,index,j,label,month,projectDuration,year,range,rowData=[];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();

    if (!ss.getSheetByName('Proforma')) {
        var response = ui.alert('Could not find a sheet named "Proforma". Create Sheet?', ui.ButtonSet.YES_NO);
        if (response == ui.Button.YES) {
            ss.insertSheet('Proforma');
            sheet = ss.getSheetByName('Proforma');
        } else {
          return;
        }
    }

    // assumptions
        var sheet = ss.getSheetByName('Assumptions');
        if (!sheet) {
            ui.alert('Could not find the "Assumptions" sheet');
            return;
        }
        var numRows = sheet.getLastRow();
        var assumptions = sheet.getRange(1,2,numRows,8).getValues();

        // start date
            var projectStartDate = sheet.getRange("C1").getValue();
            var projectStartDate = new Date(projectStartDate);
            month = projectStartDate.getMonth();
            year = projectStartDate.getFullYear();
            var lookup = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
            var headerDate = lookup[month] + ' ' + year;

        // sales
            index = ArrayLib.indexOf(assumptions, 0, "Total Condo Sales");
            var grossSales = assumptions[index][1];
            index = ArrayLib.indexOf(assumptions, 0, "Condo Presales");
            var preSales = assumptions[index][1];
            index = ArrayLib.indexOf(assumptions, 0, "Condo Sales / Month");
            var salesPerMonth = assumptions[index][1];
            index = ArrayLib.indexOf(assumptions, 0, "Listing Commission");
            var listingCommission = assumptions[index][1];
            index = ArrayLib.indexOf(assumptions, 0, "Selling Commission");
            var sellingCommission = assumptions[index][1];

        // earnest money
            index = ArrayLib.indexOf(assumptions, 0, "First EM Deposit");
            var firstEMPercent = assumptions[index][1];
            var firstEMDate = assumptions[index][2];
            var firstEMDepositMonth = monthDiff(projectStartDate,new Date(firstEMDate));

            index = ArrayLib.indexOf(assumptions, 0, "Second EM Deposit");
            var secondEMPercent = assumptions[index][1];
            var secondEMDepositTask = assumptions[index][3];

            index = ArrayLib.indexOf(assumptions, 0, "Third EM Deposit");
            var thirdEMPercent = assumptions[index][1];
            var thirdEMDepositTask = assumptions[index][3];

        // condo construction
            index = ArrayLib.indexOf(assumptions, 0, "Construction Cost");
            var constCost = assumptions[index][1];
            index = ArrayLib.indexOf(assumptions, 0, "Construction Management");
            var constMgmtPercent = assumptions[index][1];
            index = ArrayLib.indexOf(assumptions, 0, "Contingency");
            var contingencyPercent = assumptions[index][1];
            index = ArrayLib.indexOf(assumptions, 0, "Schedule Tab Name");
            var scheduleTabName = assumptions[index][1];

    // project duration
        var earliestStartMilliseconds,earliestFinishMilliseconds,earliestFinish,eariestStart;
        var sheet = ss.getSheetByName(scheduleTabName);
        if (!sheet) {
            ui.alert('Could not find a schedule sheet named "' + scheduleTabName + '"');
            return;
        }
        var scheduleData = sheet.getDataRange().getValues();
        headers = scheduleData.shift();
        var minEarliestStart; // milliseconds
        var maxEarliestFinish = 0; // milliseconds

        for (i = 0; i < scheduleData.length; i++) {
            var earliestStartIndex = headers.indexOf('Earliest Start');
            earliestStart = new Date(scheduleData[i][earliestStartIndex]);
            var earliestFinishIndex = headers.indexOf('Earliest Finish');
            earliestFinish = new Date(scheduleData[i][earliestFinishIndex]);
            earliestStartMilliseconds = earliestStart.getTime();
            earliestFinishMilliseconds = earliestFinish.getTime();
            if (earliestFinishMilliseconds > maxEarliestFinish) maxEarliestFinish = earliestFinishMilliseconds;
            if (i == 0) minEarliestStart = earliestStartMilliseconds;
            if (earliestStartMilliseconds < minEarliestStart) minEarliestStart = earliestStartMilliseconds;
        }
        var endDate = new Date(maxEarliestFinish);
        projectDuration = monthDiff(projectStartDate,endDate) + 2;
        var constDuration = monthDiff(new Date(minEarliestStart), new Date(maxEarliestFinish));

        for (i = 0; i < projectDuration; i++) {
            arrayTemplate.push(null);
        }

        var totalExpense = [...arrayTemplate];
        var totalHistoricalExpense = 0;

    // header
        var header = [];
        lookup = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        month = Number(projectStartDate.getMonth());
        year = Number(projectStartDate.getFullYear());

        for (i = 0; i < projectDuration; i ++) {
            if (month > 11) {
                month = 0;
                year++;
            }
            format_month = lookup[month];
            month++;
            header[i] = (format_month + ' ' + year);
        }
        header.unshift('Historical');
        header.unshift('Total');
        header.unshift('');

    // condo const cost
        var condoConstCost = [...arrayTemplate];
        var constMgmt = [...arrayTemplate];
        var contingency = [...arrayTemplate];

        for (i = 0; i < scheduleData.length; i++) {

            if (scheduleData[i][1] == secondEMDepositTask) {
                var secondEMDepositMonth = monthDiff(projectStartDate,new Date(scheduleData[i][earliestStartIndex]));
            }

            if (scheduleData[i][1] == thirdEMDepositTask) {
                var thirdEMDepositMonth = monthDiff(projectStartDate,new Date(scheduleData[i][earliestFinishIndex]));
            }

            var startDate = new Date(scheduleData[i][earliestStartIndex]);
            var endDate = new Date(scheduleData[i][earliestFinishIndex]);

            var startMonth = monthDiff(projectStartDate,startDate);
            var endMonth = monthDiff(projectStartDate,endDate);
            var durationMonths = endMonth - startMonth + 1;

            var percentOfCostIndex = headers.indexOf("% of Cost");
            if (percentOfCostIndex == -1) {
                ui.alert('Could not find column "% of Cost" on the "' + scheduleTabName + '" tab');
                return;
            }

            var totalCost = constCost * scheduleData[i][percentOfCostIndex]/100;
            var costPerMonth = totalCost / durationMonths;

            for (j = 0; j < durationMonths; j++) {
                condoConstCost[startMonth + j + 1] += costPerMonth;
            }
        }

        var condoConstCostTotal = 0;
        var constMgmtTotal = 0;
        var contingencyTotal = 0;

        for (i = 0; i < projectDuration; i++) {

            condoConstCostTotal += condoConstCost[i];
            if (!condoConstCost[i]) condoConstCost[i] = null;

            constMgmtTotal += condoConstCost[i] * constMgmtPercent;
            contingencyTotal += condoConstCost[i] * contingencyPercent;

            if (condoConstCost[i]) {
                constMgmt[i] = condoConstCost[i] * constMgmtPercent;
                contingency[i] = condoConstCost[i] * contingencyPercent;
            } else {
                constMgmt[i] = null;
                contingency[i] = null;
            }
        }

        for (i = 0; i < projectDuration; i ++) {
            totalExpense[i] += (condoConstCost[i] + constMgmt[i] + contingency[i]);
        }

        condoConstCost.unshift(null);
        condoConstCost.unshift(condoConstCostTotal);
        condoConstCost.unshift('Construction');

        constMgmt.unshift(null);
        constMgmt.unshift(constMgmtTotal);
        constMgmt.unshift('Mgmt');

        contingency.unshift(null);
        contingency.unshift(contingencyTotal);
        contingency.unshift('Contingency');

    // revenue (must come after condo const cost for EM triggers)

        // clone arrays
        var closingRevenue = [...arrayTemplate];
        var upfrontCommissions = [...arrayTemplate];
        var cummulativeNetRevenue = [...arrayTemplate];
        var cummulativeSales = [...arrayTemplate];
        var firstEMDeposit = [...arrayTemplate];
        var secondEMDeposit = [...arrayTemplate];
        var thirdEMDeposit = [...arrayTemplate];
        var initialSales = [...arrayTemplate];
        var balanceOfCommissions = [...arrayTemplate];
        var netRevenue = [...arrayTemplate];
        var totalSales = [...arrayTemplate];

        var sales = [];
        var upfrontCommissionsTotal = 0;
        var remainingSales = grossSales - preSales;
        var salesTotal = 0;
        var firstEMDepositTotal = 0;
        var secondEMDepositTotal = 0;
        var thirdEMDepositTotal = 0;
        var netRevenueTotal = 0;
        var totalSalesTotal = 0;

        for (i = 0; i < projectDuration; i++) {
            if (remainingSales > salesPerMonth) {
                sales[i] = salesPerMonth;
                remainingSales -= salesPerMonth;
            } else {
                sales[i] = remainingSales;
                remainingSales = null;
            }
            initialSales[i] = null;

            if (i == projectDuration - 1) {
                if (remainingSales) {
                    sales[i] += remainingSales;
                }
            }
            totalSales[i] = sales[i];
            totalSalesTotal += totalSales[i];

            salesTotal += sales[i];
            if (i == 0) cummulativeSales[0] = sales[0];
            else {
                if (cummulativeSales[i-1]) {
                    cummulativeSales[i] = cummulativeSales[i-1] + sales[i];
                } else {
                    cummulativeSales[i] = sales[i];
                }
            }

            // earnest money
                if (i != projectDuration -1) {

                    var preSalesFirstEM = preSales * firstEMPercent;
                    var preSalesSecondEM = preSales * secondEMPercent;
                    var preSalesThirdEM = preSales * thirdEMPercent;

                    if (i < firstEMDepositMonth) {
                        firstEMDeposit[i] = null;
                    }

                    if (i == firstEMDepositMonth) {
                        firstEMDeposit[i] += preSalesFirstEM;
                        firstEMDeposit[i] += cummulativeSales[i] * firstEMPercent;
                        upfrontCommissions[i] += -(preSales + cummulativeSales[i]) * sellingCommission/2;
                    }

                    if (i > firstEMDepositMonth && i < secondEMDepositMonth) {
                        firstEMDeposit[i] += sales[i] * firstEMPercent;
                        upfrontCommissions[i] += - sales[i] * sellingCommission/2;
                    }

                    if (i == secondEMDepositMonth) {
                        firstEMDeposit[i] += sales[i] * firstEMPercent;
                        secondEMDeposit[i] += preSalesSecondEM;
                        secondEMDeposit[i] += cummulativeSales[i] * secondEMPercent;
                        upfrontCommissions[i] += - sales[i] * sellingCommission/2;
                    }

                    if (i > secondEMDepositMonth && i < thirdEMDepositMonth) {
                        firstEMDeposit[i] += sales[i] * firstEMPercent;
                        secondEMDeposit[i] += sales[i] * secondEMPercent;
                        upfrontCommissions[i] += - sales[i] * sellingCommission/2;
                    }

                    if (i == thirdEMDepositMonth) {
                        thirdEMDeposit[i] += preSalesThirdEM;
                        firstEMDeposit[i] += sales[i] * firstEMPercent;
                        secondEMDeposit[i] += sales[i] * secondEMPercent;
                        thirdEMDeposit[i] += cummulativeSales[i] * thirdEMPercent;
                        upfrontCommissions[i] += - sales[i] * sellingCommission/2;
                    }

                    if (i > thirdEMDepositMonth) {
                        if (sales[i]) {
                            firstEMDeposit[i] += sales[i] * firstEMPercent;
                            secondEMDeposit[i] += sales[i] * secondEMPercent;
                            thirdEMDeposit[i] += sales[i] * thirdEMPercent;
                            upfrontCommissions[i] += - sales[i] * sellingCommission/2;
                        } else {
                            firstEMDeposit[i] = null;
                            secondEMDeposit[i] = null;
                            thirdEMDeposit[i] = null;
                            upfrontCommissions[i] = null;
                        }
                    }

                    firstEMDepositTotal += firstEMDeposit[i];
                    secondEMDepositTotal += secondEMDeposit[i];
                    thirdEMDepositTotal += thirdEMDeposit[i];
                    upfrontCommissionsTotal += upfrontCommissions[i];
                    netRevenue[i] = closingRevenue[i] + firstEMDeposit[i] + secondEMDeposit[i] + thirdEMDeposit[i] + upfrontCommissions[i] + balanceOfCommissions[i];
                    if (netRevenue[i] == 0) netRevenue[i] = null;
                    netRevenueTotal += netRevenue[i];

                } else if (i == projectDuration -1) {
                    closingRevenue[i] = (preSales + salesTotal - firstEMDepositTotal - secondEMDepositTotal - thirdEMDepositTotal);
                    balanceOfCommissions[i] = -(totalSalesTotal + preSales) * (sellingCommission + listingCommission) - upfrontCommissionsTotal;
                    netRevenue[i] = closingRevenue[i] + balanceOfCommissions[i];
                    netRevenueTotal += netRevenue[i];
                    balanceOfCommissionsTotal = balanceOfCommissions[i];
                }
        }

    // add label and totals

        initialSales.unshift(preSales);
        initialSales.unshift(preSales);
        initialSales.unshift("Presales");

        sales.unshift('');
        sales.unshift(salesTotal);
        sales.unshift("Sales");

        totalSales.unshift(preSales);
        totalSales.unshift(totalSalesTotal + preSales);
        totalSales.unshift('Total Sales');

        firstEMDeposit.unshift('');
        firstEMDeposit.unshift(firstEMDepositTotal);
        firstEMDeposit.unshift("First EM Deposit");

        secondEMDeposit.unshift('');
        secondEMDeposit.unshift(secondEMDepositTotal);
        secondEMDeposit.unshift("Second EM Deposit");

        thirdEMDeposit.unshift('');
        thirdEMDeposit.unshift(thirdEMDepositTotal);
        thirdEMDeposit.unshift('Third EM Deposit');

        closingRevenue.unshift('');
        closingRevenue.unshift(preSales + salesTotal - firstEMDepositTotal - secondEMDepositTotal - thirdEMDepositTotal);
        closingRevenue.unshift('Closing Revenue');

        upfrontCommissions.unshift('');
        upfrontCommissions.unshift(upfrontCommissionsTotal);
        upfrontCommissions.unshift('Upfront Commissions');

        balanceOfCommissions.unshift('');
        //balanceOfCommissions.unshift(- (totalSalesTotal + preSales) * (listingCommission + sellingCommission/2));
        balanceOfCommissions.unshift(balanceOfCommissionsTotal);
        balanceOfCommissions.unshift('Balance Of Commissions');

        netRevenue.unshift('');
        netRevenue.unshift(netRevenueTotal);
        netRevenue.unshift('Net Revenue');

    // display revenue
        sheet = ss.getSheetByName('Proforma');
        sheet.clear();

        var range = sheet.getRange("A1:B2");
        range.mergeAcross();

        range = sheet.getRange("A2:B2");
        range.mergeAcross();

        sheet.getRange("A1").setFontSize(14).setValue('Sommet Blanc Proforma');
        sheet.getRange('A2').setFontSize(12).setHorizontalAlignment('left').setValue(headerDate);

        sheet.setRowHeight(4,25);
        sheet.getRange(4,2,1,projectDuration + 3).setNumberFormat("#").setFontSize(11).setBackground('#256ab0').setFontColor('white').setHorizontalAlignment('center').setVerticalAlignment('middle').setValues([header]);
        sheet.getRange(4,1).setBackground('#256ab0').setFontColor('white');
        sheet.getRange('B6').setFontSize(11).setFontWeight('bold').setNumberFormat('@').setValue('SALES');
        sheet.getRange(7,2,1,projectDuration + 3).setNumberFormat("#,##0").setFontSize(10).setHorizontalAlignment('right').setValues([initialSales]);
        sheet.getRange('B7').setHorizontalAlignment('left').setNumberFormat('  @');
        sheet.getRange(8,2,1,projectDuration + 3).setNumberFormat("#,##0").setBorder(false,false,true,false,false,false).setFontSize(10).setHorizontalAlignment('right').setValues([sales]);
        sheet.getRange('B8').setHorizontalAlignment('left').setNumberFormat('  @').setBorder(false,false,false,false,false,false);

        sheet.getRange(9,2,1,projectDuration + 3).setNumberFormat("#,##0").setFontSize(10).setHorizontalAlignment('right').setFontWeight('').setValues([totalSales]);
        sheet.getRange('B9').setFontSize(11).setHorizontalAlignment('left').setFontWeight('bold');

        sheet.getRange('B11').setFontSize(11).setHorizontalAlignment('left').setFontWeight('bold').setValue('REVENUE');
        sheet.getRange('B12').setFontSize(11).setHorizontalAlignment('left').setFontWeight('bold').setValue('  Condos');

        sheet.getRange(13,2,1,projectDuration + 3).setNumberFormat("#,##0").setFontSize(10).setHorizontalAlignment('right').setValues([closingRevenue]);
        sheet.getRange('B13').setHorizontalAlignment('left').setNumberFormat('    @');
        sheet.getRange(14,2,1,projectDuration + 3).setNumberFormat("#,##0").setFontSize(10).setHorizontalAlignment('right').setValues([firstEMDeposit]);
        sheet.getRange('B14').setHorizontalAlignment('left').setNumberFormat('    @');
        sheet.getRange(15,2,1,projectDuration + 3).setNumberFormat("#,##0").setFontSize(10).setHorizontalAlignment('right').setValues([secondEMDeposit]);
        sheet.getRange('B15').setHorizontalAlignment('left').setNumberFormat('    @');
        sheet.getRange(16,2,1,projectDuration + 3).setNumberFormat("#,##0").setFontSize(10).setHorizontalAlignment('right').setValues([thirdEMDeposit]);
        sheet.getRange('B16').setHorizontalAlignment('left').setNumberFormat('    @');
        sheet.getRange(17,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setValues([upfrontCommissions]);
        sheet.getRange('B17').setHorizontalAlignment('left').setNumberFormat('    @');
        sheet.getRange(18,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setValues([balanceOfCommissions]);
        sheet.getRange('B18').setHorizontalAlignment('left').setNumberFormat('    @');

        sheet.setRowHeight(20,25);
        sheet.getRange(20,2,1,projectDuration + 3).setNumberFormat("#,##0").setFontSize(10).setHorizontalAlignment('right').setBackground('#666').setFontColor('white').setVerticalAlignment('middle').setValues([netRevenue]);
        sheet.getRange('B20').setFontSize(11).setHorizontalAlignment('left');
        sheet.getRange(20,1).setBackground('#666').setFontColor('white');

        sheet.getRange('B22').setFontSize(11).setHorizontalAlignment('left').setFontWeight('bold').setNumberFormat('@').setValue('EXPENSE');

    // expense

        var expenseRowNum = ArrayLib.indexOf(assumptions, 0, "EXPENSE",true);
        var loansRowNum = ArrayLib.indexOf(assumptions, 0, "LOANS",true);
        var arrayTemp=[],historical,expenseArray=[],j,remaining,type,startDate,milestone,monthlyAmt,startFinish,duration;

        rowNum = 23;

        for (i = expenseRowNum + 1; i < loansRowNum; i++) { // loop through EXPENSE items in the assumptions array

            rowData = assumptions[i];
            label = rowData[0];
            historical = rowData[1];
            remaining = rowData[2];
            type = rowData[3];
            startDate = rowData[4];
            milestone = rowData[5];
            startFinish = rowData[6];
            duration = rowData[7];
            if (!duration) duration = 1;
            total = 0;


            if (!label && !historical && !remaining) continue;
            if (label && !historical && !remaining) {
                arrayTemp = [...arrayTemplate];
                arrayTemp.unshift('');
                arrayTemp.unshift('');
                arrayTemp.unshift(label.trim());
                sheet.getRange(rowNum,2,1,projectDuration + 3).setFontSize(10).setNumberFormat('  @').setHorizontalAlignment('right').setValues([arrayTemp]);
                sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('  @').setFontSize(11).setFontWeight('bold');
                rowNum ++;
                arrayTemp = [];
                continue;
            }

            arrayTemp = [...arrayTemplate];

            switch(type) {

                case 'lump sum':

                    if (startDate) {

                        if (startDate == 'now' || (!startDate && !milestone)) {

                            if (duration > 0) {
                                monthlyAmt = remaining / duration;
                                if (!historical) {
                                    for (k = 0; k < duration; k++) {
                                        arrayTemp[k+1] = monthlyAmt;
                                        total += monthlyAmt;
                                        totalExpense[k+1] += arrayTemp[k+1];
                                    }

                                } else {
                                   for (k = 0; k < duration; k++) {
                                        arrayTemp[k] = monthlyAmt;
                                        total += monthlyAmt;
                                        totalExpense[k] += arrayTemp[k];
                                    }
                                }
                            } else if (!duration || duration == 1) {
                                startMonth = monthDiff(projectStartDate,new Date(startDate)) + 1;
                                arrayTemp[startMonth + 1] = remaining;
                                total = remaining;
                                totalExpense[startMonth + 1] += arrayTemp[startMonth + 1];
                            }

                        } else if (startDate) {

                            if (duration > 0) {

                                monthlyAmt = remaining / duration;
                                startMonth = monthDiff(projectStartDate,new Date(startDate)) + 1;
                                for (j = 0; j < duration; j++) {
                                    arrayTemp[startMonth + j] = monthlyAmt;
                                    total += monthlyAmt;
                                    totalExpense[startMonth + j] += arrayTemp[startMonth + j];
                                }

                            } else if (!duration) {
                                startMonth = monthDiff(projectStartDate,new Date(startDate)) + 1;
                                arrayTemp[startMonth + 1] = remaining;
                                total = remaining;
                                totalExpense[startMonth + 1] += arrayTemp[startMonth + 1];

                            }
                            else if (duration < 0) {
                                // shouldn't get here
                            }
                        }

                    } else if (milestone) {
                        index = ArrayLib.indexOf(scheduleData,1,milestone);
                        if (index == -1) continue;
                        monthlyAmt = remaining / duration;

                        startDate = new Date(scheduleData[index][earliestStartIndex]);
                        endDate = new Date(scheduleData[index][earliestFinishIndex]);

                        startMonth = monthDiff(projectStartDate,startDate);
                        endMonth = monthDiff(projectStartDate,endDate);

                        if (startFinish == 'start') {
                            for (j = 0; j < duration; j++) {
                                arrayTemp[startMonth + j] = monthlyAmt;
                                total += arrayTemp[startMonth + j];
                                totalExpense[startMonth + j] += arrayTemp[startMonth + j];
                            }


                        } else if (startFinish == 'finish') {

                            if (duration < 0) {
                                duration *= -1;
                                for (j = 0; j < duration; j++) {
                                    arrayTemp[endMonth - j + 1] = monthlyAmt * -1;
                                    total += monthlyAmt * -1;
                                    totalExpense[endMonth - j + 1] += arrayTemp[endMonth - j + 1];
                                }

                            } else if (duration > 0) {
                                for (j = 0; j < duration; j++) {
                                    arrayTemp[startMonth + j + 1] = monthlyAmt;
                                    total += monthlyAmt;
                                    totalExpense[startMonth + j + 1] += arrayTemp[startMonth + j + 1];
                                }
                            }
                        }
                    } else {
                        monthlyAmt = remaining / duration;
                        if (!monthlyAmt) monthlyAmt = null;
                        for (k = 0; k < duration; k++) {
                            arrayTemp[k] = monthlyAmt;
                            total += monthlyAmt;
                            totalExpense[k] += arrayTemp[k];
                        }
                    }
                break;

                case 'monthly':

                    if (!startDate && !milestone) {
                        for (j = 0; j < projectDuration; j++) {
                            arrayTemp[j] = remaining;
                            total += remaining;
                            totalExpense[j] += arrayTemp[j];
                        }
                    } else if (milestone) {
                        index = ArrayLib.indexOf(scheduleData,1,milestone);
                        monthlyAmt = remaining / duration;

                        startDate = new Date(scheduleData[index][earliestStartIndex]);
                        endDate = new Date(scheduleData[index][earliestFinishIndex]);

                        startMonth = monthDiff(projectStartDate,startDate);
                        endMonth = monthDiff(projectStartDate,endDate);
                        if (startFinish == 'start') {
                            startMonth ++; // delay one month for billing
                            for (j = startMonth; j < projectDuration; j++) {
                                arrayTemp[j] = remaining;
                                total += arrayTemp[j];
                                totalExpense[j] += arrayTemp[j];
                            }
                        } else if (startFinish == 'finish') {
                            endMonth ++; // delay one month for billing
                            for (j = endMonth; j < projectDuration; j++) {
                                arrayTemp[j] = remaining;
                                total += arrayTemp[j];
                                totalExpense[j] += arrayTemp[j];
                            }
                        }
                    } else if (startDate) {
                        startMonth = monthDiff(projectStartDate,new Date(startDate)) + 1; // delay 1 month for billing
                        for (j = startMonth; j < projectDuration; j++) {
                            arrayTemp[j] = remaining;
                            total += remaining;
                            totalExpense[j] += arrayTemp[j];
                        }
                    }
                break;

                case '% of sales':

                    var percentOfSales = remaining;
                    for (j = 0; j < projectDuration; j++) {
                        if (sales[j+3]) {
                            arrayTemp[j] = percentOfSales * sales[j+3];
                        }
                        total += arrayTemp[j];
                        totalExpense[j] += arrayTemp[j];

                    }
                break;

                default:
            }

            arrayTemp.unshift(historical);
            arrayTemp.unshift(historical + total);
            arrayTemp.unshift(label.trim());

            if (historical) totalHistoricalExpense += historical;

            sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0").setFontSize(10).setHorizontalAlignment('right').setValues([arrayTemp]);
            sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('    @');
            rowNum ++;

            if (label.trim() == "Land Acquisition") {

                sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('  @').setFontSize(11).setFontWeight('bold').setValue('Condo Construction');
                rowNum ++;

                sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0").setFontSize(10).setHorizontalAlignment('right').setValues([condoConstCost]);
                sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('    @');
                rowNum ++;

                sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0").setFontSize(10).setHorizontalAlignment('right').setValues([contingency]);
                sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('    @');
                rowNum ++;

                sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0").setFontSize(10).setHorizontalAlignment('right').setValues([constMgmt]);
                sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('    @');
                rowNum ++;
            }
        }

        rowNum ++;
        var totalExpenseTotal = 0;
        for (i = 0; i < totalExpense.length; i++) {
            if (totalExpense[i]) totalExpenseTotal += totalExpense[i];
        }

        totalExpense.unshift(totalHistoricalExpense);
        totalExpense.unshift(totalExpenseTotal + totalHistoricalExpense);
        totalExpense.unshift('Total Expense');

        sheet.setRowHeight(rowNum,25);
        sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0").setFontSize(10).setHorizontalAlignment('right').setBackground('#666').setFontColor('white').setVerticalAlignment('middle').setValues([totalExpense]);
        sheet.getRange(rowNum,2).setFontSize(11).setHorizontalAlignment('left');
        sheet.getRange(rowNum,1).setBackground('#666').setFontColor('white');

    // earnest money
        var earnestMoneyDeposits = [...arrayTemplate]
        var earnestMoneyWithdrawls = [...arrayTemplate];
        var earnestMoneyBalance = [...arrayTemplate];
        var earnestMoneyDepositsTotal = 0;
        var earnestMoneyWithdrawlsTotal = 0;

        for (i = 0; i < projectDuration -1; i++) {

            earnestMoneyDeposits[i] = firstEMDeposit[i+3] + secondEMDeposit[i+3] + thirdEMDeposit[i+3] + upfrontCommissions[i+3];
            if (!earnestMoneyDeposits[i]) earnestMoneyDeposits[i] = null;
            earnestMoneyDepositsTotal += earnestMoneyDeposits[i];

            if (i == 0) {
                if (totalExpense[i+3] <= earnestMoneyDeposits[i]) {
                    earnestMoneyWithdrawls[0] = totalExpense[3];
                    earnestMoneyWithdrawlsTotal += earnestMoneyWithdrawls[0];
                } else {
                    earnestMoneyWithdrawls[0] = earnestMoneyDeposits[0];
                    earnestMoneyWithdrawlsTotal += earnestMoneyWithdrawls[0];
                }
                earnestMoneyBalance[0] = earnestMoneyDeposits[0] - earnestMoneyWithdrawls[0];
                if (!earnestMoneyBalance[0]) earnestMoneyBalance[0] = null;
            } else {

                if ((earnestMoneyBalance[i-1] + earnestMoneyDeposits[i]) > totalExpense[i+3]) {
                    earnestMoneyWithdrawls[i] = totalExpense[i+3];
                    if (!earnestMoneyWithdrawls[i]) earnestMoneyWithdrawls[i] = null;
                    earnestMoneyWithdrawlsTotal += earnestMoneyWithdrawls[i];
                    earnestMoneyBalance[i] = earnestMoneyBalance[i-1] + earnestMoneyDeposits[i] - earnestMoneyWithdrawls[i];
                    if (!earnestMoneyBalance[i]) earnestMoneyBalance[i] = null;
                } else {
                    earnestMoneyWithdrawls[i] = earnestMoneyBalance[i-1] + earnestMoneyDeposits[i];
                    if (!earnestMoneyWithdrawls[i]) earnestMoneyWithdrawls[i] = null;
                    earnestMoneyWithdrawlsTotal += earnestMoneyWithdrawls[i];
                    earnestMoneyBalance[i] = earnestMoneyBalance[i-1] + earnestMoneyDeposits[i] - earnestMoneyWithdrawls[i];
                    if (!earnestMoneyBalance[i]) earnestMoneyBalance[i] = null;
                }
            }

        }

        earnestMoneyDeposits.unshift(null);
        earnestMoneyDeposits.unshift(earnestMoneyDepositsTotal);
        earnestMoneyDeposits.unshift('Deposits');

        earnestMoneyWithdrawls.unshift(null);
        earnestMoneyWithdrawls.unshift(earnestMoneyDepositsTotal);
        earnestMoneyWithdrawls.unshift('Withdrawls');

        earnestMoneyBalance.unshift(null);
        earnestMoneyBalance.unshift(null);
        earnestMoneyBalance.unshift('Balance');

        rowNum += 2;

        sheet.getRange(rowNum,2).setFontSize(11).setFontWeight('bold').setNumberFormat('@').setValue('EARNEST MONEY');
        rowNum ++;
        sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0").setFontSize(10).setHorizontalAlignment('right').setValues([earnestMoneyDeposits]);
        sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('  @');
        rowNum ++;
        sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0").setFontSize(10).setHorizontalAlignment('right').setValues([earnestMoneyWithdrawls]);
        sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('  @');
        rowNum ++;
        sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0").setFontSize(10).setHorizontalAlignment('right').setValues([earnestMoneyBalance]);
        sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('  @');

    // net cash flow
        var netCashFlow = [...arrayTemplate];
        var netCashFlowTotal = 0;
        for (i = 0; i < projectDuration -1; i++) {
            netCashFlow[i] = earnestMoneyWithdrawls[i+3] - totalExpense[i+3];
            if (!netCashFlow[i]) netCashFlow[i] = null;
            netCashFlowTotal += netCashFlow[i];
        }

        netCashFlow[projectDuration - 1] = netRevenue[projectDuration + 2] - totalExpense[projectDuration + 2];
        netCashFlowTotal += netCashFlow[projectDuration - 1];

        netCashFlow.unshift(-totalHistoricalExpense);
        netCashFlow.unshift(netCashFlowTotal - totalHistoricalExpense);
        netCashFlow.unshift('Net Cash Flow');

        rowNum +=2;
        sheet.setRowHeight(rowNum,25);
        sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setBackground('#666').setFontColor('white').setVerticalAlignment('middle').setValues([netCashFlow]);
        sheet.getRange(rowNum,2).setFontSize(11).setHorizontalAlignment('left');
        sheet.getRange(rowNum,1).setBackground('#666').setFontColor('white');

    // loans

        // assumptions
            sheet = ss.getSheetByName('assumptions');
            loansRowNum += 4;

            var acquisitionLoanBal = sheet.getRange(loansRowNum,3).getValue();
            var acquisitionLoanRate = sheet.getRange(loansRowNum + 1, 3).getValue();
            var acquisitionIntToDate = sheet.getRange(loansRowNum + 2, 3).getValue();

            var constLoanEstAmt = sheet.getRange(loansRowNum + 4, 3).getValue();
            var constLoanStartDate = sheet.getRange(loansRowNum + 5, 4).getValue();
            var constLoanStartMilestone = sheet.getRange(loansRowNum + 5, 5).getValue();
            var constLoanRate = sheet.getRange(loansRowNum + 6, 3).getValue();

            var constLoanBrokerFeeData = sheet.getRange(loansRowNum + 7,2,1,5).getValues();
            var constLoanOriginationFeeData = sheet.getRange(loansRowNum + 8,2,1,5).getValues();
            var constLoanCloseoutFeeData = sheet.getRange(loansRowNum + 9,2,1,5).getValues();

        // acquisition loan

            var acquisitionDraws = [...arrayTemplate];
            var acquisitionInt = [...arrayTemplate];
            var acquisitionPmts = [...arrayTemplate];
            var acquisitionBal = [...arrayTemplate];

            var acquisitionIntTotal = 0;
            var acquisitionPmtsTotal = 0;

            var acquisitionLoanEndMonth = monthDiff(projectStartDate,new Date(constLoanStartDate)) + 1; // month 1

            for (i = 0; i < acquisitionLoanEndMonth; i++) {
                if (i == 0) {
                    acquisitionBal[0] = acquisitionLoanBal;
                    acquisitionInt[0] = acquisitionLoanBal * acquisitionLoanRate / 12;
                    acquisitionIntTotal += acquisitionInt[i];
                    continue;
                }

                if (i != acquisitionLoanEndMonth -1) {
                    acquisitionInt[i] = acquisitionBal[i-1] * acquisitionLoanRate / 12;
                    acquisitionBal[i] = acquisitionBal[i-1];

                } else if (i == acquisitionLoanEndMonth - 1) {
                    acquisitionInt[i] = acquisitionBal[i-1] * acquisitionLoanRate / 12;
                    acquisitionPmts[i] = acquisitionBal[i-1];
                    acquisitionBal[i] = acquisitionBal[i-1] - acquisitionPmts[i];
                    if (!acquisitionBal[i]) acquisitionBal[i] = null;
                }
                acquisitionIntTotal += acquisitionInt[i];
                acquisitionPmtsTotal += acquisitionPmts[i]

            }

            acquisitionDraws.unshift(acquisitionLoanBal);
            acquisitionDraws.unshift(acquisitionLoanBal);
            acquisitionDraws.unshift('Draws');

            acquisitionInt.unshift(acquisitionIntToDate);
            acquisitionInt.unshift(acquisitionIntTotal + acquisitionIntToDate);
            acquisitionInt.unshift('Interest Payments');

            acquisitionPmts.unshift(null);
            acquisitionPmts.unshift(acquisitionPmtsTotal);
            acquisitionPmts.unshift('Principal Payments');

            acquisitionBal.unshift(acquisitionLoanBal);
            acquisitionBal.unshift(null);
            acquisitionBal.unshift('Balance');

            sheet = ss.getSheetByName('Proforma');

            rowNum +=2;
            sheet.getRange(rowNum,2).setFontSize(11).setFontWeight('bold').setNumberFormat('@').setValue('LOANS');
            rowNum ++;
            sheet.getRange(rowNum,2).setFontSize(11).setFontWeight('bold').setNumberFormat('  @').setHorizontalAlignment('left').setValue('Acquisition Loan');
            rowNum ++;
            sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setValues([acquisitionDraws]);
            sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('    @');
            rowNum ++;
            sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setValues([acquisitionInt]);
            sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('    @');
            rowNum ++;
            sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setValues([acquisitionPmts]);
            sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('    @');
            rowNum ++;
            sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setValues([acquisitionBal]);
            sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('    @');

        // construction loan

            var constOriginationFee = [...arrayTemplate];
            var constBrokerFee = [...arrayTemplate];
            var constCloseoutFee = [...arrayTemplate];
            var constDraws = [...arrayTemplate];
            var constInt = [...arrayTemplate];
            var constPmts = [...arrayTemplate];
            var constBal = [...arrayTemplate];

            var constDrawsTotal = 0;
            var constIntTotal = 0;
            var constPmtsTotal = 0;

            var constLoanCloseoutFee = constLoanEstAmt * constLoanCloseoutFee;
            var constLoanBrokerFee = constLoanEstAmt * constLoanBrokerFee;

            // origination fee
                rowData = constLoanOriginationFeeData[0];
                var originationFeeLabel = rowData[0];
                var originationFeeAmt = constLoanEstAmt * rowData[1];
                var originationFeeDate = rowData[2];
                var originationFeeMilestone = rowData[3];
                var originationFeeStartFinish = rowData[4];

                if (originationFeeDate) {
                    month = monthDiff(projectStartDate,new Date(originationFeeDate));
                } else if (originationFeeMilestone) {
                    index = ArrayLib.indexOf(scheduleData,1,originationFeeMilestone.trim());
                    rowData = scheduleData[index];

                    startDate = rowData[earliestStartIndex];
                    finishDate = rowData[earliestFinishIndex];

                    if (originationFeeStartFinish == 'start') {
                        month = monthDiff(projectStartDate, new Date(startDate));
                    } else if (originationFeeStartFinish == 'finish') {
                        month = monthDiff(projectStartDate, new Date(finishDate));
                    }
                }
                constOriginationFee[month] = originationFeeAmt;

            // broker fee
                rowData = constLoanBrokerFeeData[0];
                var brokerFeeLabel = rowData[0];
                var brokerFeeAmt = constLoanEstAmt * rowData[1];
                var brokerFeeDate = rowData[2];
                var brokerFeeMilestone = rowData[3];
                var brokerFeeStartFinish = rowData[4];

                if (brokerFeeDate) {
                    month = monthDiff(projectStartDate,new Date(brokerFeeDate));
                } else if (brokerFeeMilestone) {
                    index = ArrayLib.indexOf(scheduleData,1,brokerFeeMilestone.trim());
                    rowData = scheduleData[index];

                    startDate = rowData[earliestStartIndex];
                    finishDate = rowData[earliestFinishIndex];

                    if (brokerFeeStartFinish == 'start') {
                        month = monthDiff(projectStartDate, new Date(startDate));
                    } else if (brokerFeeStartFinish == 'finish') {
                        month = monthDiff(projectStartDate, new Date(finishDate));
                    }
                }
                constBrokerFee[month] = brokerFeeAmt;

            // closeout fee
                rowData = constLoanCloseoutFeeData[0];
                var closeoutFeeLabel = rowData[0];
                var closeoutFeeAmt = constLoanEstAmt * rowData[1];
                var closeoutFeeDate = rowData[2];
                var closeoutFeeMilestone = rowData[3];
                var closeoutFeeStartFinish = rowData[4];

                if (closeoutFeeDate) {
                    month = monthDiff(projectStartDate,new Date(closeoutFeeDate));
                } else if (closeoutFeeMilestone) {
                    index = ArrayLib.indexOf(scheduleData,1,closeoutFeeMilestone.trim());
                    rowData = scheduleData[index];

                    startDate = rowData[earliestStartIndex];
                    finishDate = rowData[earliestFinishIndex];

                    if (closeoutFeeStartFinish == 'start') {
                        month = monthDiff(projectStartDate, new Date(startDate));
                    } else if (closeoutFeeStartFinish == 'finish') {
                        month = monthDiff(projectStartDate, new Date(finishDate));
                    }
                }
                if (month == projectDuration - 2) month ++; // add one month when fee occurs on the last construction month
                constCloseoutFee[month] = closeoutFeeAmt;

            var constLoanStartMonth = monthDiff(projectStartDate,new Date(constLoanStartDate));
            var constLoanBrokerFeeStartMonth = monthDiff(projectStartDate,new Date(constLoanStartDate));

            for (i = constLoanStartMonth; i < projectDuration; i ++) {

                if (i == constLoanStartMonth) {
                    constDraws[i] = acquisitionPmts[i+3];
                    if (netCashFlow[i+3] < 0) {
                        constDraws[i] = constDraws[i] - netCashFlow[i+3];
                        constBal[i] = constDraws[i];
                    }
                    constDrawsTotal += constDraws[i];
                    continue;
                }

                constInt[i] = constBal[i-1] * constLoanRate/12;

                if (netCashFlow[i+3] < 0) {
                    constDraws[i] = -netCashFlow[i+3];
                    constDrawsTotal += constDraws[i];
                    constIntTotal += constInt[i];
                    constBal[i] = constBal[i-1] + constDraws[i] + constInt[i] + constOriginationFee[i] + constBrokerFee[i] + constCloseoutFee[i] - constPmts[i];
                }
                else { // netCashFlow > 0

                    if (netCashFlow[i+3] < constBal[i-1] + constInt[i] + constOriginationFee[i] + constBrokerFee[i] + constCloseoutFee[i]) {
                        constPmts[i] = netCashFlow[i+3];
                        constPmtsTotal += constDraws[i];
                    } else { // net cash flow > prior bal + int + fees
                        constPmts[i] = constBal[i-1] + constInt[i] + constOriginationFee[i] + constBrokerFee[i] + constCloseoutFee[i];
                        constPmtsTotal += constPmts[i];
                    }

                    constInt[i] = constBal[i-1] * constLoanRate/12;
                    constIntTotal += constInt[i];
                    constBal[i] = constBal[i-1] + constDraws[i] + constInt[i] + constOriginationFee[i] + constBrokerFee[i] + constCloseoutFee[i] - constPmts[i];
                    if (!constBal[i]) constBal[i] = null;
                }

            }


            constOriginationFee.unshift(null);
            constOriginationFee.unshift(originationFeeAmt);
            constOriginationFee.unshift('Origination Fee');

            constBrokerFee.unshift(null);
            constBrokerFee.unshift(brokerFeeAmt);
            constBrokerFee.unshift('Broker Fee');

            constCloseoutFee.unshift(null);
            constCloseoutFee.unshift(closeoutFeeAmt);
            constCloseoutFee.unshift('Closeout Fee');

            constDraws.unshift(null);
            constDraws.unshift(constDrawsTotal);
            constDraws.unshift('Draws');

            constInt.unshift(null);
            constInt.unshift(constIntTotal);
            constInt.unshift('Interest Accrued');

            constPmts.unshift(null);
            constPmts.unshift(constPmtsTotal);
            constPmts.unshift('Payments');

            constBal.unshift(null);
            constBal.unshift(null);
            constBal.unshift('Balance');

            sheet = ss.getSheetByName('Proforma');

            rowNum ++;
            sheet.getRange(rowNum,2).setFontSize(11).setFontWeight('bold').setNumberFormat('  @').setHorizontalAlignment('left').setValue('Construction Loan');
            rowNum++;
            sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setValues([constOriginationFee]);
            sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('    @');
            rowNum++;
            sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setValues([constBrokerFee]);
            sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('    @');
            rowNum++;
            sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setValues([constCloseoutFee]);
            sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('    @');
            rowNum++;
            sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setValues([constDraws]);
            sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('    @');
            rowNum++;
            sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setValues([constInt]);
            sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('    @');
            rowNum++;
            sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setValues([constPmts]);
            sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('    @');
            rowNum++;
            sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setValues([constBal]);
            sheet.getRange(rowNum,2).setHorizontalAlignment('left').setNumberFormat('    @');

        // total Loans

            var totalLoans = [...arrayTemplate];
            var totalLoansTotal = 0;
            for (i = 0; i < projectDuration; i++) {
                // totalLoans[i] = acquisitionDraws[i+3] + constDraws[i+3] - acquisitionPmts[i+3] - constPmts[i+3] - acquisitionInt[i+3];
                totalLoans[i] = constDraws[i+3] - constPmts[i+3] - acquisitionInt[i+3];
                if (!totalLoans[i]) totalLoans[i] = null;
                totalLoansTotal += totalLoans[i];
            }

            // totalLoans.unshift(null);
            totalLoans.unshift(-acquisitionIntToDate);
            totalLoans.unshift(totalLoansTotal - acquisitionIntToDate);
            totalLoans.unshift('Total Loans');

            rowNum +=2;
            sheet.setRowHeight(rowNum,25);
            sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setBackground('#666').setFontColor('white').setVerticalAlignment('middle').setValues([totalLoans]);
            sheet.getRange(rowNum,2).setFontSize(11).setHorizontalAlignment('left');
            sheet.getRange(rowNum,1).setBackground('#666').setFontColor('white');

    // total Cash Flow

        var totalCashFlow = [...arrayTemplate];
        var totalCashFlowTotal = 0;
        for (i = 0; i < projectDuration; i++) {
            totalCashFlow[i] = netCashFlow[i+3] - acquisitionPmts[i+3] + acquisitionDraws[i+3] - acquisitionInt[i+3] + constDraws[i+3] - constPmts[i+3];
            if (!totalCashFlow[i]) totalCashFlow[i] = null;
            totalCashFlowTotal += totalCashFlow[i];
        }

        totalCashFlow.unshift(netCashFlow[2] - acquisitionInt[2] + acquisitionDraws[2]);
        totalCashFlow.unshift(totalCashFlowTotal + netCashFlow[2] - acquisitionInt[2] + acquisitionDraws[2]);
        totalCashFlow.unshift('Total Cash Flow (Profit)');

        rowNum +=2;
        sheet.setRowHeight(rowNum,25);
        sheet.getRange(rowNum,2,1,projectDuration + 3).setNumberFormat("#,##0;(#,##0)").setFontSize(10).setHorizontalAlignment('right').setBackground('#256ab0').setFontColor('white').setVerticalAlignment('middle').setValues([totalCashFlow]);
        sheet.getRange(rowNum,2).setFontSize(11).setHorizontalAlignment('left');
        sheet.getRange(rowNum,1).setBackground('#256ab0').setFontColor('white');

    // summary

        sheet = ss.getSheetByName('Summary');
        if (!sheet) {
            ui.alert('Could not find "Summary" sheet');
            return;
        }
        sheet.activate();
        sheet.getRange('A3').setValue(headerDate);
        sheet.getRange('C6').setValue(totalSales[1]);
        sheet.getRange('C7').setValue(netRevenue[1]);
        sheet.getRange('C8').setValue(totalExpense[1]);
        sheet.getRange('C9').setValue(netCashFlow[1]);
        sheet.getRange('C12').setValue(constPmts[1]);
        sheet.getRange('C13').setValue(constDuration);
        sheet.getRange('C14').setValue(constOriginationFee[1] + constBrokerFee[1] + constCloseoutFee[1]);
        sheet.getRange('C15').setValue(constInt[1] + acquisitionInt[1]);
        sheet.getRange('C16').setValue((constPmts[1] + earnestMoneyDeposits[1])/(totalExpense[1] + constPmts[1] - constDraws[1]));
        sheet.getRange('C17').setValue((constPmts[1] + earnestMoneyDeposits[1])/totalSales[1]);
        sheet.getRange('C19').setValue(totalCashFlow[1]);
}

function monthDiff(dateFrom, dateTo) {

    return parseInt(dateTo.getMonth() - dateFrom.getMonth() + (12 * (dateTo.getFullYear() - dateFrom.getFullYear())));
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
