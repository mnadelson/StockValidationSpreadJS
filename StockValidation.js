window.onload = function() {
    var spread = new GC.Spread.Sheets.Workbook(document.getElementById('ss'));
    initSpread(spread);
};

function initSpread(spread) {
    var sheet = spread.getSheet(0);
    sheet.name("Stock Data");
    sheet.suspendPaint();
    loadData(spread);
    spread.options.highlightInvalidData = true;
    validateSharesOutstanding(spread);
    validateNonNegative(spread);
    validateStockSymbolLength(spread);
    validateBusinessDate(spread);
    validatePriceEarningsRatio(spread);
    validateStockSymbol(spread);
    validateStockSector(spread);
    handleErrors(spread);
    alertOnStockPrice(spread);
    sheet.resumePaint();
};

function loadData(spread) {
    var sheet = spread.getSheet(0);
    sheet.setRowCount(10, GC.Spread.Sheets.SheetArea.viewport);
    sheet.setRowVisible(7, true);
    sheet.setValue(0, 0, "Business Date");
    sheet.setValue(0, 1, "Position ID");
    sheet.setValue(0, 2, "Sector");
    sheet.setValue(0, 3, "Shares Outstanding");
    sheet.setValue(0, 4, "Price");
    sheet.setValue(0, 5, "Earnings Per Share");
    sheet.setValue(0, 6, "Price/Earnings Ratio");

    var todayDate = new Date();
    todayDate.setHours(0);
    todayDate.setMinutes(0);
    todayDate.setSeconds(0);

    sheet.setValue(1, 0, todayDate);
    sheet.setValue(1, 1, "ABC");
    sheet.setValue(1, 2, "Energy");
    sheet.setValue(1, 3, 500000);  // Invalid shares outstanding
    sheet.setValue(1, 4, 100);
    sheet.setValue(1, 5, 5);
    sheet.setValue(1, 6, 20);

    sheet.setValue(2, 0, todayDate);
    sheet.setValue(2, 1, "ABCDEF");// Invalid stock symbol length
    sheet.setValue(2, 2, "Utilities");
    sheet.setValue(2, 3, 5000000);
    sheet.setValue(2, 4, 98);
    sheet.setValue(2, 5, 4);
    sheet.setValue(2, 6, 24.5);

    sheet.setValue(3, 0, new Date(2020, 7, 18, 1, 0, 0));  // Invalid date
    sheet.setValue(3, 1, "DFG");
    sheet.setValue(3, 2, "Financials");
    sheet.setValue(3, 3, 2000000);
    sheet.setValue(3, 4, 100);
    sheet.setValue(3, 5, 5);
    sheet.setValue(3, 6, 20);

    sheet.setValue(4, 0, todayDate);
    sheet.setValue(4, 1, "HIJ");
    sheet.setValue(4, 2, "Computes"); // Invalid sector
    sheet.setValue(4, 3, 2500000);
    sheet.setValue(4, 4, 1000);
    sheet.setValue(4, 5, 5);
    sheet.setValue(4, 6, 200);

    sheet.setValue(5, 0, todayDate);
    sheet.setValue(5, 1, "KLM");
    sheet.setValue(5, 2, "Industrials");
    sheet.setValue(5, 3, 4000000);
    sheet.setValue(5, 4, -100); // Invalid price
    sheet.setValue(5, 5, -5);
    sheet.setValue(5, 6, 20);

    sheet.setValue(6, 0, todayDate);
    sheet.setValue(6, 1, "NOP");
    sheet.setValue(6, 2, "Industrials");
    sheet.setValue(6, 3, 4500000);
    sheet.setValue(6, 4, 100);
    sheet.setValue(6, 5, 5);
    sheet.setValue(6, 6, 21);  // Invalid  P/E ratio

    sheet.setValue(7, 0, todayDate);
    sheet.setValue(7, 1, "QRS"); // Invalid Stock Symbol
    sheet.setValue(7, 2, "Materials");
    sheet.setValue(7, 3, 6500000);
    sheet.setValue(7, 4, 100);
    sheet.setValue(7, 5, 5);
    sheet.setValue(7, 6, 20);

    sheet.setValue(0, 10, "Valid Stock Symbols");
    sheet.setValue(1, 10, "ABC");
    sheet.setValue(2, 10, "DFG");
    sheet.setValue(3, 10, "HIJ");
    sheet.setValue(4, 10, "KLM");
    sheet.setValue(5, 10, "NOP");
}

function validateSharesOutstanding(spread) {
    var sheet = spread.getSheet(0);

    var dv = new GC.Spread.Sheets.DataValidation.createNumberValidator(GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators.greaterThanOrEqualsTo, "1000000", null, true);
    dv.highlightStyle({
                 type: GC.Spread.Sheets.DataValidation.HighlightType.circle,
                 color: 'green'
             });
    dv.showInputMessage(true);
    dv.inputMessage("Shares outstanding must be greater than or equal to 1,000,000 shares.")
    dv.inputTitle("Shares Outstanding Tip");
    dv.showErrorMessage(true);
    dv.errorMessage("Shares outstanding value entered was not greater than or equal to 1,000,000 shares");
    sheet.setDataValidator(1, 3, 7, 1, dv, GC.Spread.Sheets.SheetArea.viewport);
}

function validateNonNegative(spread) {
    var sheet = spread.getSheet(0);

    var dv = new GC.Spread.Sheets.DataValidation.createNumberValidator(GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators.greaterThan, 0, null, true);
    dv.highlightStyle({
        type: GC.Spread.Sheets.DataValidation.HighlightType.circle,
        color: 'blue'
    });
    sheet.setDataValidator(1, 4, 7, 1, dv, GC.Spread.Sheets.SheetArea.viewport);
    sheet.setDataValidator(1, 5, 7, 1, dv, GC.Spread.Sheets.SheetArea.viewport);
}

function validateStockSymbolLength(spread) {
    var sheet = spread.getSheet(0);

    var dv = new GC.Spread.Sheets.DataValidation.createTextLengthValidator(GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators.between, 1, 5);
    sheet.setDataValidator(1, 1, 7, 1, dv, GC.Spread.Sheets.SheetArea.viewport);
}

function validateBusinessDate(spread) {
    var sheet = spread.getSheet(0);

    var todayDate = new Date();
    todayDate.setHours(0);
    todayDate.setMinutes(0);
    todayDate.setSeconds(0);
    var dv = new GC.Spread.Sheets.DataValidation.createDateValidator(GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators.equalsTo, todayDate, null);
    sheet.setDataValidator(1, 0, 7, 1, dv, GC.Spread.Sheets.SheetArea.viewport);
}

function validatePriceEarningsRatio(spread) {
    var sheet = spread.getSheet(0);

    var numRows = sheet.getRowCount();
    var rowCount;
    for(rowCount =  1; rowCount < numRows+1; rowCount++) {
        var cellRow = rowCount+1;   // Cell rows are 1 based so we must add 1 to rowCount to get the cell row
        var peFormula = "G"+cellRow+"="+"E"+cellRow+"/F" + cellRow;
        var dv = new GC.Spread.Sheets.DataValidation.createFormulaValidator(peFormula);
        dv.highlightStyle({
            type: GC.Spread.Sheets.DataValidation.HighlightType.dogEar,
            color: 'red'
        });
        sheet.setDataValidator(rowCount, 6, 1, 1, dv, GC.Spread.Sheets.SheetArea.viewport);
    }
}

function validateStockSymbol(spread) {
     var sheet = spread.getSheet(0);

     var numRows = sheet.getRowCount();
     var dv =  new GC.Spread.Sheets.DataValidation.createFormulaListValidator("$K$2:$K$" + numRows);
     sheet.setDataValidator(1, 1, 7, 1, dv, GC.Spread.Sheets.SheetArea.viewport);
}

function validateStockSector(spread) {
    var sheet = spread.getSheet(0);

    var dv =  new GC.Spread.Sheets.DataValidation.createListValidator("Energy,\
                                                                       Materials,\
                                                                       Industrials,\
                                                                       Consumer Discretionary,\
                                                                       Consumer Staples,\
                                                                       Health Care,\
                                                                       Financials,\
                                                                       Information Technology,\
                                                                       Telecommunication Services,\
                                                                       Utilities,\
                                                                       Real Estate");
    dv.highlightStyle({
        type: GC.Spread.Sheets.DataValidation.HighlightType.icon,
        color: 'black',
        position: GC.Spread.Sheets.DataValidation.HighlightPosition.outsideLeft
    });
    sheet.setDataValidator(1, 2, 7, 1, dv, GC.Spread.Sheets.SheetArea.viewport);
}

function alertOnStockPrice(spread) {
    var sheet = spread.getSheet(0);
    var cfs = sheet.conditionalFormats;
    var style = new GC.Spread.Sheets.Style();
    style.backColor = 'yellow';
    var cvRule = cfs.addCellValueRule(GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators.greaterThanOrEqualsTo, 1000, 0, style, [new GC.Spread.Sheets.Range(1, 4, 7, 1)]);
}

function handleErrors(spread) {
    var sheet = spread.getSheet(0);
    sheet.bind(GC.Spread.Sheets.Events.ValidationError, function (sender, args) {
    if (args.validator.showErrorMessage()) {
       if (confirm(args.validator.errorMessage())) {
           args.validationResult = GC.Spread.Sheets.DataValidation.DataValidationResult.retry;
       } else {
           args.validationResult = GC.Spread.Sheets.DataValidation.DataValidationResult.forceApply;
       }
    }
   });
}
