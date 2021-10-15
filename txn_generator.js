function createTxn() {
    const subDetailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Subscription Details")
    const txnDetailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions")
    const subData = subDetailSheet.getRange(2, 1, subDetailSheet.getLastRow() - 1, subDetailSheet.getLastColumn()).getDisplayValues()
    const txnData = txnDetailSheet.getRange(3, 1, txnDetailSheet.getLastRow() - 1, txnDetailSheet.getLastColumn()).getDisplayValues()
    const currentDate = new Date()


    let errors = []
    //txnData.getRange("H3").setFormula("=3+5");
    subData.forEach(row => {
        try {
            var res = checkTxn((row[0] + "/" + row[2] + "/" + new Date(row[5]).getTime().toString()).trim(), txnData)
            if (res.length === 0) {
                var writeData = [row[5], row[2], row[1], row[0], row[8], row[6]]
                txnDetailSheet.getRange(txnDetailSheet.getLastRow() + 1, 2, 1, 6).setValues([writeData])
            }

        } catch (err) {
            console.log(err)
            // errors.push(["Failed"])
        }
    });

    //Generating Invoice Number and Emails 
    applyFormulaFunc(txnDetailSheet, 1, "A3", `=CONCAT("EDLR/INV/",Row()-2)`)
    applyFormulaFunc(txnDetailSheet, 8, "H3", `=VLOOKUP($C3,'Contact Info'!$A$2:$G$1001,7,False)`)
    applyFormulaFunc(txnDetailSheet, 9, "I3", `=VLOOKUP($E3,'Subscription Details'!$A$2:$J$1001,10,False)`)
    applyFormulaFunc(txnDetailSheet, 10, "J3", `=VLOOKUP($C3,'Company Details'!$A$2:$G$1001,5,False)`)
    applyFormulaFunc(txnDetailSheet, 11, "K3", `=MONEYTEXT(G3,"INR")`)
    applyFormulaFunc(txnDetailSheet, 12, "L3", `=VLOOKUP($E3,'Subscription Details'!$A$2:$J$1001,4,False)`)


    //currentSheet.getRange(2,3,currentSheet.getLastRow()-1,1).setValues(errors)
}

function checkTxn(txnNew, txnData) {
    let match = []
    txnData.forEach(txn_row => {

        if (txnNew === (txn_row[4] + "/" + txn_row[2] + "/" + new Date(txn_row[1]).getTime().toString()).trim()) {
            match.push(txnNew);
        }

    })
    return match;
}


function applyFormulaFunc(sheet, col, cell, formula) {
    //const txnDetailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");

    sheet.getRange(cell).setFormula(formula);
    const lr = sheet.getLastRow();
    const fillDownRange = sheet.getRange(3, col, lr - 2, 1);
    sheet.getRange(cell).copyTo(fillDownRange);

}
