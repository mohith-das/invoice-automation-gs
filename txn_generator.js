function createTxn() {
   const subDetailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Subscription Details")
   const txnDetailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions")
   const subData = subDetailSheet.getRange(2,1,subDetailSheet.getLastRow()-1,subDetailSheet.getLastColumn()).getDisplayValues()
   const txnData = txnDetailSheet.getRange(3,1, txnDetailSheet.getLastRow()-1,txnDetailSheet.getLastColumn()).getDisplayValues()
   const currentDate = new Date()
   //let errors = []

  subData.forEach(row =>{
    console.log("in createtxn foreach",row)
    try{
      //SubID[0] - Company ID[2] - Invoice Date(Upcoming)[6] => txnNew
      var res = checkTxn((row[0] + "/" + row[2] + "/" + new Date(row[6]).getTime().toString()).trim(), txnData)
      if(res.length === 0){
        console.log("In createtxn if condition",row)

        var writeData = [row[0],row[1],row[2],"","","",row[3],row[4],row[6],row[7],row[8],row[9],row[10],row[11],row[12],row[13]]
        txnDetailSheet.getRange(txnDetailSheet.getLastRow() + 1, 2, 1, writeData.length).setValues([writeData])
      }

  }catch(err){
    console.log(err)
    //errors.push(["Failed"])
  }
  });


applyFormulaFunc(txnDetailSheet,1,"A3",`=CONCAT("EDLR/INV/",Row()-2)`) // Gen Invoice Number
applyFormulaFunc(txnDetailSheet,5,"E3",`=VLOOKUP($D3,'Company Details'!$A$2:$L,3,False)`) //Get POC 
applyFormulaFunc(txnDetailSheet,6,"F3",`=VLOOKUP($D3,'Company Details'!$A$2:$L,5,False)`) //Get Address
applyFormulaFunc(txnDetailSheet,7,"G3",`=VLOOKUP($D3,'Company Details'!$A$2:$L,7,False)`) //Get Merchant GST
applyFormulaFunc(txnDetailSheet,18,"R3",`=MONEYTEXT(Q3,"INR")`) // Amount to Text
applyFormulaFunc(txnDetailSheet,20,"T3",`=VLOOKUP($D3,'Company Details'!$A$2:$L,10,False)`)//Get Recipient
applyFormulaFunc(txnDetailSheet,21,"U3",`=VLOOKUP($D3,'Company Details'!$A$2:$L,11,False)`)//Get CC

  //currentSheet.getRange(2,3,currentSheet.getLastRow()-1,1).setValues(errors)
}

function checkTxn(txnNew,txnData)
{ 
  

  let match = []
txnData.forEach(txn_row=>{
   //SubID[0] - Company ID[2] - Invoice Date(Upcoming)[6] => txnNew
   //SubID[1] - Company ID[3] - Invoice Date(Old)[9] => txnData old
  if(txnNew === (txn_row[1] + "/" + txn_row[3] + "/" + new Date(txn_row[9]).getTime().toString()).trim())
  {
    match.push(txnNew);
  }

})
  return match;
}


function applyFormulaFunc(sheet,col, cell, formula)
{
   
  sheet.getRange(cell).setFormula(formula);
  const lr = sheet.getLastRow();
  const fillDownRange = sheet.getRange(3,col,lr-2,1);
  sheet.getRange(cell).copyTo(fillDownRange);

}
