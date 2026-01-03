


function createBulkPDFs(){
  const docFile = DriveApp.getFileById("1Qn5KZn1iQxFK3bzbSHi1DZMI5jGwbl1L7RtKuf3hsps")
  const tempFolder = DriveApp.getFolderById("1uQFkLOnccQA18uQ8xpHqjTGruU0SRNUj")
  const pdfFolder = DriveApp.getFolderById("1gDpzbSV8qtSKvpK9QqTuQYyOg9x7-z0O")
  
  const txnDetailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions")
  const txnData = txnDetailSheet.getRange(3,1, txnDetailSheet.getLastRow()-2,txnDetailSheet.getLastColumn()).getDisplayValues()
  
  let errors = []
  let fileURL = []

<<<<<<< Updated upstream
  txnData.forEach(row =>{
    let row_data = {}
    let tempStr

      row_data.invoiceNumber = row[0]
      row_data.subID = row[1]
      row_data.companyName = row[2]
      row_data.companyID = row[3]
      row_data.poc = row[4]
      row_data.compAddress = row[5]
      row_data.merchantGST = row[6]
      row_data.items = row[7]
      row_data.subDetails = row[8]
      row_data.invoiceDate = row[9]
      row_data.unitCost = row[10]
      row_data.quantity = row[11]
      row_data.totalCost = row[12]
      row_data.cgst = row[13]
      row_data.sgst = row[14]
      row_data.igst = row[15]
      row_data.finalAmount = row[16]
      row_data.amountText = row[17]

    try{

    console.log("hi",row_data)
    if(row[18].length==0){
      console.log("Starting PDF Gen")
    tempStr = createPDF(row_data,row[2]+"-"+row[0] +"-"+ new Date().toLocaleString(),docFile,tempFolder,pdfFolder)
    fileURL.push([tempStr])
    errors.push([""])
    }else
    {console.log("PDF Generated")}
  }catch(err){
    console.log(err)
    errors.push(["Failed"])
    //txnDetailSheet.getRange(3,19,txnDetailSheet.getLastRow()-2,1).setValues(errors)
  }
  });
  console.log(fileURL.length)
  txnDetailSheet.getRange(3,19,txnDetailSheet.getLastRow()-2,1).setValues(fileURL)
=======
    const txnDetailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions")
    const txnData = txnDetailSheet.getRange(3, 1, txnDetailSheet.getLastRow() - 2, txnDetailSheet.getLastColumn()).getDisplayValues()

    let errors = []
    let fileURL = []

    txnData.forEach(row => {
        let row_data = {}
        let tempStr

        row_data.invoiceNumber = row[0]
        row_data.subID = row[1]
        row_data.companyName = row[2]
        row_data.companyID = row[3]
        row_data.poc = row[4]
        row_data.compAddress = row[5]
        row_data.merchantGST = row[6]
        row_data.items = row[7]
        row_data.subDetails = row[8]
        row_data.invoiceDate = row[9]
        row_data.unitCost = row[10]
        row_data.quantity = row[11]
        row_data.totalCost = row[12]
        row_data.cgst = row[13]
        row_data.sgst = row[14]
        row_data.igst = row[15]
        row_data.finalAmount = row[16]
        row_data.amountText = row[17]

        try {

            console.log("hi", row_data)
            if (row[18].length == 0) {
                console.log("Starting PDF Gen")
                tempStr = createPDF(row_data, row[2] + "-" + row[0] + "-" + new Date().toLocaleString(), docFile, tempFolder, pdfFolder)
                fileURL.push([tempStr])
                errors.push([""])
            } else { console.log("PDF Generated") }
        } catch (err) {
            console.log(err)
            errors.push(["Failed"])
            //txnDetailSheet.getRange(3,19,txnDetailSheet.getLastRow()-2,1).setValues(errors)
        }
    });
    console.log(fileURL.length)
    txnDetailSheet.getRange(3, 19, txnDetailSheet.getLastRow() - 2, 1).setValues(fileURL)
>>>>>>> Stashed changes
}


function createPDF(row_data,pdfName,docFile,tempFolder,pdfFolder) {
  
  console.log(row_data)
  const tempFile = docFile.makeCopy(tempFolder)
  const tempDocFile = DocumentApp.openById(tempFile.getId())
  const body = tempDocFile.getBody()
  let pdfFile
  let legends = ["{invoiceNumber}","{subID}","{companyName}","{companyID}","{poc}","{compAddress}","{merchantGST}","{items}","{subDetails}","{invoiceDate}","{unitCost}","{quantity}","{totalCost}","{cgst}","{sgst}","{igst}","{finalAmount}","{amountText}"]

<<<<<<< Updated upstream
  body.replaceText(legends[0],row_data.invoiceNumber)
  body.replaceText(legends[1],row_data.subID)
  body.replaceText(legends[2],row_data.companyName)
  body.replaceText(legends[3],row_data.companyID)
  body.replaceText(legends[4],row_data.poc)
  body.replaceText(legends[5],row_data.compAddress)
  body.replaceText(legends[6],row_data.merchantGST)
  body.replaceText(legends[7],row_data.items)
  body.replaceText(legends[8],row_data.subDetails)
  body.replaceText(legends[9],row_data.invoiceDate)
  body.replaceText(legends[10],row_data.unitCost)
  body.replaceText(legends[11],row_data.quantity)
  body.replaceText(legends[12],row_data.totalCost)
  body.replaceText(legends[13],row_data.cgst)
  body.replaceText(legends[14],row_data.sgst)
  body.replaceText(legends[15],row_data.igst)
  body.replaceText(legends[16],row_data.finalAmount)
  body.replaceText(legends[17],row_data.amountText)

  
  tempDocFile.saveAndClose()
  const pdfContentBlob = tempFile.getAs(MimeType.PDF)
  pdfFile = pdfFolder.createFile(pdfContentBlob).setName(pdfName).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW).getUrl() //PDF File Handle
  tempFolder.removeFile(tempFile)
  //console.log("Inside Create PDF URL",pdfFile)
  return pdfFile
=======
    console.log(row_data)
    const tempFile = docFile.makeCopy(tempFolder)
    const tempDocFile = DocumentApp.openById(tempFile.getId())
    const body = tempDocFile.getBody()
    let pdfFile
    let legends = ["{invoiceNumber}", "{subID}", "{companyName}", "{companyID}", "{poc}", "{compAddress}", "{merchantGST}", "{items}", "{subDetails}", "{invoiceDate}", "{unitCost}", "{quantity}", "{totalCost}", "{cgst}", "{sgst}", "{igst}", "{finalAmount}", "{amountText}"]

    body.replaceText(legends[0], row_data.invoiceNumber)
    body.replaceText(legends[1], row_data.subID)
    body.replaceText(legends[2], row_data.companyName)
    body.replaceText(legends[3], row_data.companyID)
    body.replaceText(legends[4], row_data.poc)
    body.replaceText(legends[5], row_data.compAddress)
    body.replaceText(legends[6], row_data.merchantGST)
    body.replaceText(legends[7], row_data.items)
    body.replaceText(legends[8], row_data.subDetails)
    body.replaceText(legends[9], row_data.invoiceDate)
    body.replaceText(legends[10], row_data.unitCost)
    body.replaceText(legends[11], row_data.quantity)
    body.replaceText(legends[12], row_data.totalCost)
    body.replaceText(legends[13], row_data.cgst)
    body.replaceText(legends[14], row_data.sgst)
    body.replaceText(legends[15], row_data.igst)
    body.replaceText(legends[16], row_data.finalAmount)
    body.replaceText(legends[17], row_data.amountText)


    tempDocFile.saveAndClose()
    const pdfContentBlob = tempFile.getAs(MimeType.PDF)
    pdfFile = pdfFolder.createFile(pdfContentBlob).setName(pdfName).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW).getUrl() //PDF File Handle
    tempFolder.removeFile(tempFile)
    //console.log("Inside Create PDF URL",pdfFile)
    return pdfFile
>>>>>>> Stashed changes

}
