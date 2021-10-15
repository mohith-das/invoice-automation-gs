


function createBulkPDFs() {
    const docFile = DriveApp.getFileById("1Qn5KZn1iQxFK3bzbSHi1DZMI5jGwbl1L7RtKuf3hsps")
    const tempFolder = DriveApp.getFolderById("1uQFkLOnccQA18uQ8xpHqjTGruU0SRNUj")
    const pdfFolder = DriveApp.getFolderById("1gDpzbSV8qtSKvpK9QqTuQYyOg9x7-z0O")

    // const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("people")
    // const data = currentSheet.getRange(2,1,currentSheet.getLastRow()-1,4).getDisplayValues()

    //const subDetailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Subscription Details")
    const txnDetailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions")
    //const subData = subDetailSheet.getRange(2,1,subDetailSheet.getLastRow()-1,subDetailSheet.getLastColumn()).getDisplayValues()
    const txnData = txnDetailSheet.getRange(3, 1, txnDetailSheet.getLastRow() - 2, txnDetailSheet.getLastColumn()).getDisplayValues()


    let errors = []
    let fileURL = []
    txnData.forEach(row => {
        let row_data = {}
        let tempStr
        try {
            //console.log(row)
            row_data.invoiceNumber = row[0]
            row_data.invoiceDate = row[1]
            row_data.companyID = row[2]
            row_data.companyName = row[3]
            row_data.subID = row[4]
            row_data.paymentCycle = row[5]
            row_data.paymentAmount = row[6]
            row_data.email = row[7]
            row_data.items = row[8]
            row_data.clientgst = row[9]
            row_data.amountinwords = row[10]
            row_data.description = row[11]
            console.log(row_data)
            if (row[13].length == 0) {
                tempStr = createPDF(row_data, row[3] + "-" + row[0] + "-" + new Date().toLocaleString(), docFile, tempFolder, pdfFolder)
                fileURL.push([tempStr])
                errors.push([""])
            } else { console.log("Pdf Generated") }
        } catch (err) {
            console.log(err)
            errors.push(["Failed"])
        }
    });
    console.log(fileURL.length)
    txnDetailSheet.getRange(3, 14, txnDetailSheet.getLastRow() - 2, 1).setValues(fileURL)
}


function createPDF(row_data, pdfName, docFile, tempFolder, pdfFolder) {

    //console.log(row_data)
    const tempFile = docFile.makeCopy(tempFolder)
    const tempDocFile = DocumentApp.openById(tempFile.getId())
    const body = tempDocFile.getBody()
    let pdfFile

    body.replaceText("{Bill No.}", row_data.invoiceNumber)
    body.replaceText("{Invoice Date}", row_data.invoiceDate)
    //body.replaceText("{Comp ID.}",row_data.companyID)
    body.replaceText("{Merchant Name}", row_data.companyName)
    //body.replaceText("{Sub ID.}",row_data.subID)
    //body.replaceText("{Pay Cycle}",row_data.paymentCycle)
    body.replaceText("{Total Amount}", row_data.paymentAmount)
    //body.replaceText("{Client Email}",row_data.email)
    body.replaceText("{Item Name}", row_data.items)
    body.replaceText("{Merchant GST}", row_data.clientgst)
    body.replaceText("{Amount Text}", row_data.amountinwords)
    body.replaceText("{Item Description}", row_data.description)

    tempDocFile.saveAndClose()
    const pdfContentBlob = tempFile.getAs(MimeType.PDF)
    pdfFile = pdfFolder.createFile(pdfContentBlob).setName(pdfName).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW).getUrl() //PDF File Handle
    tempFolder.removeFile(tempFile)
    //console.log("Inside Create PDF URL",pdfFile)
    return pdfFile

}
