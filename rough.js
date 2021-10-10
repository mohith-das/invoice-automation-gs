// let cName = "Possier";
// let amount = "$100";
//Template ID - 16B8hN6ZDa83qACGRsQfTuYItuqXAaxKn-xPaCxxVDi4
//Folder ID (Temp)- 1uQFkLOnccQA18uQ8xpHqjTGruU0SRNUj
//Folder ID (PDFs)- 1gDpzbSV8qtSKvpK9QqTuQYyOg9x7-z0O    




function createBulkPDFs() {
    const docFile = DriveApp.getFileById("16B8hN6ZDa83qACGRsQfTuYItuqXAaxKn-xPaCxxVDi4")
    const tempFolder = DriveApp.getFolderById("1uQFkLOnccQA18uQ8xpHqjTGruU0SRNUj")
    const pdfFolder = DriveApp.getFolderById("1gDpzbSV8qtSKvpK9QqTuQYyOg9x7-z0O")
    const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("people")
    const data = currentSheet.getRange(2, 1, currentSheet.getLastRow() - 1, 4).getDisplayValues()
    let errors = []
    data.forEach(row => {
        try {
            createPDF(row[0], row[1], row[0] + "-" + new Date().toLocaleString(), docFile, tempFolder, pdfFolder)
            errors.push([""])
        } catch (err) {
            errors.push(["Failed"])
        }
    });
    currentSheet.getRange(2, 3, currentSheet.getLastRow() - 1, 1).setValues(errors)
}


function createPDF(cName, amount, pdfName, docFile, tempFolder, pdfFolder) {


    const tempFile = docFile.makeCopy(tempFolder)
    const tempDocFile = DocumentApp.openById(tempFile.getId())
    const body = tempDocFile.getBody()

    body.replaceText("{Company Name}", cName)
    body.replaceText("{Subscription Amount}", amount)
    tempDocFile.saveAndClose()
    const pdfContentBlob = tempFile.getAs(MimeType.PDF)
    pdfFolder.createFile(pdfContentBlob).setName(pdfName)
    tempFolder.removeFile(tempFile)

}
