function generateDocAndMove(name, email, contact, cetNumber, row, sheet) {
  try {
    const doc = DocumentApp.create(`${name}_Admission_Letter`); // Create the document with a specific name
    const body = doc.getBody();

    // Add content to the document
    body.appendParagraph("To,");
    body.appendParagraph("Subject: Admission Confirmation");
    body.appendParagraph(`Date: ${new Date().toLocaleDateString()}`);
    body.appendParagraph("\nContent:\n");
    body.appendParagraph(`Dear ${name},`);
    body.appendParagraph(`Email Address: ${email}`);
    body.appendParagraph(`Contact Number: ${contact}`);
    body.appendParagraph(`CET Registration Number: ${cetNumber}`);
    body.appendParagraph("\nThank you for your interest!");

    Logger.log("Document content:\n" + body.getText());

    if (!body.getText()) {
      Logger.log("Error: Document body is empty.");
      sheet.getRange(row, 15).setValue("Error: Document Not Generated");
      return;
    }

    // Create or access the "Admission Letters" folder in Drive
    const folderName = "Admission Letters";
    let folder = DriveApp.getFoldersByName(folderName).hasNext()
      ? DriveApp.getFoldersByName(folderName).next()
      : DriveApp.createFolder(folderName);

    // Move the document into the folder
    const file = DriveApp.getFileById(doc.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file); // Ensure the file is only in the specified folder

    Logger.log(`Document saved in folder: ${folder.getName()} (${folder.getUrl()})`);

    sheet.getRange(row, 15).setValue(`Document Generated: ${folder.getUrl()}`);
  } catch (error) {
    Logger.log("Error in generateDocAndMove: " + error.message);
    sheet.getRange(row, 15).setValue("Error: Document Not Generated");
  }
}