function onOpen() {
  DocumentApp.getUi()
    .createMenu("SLT Notes Tools")
    .addItem("Start New Meeting Template", "insertMeetingTemplate")
    .addItem("Import Latest Vote Results", "importVoteResultsToDoc")
    .addItem("Link Meeting Brief Document", "linkMeetingBriefDoc")
    .addItem("Archive Past Meetings", "archiveOldMeetings")
    .addToUi();
}

function insertMeetingTemplate() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy");

  const cursor = body.editAsText();
  
  body.insertParagraph(0, `SLT Meeting — ${dateStr}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  const metaParagraph = body.insertParagraph(1, "Attendees: Mrs. Sharples, President, Vice President, Secretary, Treasurer, Historian, and Technology Specialist. Representatives:\n");
  metaParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING3);

  body.insertParagraph(2, "Agenda & Discussion").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.insertListItem(3, "?");
  body.insertListItem(4, "?");
  body.insertListItem(5, "?");
  body.insertListItem(6, "?");

  body.insertParagraph(7, "\nAction Items").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const actionTable = body.insertTable(8, [
    ["Task / Decision", "Owner", "Due Date"],
    ["", "", ""],
    ["", "", ""]
  ]);
  actionTable.getRow(0).setBackgroundColor("#e8eaed");

  body.insertParagraph(9, "\n--- PAGE BREAK ---\n").setHeading(DocumentApp.ParagraphHeading.NORMAL);
  DocumentApp.getUi().alert("New meeting template inserted at the top of the document!");
}

function linkMeetingBriefDoc() {
  const ui = DocumentApp.getUi();
  const response = ui.prompt("Link Meeting Brief Document", "Enter the Google Doc ID or URL for the Meeting Brief:", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    let input = response.getResponseText().trim();
    let docId = input;
    if (input.includes("/d/")) {
      const match = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (match) {
        docId = match[1];
      }
    }
    if (docId) {
      PropertiesService.getDocumentProperties().setProperty("SLT_BRIEF_DOC_ID", docId);
      ui.alert("Meeting Brief document successfully linked!");
    } else {
      ui.alert("Invalid ID or URL entered.");
    }
  }
}

function importVoteResultsToDoc() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  const props = PropertiesService.getDocumentProperties();
  const votingDocId = props.getProperty("SLT_BRIEF_DOC_ID");

  if (!votingDocId) {
    DocumentApp.getUi().alert("No linked voting brief found. Use the menu option 'Link Meeting Brief Document' to connect it.");
    return;
  }

  try {
    const votingDoc = DocumentApp.openById(votingDocId);
    const votingText = votingDoc.getBody().getText();

    body.appendParagraph("\nImported Voting Results Snapshot").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(votingText);
    
    DocumentApp.getUi().alert("Latest voting results successfully imported into the document!");
  } catch (e) {
    DocumentApp.getUi().alert("Could not access the voting document. Check if it was deleted or if the ID is correct.");
  }
}

function archiveOldMeetings() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getChildren();

  let archiveDoc = null;
  const props = PropertiesService.getDocumentProperties();
  let archiveId = props.getProperty("SLT_ARCHIVE_DOC_ID");

  if (archiveId) {
    try {
      archiveDoc = DocumentApp.openById(archiveId);
    } catch (e) {
      archiveDoc = null;
    }
  }

  if (!archiveDoc) {
    archiveDoc = DocumentApp.create("SLT Meeting Notes Archive");
    props.setProperty("SLT_ARCHIVE_DOC_ID", archiveDoc.getId());
  }

  const archiveBody = archiveDoc.getBody();
  archiveBody.appendParagraph(`Archived on: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")}\n`).setHeading(DocumentApp.ParagraphHeading.HEADING3);

  let movedCount = 0;
  if (paragraphs.length > 50) {
    for (let i = 50; i < paragraphs.length; i++) {
      const el = paragraphs[i].copy();
      archiveBody.appendParagraph(el);
      paragraphs[i].removeFromParent();
      movedCount++;
    }
    DocumentApp.getUi().alert(`Archived ${movedCount} old paragraphs to the SLT Archive document.`);
  } else {
    DocumentApp.getUi().alert("Document is still short. No archiving needed yet.");
  }
}
