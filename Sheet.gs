function parseSLTData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = ss.getSheets()[0];
  const data = responseSheet.getDataRange().getValues();

  if (data.length < 2) return null;

  const headers = data[0];
  const totalSubmissions = data.length - 1;
  const summaryRows = [["Question / Activity", "Type", "Breakdown / Average", "Top Result / Status"]];
  const docSections = {
    approved: [],
    undiscussed: [],
    other: [],
    ratings: []
  };

  for (let col = 1; col < headers.length; col++) {
    const question = headers[col].toString().trim();
    if (!question) continue;

    const responses = [];
    for (let row = 1; row < data.length; row++) {
      const val = data[row][col];
      if (val !== "" && val !== null && val !== undefined) {
        responses.push(val);
      }
    }

    if (responses.length === 0) continue;

    const isNumericRating = responses.every(v => typeof v === "number" && v >= 1 && v <= 10);

    if (isNumericRating) {
      const sum = responses.reduce((acc, n) => acc + n, 0);
      const avg = (sum / responses.length).toFixed(2);
      const sortedNums = [...responses].sort((a, b) => a - b);
      const mid = Math.floor(sortedNums.length / 2);
      const median = sortedNums.length % 2 !== 0
        ? sortedNums[mid]
        : ((sortedNums[mid - 1] + sortedNums[mid]) / 2).toFixed(1);

      const breakdown = `Avg: ${avg}/10 | Median: ${median}/10 (${responses.length} ratings)`;
      summaryRows.push([question, "1-10 Rating", breakdown, `${avg} / 10`]);
      docSections.ratings.push(`${question}: Average ${avg}/10 (Median: ${median}/10)`);
    } else {
      const counts = {};
      responses.forEach(r => {
        const choice = r.toString().trim();
        counts[choice] = (counts[choice] || 0) + 1;
      });

      const sortedChoices = Object.entries(counts).sort((a, b) => b[1] - a[1]);
      const breakdown = sortedChoices
        .map(([choice, count]) => `${choice}: ${count} (${Math.round((count / responses.length) * 100)}%)`)
        .join(" | ");

      const [topChoice, topCount] = sortedChoices[0];
      const isUndiscussed = question.toLowerCase().includes("undiscussed");
      const status = isUndiscussed
        ? `Needs Discussion (${topChoice}: ${topCount}/${responses.length})`
        : `${topChoice} (${topCount}/${responses.length})`;

      summaryRows.push([question, "Multiple Choice", breakdown, status]);

      const entryText = `${question} — ${breakdown}`;
      if (isUndiscussed) {
        docSections.undiscussed.push(entryText);
      } else if (topChoice.toLowerCase().startsWith("yes") || topChoice.toLowerCase().includes("approve")) {
        docSections.approved.push(entryText);
      } else {
        docSections.other.push(entryText);
      }
    }
  }

  return { ss, totalSubmissions, summaryRows, docSections };
}

function updateSummaryTabOnly() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;

  try {
    const parsed = parseSLTData();
    if (!parsed) return;

    const { ss, summaryRows } = parsed;
    let summarySheet = ss.getSheetByName("Vote Summary") || ss.insertSheet("Vote Summary");
    summarySheet.clear();
    summarySheet.getRange(1, 1, summaryRows.length, 4).setValues(summaryRows);
    summarySheet.getRange("A1:D1").setFontWeight("bold").setBackground("#e8eaed");
    summarySheet.autoResizeColumns(1, 4);
  } finally {
    lock.releaseLock();
  }
}

function generateAnonymousSLTResults() {
  const parsed = parseSLTData();
  if (!parsed) {
    SpreadsheetApp.getUi().alert("No responses recorded yet.");
    return;
  }

  updateSummaryTabOnly();

  const props = PropertiesService.getDocumentProperties();
  let docId = props.getProperty("SLT_BRIEF_DOC_ID");
  let doc = null;

  if (docId) {
    try {
      doc = DocumentApp.openById(docId);
    } catch (e) {
      doc = null;
    }
  }

  if (!doc) {
    doc = DocumentApp.create("SLT Anonymous Vote Results (Live)");
    props.setProperty("SLT_BRIEF_DOC_ID", doc.getId());
  }

  const body = doc.getBody();
  body.clear();

  const { totalSubmissions, docSections } = parsed;
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d, yyyy 'at' h:mm a");

  body.appendParagraph("SLT Anonymous Voting Results")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Last Updated: ${timestamp} | Total Anonymous Submissions: ${totalSubmissions}\n`);

  const writeSection = (heading, items) => {
    if (items.length === 0) return;
    body.appendParagraph(heading).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    items.forEach(item => body.appendListItem(item));
  };

  writeSection("✅ Top Choices / Approved Items", docSections.approved);
  writeSection("⚠️ Undiscussed Activities (Needs Floor Debate)", docSections.undiscussed);
  writeSection("📋 Other Activity Results", docSections.other);
  writeSection("⭐ Overall Rating (1–10 Scale)", docSections.ratings);

  doc.saveAndClose();

  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:sans-serif;padding:8px;">
      <p><strong>Meeting Brief updated!</strong></p>
      <p><a href="${doc.getUrl()}" target="_blank">Open Persistent Google Doc</a></p>
    </div>`
  ).setWidth(320).setHeight(110);
  SpreadsheetApp.getUi().showModalDialog(html, "SLT Doc Updated");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("SLT Tools")
    .addItem("Refresh Vote Summary Tab", "updateSummaryTabOnly")
    .addItem("Update Meeting Brief (Google Doc)", "generateAnonymousSLTResults")
    .addToUi();
}
