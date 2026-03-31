// ─────────────────────────────────────────────────────────────
//  Expop Email Personalizer
//  Paste this entire file into Apps Script (see README for how)
// ─────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Expop Emailer")
    .addItem("Create Drafts", "createDrafts")
    .addToUi();
}

function createDrafts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 1. Read the template sheet ────────────────────────────
  var templateSheet = ss.getSheetByName("Template");
  if (!templateSheet) {
    SpreadsheetApp.getUi().alert('Sheet named "Template" not found.');
    return;
  }

  var subject = templateSheet.getRange("B1").getValue().toString().trim();
  var docUrl  = templateSheet.getRange("B2").getValue().toString().trim();

  if (!subject) {
    SpreadsheetApp.getUi().alert("Subject is empty — fill in cell B1 of the Template sheet.");
    return;
  }
  if (!docUrl) {
    SpreadsheetApp.getUi().alert("Doc URL is empty — paste your Google Doc URL into cell B2 of the Template sheet.");
    return;
  }

  // ── 2. Fetch the Google Doc as HTML ───────────────────────
  var docId = extractDocId(docUrl);
  if (!docId) {
    SpreadsheetApp.getUi().alert("Could not read the Google Doc ID from B2. Make sure you pasted the full Doc URL.");
    return;
  }

  var htmlBody = fetchDocAsHtml(docId);
  if (!htmlBody) {
    SpreadsheetApp.getUi().alert("Could not fetch the Google Doc. Make sure the Doc is shared with your Google account.");
    return;
  }

  // ── 3. Read recipients ────────────────────────────────────
  var recipientSheet = ss.getSheetByName("Recipients");
  if (!recipientSheet) {
    SpreadsheetApp.getUi().alert('Sheet named "Recipients" not found.');
    return;
  }

  var headers = recipientSheet
    .getRange(1, 1, 1, recipientSheet.getLastColumn())
    .getValues()[0]
    .map(function(h) { return h.toString().trim().toLowerCase(); });

  var colTitle     = headers.indexOf("title");
  var colFirstName = headers.indexOf("first_name");
  var colLastName  = headers.indexOf("last_name");
  var colEmail     = headers.indexOf("email");

  if ([colTitle, colFirstName, colLastName, colEmail].indexOf(-1) !== -1) {
    SpreadsheetApp.getUi().alert(
      "Recipients sheet must have columns: title, first_name, last_name, email\n" +
      "Found: " + headers.join(", ")
    );
    return;
  }

  var lastRow = recipientSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No recipients found in the Recipients sheet.");
    return;
  }

  var data = recipientSheet
    .getRange(2, 1, lastRow - 1, recipientSheet.getLastColumn())
    .getValues();

  // ── 4. Create one draft per recipient ─────────────────────
  var created = 0;
  var skipped = 0;

  for (var i = 0; i < data.length; i++) {
    var row       = data[i];
    var title     = row[colTitle].toString().trim();
    var firstName = row[colFirstName].toString().trim();
    var lastName  = row[colLastName].toString().trim();
    var email     = row[colEmail].toString().trim();

    if (!email) { skipped++; continue; }

    var salutation       = buildSalutation(title, firstName, lastName);
    var personalHtmlBody = htmlBody.replace(/{dear}/g, salutation);
    var personalPlain    = "Dear " + salutation + ",\n\n(Open in Gmail to view formatted email)";

    GmailApp.createDraft(email, subject, personalPlain, { htmlBody: personalHtmlBody });
    created++;
  }

  SpreadsheetApp.getUi().alert(
    "Done!\n" + created + " draft(s) created in your Gmail Drafts folder." +
    (skipped > 0 ? "\n" + skipped + " row(s) skipped (no email)." : "")
  );
}

// ── Helpers ───────────────────────────────────────────────────

function buildSalutation(title, firstName, lastName) {
  if (title.toLowerCase() === "dr") {
    return "Dr " + lastName;
  }
  return title + " " + firstName;
}

// Pull the file ID out of any standard Google Doc URL
function extractDocId(url) {
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

// Export the Google Doc as HTML using the Drive API
function fetchDocAsHtml(docId) {
  try {
    var exportUrl = "https://docs.google.com/feeds/download/documents/export/Export?id=" +
                    docId + "&exportFormat=html";
    var response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) return null;

    var fullHtml = response.getContentText();

    // Strip the outer <html><head>...</head><body> wrapper — keep just the body content
    var bodyMatch = fullHtml.match(/<body[^>]*>([\s\S]*)<\/body>/i);
    return bodyMatch ? bodyMatch[1].trim() : fullHtml;
  } catch (e) {
    return null;
  }
}
