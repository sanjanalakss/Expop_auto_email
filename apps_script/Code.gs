// ─────────────────────────────────────────────────────────────
//  Expop Email Personalizer
//  Paste this entire file into Apps Script (see README for how)
// ─────────────────────────────────────────────────────────────

// Adds a custom menu when the spreadsheet opens
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Expop Emailer")
    .addItem("Create Drafts", "createDrafts")
    .addToUi();
}

// Main function — reads recipients, builds personalised drafts
function createDrafts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 1. Read the template ──────────────────────────────────
  var templateSheet = ss.getSheetByName("Template");
  if (!templateSheet) {
    SpreadsheetApp.getUi().alert('Sheet named "Template" not found. Please check the tab name.');
    return;
  }

  var subject = templateSheet.getRange("B1").getValue().toString().trim();
  var body    = templateSheet.getRange("B2").getValue().toString().trim();

  if (!subject || !body) {
    SpreadsheetApp.getUi().alert("Subject (B1) or Body (B2) is empty in the Template sheet.");
    return;
  }

  // ── 2. Read recipients ────────────────────────────────────
  var recipientSheet = ss.getSheetByName("Recipients");
  if (!recipientSheet) {
    SpreadsheetApp.getUi().alert('Sheet named "Recipients" not found. Please check the tab name.');
    return;
  }

  // Find column positions from header row (case-insensitive)
  var headers = recipientSheet.getRange(1, 1, 1, recipientSheet.getLastColumn())
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
    SpreadsheetApp.getUi().alert("No recipients found (sheet has no data rows).");
    return;
  }

  var data = recipientSheet.getRange(2, 1, lastRow - 1, recipientSheet.getLastColumn()).getValues();

  // ── 3. Create one draft per recipient ─────────────────────
  var created = 0;
  var skipped = 0;

  for (var i = 0; i < data.length; i++) {
    var row       = data[i];
    var title     = row[colTitle].toString().trim();
    var firstName = row[colFirstName].toString().trim();
    var lastName  = row[colLastName].toString().trim();
    var email     = row[colEmail].toString().trim();

    if (!email) { skipped++; continue; }

    var salutation = buildSalutation(title, firstName, lastName);
    var personalBody = body.replace(/{dear}/g, salutation);

    GmailApp.createDraft(email, subject, personalBody);
    created++;
  }

  SpreadsheetApp.getUi().alert(
    "Done!\n" +
    created + " draft(s) created in your Gmail Drafts folder." +
    (skipped > 0 ? "\n" + skipped + " row(s) skipped (no email address)." : "")
  );
}

// Salutation rules:
//   Dr  → Dr {last_name}
//   Mr  → Mr {first_name}
//   Ms  → Ms {first_name}
function buildSalutation(title, firstName, lastName) {
  if (title.toLowerCase() === "dr") {
    return "Dr " + lastName;
  }
  return title + " " + firstName;
}
