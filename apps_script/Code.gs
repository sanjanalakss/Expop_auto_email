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

  // ── 2. Open the Google Doc and convert to HTML ────────────
  var docId = extractDocId(docUrl);
  if (!docId) {
    SpreadsheetApp.getUi().alert("Could not read the Google Doc ID from B2. Make sure you pasted the full Doc URL.");
    return;
  }

  var htmlBody;
  try {
    htmlBody = docToHtml(docId);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Could not open the Google Doc.\n\nError: " + e.message);
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

function extractDocId(url) {
  var match = url.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
  if (match) return match[1];
  var fallback = url.match(/[a-zA-Z0-9_-]{25,}/);
  return fallback ? fallback[0] : null;
}

// ── Google Doc → HTML conversion ──────────────────────────────
// Uses DocumentApp (always accessible, no extra permissions needed)

function docToHtml(docId) {
  var doc  = DocumentApp.openById(docId);
  var body = doc.getBody();
  var n    = body.getNumChildren();
  var html = '';
  var inList = false;
  var listTag = '';

  for (var i = 0; i < n; i++) {
    var child = body.getChild(i);
    var type  = child.getType();

    if (type === DocumentApp.ElementType.LIST_ITEM) {
      var li       = child.asListItem();
      var glyph    = li.getGlyphType();
      var ordered  = (glyph === DocumentApp.GlyphType.DECIMAL ||
                      glyph === DocumentApp.GlyphType.LATIN_UPPER ||
                      glyph === DocumentApp.GlyphType.LATIN_LOWER ||
                      glyph === DocumentApp.GlyphType.ROMAN_UPPER ||
                      glyph === DocumentApp.GlyphType.ROMAN_LOWER);
      var tag = ordered ? 'ol' : 'ul';

      if (!inList) {
        html   += '<' + tag + '>';
        inList  = true;
        listTag = tag;
      } else if (tag !== listTag) {
        html   += '</' + listTag + '><' + tag + '>';
        listTag = tag;
      }
      html += '<li>' + elementToHtml(li) + '</li>';

    } else {
      if (inList) {
        html   += '</' + listTag + '>';
        inList  = false;
        listTag = '';
      }

      if (type === DocumentApp.ElementType.PARAGRAPH) {
        var para    = child.asParagraph();
        var heading = para.getHeading();
        var inner   = elementToHtml(para);

        if      (heading === DocumentApp.ParagraphHeading.HEADING1) html += '<h1>' + inner + '</h1>';
        else if (heading === DocumentApp.ParagraphHeading.HEADING2) html += '<h2>' + inner + '</h2>';
        else if (heading === DocumentApp.ParagraphHeading.HEADING3) html += '<h3>' + inner + '</h3>';
        else if (inner.trim() === '')                               html += '<br>';
        else                                                         html += '<p>' + inner + '</p>';
      }
    }
  }

  if (inList) html += '</' + listTag + '>';
  return html;
}

function elementToHtml(element) {
  var html = '';
  for (var i = 0; i < element.getNumChildren(); i++) {
    var child = element.getChild(i);
    if (child.getType() === DocumentApp.ElementType.TEXT) {
      html += textToHtml(child.asText());
    }
  }
  return html;
}

function textToHtml(textEl) {
  var raw     = textEl.getText();
  if (!raw) return '';

  var indices = textEl.getTextAttributeIndices();
  if (indices.length === 0) indices = [0];
  indices.push(raw.length);

  var result = '';
  for (var i = 0; i < indices.length - 1; i++) {
    var start = indices[i];
    var end   = indices[i + 1];
    var chunk = raw.substring(start, end);
    if (!chunk) continue;

    // Escape HTML special characters
    chunk = chunk.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

    if (textEl.isUnderline(start)) chunk = '<u>'  + chunk + '</u>';
    if (textEl.isItalic(start))    chunk = '<i>'  + chunk + '</i>';
    if (textEl.isBold(start))      chunk = '<b>'  + chunk + '</b>';

    result += chunk;
  }
  return result;
}
