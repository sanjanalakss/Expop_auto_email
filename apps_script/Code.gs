// ─────────────────────────────────────────────────────────────
//  Expop Email Personalizer
//  Paste this entire file into Apps Script
// ─────────────────────────────────────────────────────────────

var SUBJECT = "Invite to Participate in National High School Startup Challenge (By Harvard & UPenn Students)";

var EMAIL_TEMPLATE = `<p>Dear {dear},</p>

<p>Hope you're having a great week!</p>

<p>My name is Sanjana - I'm a Wharton undergrad and entrepreneur. Every year, our team of Ivy League entrepreneurs hosts <b>Expop</b>, <b>a national 6-month startup challenge</b> where high schoolers are guided step-by-step to launch and operate their own real businesses.</p>

<p>We would love to invite your school to learn more and participate in our challenge.</p>

<p>Last year, Expop developed 100+ high school startups (with over 60% of participants having no business experience prior to this challenge). We guide students through ideation, product development, marketing, securing their first customers & pitching to investors.</p>

<p>Expop Challenge Highlights:</p>
<ul>
  <li>Live workshops by <b>Fortune 500 Executives</b> from Google, HP, LinkedIn, Slack</li>
  <li>Mentorship from Ivy League undergrad founders from <b>Harvard, UPenn, Princeton, Brown, Yale, Dartmouth</b></li>
  <li><b>$7,500 Cash Prize Pool</b> for Top 5 winners</li>
  <li>Open to students Grades 9 to 12</li>
</ul>

<p>~</p>

<p>We've currently opened registration for our 2026 cohort and would be glad to schedule a brief Zoom call with you or your Business / College & Career Readiness team to share how the program works and how it can support your students.</p>

<p>I'd be happy to share additional materials or a brief overview deck if helpful.</p>

<p>Hope to hear from you and looking forward to seeing how we can impact your students!</p>

<p>Warm regards,<br>
Sanjana<br>
<b>Operating Lead at Expop</b></p>`;

// ─────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Expop Emailer")
    .addItem("Create Drafts", "createDrafts")
    .addToUi();
}

function createDrafts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

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
    var personalHtmlBody = EMAIL_TEMPLATE.replace(/{dear}/g, salutation);
    var personalPlain    = "Dear " + salutation + ",\n\n(Open in Gmail to view formatted email)";

    GmailApp.createDraft(email, SUBJECT, personalPlain, { htmlBody: personalHtmlBody });
    created++;
  }

  SpreadsheetApp.getUi().alert(
    "Done!\n" + created + " draft(s) created in your Gmail Drafts folder." +
    (skipped > 0 ? "\n" + skipped + " row(s) skipped (no email)." : "")
  );
}

function buildSalutation(title, firstName, lastName) {
  if (title.toLowerCase() === "dr") {
    return "Dr " + lastName;
  }
  return title + " " + firstName;
}
