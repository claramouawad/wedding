const SHEET_NAME = "RSVP";

function doPost(e) {
  const sheet = getSheet();
  const payload = JSON.parse(e.postData.contents);
  const submittedAt = payload.submittedAt || new Date().toISOString();
  const language = payload.language || "";
  const guests = payload.guests || [];

  guests.forEach(function(guest, index) {
    sheet.appendRow([
      new Date(),
      submittedAt,
      language,
      index + 1,
      guest.name || "",
      guest.ceremony || "",
      guest.reception || "",
      guest.food || "",
      guest.song || ""
    ]);
  });

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Received At",
      "Submitted At",
      "Language",
      "Guest Number",
      "Name",
      "Ceremony",
      "Reception",
      "Food Allergies Or Intolerances",
      "Song Suggestion"
    ]);
  }

  return sheet;
}
