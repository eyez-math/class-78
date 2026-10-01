/**
 * שעות 7-8 — שרת נתונים מהגיליון
 * מחזיר את כל הלשוניות הגלויות: ערכים מוצגים + טווחי תאים ממוזגים.
 * פריסה: Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone
 */
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets()
    .filter(sh => !sh.isSheetHidden())
    .map(sh => {
      const range = sh.getDataRange(); // מתחיל תמיד ב-A1
      return {
        name: sh.getName(),
        values: range.getDisplayValues(),
        merges: range.getMergedRanges().map(m => [
          m.getRow() - 1, m.getColumn() - 1, m.getNumRows(), m.getNumColumns()
        ])
      };
    });

  return ContentService
    .createTextOutput(JSON.stringify({ updated: new Date().toISOString(), sheets: sheets }))
    .setMimeType(ContentService.MimeType.JSON);
}
