function onOpen() {
  SpreadsheetApp.getUi().createMenu("📩 Gmail Exporter")
    .addItem("📤 פתח סרגל ייצוא", "showSidebar")
    .addToUi();
}

/**
 * פותח סרגל צד עם כפתור לייצוא
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("📩 ייצוא מיילים")
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * פונקציית הייצוא הראשית – מופעלת מהסרגל
 */
function exportGmailToDriveIndex(query, folderUrl) {
  const folderId = extractFolderIdFromUrl(folderUrl);
  if (!folderId) throw new Error("❌ קישור לתיקיה לא תקין.");
  const folder = DriveApp.getFolderById(folderId);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Gmail Files Index") ||
                SpreadsheetApp.getActiveSpreadsheet().insertSheet("Gmail Files Index");
  sheet.clear();
  sheet.appendRow([
    "Thread ID", "Message ID", "File Type", "Date", "From", "To", "Subject",
    "Snippet", "File Name", "Drive URL", "Attachment Message ID"
  ]);

  const threads = GmailApp.search(query);
  console.log(`📥 נמצאו ${threads.length} שרשורים.`);

  threads.forEach(thread => {
    const threadId = thread.getId();
    const messages = thread.getMessages();

    messages.forEach((msg, msgIndex) => {
      const messageId = `${threadId}_${msgIndex + 1}`;
      const date = msg.getDate();
      const from = msg.getFrom();
      const to = msg.getTo();
      const subject = msg.getSubject();
      const snippet = msg.getPlainBody().substring(0, 100).replace(/\n/g, " ");

      // שמירת ההודעה כ-PDF
      const doc = DocumentApp.create(`Msg_${messageId}`);
      const body = doc.getBody();
      body.appendParagraph(`נושא: ${subject}`);
      body.appendParagraph(`מאת: ${from}`);
      body.appendParagraph(`אל: ${to}`);
      body.appendParagraph(`תאריך: ${date}`);
      body.appendParagraph(`\n${msg.getPlainBody()}`);
      doc.saveAndClose();

      const pdfBlob = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
      const pdfFile = folder.createFile(pdfBlob);
      pdfFile.setName(`Msg_${messageId}.pdf`);
      DriveApp.getFileById(doc.getId()).setTrashed(true);

      // שורת הודעה
      sheet.appendRow([
        threadId, messageId, "message", date, from, to, subject,
        snippet, pdfFile.getName(), pdfFile.getUrl(), ""
      ]);

      // קבצים מצורפים
      const attachments = msg.getAttachments();
      attachments.forEach(att => {
        try {
          if (!att || att.getBytes().length === 0) {
            console.warn(`⚠️ צרופה ריקה — דילוג`);
            return;
          }
          const file = folder.createFile(att);
          sheet.appendRow([
            threadId, "", "attachment", "", "", "", "", "",
            file.getName(), file.getUrl(), messageId
          ]);
        } catch (e) {
          console.error(`❌ שגיאה בשמירת צרופה: ${e}`);
        }
      });
    });
  });

  return `🎉 הסתיים ייצוא של ${threads.length} שרשורים!`;
}

/**
 * חילוץ מזהה תיקיה מתוך קישור דרייב
 */
function extractFolderIdFromUrl(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}
