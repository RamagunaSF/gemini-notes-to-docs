/**
 * GeminiNotesToDocs.gs
 *
 * Automatically captures Gemini meeting notes from Gmail and appends them
 * to per-project Google Docs, ready for use in NotebookLM.
 *
 * SETUP STEPS:
 *  1. Go to https://script.google.com and create a new project
 *  2. Paste this entire file into the editor
 *  3. Run setScriptProperties() once to save your folder ID securely
 *  4. Edit PROJECT_MAP below to match your meeting keywords
 *  5. Run setupTrigger() once to install the hourly trigger
 *  6. Authorize the script when prompted (Gmail + Drive + Docs access)
 */

// ─────────────────────────────────────────────────────────────────────────────
// CONFIGURATION — edit PROJECT_MAP here; store DRIVE_FOLDER_ID via setScriptProperties()
// ─────────────────────────────────────────────────────────────────────────────

const CONFIG = {

  // Label applied to emails after successful processing, to prevent re-processing.
  PROCESSED_LABEL: "gemini-notes-processed",

  // Map of subject-line keywords (case-insensitive) → project document name.
  // The FIRST matching keyword wins — put more specific phrases before generic ones.
  // Example: "Sales Team Connect" must appear before "Sales" to avoid false matches.
  PROJECT_MAP: [
    { keyword: "DSU",                docName: "Automation Testing DSU" },
    { keyword: "Huddle",             docName: "DEV Huddle" },
    { keyword: "Huddel",             docName: "DEV Huddle" },           // intentional typo variant seen in Gemini subjects
    { keyword: "Sales Team Connect", docName: "[Internal] Insulet: Sales Team Connect (Daily)" },
    { keyword: "Sales Cloud",        docName: "Insulet: Sales Cloud - Daily Sync" },
    { keyword: "Test",               docName: "Test Automation" },
  ],

  // Fallback doc name when no keyword matches the subject.
  // Set to null to skip unmatched emails entirely instead of routing them.
  UNMATCHED_DOC_NAME: "Uncategorized Meeting Notes",

  // Doc rotation period: "quarter" (e.g. 2026-Q2) or "month" (e.g. 2026-06).
  // A new doc is created each period to keep docs within NotebookLM's 200 MB / 5 M word limit.
  ROTATION: "quarter",
};


// ─────────────────────────────────────────────────────────────────────────────
// MAIN FUNCTION — called by the time-based trigger every hour
// ─────────────────────────────────────────────────────────────────────────────

function processGeminiNotes() {
  const label = getOrCreateLabel_(CONFIG.PROCESSED_LABEL);

  // Build a date-scoped query at runtime so only TODAY's emails are ever fetched.
  // This works correctly on first run, new machines, and shared setups — no stale history pulled in.
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const query = `from:gemini-notes@google.com -label:${CONFIG.PROCESSED_LABEL} after:${today}`;

  // ── BACKFILL / DATE-RANGE OVERRIDE ────────────────────────────────────────
  // To reprocess or import emails between two specific dates, comment out the
  // live query above and uncomment the line below, replacing the dates as needed.
  //
  // When to use:
  //   • The trigger missed a day or was paused
  //   • You want to rebuild a doc from a specific period
  //   • You're importing historical notes for the first time
  //
  // Note: "after" is inclusive, "before" is exclusive — set "before" to the
  // day AFTER the last date you want (e.g. before:2026/05/20 captures up to May 19).
  // Also remove -label:${CONFIG.PROCESSED_LABEL} if you need to reprocess already-labeled threads.
  //
  // const query = `from:gemini-notes@google.com -label:${CONFIG.PROCESSED_LABEL} after:2026/05/17 before:2026/05/20`;
  // ─────────────────────────────────────────────────────────────────────────

  const threads = GmailApp.search(query, 0, 50);

  if (threads.length === 0) {
    Logger.log("No new Gemini notes emails found.");
    return;
  }

  Logger.log(`Found ${threads.length} email thread(s) to process.`);

  const folderId = PropertiesService.getScriptProperties().getProperty("DRIVE_FOLDER_ID");
  if (!folderId) {
    Logger.log("ERROR: DRIVE_FOLDER_ID not set. Run setScriptProperties() first.");
    return;
  }
  const folder = DriveApp.getFolderById(folderId);

  threads.forEach(thread => {
    const messages = thread.getMessages();

    // Only label the thread if ALL messages processed successfully.
    // This ensures any failed message is retried on the next hourly run.
    let allSucceeded = true;

    messages.forEach(message => {
      try {
        processMessage_(message, folder);
      } catch (e) {
        allSucceeded = false;
        // Log message ID instead of subject to avoid leaking meeting titles
        // (client names, project names) into Cloud Logs visible to Workspace admins.
        Logger.log(`Error processing message ID "${message.getId()}": ${e.message}`);
      }
    });

    if (allSucceeded) {
      thread.addLabel(label);
    } else {
      Logger.log(`Thread ID "${thread.getId()}" NOT labeled — will retry on next run.`);
    }
  });
}


// ─────────────────────────────────────────────────────────────────────────────
// CORE LOGIC
// ─────────────────────────────────────────────────────────────────────────────

function processMessage_(message, folder) {
  const subject = message.getSubject();
  const date    = message.getDate();
  const body    = extractNotesFromEmail_(message);

  const baseDocName = resolveDocName_(subject);
  if (!baseDocName) {
    Logger.log(`Skipping unmatched email: "${subject}"`);
    return;
  }

  // Append the rotation period so each quarter/month gets its own doc.
  // Example: "DEV Huddle — 2026-Q2"
  const docName = `${baseDocName} — ${getCurrentPeriod_()}`;

  const doc = getOrCreateDoc_(docName, folder);
  appendMeetingNotes_(doc, subject, date, body);

  Logger.log(`Appended notes from message ID "${message.getId()}" → "${docName}"`);
}

/**
 * Returns the current rotation period string based on CONFIG.ROTATION.
 *   "quarter" → "2026-Q2"
 *   "month"   → "2026-06"
 */
function getCurrentPeriod_() {
  const now  = new Date();
  const year = now.getFullYear();

  if (CONFIG.ROTATION === "month") {
    const month = String(now.getMonth() + 1).padStart(2, "0");
    return `${year}-${month}`;
  }

  // Default: quarterly rotation
  const quarter = Math.ceil((now.getMonth() + 1) / 3);
  return `${year}-Q${quarter}`;
}

/**
 * Matches the email subject against PROJECT_MAP keywords (case-insensitive).
 * Returns the target doc name, UNMATCHED_DOC_NAME, or null to skip entirely.
 */
function resolveDocName_(subject) {
  const lowerSubject = subject.toLowerCase();
  for (const entry of CONFIG.PROJECT_MAP) {
    if (lowerSubject.includes(entry.keyword.toLowerCase())) {
      return entry.docName;
    }
  }
  return CONFIG.UNMATCHED_DOC_NAME;
}

/**
 * Finds an existing Google Doc by name in the folder, or creates a new one.
 * Newly created docs get a title heading and subtitle so they're ready for NotebookLM.
 */
function getOrCreateDoc_(docName, folder) {
  const files = folder.getFilesByName(docName);
  if (files.hasNext()) {
    return DocumentApp.openById(files.next().getId());
  }

  const doc  = DocumentApp.create(docName);
  const file = DriveApp.getFileById(doc.getId());

  // Move the new doc from root into the target folder.
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  const body     = doc.getBody();
  const title    = body.appendParagraph(docName);
  title.setHeading(DocumentApp.ParagraphHeading.TITLE);

  const subtitle = body.appendParagraph("Auto-generated from Gemini meeting notes");
  subtitle.setHeading(DocumentApp.ParagraphHeading.SUBTITLE);

  body.appendHorizontalRule();
  doc.saveAndClose();

  return DocumentApp.openById(doc.getId());
}

/**
 * Inserts a new meeting section at the TOP of the document (just after the title)
 * so the latest notes always appear first when the doc is opened.
 *
 * Insert order is reversed because each insertParagraph(1, ...) pushes the
 * previous insert down by one — the last call ends up at the top.
 */
function appendMeetingNotes_(doc, subject, date, notesText) {
  const body    = doc.getBody();
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy 'at' h:mm a");

  body.insertHorizontalRule(1);
  body.insertParagraph(1, notesText);

  const subheading = body.insertParagraph(1, subject);
  subheading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const heading = body.insertParagraph(1, dateStr);
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  doc.saveAndClose();
}

/**
 * Extracts plain-text body from the email.
 * Falls back to stripping HTML tags if no plain-text body is available.
 */
function extractNotesFromEmail_(message) {
  const plain = message.getPlainBody();
  if (plain && plain.trim().length > 0) {
    return plain.trim();
  }
  return message.getBody()
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .trim();
}

/**
 * Returns the Gmail label by name, creating it if it doesn't exist yet.
 */
function getOrCreateLabel_(labelName) {
  return GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
}


// ─────────────────────────────────────────────────────────────────────────────
// ONE-TIME SETUP — run these functions once from the Apps Script editor
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Run ONCE to securely store your Drive folder ID in Script Properties.
 * Keeps the folder ID out of source code so it's not exposed in version history.
 * Replace the value below with your actual folder ID before running.
 */
function setScriptProperties() {
  PropertiesService.getScriptProperties().setProperty(
    "DRIVE_FOLDER_ID",
    "1t1j9ouK86pDeP4EMBsnlHRxgajt0c6zr"  // ← replace with your folder ID
  );
  Logger.log("DRIVE_FOLDER_ID saved to Script Properties.");
}

/**
 * Run ONCE to install the hourly trigger.
 * Removes any existing trigger first to avoid duplicates.
 * After this, processGeminiNotes() runs automatically every hour.
 */
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "processGeminiNotes")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("processGeminiNotes")
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log("Trigger installed: processGeminiNotes will run every hour.");
}

/**
 * Run this to pause the automation by removing the hourly trigger.
 * Re-run setupTrigger() to resume.
 */
function removeTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "processGeminiNotes")
    .forEach(t => ScriptApp.deleteTrigger(t));

  Logger.log("Trigger removed.");
}
