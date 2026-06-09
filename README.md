# GeminiNotesToDocs

A Google Apps Script that automatically captures Gemini meeting notes from Gmail and routes them into per-project Google Docs — organized and ready for NotebookLM.

## What It Does

Every meeting with Gemini AI enabled generates a summary email from `gemini-notes@google.com`. This script runs hourly, matches each email subject against your project keywords, and inserts the notes into the right Google Doc — newest on top.

## Setup

See the full setup guide: [GeminiNotesToDocs — Setup & User Guide](https://docs.google.com/document/d/1KhUTqrbO8oxYXFffnYOMB7D_gHjgy0KvjZEMI6bJPBs/edit)

### Quick Start

1. Go to [script.google.com](https://script.google.com) and create a new project
2. Paste `GeminiNotesToDocs.gs` into the editor
3. Run `setScriptProperties()` with your Drive folder ID
4. Edit `PROJECT_MAP` in `CONFIG` to match your meetings
5. Run `setupTrigger()` to install the hourly trigger
6. Authorize when prompted

## Key Features

- **Keyword routing** — maps subject keywords to project docs
- **Safe retry** — failed messages are never skipped; retried next hour
- **Date-range backfill** — fetch notes between specific dates
- **Quarterly/monthly rotation** — keeps docs within NotebookLM's limits
- **Privacy-safe logging** — logs message IDs, not meeting titles
- **No secrets in source** — folder ID stored in Script Properties

## Configuration

Edit `CONFIG` at the top of the script:

```js
PROJECT_MAP: [
  { keyword: "DSU",                docName: "Automation Testing DSU" },
  { keyword: "Huddle",             docName: "DEV Huddle" },
  { keyword: "Sales Team Connect", docName: "[Internal] Insulet: Sales Team Connect (Daily)" },
  { keyword: "Sales Cloud",        docName: "Insulet: Sales Cloud - Daily Sync" },
  // Add your own rows here
],
ROTATION: "quarter",  // or "month"
```
