#!/usr/bin/env bun
// Test the getUnreadEmails function directly

import { runAppleScript } from 'run-applescript';

async function testUnread() {
  const limit = 10;
  const folderPath = "inbox";

  const script = `
    tell application "Microsoft Outlook"
      try
        set theFolder to ${folderPath}
        set unreadList to {}

        -- Get unread messages directly using whose clause (avoids corrupted message issues)
        set unreadMessages to messages of theFolder whose is read is false
        set totalUnread to count of unreadMessages

        return "Found " & totalUnread & " unread messages"
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  const result = await runAppleScript(script);
  console.log("Result:", result);
}

testUnread();