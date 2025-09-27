// Simple script to fetch emails using the Outlook MCP
import { runAppleScript } from 'run-applescript';

async function fetchEmails() {
    try {
        // Fetch unread emails
        const unreadScript = `
            tell application "Microsoft Outlook"
                try
                    set unreadMessages to messages of inbox whose is read is false
                    set unreadCount to count of unreadMessages
                    return "Unread emails: " & unreadCount
                on error errMsg
                    return "Error checking unread: " & errMsg
                end try
            end tell
        `;

        // Fetch recent emails
        const recentScript = `
            tell application "Microsoft Outlook"
                try
                    set allMessages to messages 1 through 5 of inbox
                    set emailCount to count of allMessages
                    set summaryList to {}

                    repeat with msg in allMessages
                        try
                            set msgSubject to subject of msg
                            set msgRead to is read of msg
                            set msgSummary to "• " & msgSubject & " (Read: " & msgRead & ")"
                            set end of summaryList to msgSummary
                        on error
                            set end of summaryList to "• [Error reading message]"
                        end try
                    end repeat

                    return "Recent " & emailCount & " emails:\\n" & (summaryList as string)
                on error errMsg
                    return "Error fetching recent: " & errMsg
                end try
            end tell
        `;

        const unreadResult = await runAppleScript(unreadScript);
        const recentResult = await runAppleScript(recentScript);

        console.log("## 📧 Direct Email Fetch\n");
        console.log(unreadResult);
        console.log("\n" + recentResult);

    } catch (error) {
        console.error("Error fetching emails:", error.message);
    }
}

fetchEmails();