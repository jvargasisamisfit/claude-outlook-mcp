// Debug script to check calendar sources
import { runAppleScript } from 'run-applescript';

async function debugCalendar() {
    try {
        // Check which calendars exist
        const calendarListScript = `
            tell application "Microsoft Outlook"
                try
                    set calendarList to {}
                    set allCalendars to every calendar

                    repeat with cal in allCalendars
                        try
                            set calName to name of cal
                            set calID to id of cal
                            set end of calendarList to "Calendar: " & calName & " (ID: " & calID & ")"
                        on error
                            set end of calendarList to "Error reading calendar"
                        end try
                    end repeat

                    return "Available Calendars:\\n" & (calendarList as string)
                on error errMsg
                    return "Error listing calendars: " & errMsg
                end try
            end tell
        `;

        // Check events specifically for Monday Sep 29
        const mondayEventsScript = `
            tell application "Microsoft Outlook"
                try
                    set targetDate to date "Monday, September 29, 2025 12:00:00 AM"
                    set endDate to targetDate + 1 * days

                    -- Get events from main calendar only
                    set mainCalendar to calendar id 119
                    set mondayEvents to every calendar event of mainCalendar whose start time >= targetDate and start time < endDate

                    set eventList to "Events for Monday Sep 29 (Main Calendar):\\n"

                    repeat with evt in mondayEvents
                        try
                            set evtSubject to subject of evt
                            set evtTime to start time of evt
                            set evtOrganizer to organizer of evt
                            set evtStatus to free busy status of evt

                            set eventList to eventList & "• " & evtSubject & " at " & (time string of evtTime) & " (Organizer: " & name of evtOrganizer & ", Status: " & evtStatus & ")\\n"
                        on error
                            set eventList to eventList & "• Error reading event\\n"
                        end try
                    end repeat

                    return eventList
                on error errMsg
                    return "Error checking Monday events: " & errMsg
                end try
            end tell
        `;

        console.log("## Calendar Debug Information\n");

        const calendars = await runAppleScript(calendarListScript);
        console.log(calendars);
        console.log("\n");

        const mondayEvents = await runAppleScript(mondayEventsScript);
        console.log(mondayEvents);

    } catch (error) {
        console.error("Debug error:", error.message);
    }
}

debugCalendar();