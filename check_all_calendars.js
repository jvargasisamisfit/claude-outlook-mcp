// Check events across ALL calendars for Monday
import { runAppleScript } from 'run-applescript';

async function checkAllCalendars() {
    try {
        const checkAllScript = `
            tell application "Microsoft Outlook"
                try
                    set targetDate to date "Monday, September 29, 2025 12:00:00 AM"
                    set endDate to targetDate + 1 * days

                    set resultList to "## Monday Sep 29 - Events Across All Calendars\\n\\n"

                    -- Check main calendar (119)
                    try
                        set mainCal to calendar id 119
                        set mainEvents to every calendar event of mainCal whose start time >= targetDate and start time < endDate
                        set resultList to resultList & "**Main Calendar (119)**: " & (count of mainEvents) & " events\\n"
                    on error
                        set resultList to resultList & "**Main Calendar (119)**: Error\\n"
                    end try

                    -- Check Calendar (13)
                    try
                        set cal13 to calendar id 13
                        set cal13Events to every calendar event of cal13 whose start time >= targetDate and start time < endDate
                        set resultList to resultList & "**Calendar (13)**: " & (count of cal13Events) & " events\\n"
                        if (count of cal13Events) > 0 then
                            repeat with evt in cal13Events
                                try
                                    set evtSubject to subject of evt
                                    set evtTime to time string of start time of evt
                                    set resultList to resultList & "  • " & evtTime & " - " & evtSubject & "\\n"
                                on error
                                    set resultList to resultList & "  • [Error reading event]\\n"
                                end try
                            end repeat
                        end if
                    on error
                        set resultList to resultList & "**Calendar (13)**: Error\\n"
                    end try

                    -- Check Noah's calendar (122)
                    try
                        set noahCal to calendar id 122
                        set noahEvents to every calendar event of noahCal whose start time >= targetDate and start time < endDate
                        set resultList to resultList & "**Noah Moss (122)**: " & (count of noahEvents) & " events\\n"
                        if (count of noahEvents) > 0 then
                            repeat with evt in noahEvents
                                try
                                    set evtSubject to subject of evt
                                    set evtTime to time string of start time of evt
                                    set resultList to resultList & "  • " & evtTime & " - " & evtSubject & "\\n"
                                on error
                                    set resultList to resultList & "  • [Error reading event]\\n"
                                end try
                            end repeat
                        end if
                    on error
                        set resultList to resultList & "**Noah Moss (122)**: Error\\n"
                    end try

                    -- Check ALL events (across all calendars)
                    set allEvents to every calendar event whose start time >= targetDate and start time < endDate
                    set resultList to resultList & "\\n**Total across ALL calendars**: " & (count of allEvents) & " events\\n"

                    if (count of allEvents) > 0 then
                        repeat with evt in allEvents
                            try
                                set evtSubject to subject of evt
                                set evtTime to time string of start time of evt
                                set evtCal to calendar of evt
                                set evtCalName to name of evtCal
                                set resultList to resultList & "  • " & evtTime & " - " & evtSubject & " (Calendar: " & evtCalName & ")\\n"
                            on error
                                try
                                    set evtSubject to subject of evt
                                    set evtTime to time string of start time of evt
                                    set resultList to resultList & "  • " & evtTime & " - " & evtSubject & " (Unknown calendar)\\n"
                                on error
                                    set resultList to resultList & "  • [Error reading event]\\n"
                                end try
                            end try
                        end repeat
                    end if

                    return resultList
                on error errMsg
                    return "Error: " & errMsg
                end try
            end tell
        `;

        const result = await runAppleScript(checkAllScript);
        console.log(result);

    } catch (error) {
        console.error("Error:", error.message);
    }
}

checkAllCalendars();