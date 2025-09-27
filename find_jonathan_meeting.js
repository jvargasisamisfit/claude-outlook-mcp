// Find the Jonathan/Jon Connect meeting
import { runAppleScript } from 'run-applescript';

async function findJonathanMeeting() {
    const script = `
        tell application "Microsoft Outlook"
            try
                -- Search for meetings with "Jonathan" or "Jon" in the subject
                set allEvents to every calendar event whose subject contains "Jonathan" or subject contains "Jon Connect"

                set output to "Found " & (count of allEvents) & " events with Jonathan/Jon:\\n"

                set counter to 0
                repeat with evt in allEvents
                    if counter < 10 then
                        try
                            set evtSubject to subject of evt
                            set evtStart to start time of evt
                            set evtEnd to end time of evt

                            -- Check if it's the Monday meeting
                            set evtDay to weekday of evtStart as string
                            set evtDate to date string of evtStart

                            set output to output & "\\n• " & evtSubject
                            set output to output & "\\n  Date: " & evtDate & " (" & evtDay & ")"
                            set output to output & "\\n  Time: " & (time string of evtStart) & " - " & (time string of evtEnd)

                            -- Try to get calendar info
                            try
                                set evtCal to calendar of evt
                                set calID to id of evtCal
                                set output to output & "\\n  Calendar ID: " & calID
                            on error
                                set output to output & "\\n  Calendar: Unable to determine"
                            end try

                            -- Check if it's recurring
                            try
                                set isRecurring to is recurring of evt
                                if isRecurring then
                                    set output to output & "\\n  Recurring: Yes"
                                else
                                    set output to output & "\\n  Recurring: No"
                                end if
                            on error
                                set output to output & "\\n  Recurring: Unknown"
                            end try

                            set output to output & "\\n"
                            set counter to counter + 1
                        on error errMsg
                            set output to output & "\\n• Error reading event: " & errMsg & "\\n"
                        end try
                    end if
                end repeat

                return output
            on error errMsg
                return "Error searching: " & errMsg
            end try
        end tell
    `;

    const result = await runAppleScript(script);
    console.log(result);
}

findJonathanMeeting();