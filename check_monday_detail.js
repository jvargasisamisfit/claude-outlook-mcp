import { runAppleScript } from 'run-applescript';

async function checkMonday() {
    const script = `
        tell application "Microsoft Outlook"
            set mondayStart to date "Monday, September 29, 2025 12:00:00 AM"
            set mondayEnd to mondayStart + 1 * days
            
            set allMondayEvents to every calendar event whose start time >= mondayStart and start time < mondayEnd
            
            set output to "Monday Events (ALL calendars):\\n"
            repeat with evt in allMondayEvents
                try
                    set evtSubject to subject of evt
                    set evtTime to time string of start time of evt
                    set evtCal to calendar of evt
                    set calID to id of evtCal
                    set calName to name of evtCal
                    set output to output & "• " & evtTime & " - " & evtSubject & " (Cal: " & calName & ", ID: " & calID & ")\\n"
                on error
                    set output to output & "• Error reading event\\n"
                end try
            end repeat
            return output
        end tell
    `;
    
    const result = await runAppleScript(script);
    console.log(result);
}

checkMonday();
