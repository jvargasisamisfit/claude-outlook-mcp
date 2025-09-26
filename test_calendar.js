import { runAppleScript } from 'run-applescript';

const script = `
  tell application "Microsoft Outlook"
    set todayEvents to {}
    set theCalendar to calendar id 119
    set todayDate to current date
    set startOfDay to todayDate - (time of todayDate)
    set endOfDay to startOfDay + 1 * days
    
    set allEvents to calendar events of theCalendar
    set eventList to {}
    
    repeat with evt in allEvents
      try
        if start time of evt >= startOfDay and start time of evt < endOfDay then
          set end of eventList to evt
        end if
      on error
        -- Skip
      end try
    end repeat
    
    repeat with i from 1 to 3
      if i > (count of eventList) then exit repeat
      set theEvent to item i of eventList
      set eventData to {subject:subject of theEvent, ¬
                   startTime:start time of theEvent, ¬
                   endTime:end time of theEvent, ¬
                   location:location of theEvent, ¬
                   eventId:id of theEvent}
      
      set end of todayEvents to eventData
    end repeat
    
    return todayEvents
  end tell
`;

async function test() {
  const result = await runAppleScript(script);
  console.log("Result:", result);
  console.log("Length:", result.length);
}

test();
