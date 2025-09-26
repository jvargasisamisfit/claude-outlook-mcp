import { runAppleScript } from 'run-applescript';

const script = `
  tell application "Microsoft Outlook"
    set contactList to {}
    set allContactsList to contacts
    
    repeat with i from 1 to 2
      if i > (count of allContactsList) then exit repeat
      
      set theContact to item i of allContactsList
      
      -- Get contact name
      set contactName to "Unknown"
      try
        set contactName to display name of theContact
      on error
        try
          set contactName to first name of theContact & " " & last name of theContact
        on error
          -- Keep Unknown
        end try
      end try
      
      -- Get email
      set contactEmail to "No email"
      try
        set emailList to email addresses of theContact
        if (count of emailList) > 0 then
          set contactEmail to address of item 1 of emailList
        end if
      on error
        -- Keep No email
      end try
      
      -- Create contact data record
      set contactData to {name:contactName, email:contactEmail, phone:"No phone"}
      set end of contactList to contactData
    end repeat
    
    return contactList
  end tell
`;

async function test() {
  const result = await runAppleScript(script);
  console.log("Raw result:", result);
  console.log("Type:", typeof result);
}

test();
