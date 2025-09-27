const { spawn } = require("child_process");

async function runAppleScript(script) {
  return new Promise((resolve, reject) => {
    const osascript = spawn("osascript", ["-e", script]);
    
    let stdout = "";
    let stderr = "";
    
    osascript.stdout.on("data", (data) => {
      stdout += data.toString();
    });
    
    osascript.stderr.on("data", (data) => {
      stderr += data.toString();
    });
    
    osascript.on("close", (code) => {
      if (code === 0) {
        console.log("Raw output:", JSON.stringify(stdout));
        resolve(stdout.trim());
      } else {
        reject(new Error(stderr || `osascript exited with code ${code}`));
      }
    });
  });
}

const script = `
  tell application "Microsoft Outlook"
    set contactList to {}
    set allContactsList to contacts
    set contactCount to count of allContactsList
    set limitCount to 5
    
    if contactCount < limitCount then
      set limitCount to contactCount
    end if
    
    repeat with i from 1 to limitCount
      try
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

        -- Get email address
        set contactEmail to "No email"
        try
          set emailList to email addresses of theContact
          if (count of emailList) > 0 then
            set contactEmail to address of item 1 of emailList
          end if
        on error
          -- Keep No email
        end try

        -- Get phone number
        set contactPhone to "No phone"
        try
          set phoneList to phone numbers of theContact
          if (count of phoneList) > 0 then
            set contactPhone to content of item 1 of phoneList
          end if
        on error
          -- Keep No phone
        end try

        -- Create contact data as string
        set contactData to "name:" & contactName & ", email:" & contactEmail & ", phone:" & contactPhone

        set end of contactList to contactData
      on error
        -- Skip contacts that can't be processed
      end try
    end repeat
    
    return contactList
  end tell
`;

runAppleScript(script)
  .then(result => {
    console.log("Result:", result);
    console.log("Split by newlines:", result.split('\n'));
  })
  .catch(err => console.error("Error:", err));
