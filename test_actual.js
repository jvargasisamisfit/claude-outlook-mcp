import { runAppleScript } from 'run-applescript';

const script = `
  tell application "Microsoft Outlook"
    try
      set targetFolder to inbox
      set messageList to {}
      set msgCount to 0
      set allMsgs to messages of targetFolder

      repeat with i from 1 to 3
        if msgCount >= 3 then exit repeat

        try
          set theMsg to item i of allMsgs
          set msgSender to sender of theMsg
          set senderName to "Unknown"
          set senderAddress to "unknown@example.com"

          try
            set senderName to name of msgSender
            set senderAddress to address of msgSender
          on error
            try
              set senderAddress to msgSender as string
            on error
              -- Keep defaults
            end try
          end try

          set msgContent to ""
          try
            set msgContent to content of theMsg
            if length of msgContent > 500 then
              set msgContent to (text 1 thru 500 of msgContent) & "..."
            end if
          on error
            set msgContent to "[Content not available]"
          end try

          set msgData to {subject:subject of theMsg, ¬
                     senderName:senderName, ¬
                     senderAddress:senderAddress, ¬
                     dateReceived:time sent of theMsg, ¬
                     msgContent:msgContent, ¬
                     msgId:id of theMsg}

          set end of messageList to msgData
          set msgCount to msgCount + 1
        on error
          -- Skip problematic messages
        end try
      end repeat

      return messageList
    on error errMsg
      return "Error: " & errMsg
    end try
  end tell
`;

async function test() {
    try {
        const result = await runAppleScript(script);
        console.log("Result substring:", result.substring(0, 200));
        console.log("Result length:", result.length);
    } catch(e) {
        console.error("Error:", e);
    }
}

test();
