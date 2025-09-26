import { runAppleScript } from 'run-applescript';

async function debug() {
    const script = `
      tell application "Microsoft Outlook"
        try
          set targetFolder to inbox
          set messageList to {}
          set allMsgs to messages of targetFolder
          
          repeat with i from 1 to 2
            set theMsg to item i of allMsgs
            set msgData to {subject:subject of theMsg, msgId:id of theMsg}
            set end of messageList to msgData
          end repeat
          
          return messageList
        on error errMsg
          return "Error: " & errMsg
        end try
      end tell
    `;
    
    const result = await runAppleScript(script);
    console.log("Raw result:", JSON.stringify(result));
    console.log("\nSplit test:");
    const splits = result.split(/(?<=msgId:\d+),\s*(?=subject:)/);
    console.log("Splits:", splits);
}

debug();
