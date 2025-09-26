import { runAppleScript } from 'run-applescript';

async function test() {
    try {
        const result = await runAppleScript('tell application "Microsoft Outlook" to count messages of inbox');
        console.log("Count result:", result);
        console.log("Type:", typeof result);
    } catch(e) {
        console.error("Error:", e);
    }
}

test();
