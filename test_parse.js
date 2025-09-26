const testData = `subject:True North All-Company Onsite: Travel Details, senderName:Alicia Schlieske, senderAddress:aschlieske@markarch.com, dateReceived:date Friday, September 26, 2025 at 2:52:06 PM, msgContent:<html><head>, msgId:83, subject:   Edit on the Miro board, senderName:Nick Jimeno via Miro, senderAddress:important@notification.miro.com, dateReceived:date Friday, September 26, 2025 at 1:32:07 PM, msgContent:<!DOCTYPE html>, msgId:84`;

// Test the splitting logic
const parts = testData.split(/,\s*msgId:\d+(?:,\s*|$)/);
console.log("Parts:", parts);
console.log("Number of parts:", parts.length);

// Better approach - split by ", subject:" after the first email
const emails = [];
const emailStrings = testData.split(/(?<=msgId:\d+),\s*(?=subject:)/);
console.log("\nEmail strings:", emailStrings);
console.log("Number of emails:", emailStrings.length);
