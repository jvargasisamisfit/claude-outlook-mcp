# MCP Outlook Tools Guide

## Available Tools

### 1. mcp__outlook-mcp__outlook_mail
Email management tool for reading, searching, sending emails and managing folders.

#### Operations:

##### `unread` - Get unread emails
```javascript
{
  "operation": "unread",
  "folder": "Inbox",  // optional, defaults to Inbox
  "limit": 10         // optional, defaults to 10
}
```

##### `read` - Read emails from a folder
```javascript
{
  "operation": "read",
  "folder": "Inbox",  // optional, defaults to Inbox
  "limit": 10         // optional
}
```

##### `search` - Search for emails
```javascript
{
  "operation": "search",
  "searchTerm": "invoice",  // required
  "folder": "Inbox",        // optional
  "limit": 10               // optional
}
```

##### `send` - Send an email
```javascript
{
  "operation": "send",
  "to": "user@example.com",           // required
  "subject": "Meeting Tomorrow",       // required
  "body": "Let's meet at 10am",       // required
  "cc": "manager@example.com",        // optional
  "bcc": "archive@example.com",       // optional
  "isHtml": false,                    // optional, defaults to false
  "attachments": ["/path/to/file.pdf"] // optional array of file paths
}
```

##### `folders` - List all mail folders
```javascript
{
  "operation": "folders"
}
```

### 2. mcp__outlook-mcp__outlook_calendar
Calendar management for viewing, creating, updating, and deleting events.

#### Operations:

##### `today` - Get today's events
```javascript
{
  "operation": "today",
  "limit": 10  // optional
}
```

##### `upcoming` - Get upcoming events
```javascript
{
  "operation": "upcoming",
  "days": 7,   // optional, number of days ahead (default: 7)
  "limit": 10  // optional
}
```

##### `search` - Search calendar events
```javascript
{
  "operation": "search",
  "searchTerm": "meeting",  // required
  "limit": 10               // optional
}
```

##### `create` - Create a new event
```javascript
{
  "operation": "create",
  "subject": "Team Meeting",                    // required
  "start": "2025-09-27T14:00:00",              // required (ISO format)
  "end": "2025-09-27T15:00:00",                // required (ISO format)
  "location": "Conference Room A",              // optional
  "body": "Quarterly review discussion",        // optional
  "attendees": "john@example.com,jane@example.com" // optional, comma-separated
}
```

##### `delete` - Delete an event
```javascript
{
  "operation": "delete",
  "eventId": "12345"  // required (get from search/today/upcoming results)
}
```

##### `update` - Update an existing event
```javascript
{
  "operation": "update",
  "eventId": "12345",                          // required
  "subject": "Updated Meeting Title",          // optional
  "start": "2025-09-27T15:00:00",             // optional
  "end": "2025-09-27T16:00:00",               // optional
  "location": "New Room",                      // optional
  "body": "Updated description"                // optional
}
```

### 3. mcp__outlook-mcp__outlook_contacts
Contact management for listing and searching contacts.

#### Operations:

##### `list` - List contacts
```javascript
{
  "operation": "list",
  "limit": 20  // optional, defaults to 20
}
```

##### `search` - Search contacts
```javascript
{
  "operation": "search",
  "searchTerm": "Smith",  // required
  "limit": 10             // optional
}
```

## Important Implementation Details

1. **Time Format**:
   - Input: Use ISO 8601 format (e.g., "2025-09-27T14:00:00")
   - Internal: Server converts to AppleScript format with 12-hour AM/PM
   - The formatTime function handles conversion automatically

2. **Classic Outlook Required**:
   - Must use Classic Outlook on macOS
   - New Outlook has limited AppleScript support

3. **Calendar Specifics**:
   - Uses calendar ID 119 as default (main Calendar)
   - Falls back to searching for calendar named "Calendar"
   - Update operations don't call `save` (causes error -1701)

4. **Email Attachments**:
   - Supports file attachments via `attachments` array
   - Paths can be absolute or relative
   - Server converts relative paths to absolute

5. **HTML Emails**:
   - Set `isHtml: true` to send HTML formatted emails
   - Content type is set appropriately

## Common Issues & Solutions

### Email Issues
- **No emails found**: Ensure you're using Classic Outlook, not New Outlook
- **Send fails with attachments**: Check file paths are correct and accessible
- **Folder not found**: Use `folders` operation to list available folders

### Calendar Issues
- **Error -1701**: Known Microsoft bug - restart Claude Desktop
- **Events at wrong time**: Server handles conversion, use ISO format
- **Can't find calendar**: Check calendar ID or name in Outlook preferences

### Contact Issues
- **Limited info returned**: AppleScript access to contacts may be restricted
- **Search not working**: Ensure contacts are synced in Outlook

### General Troubleshooting
1. Ensure Microsoft Outlook is running
2. Check you're using Classic Outlook (not New Outlook)
3. Verify AppleScript permissions are granted to Terminal/Claude
4. Restart Claude Desktop if persistent errors occur
5. Check console errors for detailed debugging info

## Error Handling

The server includes extensive error handling:
- Validates all required parameters
- Checks Outlook accessibility before operations
- Provides fallback methods for operations
- Returns detailed error messages

## Performance Notes

- Operations may be slow with large mailboxes
- Limit parameters help control response size
- Calendar operations scan all events (no native date filtering in AppleScript)
- Contact operations may have limited access based on privacy settings