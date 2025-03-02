# Claude Outlook MCP Tool

This is a Model Context Protocol (MCP) tool that allows Claude to interact with Microsoft Outlook for macOS.

## Features

- Mail:
  - Read unread and regular emails
  - Search emails by keywords
  - Send emails with to, cc, and bcc recipients
  - List mail folders
- Calendar:
  - View today's events
  - View upcoming events
  - Search for events
  - Create new calendar events
- Contacts:
  - List contacts
  - Search contacts by name

## Prerequisites

- macOS with Apple Silicon (M1/M2/M3) or Intel chip
- [Microsoft Outlook for Mac](https://apps.apple.com/us/app/microsoft-outlook/id985367838) installed and configured
- [Bun](https://bun.sh/) installed
- [Claude desktop app](https://claude.ai/desktop) installed

## Installation

1. Clone this repository:

```bash
git clone https://github.com/syedazharmbnr1/claude-outlook-mcp.git
cd claude-outlook-mcp
```

2. Install dependencies:

```bash
bun install
```

3. Make sure the script is executable:

```bash
chmod +x index.ts
```

4. Update your Claude Desktop configuration:

Edit your `claude_desktop_config.json` file (located at `~/Library/Application Support/Claude/claude_desktop_config.json`) to include this tool:

```json
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "/Users/YOURUSERNAME/.bun/bin/bun",
      "args": ["run", "/path/to/claude-outlook-mcp/index.ts"]
    }
  }
}
```

Make sure to replace `YOURUSERNAME` with your actual macOS username and adjust the path to where you cloned this repository.

5. Restart Claude Desktop app

6. Grant permissions:
   - Go to System Preferences > Privacy & Security > Privacy
   - Give Terminal (or your preferred terminal app) access to Accessibility features
   - You may see permission prompts when the tool is first used

## Troubleshooting

If you encounter the error `Cannot find module '@modelcontextprotocol/sdk/server/index.js'`:

1. Make sure you've run `bun install` to install all dependencies
2. Try installing the MCP SDK explicitly:
   ```bash
   bun add @modelcontextprotocol/sdk@^1.5.0
   ```
3. Check if the module exists in your node_modules directory:
   ```bash
   ls -la node_modules/@modelcontextprotocol/sdk/server/
   ```

If the error persists, try creating a new project with Bun:

```bash
mkdir -p ~/Desktop/others/claudeMCP/claude-outlook-mcp
cd ~/Desktop/others/claudeMCP/claude-outlook-mcp
bun init -y
```

Then copy the package.json and index.ts files to the new directory and run:

```bash
bun install
bun run index.ts
```

Update your claude_desktop_config.json to point to the new location.

## Usage

Once installed, you can use the Outlook tool directly from Claude by asking questions like:

- "Can you check my unread emails in Outlook?"
- "Search my Outlook emails for the quarterly report"
- "Send an email to john@example.com with the subject 'Meeting Tomorrow'"
- "What's on my calendar today?"
- "Create a meeting for tomorrow at 2pm"
- "Find the contact information for Jane Smith"

## Examples

### Email Operations

```
Check my unread emails in Outlook
```

```
Send an email to alex@example.com with subject "Project Update" and the following body: Here's the latest update on our project. We've completed phase 1 and are moving on to phase 2.
```

```
Search my emails for "budget meeting"
```

### Calendar Operations

```
What events do I have today?
```

```
Create a calendar event for a team meeting tomorrow from 2pm to 3pm
```

```
Show me my upcoming events for the next 2 weeks
```

### Contact Operations

```
List all my Outlook contacts
```

```
Search for contact information for Jane Smith
```

## License

MIT