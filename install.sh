#!/bin/bash

# Colors for output
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

echo -e "${GREEN}Installing Claude Outlook MCP Tool...${NC}"

# Check if Bun is installed
if ! command -v bun &> /dev/null; then
    echo -e "${RED}Bun is not installed. Please install Bun first:${NC}"
    echo -e "${YELLOW}curl -fsSL https://bun.sh/install | bash${NC}"
    exit 1
fi

# Install dependencies
echo -e "${GREEN}Installing dependencies...${NC}"
bun install

if [ $? -ne 0 ]; then
    echo -e "${RED}Failed to install dependencies. Trying with explicit MCP SDK...${NC}"
    bun add @modelcontextprotocol/sdk@^1.5.0
    bun install
fi

# Make script executable
chmod +x index.ts

# Get current username
USERNAME=$(whoami)
INSTALL_PATH=$(pwd)

# Create claude_desktop_config.json snippet
CONFIG_SNIPPET=$(cat << EOF
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "/Users/$USERNAME/.bun/bin/bun",
      "args": ["run", "$INSTALL_PATH/index.ts"]
    }
  }
}
EOF
)

echo -e "${GREEN}Installation complete!${NC}"
echo -e "${YELLOW}Please add the following to your Claude Desktop config file at:${NC}"
echo -e "${YELLOW}~/Library/Application Support/Claude/claude_desktop_config.json${NC}"
echo ""
echo -e "${GREEN}$CONFIG_SNIPPET${NC}"
echo ""
echo -e "${YELLOW}Don't forget to restart Claude Desktop app after making these changes.${NC}"
echo -e "${YELLOW}You may need to grant Terminal access to Accessibility features in System Preferences.${NC}"