#!/usr/bin/env node

const { spawn } = require('child_process');

// Start the MCP server
const mcpProcess = spawn('bun', ['run', '/Users/jvargas/Documents/GitHub/claude-outlook-mcp/index.ts'], {
  stdio: ['pipe', 'pipe', 'pipe'],
  cwd: '/Users/jvargas/Documents/GitHub/claude-outlook-mcp'
});

// Log stderr output
mcpProcess.stderr.on('data', (data) => {
  console.error('[MCP STDERR]:', data.toString());
});

// Log stdout output
mcpProcess.stdout.on('data', (data) => {
  console.log('[MCP STDOUT]:', data.toString());
});

// Send initialize request after a delay
setTimeout(() => {
  const initRequest = JSON.stringify({
    jsonrpc: '2.0',
    id: 1,
    method: 'initialize',
    params: {
      protocolVersion: '2024-11-05',
      capabilities: {},
      clientInfo: {
        name: 'test-client',
        version: '1.0.0'
      }
    }
  });

  console.log('Sending initialize request...');
  mcpProcess.stdin.write(initRequest + '\n');

  // Send tool list request after initialization
  setTimeout(() => {
    const toolsRequest = JSON.stringify({
      jsonrpc: '2.0',
      id: 2,
      method: 'tools/list',
      params: {}
    });

    console.log('Sending tools/list request...');
    mcpProcess.stdin.write(toolsRequest + '\n');

    // Send call tool request
    setTimeout(() => {
      const callRequest = JSON.stringify({
        jsonrpc: '2.0',
        id: 3,
        method: 'tools/call',
        params: {
          name: 'outlook_mail',
          arguments: {
            operation: 'unread',
            limit: 5
          }
        }
      });

      console.log('Sending outlook_mail unread request...');
      mcpProcess.stdin.write(callRequest + '\n');

      // Wait for response and then exit
      setTimeout(() => {
        console.log('Killing MCP process...');
        mcpProcess.kill();
        process.exit(0);
      }, 5000);
    }, 1000);
  }, 1000);
}, 1000);