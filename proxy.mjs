#!/usr/bin/env node
/**
 * Stdio ↔ HTTP bridge for the Outlook MCP Worker.
 * Claude Desktop spawns this as a child process and communicates via stdio.
 * This script forwards each JSON-RPC message to the Worker's /mcp endpoint
 * and writes the response back to stdout.
 */

import * as readline from 'readline';

const WORKER_URL = 'https://outlook-mcp.bashar-basheer.workers.dev/mcp';
const MCP_SECRET = process.env.MCP_SECRET ?? '';

const rl = readline.createInterface({
  input: process.stdin,
  terminal: false,
});

rl.on('line', async (line) => {
  const trimmed = line.trim();
  if (!trimmed) return;

  try {
    const response = await fetch(WORKER_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'X-MCP-Secret': MCP_SECRET,
      },
      body: trimmed,
    });

    // 202 Accepted = notification acknowledged, no response body expected
    if (response.status === 202) return;

    const text = await response.text();
    process.stdout.write(text + '\n');
  } catch (err) {
    process.stderr.write(`[outlook-mcp proxy] ${err.message}\n`);
  }
});

process.stdin.on('end', () => process.exit(0));
