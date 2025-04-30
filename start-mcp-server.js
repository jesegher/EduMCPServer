import { createRequire } from 'module';
const require = createRequire(import.meta.url);

import 'dotenv/config';
import express from 'express';
import axios from 'axios';
import crypto from 'crypto';

const msal = require('@azure/msal-node');
const { ConfidentialClientApplication } = msal;

import { McpServer, ResourceTemplate } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from 'zod';
import { Console } from 'console';

import registerAssignmentTools from './Tools/assignment.js';
import registerRubricTools from './Tools/rubric.js';
import registerClassTools from './Tools/class.js';
import registerUserTools from './Tools/users.js';
import registerGroupTools from './Tools/group.js';
import registerAuthTools from './Tools/auth.js';


let accessToken = null;
let isAuthenticated = false;
const pendingAuthStates = new Set();

const msalClient = new ConfidentialClientApplication({
  auth: {
    clientId: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
  },
});

const graphScopes = [
  "https://graph.microsoft.com/EduRoster.ReadWrite",
  "https://graph.microsoft.com/EduAssignments.ReadWrite",
  "https://graph.microsoft.com/User.ReadWrite.All",
  "https://graph.microsoft.com/Group.ReadWrite.All",
  "https://graph.microsoft.com/Directory.ReadWrite.All"
];

async function createMCPServer() {
  console.error("🚀 Starting MCP Education server...");

  const server = new McpServer({
    name: "education-server",
    version: "1.0.0",
    capabilities: {
      resources: true,
      tools: true,
      prompts: true
    }
  });

  const stdioTransport = new StdioServerTransport();
  
  console.error("📝 Registering tools...");
  const auth = { accessToken: null, isAuthenticated: false };

  registerAuthTools(server, auth, msalClient, pendingAuthStates, graphScopes);
  registerAssignmentTools(server, auth);
  registerRubricTools(server, auth);
  registerClassTools(server, auth);
  registerUserTools(server,auth);
  registerGroupTools(server,auth);
  
  

 
  server.resource("config", "config://app", async (uri) => ({
    contents: [{ uri: uri.href, text: "Entra ID Stuff" }]
  }));

  server.resource(
    "greeting",
    new ResourceTemplate("greeting://{name}", { list: undefined }),
    async (uri, { name }) => ({
      contents: [{ uri: uri.href, text: `Hello, ${name}!` }]
    })
  );

  server.resource(
    "echo",
    new ResourceTemplate("echo://{message}", { list: undefined }),
    async (uri, { message }) => ({
      contents: [{ uri: uri.href, text: `Resource echo: ${message}` }]
    })
  );

  server.prompt("class-summary", {}, () => ({
    messages: [{
      role: "user",
      content: { type: "text", text: "Get my assignments and organize them by class, showing a summary for each course." }
    }]
  }));

  const app = express();
  const PORT = process.env.PORT || 3000;

  app.get('/auth/callback', async (req, res) => {
    console.error("📥 Received auth callback");

    const state = req.query.state;
    if (!state || !pendingAuthStates.has(state)) {
      console.error("❌ Invalid state parameter in callback");
      if (!res.headersSent) {
        return res.status(400).send("Invalid state parameter");
      }
      return;
    }

    pendingAuthStates.delete(state);

    try {
      console.error("🔄 Acquiring token...");
      const tokenResponse = await msalClient.acquireTokenByCode({
        code: req.query.code,
        scopes: graphScopes,
        redirectUri: process.env.REDIRECT_URI
      });

      accessToken = tokenResponse.accessToken;
      isAuthenticated = true;
      auth.accessToken = accessToken;
      auth.isAuthenticated = true;

      console.error("✅ Authentication successful!");

      if (!res.headersSent) {
        res.send(`
          <h2>Authentication successful</h2>
          <p>You can now close this window and return to Claude Desktop.</p>
          <script>window.close();</script>
        `);
      }
    } catch (error) {
      console.error("❌ Callback error:", error);
      if (!res.headersSent) {
        res.status(500).send("Error during authentication");
      }
    }
  });

  const server_app = app.listen(PORT)
    .on('error', (err) => {
      if (err.code === 'EADDRINUSE') {
        console.error(`❌ Port ${PORT} is already in use`);
      } else {
        console.error('❌ Auth server error:', err);
      }
    })
    .on('listening', () => {
      console.error(`✅ Auth server running on port ${PORT}`);
    });

  server_app.timeout = 10000;

  console.error("🔌 Connecting MCP server to transport...");

  try {
    const connectWithTimeout = Promise.race([
      server.connect(stdioTransport), 
      new Promise((_, reject) =>
        setTimeout(() => reject(new Error("MCP connection timeout")), 10000)
      )
    ]);

    await connectWithTimeout;
    console.error("✅ MCP server connected and ready on both stdio and SSE!");
  } catch (err) {
    console.error("❌ Failed to connect MCP server:", err);
    process.exit(1);
  }
}


process.on('uncaughtException', (err) => {
  console.error('❌ Uncaught exception:', err);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('❌ Unhandled promise rejection:', reason);
});

createMCPServer().catch(err => {
  console.error("❌ Fatal MCP error:", err);
  process.exit(1);
});
