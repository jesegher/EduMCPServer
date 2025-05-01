console.log("🟢 MCP server entry file loaded");

import { createRequire } from 'module';
const require = createRequire(import.meta.url);

import 'dotenv/config';

import express from "express";
const { Request, Response } = express;

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";

import { z } from 'zod';

const msal = require('@azure/msal-node');
const { ConfidentialClientApplication } = msal;

import registerAssignmentTools from './Tools/assignment.js';
import registerRubricTools from './Tools/rubric.js';
import registerClassTools from './Tools/class.js';
import registerUserTools from './Tools/users.js';
import registerGroupTools from './Tools/group.js';
import registerAuthTools from './Tools/auth.js';

const pendingAuthStates = new Set();
let accessToken = null;
let isAuthenticated = false;

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

const auth = { accessToken: null, isAuthenticated: false };

const server = new McpServer({
  name: "mcp-edu-hosted",
  description: "A Microsoft Edu hosted MCP server.",
  version: "1.0.0",
  tools: [],
});

console.error("📝 Registering tools...");
  

registerAuthTools(server, auth, msalClient, pendingAuthStates, graphScopes);
registerAssignmentTools(server, auth);
registerRubricTools(server, auth);
registerClassTools(server, auth);
registerUserTools(server,auth);
registerGroupTools(server,auth);

const app = express();

// to support multiple simultaneous connections we have a lookup object from
// sessionId to transport
const transports = {};

app.get("/sse", async (req, res) => {
  // Get the full URI from the request
  const host = req.get("host");

  const fullUri = `https://${host}/messages`;
  const transport = new SSEServerTransport(fullUri, res);

  transports[transport.sessionId] = transport;
  res.on("close", () => {
    delete transports[transport.sessionId];
  });
  await server.connect(transport);
});

app.post("/messages", async (req, res) => {
  const sessionId = String(req.query.sessionId);
  const transport = transports[sessionId];
  if (transport) {
    await transport.handlePostMessage(req, res);
  } else {
    res.status(400).send("No transport found for sessionId");
  }
});

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

app.get("/", (req, res) => {
  res.send("The Education MCP server is running!");
});

const PORT = process.env.PORT || 3001;

app.listen(PORT, () => {
  console.log(`✅ Server is running at http://localhost:${PORT}`);
});
