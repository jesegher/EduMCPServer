console.log("🟢 MCP server entry file loaded");

import { createRequire } from 'module';
const require = createRequire(import.meta.url);

import express from 'express';
import 'dotenv/config';

import { McpServer, ResourceTemplate } from "@modelcontextprotocol/sdk/server/mcp.js";
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


/**
 * This example server demonstrates the deprecated HTTP+SSE transport 
 * (protocol version 2024-11-05). It is mainly used for testing backward-compatible clients.
 * 
 * The server exposes two endpoints:
 * - /mcp: For establishing the SSE stream (GET)
 * - /messages: For receiving client messages (POST)
 */

// Create an MCP server instance
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

const getServer = () => {
  const server = new McpServer({
    name: 'simple-sse-server',
    version: '1.0.0',
  }, { capabilities: { logging: {} } });

 

  console.error("📝 Registering tools...");
  

  registerAuthTools(server, auth, msalClient, pendingAuthStates, graphScopes);
  registerAssignmentTools(server, auth);
  registerRubricTools(server, auth);
  registerClassTools(server, auth);
  registerUserTools(server,auth);
  registerGroupTools(server,auth);

  return server;
};

const app = express();
app.use(express.json());

// Store transports by session ID
const transports = {};

// SSE endpoint for establishing the stream
app.get('/mcp', async (req, res) => {
  console.log('Received GET request to /sse (establishing SSE stream)');

  try {
    // Create a new SSE transport for the client
    // The endpoint for POST messages is '/messages'
    const transport = new SSEServerTransport('/messages', res);

    // Store the transport by session ID
    const sessionId = transport.sessionId;
    transports[sessionId] = transport;

    // Set up onclose handler to clean up transport when closed
    transport.onclose = () => {
      console.log(`SSE transport closed for session ${sessionId}`);
      delete transports[sessionId];
    };

    // Connect the transport to the MCP server
    const server = getServer();
    await server.connect(transport);

    console.log(`Established SSE stream with session ID: ${sessionId}`);
  } catch (error) {
    console.error('Error establishing SSE stream:', error);
    if (!res.headersSent) {
      res.status(500).send('Error establishing SSE stream');
    }
  }
});

// Messages endpoint for receiving client JSON-RPC requests
app.post('/messages', async (req, res) => {
  console.log('Received POST request to /messages');

  // Extract session ID from URL query parameter
  const sessionId = req.query.sessionId;

  if (!sessionId) {
    console.error('No session ID provided in request URL');
    res.status(400).send('Missing sessionId parameter');
    return;
  }

  const transport = transports[sessionId];
  if (!transport) {
    console.error(`No active transport found for session ID: ${sessionId}`);
    res.status(404).send('Session not found');
    return;
  }

  try {
    // Handle the POST message with the transport
    await transport.handlePostMessage(req, res, req.body);
  } catch (error) {
    console.error('Error handling request:', error);
    if (!res.headersSent) {
      res.status(500).send('Error handling request');
    }
  }
});

// Start the server

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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✅ MCP Server is listening on port ${PORT}`);
});

// Handle server shutdown
process.on('SIGINT', async () => {
  console.log('Shutting down server...');

  // Close all active transports to properly clean up resources
  for (const sessionId in transports) {
    try {
      console.log(`Closing transport for session ${sessionId}`);
      await transports[sessionId].close();
      delete transports[sessionId];
    } catch (error) {
      console.error(`Error closing transport for session ${sessionId}:`, error);
    }
  }
  console.log('Server shutdown complete');
  process.exit(0);
});