// MCP Education Assignments Server with User-Delegated Auth
require('dotenv').config();
const express = require('express');
const axios = require('axios');
const crypto = require('crypto');
const msal = require('@azure/msal-node');
const { McpServer } = require('@modelcontextprotocol/sdk/server/mcp.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const { z } = require('zod'); // Add this for parameter validation
const { Console } = require('console');

const registerAssignmentTools = require('./Tools/assignment.js');
const registerRubricTools = require('./Tools/rubric.js');
const registerClassTools = require('./Tools/class.js');


let accessToken = null;
let isAuthenticated = false;

// Add a state store to validate auth callbacks
const pendingAuthStates = new Set();

const msalClient = new msal.ConfidentialClientApplication({
  auth: {
    clientId: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
  },
});

// Update this line in your code
const graphScopes = ["https://graph.microsoft.com/EduRoster.ReadWrite","https://graph.microsoft.com/EduAssignments.ReadWrite","https://graph.microsoft.com/User.ReadWrite.All"];

async function createMCPServer() {
  console.error("🚀 Starting MCP Education server...");
  
  // Create the server with the new McpServer class
  const server = new McpServer({ 
    name: "education-server", 
    version: "1.0.0" 
  });

    // Create a transport
  const transport = new StdioServerTransport();
  
  console.error("📝 Registering tools...");
  
  
  // Register your assignment tools
  
  const auth = {
    accessToken: null,
    isAuthenticated: false
  };

  console.error("📝 Registering assignment tools...");

  registerAssignmentTools(server, auth);

  console.error("📝 Registering rubric tools...");
  registerRubricTools(server, auth);

  console.error("📝 Registering class tools...");
  registerClassTools(server, auth);

  // Register tools with the new API
  server.tool(
    "auth-login",
    {}, // Empty schema for no parameters
    async () => {
      console.error("🔑 microsoft-login tool called");
      try {
        const state = crypto.randomBytes(16).toString("hex");
        pendingAuthStates.add(state);
        
        const url = await msalClient.getAuthCodeUrl({
          scopes: graphScopes,
          redirectUri: process.env.REDIRECT_URI,
          state,
        });
        
        console.error(`📤 Generated auth URL: ${url.substring(0, 50)}...`);
        return { 
          content: [{ 
            type: "text", 
            text: JSON.stringify({
              status: "authentication_required", 
              url,
              message: "Please open this URL in your browser to authenticate"
            })
          }]
        };
      } catch (error) {
        console.error("❌ Error generating auth URL:", error);
        return { 
          content: [{ 
            type: "text", 
            text: JSON.stringify({
              status: "error", 
              message: "Failed to generate authentication URL"
            })
          }]
        };
      }
    }
  );
  
  server.tool(
    "auth-status-get",
    {}, // Empty schema for no parameters
    async () => {
      console.error("🔍 get-auth-status tool called");
      console.error(accessToken);
      
      return { 
        content: [{ 
          type: "text", 
          text: JSON.stringify({
            authenticated: isAuthenticated,
            message: isAuthenticated ? 
              "User is authenticated" : 
              "User is not authenticated. Please call microsoft-login first"
          })
        }]
      };
    }
  );
  
  
  
  server.tool(
    "user-get",
    {
      userId: z.string().optional().describe("The ID of the user to retrieve"),
      userPrincipalName: z.string().optional().describe("The email (UPN) of the user to retrieve"),
      search: z.string().optional().describe("Optional: Search query to look up users by name or email (e.g. 'john')")
    },
    async ({ userId, userPrincipalName, search }) => {
      console.error("🔍 get-user tool called");
  
      if (!isAuthenticated) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "error",
              message: "❌ User not authenticated. Please use the microsoft-login tool first."
            })
          }]
        };
      }
  
      try {
        let userResponse;
  
        // 1. Fetch by ID
        if (userId) {
          userResponse = await axios.get(
            `https://graph.microsoft.com/v1.0/users/${userId}`,
            {
              headers: { Authorization: `Bearer ${accessToken}` }
            }
          );
  
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                status: "success",
                message: "User retrieved by ID.",
                user: userResponse.data
              })
            }]
          };
        }
  
        // 2. Fetch by UPN/email
        if (userPrincipalName) {
          userResponse = await axios.get(
            `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userPrincipalName)}`,
            {
              headers: { Authorization: `Bearer ${accessToken}` }
            }
          );
  
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                status: "success",
                message: "User retrieved by userPrincipalName.",
                user: userResponse.data
              })
            }]
          };
        }
  
        // 3. Search mode
        if (search) {
          const searchResponse = await axios.get(
            `https://graph.microsoft.com/v1.0/users?$search="displayName:${search}"&$count=true`,
            {
              headers: {
                Authorization: `Bearer ${accessToken}`,
                ConsistencyLevel: "eventual"
              }
            }
          );
  
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                status: "success",
                message: "Users matching search query retrieved.",
                count: searchResponse.data?.value?.length || 0,
                users: searchResponse.data?.value
              })
            }]
          };
        }
  
        // 4. Fallback: list first page of all users
        const allResponse = await axios.get(
          `https://graph.microsoft.com/v1.0/users?$top=10`,
          {
            headers: { Authorization: `Bearer ${accessToken}` }
          }
        );
  
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "success",
              message: "Returning first page of users.",
              users: allResponse.data.value
            })
          }]
        };
      } catch (error) {
        let errorMessage = "Unknown error occurred";
        if (error.response) {
          errorMessage = `API error: ${error.response.status} - ${error.response.data?.error?.message || 'Unknown API error'}`;
        } else if (error.request) {
          errorMessage = "Network error: No response received from server";
        } else {
          errorMessage = `Request error: ${error.message}`;
        }
  
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "error",
              message: errorMessage
            })
          }]
        };
      }
    },
    {
      description: "Fetches a user by ID, email (userPrincipalName), or search query. Returns first page of users if no parameters are provided."
    }
  );
  

  // Register prompts
  server.prompt(
    "get-assignments",
    {},
    () => ({
      messages: [
        {
          role: "user",
          content: {
            type: "text",
            text: "Please get my education assignments and format them for easy reading."
          }
        }
      ]
    })
  );
  
  server.prompt(
    "assignments-by-date",
    {},
    () => ({
      messages: [
        {
          role: "user",
          content: {
            type: "text",
            text: "Get my assignments and organize them by due date, with the closest deadlines first."
          }
        }
      ]
    })
  );
  
  server.prompt(
    "upcoming-deadlines",
    {},
    () => ({
      messages: [
        {
          role: "user",
          content: {
            type: "text",
            text: "Please show me assignments that are due within the next 7 days."
          }
        }
      ]
    })
  );
  
  server.prompt(
    "class-summary",
    {},
    () => ({
      messages: [
        {
          role: "user",
          content: {
            type: "text",
            text: "Get my assignments and organize them by class, showing a summary for each course."
          }
        }
      ]
    })
  );

  // ✅ Auth callback Express server
  const app = express();
  const PORT = process.env.PORT || 3000;

  app.get('/auth/callback', async (req, res) => {
    console.error("📥 Received auth callback");
    
    // Validate state parameter
    const state = req.query.state;
    if (!state || !pendingAuthStates.has(state)) {
      console.error("❌ Invalid state parameter in callback");
      return res.status(400).send("Invalid state parameter");
    }
    
    // Remove used state
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

      res.send(`
        <h2>Authentication successful</h2>
        <p>You can now close this window and return to Claude Desktop.</p>
        <script>window.close();</script>
      `);
    } catch (error) {
      console.error("❌ Callback error:", error);
      res.status(500).send("Error during authentication");
    }
  });

  // Add proper error handling and timeout for Express server
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
  
  // Set a timeout for the Express server
  server_app.timeout = 10000; // 10 seconds

  // Connect the MCP server
  console.error("🔌 Connecting MCP server to transport...");
  
  // Add more verbose error handling for the connection
  try {
    // Add a timeout promise to avoid hanging
    const connectWithTimeout = Promise.race([
      server.connect(transport),
      new Promise((_, reject) => 
        setTimeout(() => reject(new Error("MCP connection timeout")), 10000)
      )
    ]);
    
    await connectWithTimeout;
    console.error("✅ MCP server connected and ready!");
  } catch (err) {
    console.error("❌ Failed to connect MCP server:", err);
    process.exit(1);
  }
}

// Set up global error handler
process.on('uncaughtException', (err) => {
  console.error('❌ Uncaught exception:', err);
});

// Add an unhandled promise rejection handler
process.on('unhandledRejection', (reason, promise) => {
  console.error('❌ Unhandled promise rejection:', reason);
});

createMCPServer().catch(err => {
  console.error("❌ Fatal MCP error:", err);
  process.exit(1);
});