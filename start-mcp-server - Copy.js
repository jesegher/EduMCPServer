require('dotenv').config();
const express = require('express');
const axios = require('axios');
const crypto = require('crypto');
const msal = require('@azure/msal-node');
const { McpServer } = require('@modelcontextprotocol/sdk/server/mcp.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const { z } = require('zod'); // Add this for parameter validation

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

const graphScopes = ["https://graph.microsoft.com/EduRoster.ReadWrite", "https://graph.microsoft.com/EduAssignments.ReadWrite"];

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
  
  // Register authentication tool
  server.tool(
    "microsoft-login",
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
          }]}
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
          }]}
        };
      }
    }
  );

  // Register get-auth-status tool
  server.tool(
    "get-auth-status",
    {}, // Empty schema for no parameters
    async () => {
      console.error("🔍 get-auth-status tool called");
      return { 
        content: [{ 
          type: "text", 
          text: JSON.stringify({
            authenticated: isAuthenticated,
            message: isAuthenticated ? 
              "User is authenticated" : 
              "User is not authenticated. Please call microsoft-login first"
          })
        }]}
      };
    }
  );
  
  // Register education-assignments resource
  server.resource(
    "education-assignments",
    {
      // List assignments (combined from the tool functionality)
      list: async () => {
        if (!isAuthenticated) {
          return {
            contents: [
              { 
                uri: "education-assignments://status", 
                text: "User not authenticated. Please use the microsoft-login tool first."
              }
            ]
          };
        }
        
        try {
          const res = await axios.get('https://graph.microsoft.com/v1.0/education/me/assignments', {
            headers: { Authorization: `Bearer ${accessToken}` },
            timeout: 5000
          });
          
          return {
            contents: res.data.value.map(assignment => ({
              uri: `education-assignments://${assignment.id}`,
              text: assignment.displayName
            }))
          };
        } catch (error) {
          return {
            contents: [
              { 
                uri: "education-assignments://error", 
                text: `Error fetching assignments: ${error.message}`
              }
            ]
          };
        }
      },

      // Retrieve specific assignment details
      async (uri) => {
        const assignmentId = uri.pathname.substring(2); // Remove leading //
        
        if (!isAuthenticated) {
          return {
            contents: [
              { 
                uri: uri.href, 
                text: "User not authenticated. Please use the microsoft-login tool first."
              }
            ]
          };
        }
        
        try {
          const res = await axios.get(`https://graph.microsoft.com/v1.0/education/me/assignments/${assignmentId}`, {
            headers: { Authorization: `Bearer ${accessToken}` },
            timeout: 5000
          });
          
          return {
            contents: [
              { 
                uri: uri.href, 
                text: JSON.stringify(res.data, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            contents: [
              { 
                uri: uri.href, 
                text: `Error fetching assignment details: ${error.message}`
              }
            ]
          };
        }
      },

      // Update assignment
      update: async (uri, updateData) => {
        const assignmentId = uri.pathname.substring(2); // Remove leading //
  
        if (!isAuthenticated) {
          return {
            contents: [
              { 
                uri: uri.href, 
                text: "User not authenticated. Please use the microsoft-login tool first."
              }
            ]
          };
        }

        try {
          const res = await axios.patch(
            `https://graph.microsoft.com/v1.0/education/me/assignments/${assignmentId}`,
            updateData, // This is the data to be updated (e.g., { dueDateTime: '2025-04-10T00:00:00Z' })
            {
              headers: { Authorization: `Bearer ${accessToken}` },
              timeout: 5000
            }
          );
  
          return {
            contents: [
              { 
                uri: uri.href, 
                text: `Assignment updated successfully: ${res.data.displayName}`
              }
            ]
          };
        } catch (error) {
          return {
            contents: [
              { 
                uri: uri.href, 
                text: `Error updating assignment: ${error.message}`
              }
            ]
          };
        }
      }
    }
  );

  // Optional: Add a new resource for class details
  server.resource(
    "class-details",
    { 
      list: async () => {
        if (!isAuthenticated) {
          return {
            contents: [
              { 
                uri: "class-details://status", 
                text: "User not authenticated. Please use the microsoft-login tool first."
              }
            ]
          };
        }
        
        try {
          const res = await axios.get('https://graph.microsoft.com/v1.0/education/me/classes', {
            headers: { Authorization: `Bearer ${accessToken}` },
            timeout: 5000
          });
          
          return {
            contents: res.data.value.map(cls => ({
              uri: `class-details://${cls.id}`,
              text: cls.displayName
            }))
          };
        } catch (error) {
          return {
            contents: [
              { 
                uri: "class-details://error", 
                text: `Error fetching class details: ${error.message}`
              }
            ]
          };
        }
      }
    }
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
      console.error("✅ Authentication successful!");

      res.send(`
        <h2>Authentication successful</h2>
        <p>You can now close this window and return to the application.</p>
        <script>window.close();</script>
      `);
    } catch (error) {
      console.error("❌ Callback error:", error);
      res.status(500).send("Error during authentication");
    }
  });

  // Start server
  const server_app = app.listen(PORT)
    .on('error', (err) => {
      console.error('❌ Auth server error:', err);
    })
    .on('listening', () => {
      console.error(`✅ Auth server running on port ${PORT}`);
    });

  // Connect the MCP server
  console.error("🔌 Connecting MCP server to transport...");
  
  try {
    // Connect MCP server
    await server.connect(transport);
    console.error("✅ MCP server connected and ready!");
  } catch (err) {
    console.error("❌ Failed to connect MCP server:", err);
    process.exit(1);
  }
}

// Run the server
createMCPServer().catch(err => {
  console.error("❌ Fatal MCP error:", err);
  process.exit(1);
});
