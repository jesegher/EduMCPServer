// MCP Education Assignments Server with User-Delegated Auth
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

// Update this line in your code
const graphScopes = ["https://graph.microsoft.com/EduRoster.Read", "https://graph.microsoft.com/EduAssignments.Read"];

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
  
  // Register tools with the new API
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
        }]
      };
    }
  );
  
  server.tool(
    "get-my-assignments",
    {}, // Empty schema for no parameters
    async () => {
      console.error("📚 get-my-assignments tool called");
      
      if (!accessToken) {
        console.error("❌ User not authenticated");
        return { 
          content: [{ 
            type: "text", 
            text: JSON.stringify({
              status: "error", 
              message: "User not authenticated. Please call microsoft-login first"
            })
          }]
        };
      }

      try {
        console.error("🔄 Fetching assignments from Microsoft Graph API");
        // Add timeout config to prevent hanging requests
        const res = await axios.get('https://graph.microsoft.com/v1.0/education/me/assignments', {
          headers: { Authorization: `Bearer ${accessToken}` },
          timeout: 5000 // 5 second timeout to prevent hanging
        });
        
        console.error(`✅ Successfully fetched ${res.data.value?.length || 0} assignments`);
        return { 
          content: [{ 
            type: "text", 
            text: JSON.stringify({
              status: "success", 
              value: res.data.value || []
            })
          }]
        };
      } catch (error) {
        console.error("❌ Error fetching assignments:", error.message);
        
        // Check if token expired
        if (error.response?.status === 401) {
          accessToken = null;
          isAuthenticated = false;
          return { 
            content: [{ 
              type: "text", 
              text: JSON.stringify({
                status: "error", 
                message: "Authentication expired. Please call microsoft-login again"
              })
            }]
          };
        }
        
        return { 
          content: [{ 
            type: "text", 
            text: JSON.stringify({
              status: "error", 
              message: `Failed to fetch assignments: ${error.message}`
            })
          }]
        };
      }
    }
  );

  // Define resources (education assignments)
  server.resource(
    "education-assignments",
    { list: async () => {
      // Return a list format that matches what was previously returned
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
        // Get assignments to list as resources
        const res = await axios.get('https://graph.microsoft.com/v1.0/education/me/assignments', {
          headers: { Authorization: `Bearer ${accessToken}` },
          timeout: 5000
        });
        
        // Return as resource contents
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
    }},
    async (uri) => {
      // Handle retrieving a specific assignment by ID
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
    }
  );
  
  // Define student info resource
  server.resource(
    "student-info",
    { list: async () => {
      if (!isAuthenticated) {
        return {
          contents: [
            { 
              uri: "student-info://status", 
              text: "User not authenticated. Please use the microsoft-login tool first."
            }
          ]
        };
      }
      
      return {
        contents: [
          { uri: "student-info://profile", text: "Student Profile" },
          { uri: "student-info://classes", text: "Enrolled Classes" }
        ]
      };
    }},
    async (uri) => {
      const resourceType = uri.pathname.substring(2); // Remove leading //
      
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
        let endpoint = '';
        let title = '';
        
        if (resourceType === 'profile') {
          endpoint = 'https://graph.microsoft.com/v1.0/education/me';
          title = 'Student Profile';
        } else if (resourceType === 'classes') {
          endpoint = 'https://graph.microsoft.com/v1.0/education/me/classes';
          title = 'Enrolled Classes';
        } else {
          return {
            contents: [
              { 
                uri: uri.href, 
                text: "Unknown resource type"
              }
            ]
          };
        }
        
        const res = await axios.get(endpoint, {
          headers: { Authorization: `Bearer ${accessToken}` },
          timeout: 5000
        });
        
        return {
          contents: [
            { 
              uri: uri.href, 
              text: `# ${title}\n\n${JSON.stringify(res.data, null, 2)}`
            }
          ]
        };
      } catch (error) {
        return {
          contents: [
            { 
              uri: uri.href, 
              text: `Error fetching student info: ${error.message}`
            }
          ]
        };
      }
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