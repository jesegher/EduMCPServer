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
  
  // CONVERTED: class-details resource to list-classes tool
  server.tool(
    "list-classes",
    {}, // Empty schema for no parameters
    async () => {
      console.error("📚 list-classes tool called");
      
      if (!isAuthenticated) {
        return { 
          content: [{ 
            type: "text", 
            text: JSON.stringify({
              status: "error",
              message: "User not authenticated. Please use the microsoft-login tool first."
            })
          }]
        };
      }
      
      try {
        const res = await axios.get('https://graph.microsoft.com/v1.0/education/me/classes', {
          headers: { Authorization: `Bearer ${accessToken}` },
          timeout: 5000
        });
        
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "success",
              classes: res.data.value.map(cls => ({
                id: cls.id,
                displayName: cls.displayName
              }))
            })
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "error",
              message: `Error fetching class details: ${error.message}`
            })
          }]
        };
      }
    }
  );
  
  server.tool(
    "list-assignments",
    {}, // Empty schema for no parameters
    async () => {
      console.error("📝 list-assignments tool called");
      
      if (!isAuthenticated) {
        return { 
          content: [{ 
            type: "text", 
            text: JSON.stringify({
              status: "error",
              message: "User not authenticated. Please use the microsoft-login tool first."
            })
          }]
        };
      }
      
      try {
        const res = await axios.get('https://graph.microsoft.com/v1.0/education/me/assignments', {
          headers: { Authorization: `Bearer ${accessToken}` },
          timeout: 5000
        });
        
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "success",
              assignments: res.data.value.map(assignment => ({
                id: assignment.id,
                displayName: assignment.displayName
              }))
            })
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "error",
              message: `Error fetching assignments: ${error.message}`
            })
          }]
        };
      }
    }
  );
  
  server.tool(
    "get-assignment-details",
    {
      classId: z.string().describe("The ID of the class to get assignments from"),
      assignmentId: z.string().describe("The ID of the assignment to get details for"),           
    },
    async ({ classId, assignmentId }) => {
      console.error("📝 get-assignment-details tool called");
      
      if (!isAuthenticated) {
        return { 
          content: [{ 
            type: "text", 
            text: JSON.stringify({
              status: "error",
              message: "User not authenticated. Please use the microsoft-login tool first."
            })
          }]
        };
      }

      try {
        // Get details for the specified assignment
        const detailsRes = await axios.get(
          `https://graph.microsoft.com/v1.0/education/classes/${classId}/assignments/${assignmentId}`, 
          {
            headers: { Authorization: `Bearer ${accessToken}` },
            timeout: 5000
          }
        );
        
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "success",
              assignmentDetails: detailsRes.data
            })
          }]
        };
      } catch (error) {
        // More specific error handling
        let errorMessage = "Unknown error occurred";
        
        if (error.response) {
          // The request was made and the server responded with a status code
          // that falls out of the range of 2xx
          errorMessage = `API error: ${error.response.status} - ${error.response.data?.error?.message || 'Unknown API error'}`;
        } else if (error.request) {
          // The request was made but no response was received
          errorMessage = "Network error: No response received from server";
        } else {
          // Something happened in setting up the request that triggered an Error
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
    }
  );

  server.tool(
    "get-assignment-rubric",
    {
      classId: z.string().describe("The ID of the class to get assignments from"),
      assignmentId: z.string().describe("The ID of the assignment to get details for"),           
    },
    async ({ classId, assignmentId }) => {
      console.error("📝 get-assignment-rubric tool called");
      
      if (!isAuthenticated) {
        return { 
          content: [{ 
            type: "text", 
            text: JSON.stringify({
              status: "error",
              message: "User not authenticated. Please use the microsoft-login tool first."
            })
          }]
        };
      }

      try {
        // Get details for the specified assignment
        const detailsRes = await axios.get(
          `https://graph.microsoft.com/v1.0/education/classes/${classId}/assignments/${assignmentId}/rubric`, 
          {
            headers: { Authorization: `Bearer ${accessToken}` },
            timeout: 5000
          }
        );
        
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "success",
              assignmentDetails: detailsRes.data
            })
          }]
        };
      } catch (error) {
        // More specific error handling
        let errorMessage = "Unknown error occurred";
        
        if (error.response) {
          // The request was made and the server responded with a status code
          // that falls out of the range of 2xx
          errorMessage = `API error: ${error.response.status} - ${error.response.data?.error?.message || 'Unknown API error'}`;
        } else if (error.request) {
          // The request was made but no response was received
          errorMessage = "Network error: No response received from server";
        } else {
          // Something happened in setting up the request that triggered an Error
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
    }
  );

  server.tool(
    "get-assignment-submissions",
    {
      classId: z.string().describe("The ID of the class to get assignments from"),
      assignmentId: z.string().describe("The ID of the assignment to get details for"),           
    },
    async ({ classId, assignmentId }) => {
      console.error("📝 get-assignment-rubric tool called");
      
      if (!isAuthenticated) {
        return { 
          content: [{ 
            type: "text", 
            text: JSON.stringify({
              status: "error",
              message: "User not authenticated. Please use the microsoft-login tool first."
            })
          }]
        };
      }

      try {
        // Get details for the specified assignment
        const detailsRes = await axios.get(
          `https://graph.microsoft.com/v1.0/education/classes/${classId}/assignments/${assignmentId}/submissions`, 
          {
            headers: { Authorization: `Bearer ${accessToken}` },
            timeout: 5000
          }
        );
        
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "success",
              assignmentDetails: detailsRes.data
            })
          }]
        };
      } catch (error) {
        // More specific error handling
        let errorMessage = "Unknown error occurred";
        
        if (error.response) {
          // The request was made and the server responded with a status code
          // that falls out of the range of 2xx
          errorMessage = `API error: ${error.response.status} - ${error.response.data?.error?.message || 'Unknown API error'}`;
        } else if (error.request) {
          // The request was made but no response was received
          errorMessage = "Network error: No response received from server";
        } else {
          // Something happened in setting up the request that triggered an Error
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
    }
  );

  server.tool(
    "get-assignment-submissions-outcome",
    {
      classId: z.string().describe("The ID of the class to get assignments from"),
      assignmentId: z.string().describe("The ID of the assignment to get details for"),           
      submissionId: z.string().describe("The ID of the submission to get outcome for"),           
    },
    async ({ classId, assignmentId,submissionId }) => {
      console.error("📝 get-assignment-rubric tool called");
      
      if (!isAuthenticated) {
        return { 
          content: [{ 
            type: "text", 
            text: JSON.stringify({
              status: "error",
              message: "User not authenticated. Please use the microsoft-login tool first."
            })
          }]
        };
      }

      try {
        // Get details for the specified assignment
        const detailsRes = await axios.get(
          `https://graph.microsoft.com/v1.0/education/classes/${classId}/assignments/${assignmentId}/submissions/${submissionId}/outcomes`, 
          {
            headers: { Authorization: `Bearer ${accessToken}` },
            timeout: 5000
          }
        );
        
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "success",
              assignmentDetails: detailsRes.data
            })
          }]
        };
      } catch (error) {
        // More specific error handling
        let errorMessage = "Unknown error occurred";
        
        if (error.response) {
          // The request was made and the server responded with a status code
          // that falls out of the range of 2xx
          errorMessage = `API error: ${error.response.status} - ${error.response.data?.error?.message || 'Unknown API error'}`;
        } else if (error.request) {
          // The request was made but no response was received
          errorMessage = "Network error: No response received from server";
        } else {
          // Something happened in setting up the request that triggered an Error
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
    }
  );
  
  server.tool(
    "create-assignment",
    {
      classId: z.string().describe("The ID of the class to create the assignment in"),
      displayName: z.string().describe("The display name of the assignment"),
      dueDateTime: z.string().describe("The due date and time for the assignment in ISO format"),
      instructions: z.string().optional().describe("Optional: The instructions for the assignment"),
      status: z.string().optional().default("draft").describe("Optional: Status of the assignment (draft, published)"),
      allowLateSubmissions: z.boolean().optional().describe("Optional: Whether late submissions are allowed"),
      allowStudentsToAddResourcesToSubmission: z.boolean().optional().describe("Optional: Whether students can add resources to their submission"),
      assignDateTime: z.string().optional().describe("Optional: The date when the assignment should be assigned"),
      addedStudentAction: z.string().optional().describe("Optional: Action to take when students are added to the class"),
      addToCalendarAction: z.string().optional().describe("Optional: Whether to add the assignment to student calendars"),
      grading: z.object({
        "@odata.type": z.string().default("#microsoft.graph.educationAssignmentPointsGradeType"),
        maxPoints: z.number()
      }).optional().describe("Optional: Grading type and parameters for the assignment"),
      feedbackResourcesFolderUrl: z.string().optional().describe("Optional: URL for feedback resources folder"),
      notificationChannelUrl: z.string().optional().describe("Optional: URL for notification channel"),
      resourcesFolderUrl: z.string().optional().describe("Optional: URL for resources folder"),
      categories: z.array(z.string()).optional().describe("Optional: Categories for the assignment")
    },
    async ({ 
      classId, 
      displayName, 
      dueDateTime, 
      instructions, 
      status,
      allowLateSubmissions,
      allowStudentsToAddResourcesToSubmission,
      assignDateTime,
      addedStudentAction,
      addToCalendarAction,
      grading,
      feedbackResourcesFolderUrl,
      notificationChannelUrl,
      resourcesFolderUrl,
      categories
    }) => {
      console.error("📝 create-assignment tool called");
     
      if (!isAuthenticated) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "error",
              message: "User not authenticated. Please use the microsoft-login tool first."
            })
          }]
        };
      }
      
      try {
        // Create base payload with required fields
        const assignmentPayload = {
          displayName,
          dueDateTime,
          status: status || "draft"
        };
        
       
        // Add optional fields if provided
        if (instructions) {
          assignmentPayload.instructions = {
            contentType: "text",
            content: instructions
          };
        }
        
        if (allowLateSubmissions !== undefined) {
          assignmentPayload.allowLateSubmissions = allowLateSubmissions;
        }
        
        if (allowStudentsToAddResourcesToSubmission !== undefined) {
          assignmentPayload.allowStudentsToAddResourcesToSubmission = allowStudentsToAddResourcesToSubmission;
        }
        
        if (assignDateTime) {
          assignmentPayload.assignDateTime = assignDateTime;
        }
        
        if (addedStudentAction) {
          assignmentPayload.addedStudentAction = addedStudentAction;
        }
        
        if (addToCalendarAction) {
          assignmentPayload.addToCalendarAction = addToCalendarAction;
        }
        
        if (grading) {
          assignmentPayload.grading = grading;
        }
        
        if (feedbackResourcesFolderUrl) {
          assignmentPayload.feedbackResourcesFolderUrl = feedbackResourcesFolderUrl;
        }
        
        if (notificationChannelUrl) {
          assignmentPayload.notificationChannelUrl = notificationChannelUrl;
        }
        
        if (resourcesFolderUrl) {
          assignmentPayload.resourcesFolderUrl = resourcesFolderUrl;
        }
        
        if (categories && categories.length > 0) {
          assignmentPayload.categories = categories;
        }
        
        // Create the assignment
        const createRes = await axios.post(
          `https://graph.microsoft.com/v1.0/education/classes/${classId}/assignments`,
          assignmentPayload,
          {
            headers: { Authorization: `Bearer ${accessToken}` },
            timeout: 5000
          }
        );
       
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "success",
              assignmentDetails: createRes.data
            })
          }]
        };
      } catch (error) {
        // Error handling
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
    }
  );

  server.tool(
    "update-assignment",
    {
      classId: z.string().describe("The ID of the class to create the assignment in"),
      assignmentId: z.string().describe("The ID of the assignment to update"),
      displayName: z.string().describe("The display name of the assignment"),
      dueDateTime: z.string().describe("The due date and time for the assignment in ISO format"),
      assignTo: z.object({
        "@odata.type": z.string().describe("The type of recipient (e.g., #microsoft.graph.educationAssignmentClassRecipient, #microsoft.graph.educationAssignmentIndividualRecipient)"),
        recipients: z.array(z.string()).optional().describe("Optional: Student IDs to assign the assignment to (required if using IndividualRecipient)")
      }).optional().describe("Optional: Specify which students to assign to. Default is the whole class"),
      instructions: z.string().optional().describe("Optional: The instructions for the assignment"),
    },
    async ({ 
      classId,
      assignmentId,
      displayName, 
      dueDateTime, 
      assignTo,
      instructions
     }) => {
      console.error("📝 update-assignment tool called");
     
      if (!isAuthenticated) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "error",
              message: "User not authenticated. Please use the microsoft-login tool first."
            })
          }]
        };
      }
      
      try {
        // Create base payload with required fields
        const assignmentPayload = {
          displayName,
          dueDateTime
        };
        
        // Add assignTo - if not provided, default to whole class
        assignmentPayload.assignTo = assignTo || {
          "@odata.type": "#microsoft.graph.educationAssignmentClassRecipient"
        };
        
        // Add optional fields if provided
        if (instructions) {
          assignmentPayload.instructions = {
            contentType: "text",
            content: instructions
          };
        }
        
        const createRes = await axios.patch(
          `https://graph.microsoft.com/v1.0/education/classes/${classId}/assignments/${assignmentId}`,
          assignmentPayload,
          {
            headers: { Authorization: `Bearer ${accessToken}` },
            timeout: 5000
          }
        );
       
        console.error(assignmentPayload);    

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "success",
              assignmentDetails: createRes.data
            })
          }]
        };
      } catch (error) {
        // Error handling
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