import axios from "axios";
import { z } from 'zod';

function registerUserTools(server, auth) {
  
 server.tool("user-get","Fetches a user based on userid, UPN or search string.", {
    userId: z.string().optional(),
    userPrincipalName: z.string().optional(),
    search: z.string().optional()
  }, async ({ userId, userPrincipalName, search }) => {
    console.error("🔍 get-user tool called");

    if (!auth.isAuthenticated) {
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
      if (userId) {
        userResponse = await axios.get(
          `https://graph.microsoft.com/v1.0/users/${userId}`,
          { headers: { Authorization: `Bearer ${auth.accessToken}` } }
        );
        return {
          content: [{
            type: "text",
            text: JSON.stringify({ status: "success", message: "User retrieved by ID.", user: userResponse.data })
          }]
        };
      }

      if (userPrincipalName) {
        userResponse = await axios.get(
          `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userPrincipalName)}`,
          { headers: { Authorization: `Bearer ${auth.accessToken}` } }
        );
        return {
          content: [{
            type: "text",
            text: JSON.stringify({ status: "success", message: "User retrieved by UPN.", user: userResponse.data })
          }]
        };
      }

      if (search) {
        const searchResponse = await axios.get(
          `https://graph.microsoft.com/v1.0/users?$search="displayName:${search}"&$count=true`,
          {
            headers: {
              Authorization: `Bearer ${auth.accessToken}`,
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

      const allResponse = await axios.get(
        `https://graph.microsoft.com/v1.0/users?$top=10`,
        { headers: { Authorization: `Bearer ${auth.accessToken}` } }
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
        errorMessage = "Network error: No response received";
      } else {
        errorMessage = `Request error: ${error.message}`;
      }
      return {
        content: [{
          type: "text",
          text: JSON.stringify({ status: "error", message: errorMessage })
        }]
      };
    }
  }
);

server.tool("user-update", "Updates a user based on userId. Only fields included in the input will be changed.", {
    userId: z.string(),
    updates: z.record(z.any()) // Flexible update payload
  }, async ({ userId, updates }) => {
    console.error("✏️ user-update tool called");
  
    if (!auth.isAuthenticated) {
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
      const updateResponse = await axios.patch(
        `https://graph.microsoft.com/v1.0/users/${userId}`,
        updates,
        { headers: { Authorization: `Bearer ${auth.accessToken}` } }
      );
  
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: "success",
            message: "✅ User updated successfully.",
            user: updateResponse.data
          })
        }]
      };
    } catch (error) {
      let errorMessage = "Unknown error occurred";
      if (error.response) {
        errorMessage = `API error: ${error.response.status} - ${error.response.data?.error?.message || 'Unknown API error'}`;
      } else if (error.request) {
        errorMessage = "Network error: No response received";
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
  });
  
 

  

}

export default registerUserTools;
