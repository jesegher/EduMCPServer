import axios from "axios";
import { z } from 'zod';

function registerGroupTools(server, getAuth) {

  // Tool to fetch a group by ID or display name
  server.tool("group-get", "Fetches a group based on groupId or display name.", {
    groupId: z.string().optional(),
    displayName: z.string().optional()
  }, async ({ groupId, displayName }) => {
    console.error("🔍 group-get tool called");

    const auth = getAuth(); // Get current request auth
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
      let groupResponse;
      if (groupId) {
        groupResponse = await axios.get(
          `https://graph.microsoft.com/v1.0/groups/${groupId}`,
          { headers: { Authorization: `Bearer ${auth.accessToken}` } }
        );
        return {
          content: [{
            type: "text",
            text: JSON.stringify({ status: "success", message: "Group retrieved by ID.", group: groupResponse.data })
          }]
        };
      }

      if (displayName) {
        const searchResponse = await axios.get(
          `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${displayName}'`,
          { headers: { Authorization: `Bearer ${auth.accessToken}` } }
        );
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "success",
              message: "Group retrieved by display name.",
              groups: searchResponse.data.value
            })
          }]
        };
      }

      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: "error",
            message: "❌ Please provide either groupId or displayName."
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
  });

  // Tool to update a group's attributes
  server.tool("group-update", "Updates a group based on groupId. Only fields included in the input will be changed.", {
    groupId: z.string(),
    updates: z.record(z.any()) // Flexible update payload
  }, async ({ groupId, updates }) => {
    console.error("✏️ group-update tool called");

    const auth = getAuth(); // Get current request auth
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
        `https://graph.microsoft.com/v1.0/groups/${groupId}`,
        updates,
        { headers: { Authorization: `Bearer ${auth.accessToken}` } }
      );

      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: "success",
            message: "✅ Group updated successfully.",
            group: updateResponse.data
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

  // Tool to fetch members of a group
  server.tool("group-members-get", "Fetches the members of a group based on groupId.", {
    groupId: z.string()
  }, async ({ groupId }) => {
    console.error("👥 group-members-get tool called");

    const auth = getAuth(); // Get current request auth
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
      const membersResponse = await axios.get(
        `https://graph.microsoft.com/v1.0/groups/${groupId}/members`,
        { headers: { Authorization: `Bearer ${auth.accessToken}` } }
      );

      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: "success",
            message: "✅ Group members retrieved successfully.",
            members: membersResponse.data.value
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

export default registerGroupTools;