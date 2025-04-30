import axios from "axios";
import { z } from 'zod';

function registerUserTools(server, auth) {
  
  server.tool("user-get", "Fetches a user based on userid, UPN or search string.", {
    userId: z.string().optional().nullable(),
    userPrincipalName: z.string().optional().nullable(),
    search: z.string().optional().nullable()
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
      if (typeof userId === 'string' && userId.trim() !== '') {
        const userResponse = await axios.get(
          `https://graph.microsoft.com/v1.0/users/${userId}`,
          { headers: { Authorization: `Bearer ${auth.accessToken}` } }
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
  
      if (typeof userPrincipalName === 'string' && userPrincipalName.trim() !== '') {
        const userResponse = await axios.get(
          `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userPrincipalName)}`,
          { headers: { Authorization: `Bearer ${auth.accessToken}` } }
        );
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "success",
              message: "User retrieved by UPN.",
              user: userResponse.data
            })
          }]
        };
      }
  
      if (typeof search === 'string' && search.trim() !== '') {
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
  
      // Fallback: fetch all users (paginated)
      let allUsers = [];
      let nextLink = `https://graph.microsoft.com/v1.0/users?$top=50`;
  
      do {
        const response = await axios.get(nextLink, {
          headers: { Authorization: `Bearer ${auth.accessToken}` }
        });
        allUsers = allUsers.concat(response.data.value);
        nextLink = response.data["@odata.nextLink"];
      } while (nextLink);
  
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: "success",
            message: "All users retrieved successfully.",
            count: allUsers.length,
            users: allUsers
          })
        }]
      };
  
    } catch (error) {
      const errorMessage =
        error.response?.data?.error?.message ||
        error.message ||
        "Unknown error occurred";
  
      return {
        content: [{
          type: "text",
          text: JSON.stringify({ status: "error", message: errorMessage })
        }]
      };
    }
  });
  

server.tool("user-attributes-select-get", "Retrieves only specific attributes from Entra ID. Specify which fields you need (e.g., displayName, jobTitle, skills, AgeGroup).", {
  userId: z.string().optional(),
  fields: z.array(z.string()).describe("List of field names to retrieve (e.g., displayName, userPrincipalName, jobTitle)")
}, async ({ userId, fields }) => {
  console.error("🎯 user-attributes-select-get called with:", fields);

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

  const fieldQuery = fields.join(",");
  const endpoint = `https://graph.microsoft.com/v1.0/users/${userId || 'me'}?$select=${fieldQuery}`;

  try {
    const response = await axios.get(endpoint, {
      headers: {
        Authorization: `Bearer ${auth.accessToken}`
      }
    });

    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: "success",
          message: "✅ Selected attributes retrieved.",
          attributes: response.data
        }, null, 2)
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
   
server.tool("user-groups-get", "Fetches the groups a user belongs to based on userId.", {
  userId: z.string()
}, async ({ userId }) => {
  console.error("👥 user-groups-get tool called");

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
    const groupsResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${userId}/memberOf`,
      { headers: { Authorization: `Bearer ${auth.accessToken}` } }
    );

    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: "success",
          message: "✅ User groups retrieved successfully.",
          groups: groupsResponse.data.value
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
