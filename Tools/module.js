import axios from "axios";
import { z } from 'zod';

function registerModuleTools(server, getAuth) {

  // 📦 List modules or get a specific module
  server.tool(
    "module-get", "Gets all modules for a class or a specific module. Provide classId + moduleId for a specific module, or classId only for all modules in a class.",
    {
      classId: z.string().describe("The ID of the class to get modules from"),
      moduleId: z.string().optional().describe("Optional: The ID of a specific module to retrieve")
    },
    async ({ classId, moduleId }) => {
      const auth = getAuth();
      console.error(`📦 module-get tool called [${auth.requestId}] for user ${auth.userId}`);

      if (!auth.isAuthenticated) {
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
        if (moduleId) {
          // Get specific module
          const res = await axios.get(
            `https://graph.microsoft.com/v1.0/education/classes/${classId}/modules/${moduleId}`,
            {
              headers: { Authorization: `Bearer ${auth.accessToken}` },
              timeout: 5000
            }
          );

          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                status: "success",
                module: res.data
              })
            }]
          };
        }

        // List all modules for a class
        const res = await axios.get(
          `https://graph.microsoft.com/v1.0/education/classes/${classId}/modules`,
          {
            headers: { Authorization: `Bearer ${auth.accessToken}` },
            timeout: 5000
          }
        );

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "success",
              count: res.data.value.length,
              modules: res.data.value.map(m => ({
                id: m.id,
                displayName: m.displayName,
                description: m.description,
                status: m.status,
                isPinned: m.isPinned,
                createdDateTime: m.createdDateTime,
                lastModifiedDateTime: m.lastModifiedDateTime
              }))
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
    }
  );

  // 📂 List module resources or get a specific resource
  server.tool(
    "module-resources-get", "Gets all resources for a module or a specific resource. Provide classId + moduleId + resourceId for a specific resource, or classId + moduleId for all resources.",
    {
      classId: z.string().describe("The ID of the class"),
      moduleId: z.string().describe("The ID of the module to get resources from"),
      resourceId: z.string().optional().describe("Optional: The ID of a specific resource to retrieve")
    },
    async ({ classId, moduleId, resourceId }) => {
      const auth = getAuth();
      console.error(`📂 module-resources-get tool called [${auth.requestId}] for user ${auth.userId}`);

      if (!auth.isAuthenticated) {
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
        if (resourceId) {
          // Get specific module resource
          const res = await axios.get(
            `https://graph.microsoft.com/v1.0/education/classes/${classId}/modules/${moduleId}/resources/${resourceId}`,
            {
              headers: { Authorization: `Bearer ${auth.accessToken}` },
              timeout: 5000
            }
          );

          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                status: "success",
                resource: res.data
              })
            }]
          };
        }

        // List all resources for a module
        const res = await axios.get(
          `https://graph.microsoft.com/v1.0/education/classes/${classId}/modules/${moduleId}/resources`,
          {
            headers: { Authorization: `Bearer ${auth.accessToken}` },
            timeout: 5000
          }
        );

        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "success",
              count: res.data.value.length,
              resources: res.data.value
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
    }
  );

}

export default registerModuleTools;
