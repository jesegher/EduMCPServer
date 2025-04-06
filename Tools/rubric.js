import axios from "axios";
import { z } from 'zod';

function registerRubricTools(server, auth) {
  
server.tool(
    "rubric-create-get",
    {
      displayName: z.string().describe("The display name of the rubric"),
      maxPoints: z.number().describe("The maximum number of points the rubric is worth"),
      qualities: z.array(z.object({
        description: z.object({
          content: z.string(),
          contentType: z.string()
        }),
        criteria: z.array(z.object({
          description: z.object({
            content: z.string(),
            contentType: z.string()
          })
        }))
      }))
    },
    async ({ displayName, maxPoints, qualities }) => {
      console.error("🔁 create-or-get-rubric tool called");
  
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
        // Step 1: Check for existing rubric
        const checkRes = await axios.get(
          `https://graph.microsoft.com/v1.0/education/me/rubrics?$filter=displayName eq '${displayName.replace(/'/g, "''")}'`,
          {
            headers: { Authorization: `Bearer ${auth.accessToken}` },
            timeout: 5000
          }
        );
  
        const existing = checkRes.data?.value?.[0];
  
        if (existing) {
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                status: "exists",
                message: "ℹ️ Rubric already exists. Returning existing rubric.",
                rubricId: existing.id,
                rubric: existing
              })
            }]
          };
        }
  
        // Step 2: Build levels
        const levels = qualities[0].criteria.map((_, index) => ({
          displayName: `Level ${index + 1}`,
          description: {
            content: "",
            contentType: "text"
          }
        }));
  
        // Step 3: Create new rubric
        const rubricPayload = {
          displayName,
          description: {
            content: `Rubric worth ${maxPoints} points`,
            contentType: "text"
          },
          levels,
          qualities,
          grading: {
            "@odata.type": "#microsoft.graph.educationAssignmentPointsGradeType",
            maxPoints
          }
        };
  
        const createRes = await axios.post(
          `https://graph.microsoft.com/v1.0/education/me/rubrics`,
          rubricPayload,
          {
            headers: { Authorization: `Bearer ${auth.accessToken}` },
            timeout: 5000
          }
        );
  
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "created",
              message: "✅ Rubric created successfully.",
              rubricId: createRes.data.id,
              rubric: createRes.data
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
      description: "Creates a rubric if one doesn't already exist with the same display name. Prevents duplication and returns the existing rubric if found."
    }
  );
  
  server.tool(
    "rubric_list",
    {},
    async () => {
      console.error("📋 rubric_list tool called");
  
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
        const res = await axios.get(
          `https://graph.microsoft.com/v1.0/education/me/rubrics`,
          {
            headers: { Authorization: `Bearer ${auth.accessToken}` },
            timeout: 5000
          }
        );
  
        const rubrics = res.data.value.map(rubric => ({
          id: rubric.id,
          displayName: rubric.displayName,
          maxPoints: rubric.grading?.maxPoints ?? null
        }));
  
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "success",
              count: rubrics.length,
              rubrics
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
      description: "Lists all rubrics created by the authenticated user."
    }
  );
  

}

export default registerRubricTools;
