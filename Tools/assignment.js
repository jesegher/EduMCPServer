import axios from "axios";
import { z } from 'zod';

function registerAssignmentTools(server, auth) {
  
  // 📄 Get assignment details or list all
  server.tool(
    "assignment_get","Use this function when you need to get assignment details or list all assignments",
    {
      classId: z.string().optional().describe("The class ID of the assignment (required for getting a specific assignment)"),
      assignmentId: z.string().optional().describe("The assignment ID to fetch details for"),
      includeRubric: z.boolean().optional().default(true).describe("Whether to attempt retrieving the rubric if grading is null")
    },
    async ({ classId, assignmentId, includeRubric }) => {
      console.error("📝 assignment_get tool called");

      if (!auth.isAuthenticated) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: "error",
              message: "User not authenticated"
            })
          }]
        };
      }

      // 📄 Get specific assignment
      if (classId && assignmentId) {
        try {
          const assignmentRes = await axios.get(
            `https://graph.microsoft.com/v1.0/education/classes/${classId}/assignments/${assignmentId}`,
            {
              headers: { Authorization: `Bearer ${auth.accessToken}` },
              timeout: 5000
            }
          );

          const assignment = assignmentRes.data;
          let rubric = null;

          if (includeRubric && !assignment.grading) {
            try {
              const rubricRes = await axios.get(
                `https://graph.microsoft.com/v1.0/education/classes/${classId}/assignments/${assignmentId}/rubric`,
                {
                  headers: { Authorization: `Bearer ${auth.accessToken}` },
                  timeout: 5000
                }
              );
              rubric = rubricRes.data;
            } catch (rubricError) {
              if (rubricError.response?.status !== 404) {
                console.warn("⚠️ Rubric fetch failed", rubricError.message);
              }
            }
          }

          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                status: "success",
                assignmentDetails: assignment,
                rubric: rubric || null
              })
            }]
          };
        } catch (error) {
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                status: "error",
                message: `Error fetching assignment details: ${error.message}`
              })
            }]
          };
        }
      }

      // 📋 List all assignments
      try {
        const res = await axios.get('https://graph.microsoft.com/v1.0/education/me/assignments', {
          headers: { Authorization: `Bearer ${auth.accessToken}` },
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
    },
    {
      description: "Lists all assignments or fetches a specific assignment's details (optionally including its rubric)."
    }
  );

  server.tool(
    "assignment-submissions-get",
    {
      classId: z.string().describe("The ID of the class to get assignments from"),
      assignmentId: z.string().describe("The ID of the assignment to get details for"),           
    },
    async ({ classId, assignmentId }) => {
      console.error("📝 get-assignment-rubric tool called");
      
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
        // Get details for the specified assignment
        const detailsRes = await axios.get(
          `https://graph.microsoft.com/v1.0/education/classes/${classId}/assignments/${assignmentId}/submissions`, 
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
    "assignment-submissions-outcome-get",
    {
      classId: z.string().describe("The ID of the class to get assignments from"),
      assignmentId: z.string().describe("The ID of the assignment to get details for"),           
      submissionId: z.string().describe("The ID of the submission to get outcome for"),           
    },
    async ({ classId, assignmentId,submissionId }) => {
      console.error("📝 get-assignment-rubric tool called");
      
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
        // Get details for the specified assignment
        const detailsRes = await axios.get(
          `https://graph.microsoft.com/v1.0/education/classes/${classId}/assignments/${assignmentId}/submissions/${submissionId}/outcomes`, 
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
    "assignment-create",
    {
      classId: z.string().describe("The ID of the class to create the assignment in"),
      displayName: z.string().describe("The display name of the assignment"),
      dueDateTime: z.string().describe("The due date and time for the assignment in ISO format"),
      instructions: z.string().optional().describe("Optional: The instructions for the assignment"),
      assignTo: z.object({
        "@odata.type": z.string().describe("The type of recipient (e.g., #microsoft.graph.educationAssignmentClassRecipient, #microsoft.graph.educationAssignmentIndividualRecipient)"),
        recipients: z.array(z.string()).optional().describe("Optional: Student IDs to assign the assignment to (required if using IndividualRecipient)")
      }).optional().describe("Optional: Specify which students to assign to. Default is the whole class"),
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
      assignTo,
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
        // Create base payload with required fields
        const assignmentPayload = {
          displayName,
          dueDateTime,
          status: "draft"
        };
        
        assignmentPayload.assignTo = assignTo || {
          "@odata.type": "#microsoft.graph.educationAssignmentClassRecipient"
        };
        
       
        // Add optional fields if provided
        if (instructions) {
          assignmentPayload.instructions = {
            contentType: "html",
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
            headers: { Authorization: `Bearer ${auth.accessToken}` },
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
    "assignment-update",
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
            contentType: "html",
            content: instructions
          };
        }
        
        const createRes = await axios.patch(
          `https://graph.microsoft.com/v1.0/education/classes/${classId}/assignments/${assignmentId}`,
          assignmentPayload,
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

  //Attach rubric to assignment
  server.tool(
    "assignment-rubric-attach",
    {
      classId: z.string().describe("The ID of the class that contains the assignment"),
      assignmentId: z.string().describe("The ID of the assignment to attach the rubric to"),
      rubricId: z.string().describe("The ID of the rubric to attach to the assignment")
    },
    async ({ classId, assignmentId, rubricId }) => {
      console.error("📎 attach-rubric-to-assignment tool called");
  
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
        const odataRef = {
          "@odata.id": `https://graph.microsoft.com/v1.0/education/me/rubrics/${rubricId}`
        };
  
        const response = await axios.put(
          `https://graph.microsoft.com/v1.0/education/classes/${classId}/assignments/${assignmentId}/rubric/$ref`,
          odataRef,
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
              message: "Rubric attached successfully.",
              result: response.data
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

export default registerAssignmentTools;
