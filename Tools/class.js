import axios from "axios";
import { z } from 'zod';

function registerClassTools(server, auth) {
  
    server.tool(
        "class_get","Gets all classes or a specific class with optional search.",
        {
            classId: z.string().optional().describe("Optional: The ID of the class to retrieve"),
            search: z.string().optional().describe("Optional: Filter classes by display name (contains match)")
        },
        async ({ classId, search }) => {
            console.error("📚 class_get tool called");
        
            if (!auth.isAuthenticated || !auth.accessToken) {
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
        
            const headers = {
            Authorization: `Bearer ${auth.accessToken}`
            };
        
            try {
            let classes = [];
        
            if (classId) {
                // 🔍 Get single class
                const classRes = await axios.get(
                `https://graph.microsoft.com/v1.0/education/classes/${classId}`,
                { headers, timeout: 5000 }
                );
        
                const classData = classRes.data;
        
                // 👥 Also fetch roster
                const rosterRes = await axios.get(
                `https://graph.microsoft.com/v1.0/education/classes/${classId}/members`,
                { headers, timeout: 5000 }
                );
        
                const students = [];
                const teachers = [];
        
                for (const member of rosterRes.data.value) {
                const role = member?.roles?.includes("Teacher") ? "teacher" : "student";
                const entry = {
                    id: member.id,
                    displayName: member.displayName,
                    email: member.userPrincipalName,
                    role
                };
                role === "teacher" ? teachers.push(entry) : students.push(entry);
                }
        
                classes = [{
                ...classData,
                roster: {
                    students,
                    teachers
                }
                }];
            } else {
                // 📋 Get all classes
                const res = await axios.get(`https://graph.microsoft.com/v1.0/education/me/classes`, {
                headers,
                timeout: 5000
                });
                classes = res.data.value;
        
                // 🔎 Optional: Filter by search
                if (search) {
                const lowerSearch = search.toLowerCase();
                classes = classes.filter(cls =>
                    cls.displayName?.toLowerCase().includes(lowerSearch)
                );
                }
            }
        
            return {
                content: [{
                type: "text",
                text: JSON.stringify({
                    status: "success",
                    classes: classes.map(cls => ({
                    id: cls.id,
                    displayName: cls.displayName,
                    roster: cls.roster ?? undefined // only present for single class lookups
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
        },
        {
            description: "Gets all classes or a specific class with optional search. If classId is provided, includes full student and teacher roster."
        }
        );
      

}

export default registerClassTools;