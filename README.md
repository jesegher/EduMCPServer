# 📚 Microsoft Education MCP Server

> ⚠️ **This is a personal project. It is not affiliated with or maintained by Microsoft.**

This project is a custom **Model Context Protocol (MCP) server** built to integrate with **Microsoft Graph API** for Education.

It enables Claude Desktop or other MCP-compatible tools to manage:
- ✅ Microsoft Education Classes
- ✅ Assignments (create, update, target students)
- ✅ Rubrics (create, attach, list)
- ✅ Students and teachers (roster)
- ✅ Submissions and grading

Built for AI-driven tools, testing, and intelligent prompt integration.

---

## 🚀 Features

- 🔐 Microsoft delegated authentication (OAuth via MSAL)
- 🧑‍🏫 Class & roster exploration
- 📝 Assignment creation, updating, and student targeting
- 🎓 Rubric creation and re-use
- 📤 View assignment submissions & outcomes
- 🧠 Designed for Claude Desktop & Model Context clients
---
## 📂 Structure
### 🔍 Key Components

- **`start-mcp-server.js`**
  - Loads and registers all tools.
  - Manages Microsoft authentication via MSAL.
  - Starts the MCP server using the Model Context Protocol SDK.

- **`tools/` folder**
  - Each file defines a set of related tools and registers them with the server.
  - Tools follow the MCP standard (`server.tool(...)`) with Zod schema validation and Microsoft Graph integration.

- **`.env`**
  - Stores sensitive Microsoft app credentials and config.
  - Required to authenticate with Microsoft Graph API.

- **`Example.env`**
  - A safe template for sharing or onboarding collaborators.

---

This layout makes it easy to:
- Add or modify functionality (just add a file in `tools/`)
- Keep logic separated by domain (e.g., assignments vs rubrics)
- Support clean and scalable MCP integration

---

## 🏗 How to Register the Application in Entra ID (Azure AD)

To use Microsoft Graph API with this project, you need to register an app in Entra ID:

### 1. Go to Microsoft Entra Admin Center

- Visit: [https://entra.microsoft.com](https://entra.microsoft.com)
- Sign in with an admin account.

---

### 2. Register a New Application

1. Navigate to **"Applications" → "App registrations"**
2. Click **"New registration"**
3. Fill in:
   - **Name**: `Microsoft Education MCP Server` (or any name you prefer)
   - **Supported account types**: Choose based on your scenario (usually "Accounts in this organizational directory only")
   - **Redirect URI**:  
     - Platform: `Web`  
     - URI: `http://localhost:3000/auth/callback` (or your custom URI)

4. Click **"Register"**

---

### 3. Configure API Permissions

1. After registration, go to **"API permissions"**
2. Click **"Add a permission" → Microsoft Graph → Delegated permissions**
3. Add the following:
   - `User.ReadWrite.All`
   - `EduAssignments.ReadWrite.All`
   - `EduRoster.ReadWrite.All`
   - `EduRubrics.ReadWrite`
   - `offline_access`
   - `openid`
   - `profile`

4. Click **"Grant admin consent"** to approve them for your tenant.

---

### 4. Generate a Client Secret

1. Go to **"Certificates & secrets"**
2. Under **Client secrets**, click **"New client secret"**
3. Add a description and choose an expiration (e.g. 6 months or 12 months)
4. Click **"Add"**
5. **Copy the value** — you won’t be able to see it again!

---

### 5. Save These Values in a notepad

After registration, go to **"Overview"** and copy these values:

```env
CLIENT_ID=your-application-id
CLIENT_SECRET=your-client-secret
TENANT_ID=your-directory-id
REDIRECT_URI=http://localhost:3000/auth/callback
```
---

## 📥 Clone the Repository and Install Dependencies

Follow these steps to get the project running locally:

### 1. Clone the repository

```bash
git clone https://github.com/jesegher/EduMCPServer.git
cd EDUMCPServer
```
### 2. Install dependencies
```bash
npm install
```
or manually
```bash
npm install @modelcontextprotocol/sdk axios zod dotenv @azure/msal-node
```

## 🧠 Run with Claude Desktop

To use this MCP server with Claude Desktop, add the following to your Claude Desktop `claude.settings.json` file:

```json
{
  "mcpServers": {
    "Education-Data": {
      "command": "node",
      "args": [
        "C:\\path\\to\\your\\project\\start-mcp-server.js"
      ],
      "env": {
        "TENANT_ID": "your-tenant-id",
        "CLIENT_ID": "your-client-id",
        "CLIENT_SECRET": "your-client-secret",
        "REDIRECT_URI": "http://localhost:3000/auth/callback",
        "PORT": "3000"
      }
    }
  }
}
```
💡 Replace the path and environment variables with your actual configuration details.

⚠️ Make sure to shutdown Claude every time you make a change. You need to kill it in the task manager. 

Claude will automatically detect the MCP Server and the registered tools.

<img width="362" alt="image" src="https://github.com/user-attachments/assets/924c9da6-0e0b-4c5d-9927-c74517702a5c" />

As soon as you request data, an authentication flow will be started. You can always trigger it manually by calling auth-login. I've noticed that Claude not always give you the url, you can find, and copy it from the call it makes to the authentication server if needed. [Fix in progress]

---


## 🧠 Scenario: Remediation Flow for Underperforming Students

This is a structured walkthrough based on original prompts for data analysis and follow-up instruction.


### 📝 Original Prompts

1. **Find my assignment 'lineair equation' in my algebra course. Give me the name, instructions and the due date.**

2. **I want an overview in table format of all underperforming students. I want their name, their feedback I provided, and the rubric if there is one associated. I want one line per student. Split up the rubric separate items.**

3. **Can you analyze the rubric feedback and find commonalities.**

4. **Based on this feedback, what are some ideas to remediate this.**

5. **I want them to practice variations.**

6. **Translate this into an assignment and rubric.**

7. **Create for those students a new draft assignment with this information ... Give them two weeks to complete it.**

---

Each prompt serves as a step in a larger automated or AI-assisted education workflow:

- Retrieving targeted assignments
- Analyzing rubric-aligned feedback
- Identifying patterns in student work
- Recommending learning interventions
- Creating and assigning personalized remediation tasks

