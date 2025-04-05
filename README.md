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
git clone https://github.com/your-username/your-repo.git
cd your-repo
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




   
