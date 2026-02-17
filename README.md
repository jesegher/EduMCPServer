# 📚 Microsoft Education MCP Server

> ⚠️ **This is a personal project. It is not affiliated with or maintained by Microsoft.**

This project is a custom **Model Context Protocol (MCP) server** built to integrate with **Microsoft Graph API** for Education.

It enables VSCode or Copilot Studio Agents to manage:
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
- 🧠 Designed for VSCode and Copilot Studio
---
## 📂 Structure
### 🔍 Key Components

- **`start-mcp-server-streaming.js`**
  - Loads and registers all tools with OAuth 2.0 authentication layer.
  - Implements OAuth 2.0 server with Windows account picker integration.
  - Manages Microsoft Graph authentication via MSAL for backend API access.
  - Starts the MCP server using the Model Context Protocol SDK with streaming transport.

- **`tools/` folder**
  - Each file defines a set of related tools and registers them with the server.
  - Tools follow the MCP standard (`server.tool(...)`) with Zod schema validation and Microsoft Graph integration.

- **`.env`**
  - Stores sensitive Microsoft app credentials and config.
  - Required for Microsoft Graph API authentication.

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
     - URI: `https://<YOUR_APP_NAME>.azurewebsites.net/oauth/callback`

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
REDIRECT_URI=https://<YOUR_APP_NAME>.azurewebsites.net/oauth/callback
```
---


## 🚀 Azure Deployment Guide

Deploy your MCP server to Azure App Service for production use with OAuth 2.0 authentication.

### 1. 📋 Prerequisites

- ✅ Azure subscription
- ✅ Azure App registration (completed above)
- ✅ Code pushed to GitHub repository

### 2. 🏗️ Create Azure App Service

#### Option A: Azure Portal
1. Go to [Azure Portal](https://portal.azure.com)
2. **Create a resource** → **Web App**
3. Configure:
   - **Subscription**: Your subscription
   - **Resource Group**: Create new or use existing
   - **Name**: `your-app-name` (e.g., `edumcp-server-prod`)
   - **Runtime stack**: `Node 22 LTS`
   - **Operating System**: `Linux`
   - **Region**: Choose closest to users


### 3. 🔐 Configure Environment Variables

In Azure Portal → Your App Service → **Configuration** → **Application settings**, add:

```
CLIENT_ID = your-azure-app-client-id
TENANT_ID = your-tenant-id  
CLIENT_SECRET = your-client-secret
REDIRECT_URI = https://<YOUR_APP_NAME>.azurewebsites.net/oauth/callback
```

### 4. 📦 Deploy Your Code

#### Option A: External Git (Recommended)
1. **Azure Portal** → Your App Service → **Deployment Center**
2. **Source**: External Git
3. **Repository**: `https://github.com/jesegher/EduMCPServer.git`
4. **Branch**: `main`
5. **Repository Type**: Public
6. **Save**

#### Option B: VS Code Extension
1. Install **Azure App Service** extension
2. Right-click your project → **Deploy to Web App**
3. Select your subscription and app service

### 5. 🎯 Configure MCP Clients

Update your MCP client configurations for production:

#### VS Code MCP Client Config:

Add to your VS Code `mcp.json` (User settings):
```json
{
  "servers": {
    "EDUMCP Server": {
      "url": "https://your-app-name.azurewebsites.net/mcp",
      "type": "http"
    }
  }
}
```

VS Code automatically discovers OAuth configuration via the `/.well-known/oauth-authorization-server` endpoint.

### 6. ✅ Verify Deployment

1. **Test Endpoints**:
   - Health check: `https://your-app-name.azurewebsites.net`
   - OAuth discovery: `https://your-app-name.azurewebsites.net/.well-known/oauth-authorization-server`

2. **Check Logs** in Azure Portal → App Service → **Log stream**

3. **Test Authentication** with your MCP client

### 7. 🛡️ Production Security

- ✅ Use **Key Vault** for sensitive secrets (optional)
- ✅ Enable **HTTPS only** in App Service → TLS/SSL settings
- ✅ Configure **custom domain** (optional)
- ✅ Set up **Application Insights** for monitoring

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

