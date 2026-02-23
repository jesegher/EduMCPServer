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
   - **Redirect URI**: Leave blank for now (will be added after Azure App Service is created)

4. Click **"Register"**

---

### 3. Configure API Permissions

1. After registration, go to **"API permissions"**
2. Click **"Add a permission" → Microsoft Graph → Delegated permissions**
3. Add the following permissions (use the search/filter to find each one):

   **OpenID permissions:**
   - `offline_access`
   - `openid`
   - `profile`

   **EduAssignments:**
   - `EduAssignments.ReadWrite`

   **EduRoster:**
   - `EduRoster.ReadWrite`

   **User:**
   - `User.ReadWrite.All`

4. Click **"Add permissions"**
5. Then, in **"Configured permissions"**, click **"Grant admin consent for [your tenant]"** to approve them for all users in your tenant.

> 💡 **Tip**: The "Grant admin consent" button only appears if you are signed in with an admin account. This step ensures all users (not just the admin) can use these permissions.

---

### 4. Generate a Client Secret

1. Go to **"Certificates & secrets"**
2. Under **Client secrets**, click **"New client secret"**
3. **Description**: Type any label you like (e.g., `mcp-server-secret`) — this is just a friendly name to help you identify it later.
4. **Expires**: Choose an expiration period (e.g., 6 months or 12 months)
5. Click **"Add"**
6. **Copy the "Value" immediately** — this is your `CLIENT_SECRET`. You won't be able to see it again after leaving this page!

> ⚠️ **Important**: The "Value" column is your secret, not the "Secret ID". Save it somewhere safe (e.g., a password manager or a notepad) — you'll need it later.

---

### 5. Save These Values

Go to your app registration's **"Overview"** page and copy these values. Save them in a notepad — you'll need them during Azure deployment.

| Name | Where to find it |
|------|------------------|
| `CLIENT_ID` | **Application (client) ID** on the Overview page |
| `CLIENT_SECRET` | The **Value** you copied in Step 4 |
| `TENANT_ID` | **Directory (tenant) ID** on the Overview page |
| `REDIRECT_URI` | We'll create this in the Azure Deployment section below — skip it for now |

---


## 🚀 Azure Deployment Guide

Deploy your MCP server to Azure App Service for production use with OAuth 2.0 authentication.

### 1. 📋 Prerequisites

- ✅ Azure subscription
- ✅ Azure App registration (completed above)
- ✅ Code pushed to GitHub repository

### 2. 🏗️ Create Azure App Service

#### Azure Portal
1. Go to [Azure Portal](https://portal.azure.com)
2. **Create a resource** → **Web App**
3. Configure:
   - **Subscription**: Your subscription
   - **Resource Group**: Create new or use existing (e.g., `edumcp-resource-group`)
   - **Name**: `your-app-name` (choose a name that makes sense for your organization. It has to be a name that is unique in the Azure stack.)
   - **Publish**: `Code`
   - **Runtime stack**: `Node 22 LTS`
   - **Operating System**: `Linux`
   - **Region**: Choose closest to your users
   - **Linux Plan**: Default
   - **Pricing Plan**: Default
4. Click **"Review + create"**, then after validation click **"Create"**. It should deploy in a minute or two.


### 3. 🔐 Configure Environment Variables

In Azure Portal → **App Services** → **[Your App]** → **Settings** → **Environment Variables** → click **"+ Add"**:

Add the following four environment variables (Name → Value):

| Name | Value |
|------|-------|
| `CLIENT_ID` | Your Application (client) ID from Entra |
| `TENANT_ID` | Your Directory (tenant) ID from Entra |
| `CLIENT_SECRET` | The client secret Value you saved earlier |
| `REDIRECT_URI` | `https://<YOUR_APP_NAME>.azurewebsites.net/oauth/callback` |

> Replace `<YOUR_APP_NAME>` with the App Service name you chose in Step 2 (e.g., `edumcp-server-prod`).

Once all four are added, click **"Apply"** and then **"Confirm"** to restart the app service.

### 4. � Add Redirect URI in Entra ID (Azure AD)

Now that your App Service is created and you have the URL, you **must** go back to your app registration and add the redirect URI.

> ⚠️ **Important**: If you skipped adding the redirect URI during app registration, authentication will fail until you complete this step.

#### Steps to Add the Redirect URI:

1. Go to [Microsoft Entra Admin Center](https://entra.microsoft.com)
2. Navigate to **"Entra ID" → "App registrations"**
3. Select your registered application from the list
4. In the left sidebar, click **"Authentication"**
5. Under **"Platform configurations"**, click **"Add a platform"**
6. Select **"Web"**
7. In the **"Redirect URIs"** field, enter:
   ```
   https://<YOUR_APP_NAME>.azurewebsites.net/oauth/callback
   ```
   Replace `<YOUR_APP_NAME>` with your actual Azure App Service name (e.g., `edumcp-server-prod`).
8. Click **"Configure"** to save

### 5. �📦 Deploy Your Code

#### Option A: External Git (Recommended)

> ⚠️ **Note**: Before deploying, Azure requires SCM Basic Authentication to be enabled.  
> Go to **App Service** → **Configuration** → **General settings** → **SCM Basic Auth Publishing Credentials** → **On** → click **"Apply"**

Then proceed with the deployment:

1. **Azure Portal** → **App Services** → **[Your App]** → **Deployment** → **Deployment Center**
2. **Source**: Manual Deployment → **External Git**
3. **Repository**: `https://github.com/jesegher/EduMCPServer.git`
4. **Branch**: `main`
5. **Repository Type**: Public
6. Click **"Save"**


#### Option B: VS Code Extension
1. Install **Azure App Service** extension
2. Right-click your project → **Deploy to Web App**
3. Select your subscription and app service

### 6. 🎯 Configure MCP Clients

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

### 7. ✅ Verify Deployment

1. **Test Endpoints**:
   - Health check: `https://your-app-name.azurewebsites.net`
   - OAuth discovery: `https://your-app-name.azurewebsites.net/.well-known/oauth-authorization-server`

2. **Check Logs** in Azure Portal → App Service → **Log stream**

3. **Test Authentication** with your MCP client

### 8. 🛡️ Production Security

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

## 🖥️ Scenario: In Copilot Studio, prep for a Parent-Teacher Conference
**Prompt:** Can you give me a full performance overview of Al Frederickson in preparation of my parent-teacher conference with his parents? 
**Visualization in Copilot Studio**
<img width="920" height="604" alt="image" src="https://github.com/user-attachments/assets/e08bd525-4ecc-45b2-9739-2a620b6f9f4a" />



