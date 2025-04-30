import crypto from 'crypto';

/**
 * Registers authentication tools to the MCP server.
 * @param {McpServer} server - The MCP server instance.
 * @param {Object} auth - The authentication state object.
 * @param {ConfidentialClientApplication} msalClient - The MSAL client instance.
 * @param {Set} pendingAuthStates - A set to track pending authentication states.
 * @param {Array} graphScopes - The Microsoft Graph API scopes.
 */

export function registerAuthTools(server, auth, msalClient, pendingAuthStates, graphScopes) {
  server.tool("auth-login", {}, async () => {
    console.error("🔑 microsoft-login tool called");
    try {
      const state = crypto.randomBytes(16).toString("hex");
      pendingAuthStates.add(state);

      const url = await msalClient.getAuthCodeUrl({
        scopes: graphScopes,
        redirectUri: process.env.REDIRECT_URI,
        state,
      });

      console.error(`📤 Generated auth URL: ${url.substring(0, 50)}...`);
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: "authentication_required",
            url,
            message: "Please open this URL in your browser to authenticate"
          })
        }]
      };
    } catch (error) {
      console.error("❌ Error generating auth URL:", error);
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: "error",
            message: "Failed to generate authentication URL"
          })
        }]
      };
    }
  });

  server.tool("auth-status-get", "get the authentication status of the user.", async () => {
    console.error("🔍 get-auth-status tool called");
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          authenticated: auth.isAuthenticated,
          message: auth.isAuthenticated ?
            "User is authenticated" :
            "User is not authenticated. Please call microsoft-login first"
        })
      }]
    };
  });
}

export default registerAuthTools;