/**
 * OAuth 2 authentication tools for MCP server.
 * Authentication is now handled by OAuth 2 flow, these tools provide status information.
 */

export function registerAuthTools(server, auth, msalClient, pendingAuthStates, graphScopes) {
  
  server.tool("auth-status-get", "get the authentication status of the user.", async () => {
    console.error("🔍 OAuth auth-status-get tool called");
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          authenticated: auth.isAuthenticated,
          message: auth.isAuthenticated ? 
            "User is authenticated via OAuth 2" : 
            "User is not authenticated - OAuth 2 required",
          authentication_method: "OAuth 2.0 with Microsoft Entra ID",
          oauth_info: auth.isAuthenticated ? 
            "Authentication handled by OAuth 2 flow with Windows account picker" :
            "Use OAuth 2 authorization endpoint to authenticate"
        })
      }]
    };
  });
  
  // Note: auth-login is no longer needed since OAuth 2 handles authentication
  // Clients use the OAuth 2 endpoints directly
}

export default registerAuthTools;