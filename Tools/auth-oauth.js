/**
 * OAuth 2 authentication tools for MCP server.
 * Authentication is now handled by OAuth 2 flow, these tools provide status information.
 */

export function registerAuthTools(server, getAuth, msalClient, pendingAuthStates, graphScopes) {
  
  server.tool("auth-status-get", "get the authentication status of the user.", { readOnlyHint: true }, async () => {
    console.error("🔍 OAuth auth-status-get tool called");
    
    const auth = getAuth(); // Get current request auth
    console.error(`🔍 DEBUG - Auth object:`, JSON.stringify(auth, null, 2));
    console.error(`🔍 DEBUG - isAuthenticated: ${auth.isAuthenticated}`);
    console.error(`🔍 DEBUG - accessToken exists: ${!!auth.accessToken}`);
    console.error(`🔍 DEBUG - accessToken length: ${auth.accessToken ? auth.accessToken.length : 'null'}`);
    
    const hasValidToken = auth.isAuthenticated && auth.accessToken;
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          authenticated: hasValidToken,
          message: hasValidToken ? 
            "User is authenticated via OAuth 2 with valid Microsoft Graph token" : 
            "User is not authenticated - OAuth 2 required or missing Graph token",
          authentication_method: "OAuth 2.0 with Microsoft Entra ID",
          oauth_info: hasValidToken ? 
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