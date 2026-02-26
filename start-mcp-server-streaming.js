console.log("🟢 MCP Streaming HTTPS server entry file loaded");

import { createRequire } from 'module';
const require = createRequire(import.meta.url);

import 'dotenv/config';
import express from "express";
import crypto from "crypto";
import { promisify } from 'util';

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";

import { z } from 'zod';

const msal = require('@azure/msal-node');
const { ConfidentialClientApplication } = msal;

import registerAssignmentTools from './Tools/assignment.js';
import registerRubricTools from './Tools/rubric.js';
import registerClassTools from './Tools/class.js';
import registerUserTools from './Tools/users.js';
import registerGroupTools from './Tools/group.js';
import registerModuleTools from './Tools/module.js';
import registerAuthTools from './Tools/auth-oauth.js';

const pendingAuthStates = new Set();
let accessToken = null;
let isAuthenticated = false;

const msalClient = new ConfidentialClientApplication({
  auth: {
    clientId: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
  },
});

const graphScopes = [
  "https://graph.microsoft.com/EduRoster.ReadWrite",
  "https://graph.microsoft.com/EduAssignments.ReadWrite",
  "https://graph.microsoft.com/User.ReadWrite.All",
  "https://graph.microsoft.com/Group.ReadWrite.All",
  "https://graph.microsoft.com/Directory.ReadWrite.All"
];

// OAuth 2 Server setup
const oauthTokens = new Map(); // Store issued OAuth tokens
const oauthCodes = new Map();  // Store authorization codes
const oauthRefreshTokens = new Map(); // Store refresh tokens

// Pre-registered OAuth clients - VSCode, MCP Inspector, CPS, Claude, and Microsoft AI Foundry
const oauthClients = {
  'vscode-mcp-client': {
    clientId: 'vscode-mcp-client',
    clientSecret: null, // Public client
    redirectUris: [
      'http://127.0.0.1:33418',
      'http://127.0.0.1:33418/',
      'https://vscode.dev/redirect',
      'http://localhost:3000/callback',
      'http://127.0.0.1:3000/callback'
    ],
    name: 'VS Code MCP Client'
  },
  'mcp-client': {
    clientId: 'mcp-client',
    clientSecret: null,
    redirectUris: [
      'https://global.consent.azure-apim.net/redirect/*' // Pattern for any CPS tool
    ],
    name: 'MCP Client (Copilot Studio)'
  },
  'mcp-inspector-client': {
    clientId: 'mcp-inspector-client',
    clientSecret: null,
    redirectUris: [
      'http://localhost:6274/oauth/callback',
      'http://127.0.0.1:6274/oauth/callback',
      'http://localhost:8080/callback',
      'http://127.0.0.1:8080/callback',
      'http://localhost:5173/callback',
      'http://127.0.0.1:5173/callback'
    ],
    name: 'MCP Inspector Client'
  },
  'claude-ai': {
    clientId: 'claude-ai',
    clientSecret: null,
    redirectUris: [
      'https://claude.ai/oauth/callback',
      'https://claude.ai/api/oauth/callback'
    ],
    name: 'Claude (Anthropic)'
  },
  'microsoft-foundry': {
    clientId: 'microsoft-foundry',
    clientSecret: null,
    redirectUris: [
      'https://ai.azure.com/oauth/callback' // Pattern - wildcard validated below
    ],
    name: 'Microsoft AI Foundry'
  }
};

// Validate client and redirect URI
const validateClient = (clientId, redirectUri) => {
  const client = oauthClients[clientId];
  if (!client) return false;
  
  if (redirectUri) {
    // Special handling for CPS - accept any Azure APIM redirect URI
    if (clientId === 'mcp-client' && redirectUri.startsWith('https://global.consent.azure-apim.net/redirect/')) {
      return client;
    }

    // Special handling for Claude - accept any claude.ai callback URI
    if (clientId === 'claude-ai' && redirectUri.startsWith('https://claude.ai/')) {
      return client;
    }

    // Special handling for Microsoft AI Foundry - accept any ai.azure.com callback URI
    if (clientId === 'microsoft-foundry' && /^https:\/\/.*\.?ai\.azure\.com\//.test(redirectUri)) {
      return client;
    }
    
    // For other clients, require exact match
    if (!client.redirectUris.includes(redirectUri)) {
      return false;
    }
  }
  
  return client;
};

// Generate secure tokens
const generateToken = () => crypto.randomBytes(32).toString('base64url');
const generateCode = () => crypto.randomBytes(16).toString('base64url');

// OAuth 2 Token validation middleware
const validateOAuthToken = async (req, res, next) => {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return res.status(401).json({ error: 'invalid_token', error_description: 'Missing or invalid authorization header' });
  }

  const token = authHeader.substring(7);
  const tokenData = oauthTokens.get(token);

  if (!tokenData || tokenData.expiresAt < Date.now()) {
    return res.status(401).json({ error: 'invalid_token', error_description: 'Token expired or invalid' });
  }

  // Attach user context to request
  req.oauth = {
    token: token,
    userId: tokenData.userId,
    scopes: tokenData.scopes,
    msGraphToken: tokenData.msGraphToken
  };

  next();
};

const server = new McpServer({
  name: "mcp-edu-streaming",
  description: "A Microsoft Edu streaming HTTPS MCP server.",
  version: "1.0.0",
  tools: [],
});

console.error("📝 Registering tools...");

// Simple global auth state that persists
global.mcpAuth = { 
  accessToken: null, 
  isAuthenticated: false,
  requestId: null,
  userId: null
};

// Simple getAuth function
const getAuth = () => global.mcpAuth;

registerAuthTools(server, getAuth, msalClient, pendingAuthStates, graphScopes);
registerAssignmentTools(server, getAuth);
registerRubricTools(server, getAuth);
registerClassTools(server, getAuth);
registerUserTools(server, getAuth);
registerGroupTools(server, getAuth);
registerModuleTools(server, getAuth);

const app = express();

// Add CORS headers for cross-origin requests
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, x-mcp-session-id, mcp-session-id, Cache-Control');
  res.header('Access-Control-Expose-Headers', 'x-mcp-session-id, mcp-session-id');
  
  if (req.method === 'OPTIONS') {
    res.sendStatus(200);
  } else {
    next();
  }
});

// Parse JSON request bodies
app.use(express.json());

// Parse URL-encoded request bodies (for OAuth token requests)
app.use(express.urlencoded({ extended: true }));

// OAuth 2 Discovery endpoint
app.get('/.well-known/oauth-authorization-server', (req, res) => {
  const baseUrl = `https://${req.get('host')}`;
  console.log('🔍 OAuth discovery requested from:', req.get('user-agent'));
  
  const discovery = {
    issuer: baseUrl,
    authorization_endpoint: `${baseUrl}/oauth/authorize`,
    token_endpoint: `${baseUrl}/oauth/token`,
    response_types_supported: ["code"],
    grant_types_supported: ["authorization_code", "refresh_token"],
    scopes_supported: ["mcp:read", "mcp:write"],
    code_challenge_methods_supported: ["S256"],
    token_endpoint_auth_methods_supported: ["none"],
    registration_endpoint: `${baseUrl}/oauth/register`,
    supported_clients: Object.keys(oauthClients),
    // Force VS Code to use authorization code flow
    response_modes_supported: ["query"],
    ui_locales_supported: ["en-US"],
    service_documentation: `${baseUrl}/oauth/docs`,
    // Indicate this server requires user authorization
    introspection_endpoint_auth_methods_supported: ["none"],
    revocation_endpoint_auth_methods_supported: ["none"]
  };
  
  console.log('📋 Returning OAuth discovery:', discovery);
  res.json(discovery);
});

// VS Code OAuth resource metadata endpoint
app.get('/.well-known/oauth-protected-resource', (req, res) => {
  const baseUrl = `https://${req.get('host')}`;
  console.log('🔍 OAuth resource metadata requested');
  
  res.json({
    resource: `${baseUrl}/mcp`,
    authorization_servers: [`${baseUrl}`],
    scopes_supported: ["mcp:read", "mcp:write"]
  });
});

// VS Code MCP-specific OAuth resource metadata
app.get('/.well-known/oauth-protected-resource/mcp', (req, res) => {
  const baseUrl = `https://${req.get('host')}`;
  console.log('🔍 MCP OAuth resource metadata requested');
  
  res.json({
    resource: `${baseUrl}/mcp`,
    authorization_servers: [baseUrl],
    scopes_required: ["mcp:read", "mcp:write"]
  });
});

// OAuth 2 Client Registration endpoint (for VS Code dynamic registration)
app.post('/oauth/register', (req, res) => {
  try {
    const { redirect_uris, client_name } = req.body;
    
    // For VS Code, auto-approve with known client ID
    if (redirect_uris && redirect_uris.some(uri => 
      uri.includes('127.0.0.1:33418') || uri.includes('vscode.dev'))) {
      
      return res.json({
        client_id: 'vscode-mcp-client',
        client_name: client_name || 'VS Code MCP Client',
        redirect_uris: oauthClients['vscode-mcp-client'].redirectUris,
        grant_types: ['authorization_code'],
        response_types: ['code'],
        token_endpoint_auth_method: 'none'
      });
    }
    
    // For MCP Inspector, check for common dev server ports
    if (redirect_uris && redirect_uris.some(uri => 
      uri.includes('localhost:6274') || uri.includes('127.0.0.1:6274') ||
      uri.includes('localhost:8080') || uri.includes('localhost:5173') ||
      uri.includes('127.0.0.1:8080') || uri.includes('127.0.0.1:5173'))) {
      
      return res.json({
        client_id: 'mcp-inspector-client',
        client_name: client_name || 'MCP Inspector Client',
        redirect_uris: oauthClients['mcp-inspector-client'].redirectUris,
        grant_types: ['authorization_code'],
        response_types: ['code'],
        token_endpoint_auth_method: 'none'
      });
    }

    // For Claude (Anthropic)
    if (redirect_uris && redirect_uris.some(uri => uri.includes('claude.ai'))) {
      return res.json({
        client_id: 'claude-ai',
        client_name: client_name || 'Claude (Anthropic)',
        redirect_uris: oauthClients['claude-ai'].redirectUris,
        grant_types: ['authorization_code'],
        response_types: ['code'],
        token_endpoint_auth_method: 'none'
      });
    }

    // For Microsoft AI Foundry
    if (redirect_uris && redirect_uris.some(uri => /.*\.?ai\.azure\.com/.test(uri))) {
      return res.json({
        client_id: 'microsoft-foundry',
        client_name: client_name || 'Microsoft AI Foundry',
        redirect_uris: redirect_uris,
        grant_types: ['authorization_code'],
        response_types: ['code'],
        token_endpoint_auth_method: 'none'
      });
    }
    
    // For other clients, return generic client ID
    res.json({
      client_id: 'mcp-client',
      client_name: client_name || 'MCP Client',
      redirect_uris: redirect_uris,
      grant_types: ['authorization_code'],
      response_types: ['code'],
      token_endpoint_auth_method: 'none'
    });
    
  } catch (error) {
    console.error('❌ Client registration error:', error);
    res.status(400).json({
      error: 'invalid_request',
      error_description: 'Client registration failed'
    });
  }
});

// OAuth 2 Authorization endpoint
app.get('/oauth/authorize', async (req, res) => {
  try {
    console.log('🔐 OAuth authorization request:', req.query);
    
    const {
      client_id,
      response_type,
      redirect_uri,
      scope,
      state,
      code_challenge,
      code_challenge_method
    } = req.query;

    // Basic validation
    if (!client_id || response_type !== 'code' || !redirect_uri) {
      return res.status(400).json({
        error: 'invalid_request',
        error_description: 'Missing required parameters'
      });
    }

    // Validate client and redirect URI
    console.log('🔍 Validating client_id:', client_id);
    console.log('🔍 Validating redirect_uri:', redirect_uri);
    console.log('🔍 Available clients:', Object.keys(oauthClients));
    console.log('🔍 Expected redirect URIs:', oauthClients[client_id]?.redirectUris);
    
    const client = validateClient(client_id, redirect_uri);
    if (!client) {
      console.log('❌ Client validation failed');
      return res.status(400).json({
        error: 'invalid_client',
        error_description: 'Invalid client ID or redirect URI'
      });
    }
    
    console.log('✅ Client validation passed');

    // Generate authorization code and store challenge
    const authCode = generateCode();
    const codeData = {
      clientId: client_id,
      redirectUri: redirect_uri,
      scope: scope || 'mcp:read mcp:write',
      state: state,
      codeChallenge: code_challenge,
      codeChallengeMethod: code_challenge_method,
      createdAt: Date.now(),
      expiresAt: Date.now() + (10 * 60 * 1000) // 10 minutes
    };
    
    oauthCodes.set(authCode, codeData);

    // Redirect to Microsoft for authentication with Windows account picker
    const redirectUri = `https://${req.get('host')}/oauth/callback`;
    console.log('🔗 Detected host:', req.get('host'));
    console.log('🔗 Building redirect URI:', redirectUri);
    
    const msalAuthUrl = await msalClient.getAuthCodeUrl({
      scopes: graphScopes,
      redirectUri: redirectUri,
      state: authCode, // Use our auth code as state
      prompt: 'select_account', // Forces Windows account picker
      responseMode: 'query'
    });

    console.log(`🔗 Redirecting to Microsoft auth: ${msalAuthUrl}`);
    res.redirect(msalAuthUrl);

  } catch (error) {
    console.error('❌ OAuth authorization error:', error);
    res.status(500).json({
      error: 'server_error',
      error_description: 'Authorization server error'
    });
  }
});

// OAuth 2 Microsoft callback handler
app.get('/oauth/callback', async (req, res) => {
  try {
    console.log('📥 OAuth callback from Microsoft:', req.query);
    
    const { code, state: authCode, error } = req.query;

    if (error) {
      console.error('❌ Microsoft auth error:', error);
      return res.status(400).json({
        error: 'access_denied',
        error_description: 'User denied authorization'
      });
    }

    const codeData = oauthCodes.get(authCode);
    if (!codeData || codeData.expiresAt < Date.now()) {
      console.error('❌ Invalid or expired authorization code');
      return res.status(400).json({
        error: 'invalid_grant',
        error_description: 'Authorization code expired or invalid'
      });
    }

    // Exchange Microsoft code for Graph token
    const callbackRedirectUri = `https://${req.get('host')}/oauth/callback`;
    console.log('🔃 Callback host:', req.get('host'));
    console.log('🔃 Callback redirect URI:', callbackRedirectUri);
    console.log('🔃 Microsoft callback - code:', code);
    console.log('🔃 Microsoft callback - state (authCode):', authCode);
    console.log('🔃 Microsoft callback - code:', code);
    console.log('🔃 Microsoft callback - state (authCode):', authCode);
    
    const tokenResponse = await msalClient.acquireTokenByCode({
      code: code,
      scopes: graphScopes,
      redirectUri: callbackRedirectUri
    });

    console.log('✅ Microsoft token acquired successfully');

    // Build redirect URL with our authorization code
    const redirectUrl = new URL(codeData.redirectUri);
    redirectUrl.searchParams.set('code', authCode);
    if (codeData.state) {
      redirectUrl.searchParams.set('state', codeData.state);
    }

    // Store Microsoft Graph token with the authorization code
    codeData.msGraphToken = tokenResponse.accessToken;
    codeData.userId = tokenResponse.account?.homeAccountId || 'unknown';
    oauthCodes.set(authCode, codeData);

    console.log(`🏁 Final redirect back to VS Code: ${redirectUrl.toString()}`);
    console.log(`🏁 Redirect URI: ${codeData.redirectUri}`);
    console.log(`🏁 Authorization code: ${authCode}`);
    res.redirect(redirectUrl.toString());

  } catch (error) {
    console.error('❌ OAuth callback error:', error);
    res.status(500).json({
      error: 'server_error', 
      error_description: 'Callback processing failed'
    });
  }
});

// OAuth 2 Token endpoint
app.post('/oauth/token', async (req, res) => {
  try {
    console.log('🎫 ===== TOKEN REQUEST RECEIVED =====');
    console.log('🎫 Token request from:', req.get('user-agent'));
    console.log('🎫 Token request body:', JSON.stringify(req.body, null, 2));
    console.log('🎫 Token request headers:', JSON.stringify(req.headers, null, 2));
    console.log('🎫 =====================================');
    
    const {
      grant_type,
      code,
      redirect_uri,
      client_id,
      client_secret,
      code_verifier,
      refresh_token
    } = req.body;

    // Handle refresh token grant
    if (grant_type === 'refresh_token') {
      console.log('🔄 Processing refresh token request');
      
      const refreshData = oauthRefreshTokens.get(refresh_token);
      if (!refreshData) {
        console.error('❌ Invalid refresh token');
        return res.status(400).json({
          error: 'invalid_grant',
          error_description: 'Invalid refresh token'
        });
      }

      // Check if refresh token is expired (30 days)
      if (refreshData.expiresAt < Date.now()) {
        console.error('❌ Refresh token expired');
        oauthRefreshTokens.delete(refresh_token);
        return res.status(400).json({
          error: 'invalid_grant',
          error_description: 'Refresh token expired'
        });
      }

      // Try to get a fresh MS Graph token using MSAL cache
      let msGraphToken = refreshData.msGraphToken;
      try {
        const accounts = await msalClient.getTokenCache().getAllAccounts();
        const account = accounts.find(a => a.homeAccountId === refreshData.userId);
        if (account) {
          const silentResult = await msalClient.acquireTokenSilent({
            scopes: graphScopes,
            account: account
          });
          msGraphToken = silentResult.accessToken;
          console.log('✅ Refreshed MS Graph token silently');
        }
      } catch (silentError) {
        console.log('⚠️ Could not refresh MS Graph token silently, using cached token');
      }

      // Generate new tokens
      const newAccessToken = generateToken();
      const newRefreshToken = generateToken();
      
      const newTokenData = {
        accessToken: newAccessToken,
        refreshToken: newRefreshToken,
        tokenType: 'Bearer',
        expiresAt: Date.now() + (60 * 60 * 1000), // 1 hour
        scopes: refreshData.scopes,
        userId: refreshData.userId,
        msGraphToken: msGraphToken,
        createdAt: Date.now()
      };

      // Store new tokens
      oauthTokens.set(newAccessToken, newTokenData);
      oauthRefreshTokens.set(newRefreshToken, {
        userId: refreshData.userId,
        scopes: refreshData.scopes,
        msGraphToken: msGraphToken,
        clientId: refreshData.clientId,
        expiresAt: Date.now() + (30 * 24 * 60 * 60 * 1000) // 30 days
      });

      // Invalidate old refresh token (rotation)
      oauthRefreshTokens.delete(refresh_token);

      console.log(`✅ Token refreshed successfully for user: ${refreshData.userId}`);

      return res.json({
        access_token: newAccessToken,
        refresh_token: newRefreshToken,
        token_type: 'Bearer',
        expires_in: 3600,
        scope: refreshData.scopes.join(' ')
      });
    }

    // Handle authorization code grant
    if (grant_type !== 'authorization_code') {
      return res.status(400).json({
        error: 'unsupported_grant_type',
        error_description: 'Only authorization_code and refresh_token grant types are supported'
      });
    }

    const codeData = oauthCodes.get(code);
    if (!codeData || codeData.expiresAt < Date.now()) {
      return res.status(400).json({
        error: 'invalid_grant',
        error_description: 'Authorization code expired or invalid'
      });
    }

    // Validate client and redirect URI
    const client = validateClient(client_id, redirect_uri);
    if (!client) {
      return res.status(400).json({
        error: 'invalid_client',
        error_description: 'Invalid client ID or redirect URI'
      });
    }

    // Validate request matches authorization
    if (codeData.clientId !== client_id || codeData.redirectUri !== redirect_uri) {
      return res.status(400).json({
        error: 'invalid_grant',
        error_description: 'Client ID or redirect URI mismatch'
      });
    }

    // Generate access token
    const accessToken = generateToken();
    const refreshToken = generateToken();
    
    const tokenData = {
      accessToken: accessToken,
      refreshToken: refreshToken,
      tokenType: 'Bearer',
      expiresAt: Date.now() + (60 * 60 * 1000), // 1 hour
      scopes: codeData.scope.split(' '),
      userId: codeData.userId,
      msGraphToken: codeData.msGraphToken,
      createdAt: Date.now()
    };

    oauthTokens.set(accessToken, tokenData);
    
    // Store refresh token for later use
    oauthRefreshTokens.set(refreshToken, {
      userId: codeData.userId,
      scopes: codeData.scope.split(' '),
      msGraphToken: codeData.msGraphToken,
      clientId: client_id,
      expiresAt: Date.now() + (30 * 24 * 60 * 60 * 1000) // 30 days
    });

    // Clean up authorization code
    oauthCodes.delete(code);

    console.log(`✅ OAuth token issued successfully for user: ${codeData.userId}`);

    res.json({
      access_token: accessToken,
      refresh_token: refreshToken,
      token_type: 'Bearer',
      expires_in: 3600,
      scope: codeData.scope
    });

  } catch (error) {
    console.error('❌ OAuth token error:', error);
    res.status(500).json({
      error: 'server_error',
      error_description: 'Token generation failed'
    });
  }
});

// Request logging middleware
app.use((req, res, next) => {
  console.log(`📥 ${req.method} ${req.path} from ${req.ip}`);
  next();
});

// Transport options with session management - stateless mode for MCP Inspector
const transportOptions = {
  sessionIdGenerator: undefined, // Disable session management for stateless mode
  onsessioninitialized: (sessionId) => {
    console.log(`🆔 New session initialized: ${sessionId}`);
  },
  enableJsonResponse: true, // Enable JSON responses for MCP Inspector compatibility
};

// Main MCP endpoint with OAuth 2 protection
app.all("/mcp", validateOAuthToken, async (req, res) => {
  const requestId = `req_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  try {
    console.log(`📨 ${req.method} authenticated MCP request from user: ${req.oauth.userId} [${requestId}]`);
    console.log(`📋 Headers:`, req.headers);
    if (req.body) {
      console.log(`📦 Body:`, JSON.stringify(req.body, null, 2));
    }
    
    // Set auth context for this request BEFORE tool execution
    global.mcpAuth.accessToken = req.oauth.msGraphToken;
    global.mcpAuth.isAuthenticated = true; 
    global.mcpAuth.requestId = requestId;
    global.mcpAuth.userId = req.oauth.userId;
    
    console.log(`🔐 Auth context set [${requestId}] for user ${req.oauth.userId}`);
    console.log(`🔐 Token available: ${!!global.mcpAuth.accessToken}`);
    console.log(`🔐 Token length: ${global.mcpAuth.accessToken ? global.mcpAuth.accessToken.length : 'null'}`);
    
    // Create a fresh transport for each request (stateless mode)
    console.log(`🔄 Creating stateless transport for authenticated request [${requestId}]`);
    const transport = new StreamableHTTPServerTransport(transportOptions);
    
    // Connect the server to the transport
    console.log(`🔗 Connecting server to transport...`);
    await server.connect(transport);
    console.log(`✅ Server connected to transport`);
    
    // Handle the request with the transport
    console.log(`⚡ Handling authenticated request with transport...`);
    await transport.handleRequest(req, res, req.body);
    console.log(`✅ Authenticated request handled successfully`);
    
  } catch (error) {
    console.error("❌ Error handling authenticated MCP request:", error);
    console.error("❌ Error stack:", error.stack);
    if (!res.headersSent) {
      res.status(500).json({ error: "Internal server error", details: error.message });
    }
  } finally {
    // Don't clean up auth context - keep it for subsequent requests
    console.log(`✅ Request completed [${requestId}] - auth context preserved`);
  }
});

// Health check endpoint
app.get("/", (req, res) => {
  const baseUrl = `https://${req.get('host')}`;
  
  res.json({
    name: "Education MCP Server with OAuth 2",
    version: "1.0.0", 
    protocol: "Streamable HTTPS with OAuth 2 Authentication",
    status: "running",
    oauth2: {
      authorization_endpoint: `${baseUrl}/oauth/authorize`,
      token_endpoint: `${baseUrl}/oauth/token`,
      discovery_endpoint: `${baseUrl}/.well-known/oauth-authorization-server`
    },
    endpoints: {
      mcp: `${baseUrl}/mcp`,
      oauth_authorize: `${baseUrl}/oauth/authorize`,
      oauth_token: `${baseUrl}/oauth/token`,
      oauth_discovery: `${baseUrl}/.well-known/oauth-authorization-server`
    },
    scopes_supported: ["mcp:read", "mcp:write"],
    authentication: "Microsoft Entra ID with Windows account picker"
  });
});

const PORT = process.env.PORT || process.env.WEBSITES_PORT || 3001;
const BASE_URL = process.env.WEBSITE_HOSTNAME 
  ? `https://${process.env.WEBSITE_HOSTNAME}` 
  : `http://localhost:${PORT}`;

app.listen(PORT, () => {
  console.log(`✅ MCP OAuth 2 Server is running on port ${PORT}`);
  console.log(`🔗 MCP endpoint: ${BASE_URL}/mcp`);
  console.log(`🔐 OAuth Authorization: ${BASE_URL}/oauth/authorize`);
  console.log(`🎫 OAuth Token: ${BASE_URL}/oauth/token`);
  console.log(`📋 OAuth Discovery: ${BASE_URL}/.well-known/oauth-authorization-server`);
  console.log(`🪟 Authentication: Microsoft Entra ID with Windows account picker`);
  console.log(`📋 Protocol: Streamable HTTPS with OAuth 2 protection`);
});