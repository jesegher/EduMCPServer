console.log("🟢 MCP Streaming HTTPS server entry file loaded");

// Global error handlers
process.on('unhandledRejection', (reason, promise) => {
  console.error('🚨 UNHANDLED REJECTION:', reason);
  console.error('🚨 Promise:', promise);
  if (reason instanceof Error) {
    console.error('🚨 Stack:', reason.stack);
  }
});

process.on('uncaughtException', (error) => {
  console.error('🚨 UNCAUGHT EXCEPTION:', error.message);
  console.error('🚨 Stack:', error.stack);
});

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
const jwt = require('jsonwebtoken');
const jwksRsa = require('jwks-rsa');

// Log MCP SDK version and supported protocol versions at startup
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const mcpPkg = JSON.parse(readFileSync(join(__dirname, 'node_modules', '@modelcontextprotocol', 'sdk', 'package.json'), 'utf8'));
const { LATEST_PROTOCOL_VERSION, SUPPORTED_PROTOCOL_VERSIONS } = require('@modelcontextprotocol/sdk/types.js');
console.log(`📦 MCP SDK version: ${mcpPkg.version}`);
console.log(`📦 Latest protocol version: ${LATEST_PROTOCOL_VERSION}`);
console.log(`📦 Supported protocol versions: ${SUPPORTED_PROTOCOL_VERSIONS.join(', ')}`);

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

// Pre-registered OAuth clients - VSCode, MCP Inspector, CPS, Claude, Microsoft AI Foundry, and Microsoft Teams
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
  },
  'microsoft-teams': {
    clientId: 'microsoft-teams',
    clientSecret: null,
    redirectUris: [
      'https://teams.microsoft.com/api/platform/v1.0/oAuthRedirect',
      'https://teams.microsoft.com/api/platform/v1.0/oAuthConsentRedirect',
      'https://teams.microsoft.com/oauth/callback'
    ],
    name: 'Microsoft Teams'
  },
  'd0e9fbdf-8e8f-4694-816f-c01c86c8c7a1': {
    clientId: 'd0e9fbdf-8e8f-4694-816f-c01c86c8c7a1',
    clientSecret: null,
    redirectUris: [
      'https://teams.microsoft.com/api/platform/v1.0/oAuthRedirect',
      'https://teams.microsoft.com/api/platform/v1.0/oAuthConsentRedirect',
      'https://teams.microsoft.com/oauth/callback',
      'https://edumcpserver-streaming.azurewebsites.net/oauth/callback'
    ],
    name: 'Microsoft Teams (Azure App Registration)'
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

    // Special handling for Microsoft Teams - accept any teams.microsoft.com callback URI
    if (clientId === 'microsoft-teams' && redirectUri.startsWith('https://teams.microsoft.com/')) {
      return client;
    }

    // Special handling for Microsoft Teams with Azure App Registration ID
    if (clientId === 'd0e9fbdf-8e8f-4694-816f-c01c86c8c7a1' && redirectUri.startsWith('https://teams.microsoft.com/')) {
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
    msGraphToken: tokenData.msGraphToken,
    clientId: tokenData.clientId
  };

  next();
};

// Simple global auth state that persists
global.mcpAuth = { 
  accessToken: null, 
  isAuthenticated: false,
  requestId: null,
  userId: null
};

// Simple getAuth function
const getAuth = () => global.mcpAuth;

// Factory: create a fresh full server with all tools registered
function createFullServer() {
  const server = new McpServer({
    name: "mcp-edu-full",
    description: "A Microsoft Edu streaming HTTPS MCP server (full access).",
    version: "1.0.0",
    tools: [],
  });
  registerAuthTools(server, getAuth, msalClient, pendingAuthStates, graphScopes);
  registerAssignmentTools(server, getAuth);
  registerRubricTools(server, getAuth);
  registerClassTools(server, getAuth);
  registerUserTools(server, getAuth);
  registerGroupTools(server, getAuth);
  registerModuleTools(server, getAuth);
  return server;
}

// Factory: create a fresh read-only server (M365 federated connector)
function createReadOnlyServer() {
  const server = new McpServer({
    name: "mcp-edu-readonly",
    description: "A Microsoft Edu streaming HTTPS MCP server (read-only).",
    version: "1.0.0",
    tools: [],
  });
  registerAuthTools(server, getAuth, msalClient, pendingAuthStates, graphScopes);
  registerAssignmentTools(server, getAuth);
  registerRubricTools(server, getAuth);
  registerClassTools(server, getAuth);
  registerUserTools(server, getAuth);
  registerGroupTools(server, getAuth);
  registerModuleTools(server, getAuth);
  return server;
}

console.error("📝 Server factories ready.");

// Deep validation and sanitization of tools/list response payload
function validateToolsListPayload(tools, requestId) {
  console.log(`🔍 [${requestId}] ===== TOOLS/LIST VALIDATION START =====`);
  console.log(`🔍 [${requestId}] Total tools: ${tools.length}`);
  
  const issues = [];
  
  for (let i = 0; i < tools.length; i++) {
    const tool = tools[i];
    const prefix = `Tool[${i}] "${tool.name || 'UNNAMED'}"`;
    
    // Check required fields
    if (!tool.name) {
      issues.push(`${prefix}: MISSING name`);
    } else {
      // Validate name is simple lowercase (alphanumeric, hyphens, underscores)
      if (!/^[a-z][a-z0-9_-]*$/.test(tool.name)) {
        issues.push(`${prefix}: name contains invalid characters (must be lowercase alphanumeric/hyphens/underscores, start with letter)`);
      }
    }
    
    if (!tool.description) {
      issues.push(`${prefix}: MISSING description`);
    }
    
    if (!tool.inputSchema) {
      issues.push(`${prefix}: MISSING inputSchema`);
    } else {
      // Validate inputSchema structure
      if (tool.inputSchema.type !== 'object') {
        issues.push(`${prefix}: inputSchema.type is "${tool.inputSchema.type}" (expected "object")`);
      }
      if (tool.inputSchema.properties === undefined || tool.inputSchema.properties === null) {
        issues.push(`${prefix}: inputSchema.properties is ${tool.inputSchema.properties}`);
      }
      if (typeof tool.inputSchema.properties !== 'object') {
        issues.push(`${prefix}: inputSchema.properties is not an object`);
      }
      
      // Flag $schema field (may confuse strict MCP clients)
      if (tool.inputSchema.$schema) {
        issues.push(`${prefix}: inputSchema contains $schema field: "${tool.inputSchema.$schema}" (may be rejected by M365 connector)`);
      }
      
      // Check for undefined/null values in properties
      if (tool.inputSchema.properties) {
        for (const [propName, propSchema] of Object.entries(tool.inputSchema.properties)) {
          if (propSchema === undefined || propSchema === null) {
            issues.push(`${prefix}: property "${propName}" has null/undefined schema`);
          } else {
            if (!propSchema.type && !propSchema.anyOf && !propSchema.oneOf && !propSchema.allOf && !propSchema.$ref) {
              issues.push(`${prefix}: property "${propName}" has no type/anyOf/oneOf/allOf/$ref`);
            }
            // Check for unsupported types
            if (propSchema.type && !['string', 'number', 'integer', 'boolean', 'array', 'object', 'null'].includes(propSchema.type)) {
              issues.push(`${prefix}: property "${propName}" has unsupported type "${propSchema.type}"`);
            }
            // Check for recursive $ref
            if (propSchema.$ref) {
              issues.push(`${prefix}: property "${propName}" uses $ref (potential recursion): ${propSchema.$ref}`);
            }
          }
        }
      }
    }
    
    // Flag unexpected top-level fields on tool (execution, _meta, etc.)
    const expectedToolFields = new Set(['name', 'description', 'inputSchema', 'annotations']);
    for (const key of Object.keys(tool)) {
      if (!expectedToolFields.has(key)) {
        issues.push(`${prefix}: unexpected field "${key}": ${JSON.stringify(tool[key])} (may be rejected by M365 connector)`);
      }
    }
    
    // Log the full schema for each tool
    console.log(`🔍 [${requestId}] ${prefix}:`);
    console.log(`     description: ${(tool.description || '').substring(0, 100)}`);
    console.log(`     inputSchema: ${JSON.stringify(tool.inputSchema)}`);
    if (tool.execution) {
      console.log(`     ⚠️ execution: ${JSON.stringify(tool.execution)}`);
    }
  }
  
  if (issues.length > 0) {
    console.error(`🚨 [${requestId}] ===== TOOLS VALIDATION ISSUES (${issues.length}) =====`);
    for (const issue of issues) {
      console.error(`🚨 [${requestId}]   ${issue}`);
    }
  } else {
    console.log(`✅ [${requestId}] All ${tools.length} tools passed validation`);
  }
  
  console.log(`🔍 [${requestId}] ===== TOOLS/LIST VALIDATION END =====`);
  return issues;
}

// Sanitize tool schemas - remove undefined/null fields recursively
function sanitizeToolSchema(obj) {
  if (obj === null || obj === undefined) return undefined;
  if (typeof obj !== 'object') return obj;
  if (Array.isArray(obj)) return obj.map(sanitizeToolSchema).filter(v => v !== undefined);
  
  const result = {};
  for (const [key, value] of Object.entries(obj)) {
    if (value === undefined || value === null) continue;
    const sanitized = sanitizeToolSchema(value);
    if (sanitized !== undefined) {
      result[key] = sanitized;
    }
  }
  return result;
}

const app = express();

// Add CORS headers for cross-origin requests
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, x-mcp-session-id, mcp-session-id, Cache-Control');
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
    scopes_supported: ["mcp:read"],
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
    scopes_supported: ["mcp:read"]
  });
});

// VS Code MCP-specific OAuth resource metadata
app.get('/.well-known/oauth-protected-resource/mcp', (req, res) => {
  const baseUrl = `https://${req.get('host')}`;
  console.log('🔍 MCP OAuth resource metadata requested');
  
  res.json({
    resource: `${baseUrl}/mcp`,
    authorization_servers: [baseUrl],
    scopes_required: ["mcp:read"]
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

    // For Microsoft Teams
    if (redirect_uris && redirect_uris.some(uri => uri.includes('teams.microsoft.com'))) {
      return res.json({
        client_id: 'microsoft-teams',
        client_name: client_name || 'Microsoft Teams',
        redirect_uris: oauthClients['microsoft-teams'].redirectUris,
        grant_types: ['authorization_code'],
        response_types: ['code'],
        token_endpoint_auth_method: 'none'
      });
    }

    // For Microsoft Teams with specific Azure App Registration ID
    if (redirect_uris && redirect_uris.some(uri => uri.includes('teams.microsoft.com'))) {
      return res.json({
        client_id: 'd0e9fbdf-8e8f-4694-816f-c01c86c8c7a1',
        client_name: client_name || 'Microsoft Teams (Azure App Registration)',
        redirect_uris: oauthClients['d0e9fbdf-8e8f-4694-816f-c01c86c8c7a1'].redirectUris,
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

    console.log("🔍 Incoming OAuth authorize request");
    console.log("client_id:", client_id);
    console.log("redirect_uri:", redirect_uri);
    console.log("scope:", scope);
    console.log("state:", state);

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
      scope: scope || 'mcp:read',
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
        clientId: refreshData.clientId,
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
      clientId: client_id,
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

// Wrap transport.send() to log the exact JSON-RPC message before serialization
function wrapTransportSend(transport, requestId) {
  const originalSend = transport.send.bind(transport);
  transport.send = async function(message) {
    // Detect tools/list response
    if (message?.result?.tools) {
      console.log(`🔬 [${requestId}] ===== TRANSPORT.SEND: tools/list PAYLOAD =====`);
      console.log(`🔬 [${requestId}] Message ID: ${message.id}`);
      console.log(`🔬 [${requestId}] Tool count: ${message.result.tools.length}`);
      
      // Check for non-serializable values
      const serialized = JSON.stringify(message.result, (key, value) => {
        if (typeof value === 'function') {
          console.error(`🚨 [${requestId}] FUNCTION found at key "${key}"`);
          return '[FUNCTION]';
        }
        if (typeof value === 'symbol') {
          console.error(`🚨 [${requestId}] SYMBOL found at key "${key}"`);
          return '[SYMBOL]';
        }
        if (value === undefined) {
          console.error(`🚨 [${requestId}] UNDEFINED found at key "${key}"`);
          return null;
        }
        return value;
      }, 2);
      
      console.log(`🔬 [${requestId}] EXACT SERIALIZED tools/list result:\n${serialized}`);
      
      // Validate each tool
      for (let i = 0; i < message.result.tools.length; i++) {
        const tool = message.result.tools[i];
        const checks = [];
        if (!tool.name) checks.push('MISSING name');
        if (!tool.description) checks.push('MISSING description');
        if (!tool.inputSchema) checks.push('MISSING inputSchema');
        else {
          if (tool.inputSchema.type !== 'object') checks.push(`inputSchema.type="${tool.inputSchema.type}"`);
          if (!tool.inputSchema.properties || typeof tool.inputSchema.properties !== 'object') checks.push('inputSchema.properties invalid');
        }
        if (tool.execution) checks.push(`has "execution" field (SDK-injected, may cause 400)`);
        if (tool.inputSchema?.$schema) checks.push(`has "$schema" field in inputSchema`);
        
        if (checks.length > 0) {
          console.warn(`⚠️ [${requestId}] Tool[${i}] "${tool.name}": ${checks.join(', ')}`);
        } else {
          console.log(`✅ [${requestId}] Tool[${i}] "${tool.name}": VALID`);
        }
      }
      
      // Verify the final payload is valid JSON by round-tripping
      try {
        const roundTrip = JSON.parse(serialized);
        console.log(`✅ [${requestId}] tools/list payload round-trip OK (${serialized.length} bytes)`);
      } catch (e) {
        console.error(`🚨 [${requestId}] tools/list payload FAILS JSON round-trip: ${e.message}`);
      }
      
      console.log(`🔬 [${requestId}] ===== END TRANSPORT.SEND: tools/list =====`);
    } else if (message?.result?.protocolVersion) {
      console.log(`🔬 [${requestId}] TRANSPORT.SEND initialize response: ${JSON.stringify(message.result)}`);
    } else if (message?.error) {
      console.error(`🔬 [${requestId}] TRANSPORT.SEND ERROR: ${JSON.stringify(message.error, null, 2)}`);
    } else {
      console.log(`🔬 [${requestId}] TRANSPORT.SEND: ${JSON.stringify(message).substring(0, 300)}`);
    }
    
    // Sanitize tools/list payload: strip SDK-injected fields that cause M365 connector 400 errors
    if (message?.result?.tools) {
      for (const tool of message.result.tools) {
        delete tool.execution;
        delete tool._meta;
        delete tool.title;
        if (tool.inputSchema) {
          delete tool.inputSchema.$schema;
        }
      }
    }

    return originalSend(message);
  };
}

// Request logging middleware
app.use((req, res, next) => {
  console.log(`📥 ${req.method} ${req.path} from ${req.ip}`);
  next();
});

// Response interceptor to log MCP responses with deep tools/list validation
function interceptMcpResponse(res, requestId) {
  const originalJson = res.json.bind(res);
  const originalWrite = res.write.bind(res);
  const originalEnd = res.end.bind(res);

  // Helper to deeply inspect a parsed MCP payload
  function inspectPayload(payload) {
    if (payload.result?.tools) {
      // DEEP LOG: Full tools/list response
      console.log(`📤 [${requestId}] ===== TOOLS/LIST FULL RESPONSE =====`);
      console.log(`📤 [${requestId}] tools count: ${payload.result.tools.length}`);
      console.log(`📤 [${requestId}] FULL PAYLOAD:\n${JSON.stringify(payload, null, 2)}`);
      console.log(`📤 [${requestId}] ===== END TOOLS/LIST FULL RESPONSE =====`);
      
      // Run validation
      validateToolsListPayload(payload.result.tools, requestId);
    } else if (payload.result?.protocolVersion) {
      console.log(`📤 [${requestId}] initialize response:`, JSON.stringify(payload.result, null, 2));
    } else if (payload.result?.content) {
      console.log(`📤 [${requestId}] tool call response:`, JSON.stringify(payload.result, null, 2).substring(0, 500));
    } else if (payload.error) {
      console.error(`📤 [${requestId}] MCP ERROR response:`, JSON.stringify(payload.error, null, 2));
    } else {
      console.log(`📤 [${requestId}] MCP response:`, JSON.stringify(payload, null, 2).substring(0, 500));
    }
  }

  res.json = function(body) {
    console.log(`📤 [${requestId}] JSON response (status ${res.statusCode}):`, JSON.stringify(body, null, 2));
    return originalJson(body);
  };

  res.write = function(chunk, ...args) {
    const data = typeof chunk === 'string' ? chunk : chunk?.toString();
    if (data) {
      const lines = data.split('\n').filter(l => l.trim());
      for (const line of lines) {
        if (line.startsWith('data:')) {
          try {
            const payload = JSON.parse(line.substring(5).trim());
            inspectPayload(payload);
          } catch (e) {
            console.log(`📤 [${requestId}] raw SSE data:`, line.substring(0, 500));
          }
        } else if (line.startsWith('{')) {
          try {
            const payload = JSON.parse(line);
            inspectPayload(payload);
          } catch (e) {
            console.log(`📤 [${requestId}] raw data:`, line.substring(0, 500));
          }
        }
      }
    }
    return originalWrite(chunk, ...args);
  };

  res.end = function(chunk, ...args) {
    if (chunk) {
      const data = typeof chunk === 'string' ? chunk : chunk?.toString();
      if (data && data.trim()) {
        // Also check end chunk for tools/list payload
        const lines = data.split('\n').filter(l => l.trim());
        for (const line of lines) {
          if (line.startsWith('data:')) {
            try {
              const payload = JSON.parse(line.substring(5).trim());
              inspectPayload(payload);
            } catch (e) { /* ignore */ }
          } else if (line.startsWith('{')) {
            try {
              const payload = JSON.parse(line);
              inspectPayload(payload);
            } catch (e) { /* ignore */ }
          }
        }
        console.log(`📤 [${requestId}] end chunk (${data.length} bytes):`, data.substring(0, 500));
      }
    }
    console.log(`📤 [${requestId}] Response ended with status: ${res.statusCode}`);
    return originalEnd(chunk, ...args);
  };
}

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
    
    // Choose server based on OAuth client - M365 federated connector gets read-only
    const isCopilotConnector =
      req.oauth.clientId === 'd0e9fbdf-8e8f-4694-816f-c01c86c8c7a1' ||
      req.oauth.clientId === 'microsoft-teams' ||
      req.get('user-agent')?.includes('Teams_Platform');
    
    // Create a fresh server instance per request
    const activeServer = isCopilotConnector ? createReadOnlyServer() : createFullServer();
    console.log(`🔀 Using fresh ${isCopilotConnector ? 'read-only' : 'full'} server for client: ${req.oauth.clientId}`);
    
    // Intercept response to log MCP payloads
    interceptMcpResponse(res, requestId);
    
    // Create a fresh transport for each request
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
      enableJsonResponse: false
    });
    
    // Wrap transport.send to log exact outgoing JSON-RPC messages
    wrapTransportSend(transport, requestId);
    
    // Cleanup transport when response closes
    res.on("close", async () => {
      try {
        await transport.close();
      } catch (e) {
        console.warn("Transport close failed:", e.message);
      }
    });
    
    // Connect fresh server and handle request
    console.log("fresh server instance:", activeServer.constructor.name);
    console.log("request id:", requestId);
    await activeServer.connect(transport);
    await transport.handleRequest(req, res, req.body);
    console.log(`✅ transport.handleRequest() completed [${requestId}]`);
    
  } catch (error) {
    console.error(`❌ [${requestId}] Error handling authenticated MCP request:`, error);
    console.error(`❌ [${requestId}] Error stack:`, error.stack);
    if (!res.headersSent) {
      res.status(500).json({ error: "Internal server error", details: error.message });
    }
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
      m365_mcp: `${baseUrl}/m365/mcp`,
      oauth_authorize: `${baseUrl}/oauth/authorize`,
      oauth_token: `${baseUrl}/oauth/token`,
      oauth_discovery: `${baseUrl}/.well-known/oauth-authorization-server`
    },
    scopes_supported: ["mcp:read"],
    authentication: "Microsoft Entra ID with Windows account picker"
  });
});

// ========== Microsoft 365 Entra Token Validation ==========

// Accepted audiences for Entra JWT validation
const entraAudiences = [
  process.env.ENTRA_APP_ID,
  process.env.CLIENT_ID,
  `api://${process.env.CLIENT_ID}`,
  process.env.APPLICATION_ID_URI
].filter(Boolean); // Remove undefined/null entries

// JWKS client for validating Entra JWTs
const jwksClient = jwksRsa({
  jwksUri: `https://login.microsoftonline.com/${process.env.TENANT_ID}/discovery/v2.0/keys`,
  cache: true,
  cacheMaxAge: 86400000, // 24 hours
  rateLimit: true
});

function getSigningKey(header, callback) {
  jwksClient.getSigningKey(header.kid, (err, key) => {
    if (err) return callback(err);
    const signingKey = key.getPublicKey();
    callback(null, signingKey);
  });
}

// Entra token validation middleware for M365 Federated Connector
const validateEntraToken = async (req, res, next) => {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    res.set('WWW-Authenticate', `Bearer realm="${process.env.ENTRA_APP_ID || process.env.CLIENT_ID}", error="invalid_token", error_description="Missing or invalid authorization header"`);
    return res.status(401).json({ error: 'invalid_token', error_description: 'Missing or invalid authorization header' });
  }

  const token = authHeader.substring(7);

  try {
    const decoded = await new Promise((resolve, reject) => {
      jwt.verify(token, getSigningKey, {
        algorithms: ['RS256'],
        issuer: `https://login.microsoftonline.com/${process.env.TENANT_ID}/v2.0`,
        audience: entraAudiences
      }, (err, decoded) => {
        if (err) reject(err);
        else resolve(decoded);
      });
    });

    // Validate scopes - check for MCP.Access or access_as_user
    const tokenScopes = decoded.scp ? decoded.scp.split(' ') : [];
    const hasValidScope = tokenScopes.some(s => 
      s === 'MCP.Access' || s === 'access_as_user'
    );

    if (!hasValidScope) {
      console.error('❌ Entra token missing required scope. Found:', tokenScopes);
      res.set('WWW-Authenticate', `Bearer realm="${process.env.ENTRA_APP_ID || process.env.CLIENT_ID}", error="insufficient_scope", error_description="Token missing required scope (MCP.Access or access_as_user)"`);
      return res.status(403).json({ error: 'insufficient_scope', error_description: 'Token missing required scope (MCP.Access or access_as_user)' });
    }

    // Populate req.oauth
    req.oauth = {
      userId: decoded.oid || decoded.preferred_username || decoded.sub,
      scopes: tokenScopes,
      entraToken: token,
      clientId: decoded.azp || decoded.appid || 'entra-m365'
    };

    console.log(`✅ Entra token validated for user: ${req.oauth.userId}, scopes: ${tokenScopes.join(', ')}`);
    next();

  } catch (err) {
    console.error('❌ Entra token validation failed:', err.message);
    res.set('WWW-Authenticate', `Bearer realm="${process.env.ENTRA_APP_ID || process.env.CLIENT_ID}", error="invalid_token", error_description="${err.message}"`);
    return res.status(401).json({ error: 'invalid_token', error_description: err.message });
  }
};

// M365 MCP endpoint - Entra OAuth, read-only server only
app.all("/m365/mcp", validateEntraToken, async (req, res) => {
  const requestId = `m365_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  try {
    console.log(`📨 ${req.method} M365 MCP request from user: ${req.oauth.userId} [${requestId}]`);
    if (req.body) {
      console.log(`📦 [${requestId}] Body:`, JSON.stringify(req.body, null, 2));
    }
    
    // Perform On-Behalf-Of token exchange to get a Graph token
    let graphToken;
    try {
      const oboResponse = await msalClient.acquireTokenOnBehalfOf({
        oboAssertion: req.oauth.entraToken,
        scopes: graphScopes
      });
      graphToken = oboResponse.accessToken;
      console.log(`✅ OBO token exchange succeeded [${requestId}]`);
    } catch (oboError) {
      console.error(`❌ OBO token exchange failed [${requestId}]:`, oboError.message);
      console.error(`❌ OBO stack [${requestId}]:`, oboError.stack);
      return res.status(403).json({
        error: 'obo_exchange_failed',
        error_description: 'Graph delegated token exchange failed. Ensure the app registration has the required Graph API permissions and admin consent is granted.',
        details: oboError.message
      });
    }

    // Set auth context with the Graph token (NOT the incoming MCP API token)
    global.mcpAuth.accessToken = graphToken;
    global.mcpAuth.isAuthenticated = true;
    global.mcpAuth.requestId = requestId;
    global.mcpAuth.userId = req.oauth.userId;
    
    console.log(`🔐 M365 Auth context set [${requestId}] for user ${req.oauth.userId} (Graph token via OBO)`);
    
    // Intercept response to log MCP payloads
    interceptMcpResponse(res, requestId);
    
    // Create a fresh read-only server instance per request
    const activeServer = createReadOnlyServer();
    
    // Create a fresh transport for each request
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
      enableJsonResponse: false
    });
    
    // Wrap transport.send to log exact outgoing JSON-RPC messages
    wrapTransportSend(transport, requestId);
    
    // Cleanup transport when response closes
    res.on("close", async () => {
      try {
        await transport.close();
      } catch (e) {
        console.warn("Transport close failed:", e.message);
      }
    });
    
    // Connect fresh server and handle request
    console.log("fresh server instance:", activeServer.constructor.name);
    console.log("request id:", requestId);
    await activeServer.connect(transport);
    await transport.handleRequest(req, res, req.body);
    console.log(`✅ M365 transport.handleRequest() completed [${requestId}]`);
  } catch (error) {
    console.error(`❌ [${requestId}] Error handling M365 MCP request:`, error);
    console.error(`❌ [${requestId}] Stack:`, error.stack);
    if (!res.headersSent) {
      res.status(500).json({ error: "Internal server error", details: error.message });
    }
  }
});

// M365 health check endpoint
app.get("/m365/health", (req, res) => {
  const baseUrl = `https://${req.get('host')}`;
  res.json({
    name: "Education MCP Server - M365 Endpoint",
    version: "1.0.0",
    status: "running",
    endpoint: `${baseUrl}/m365/mcp`,
    auth: "Microsoft Entra ID (direct JWT validation + OBO for Graph)",
    accepted_audiences: entraAudiences,
    scope_required: "MCP.Access or access_as_user",
    graph_scopes_via_obo: graphScopes,
    server_mode: "read-only"
  });
});

const PORT = process.env.PORT || process.env.WEBSITES_PORT || 3001;
const BASE_URL = process.env.WEBSITE_HOSTNAME 
  ? `https://${process.env.WEBSITE_HOSTNAME}` 
  : `http://localhost:${PORT}`;

app.listen(PORT, () => {
  console.log(`✅ MCP OAuth 2 Server is running on port ${PORT}`);
  console.log(`🔗 MCP endpoint: ${BASE_URL}/mcp`);
  console.log(`� M365 MCP endpoint: ${BASE_URL}/m365/mcp`);
  console.log(`🔐 OAuth Authorization: ${BASE_URL}/oauth/authorize`);
  console.log(`🎫 OAuth Token: ${BASE_URL}/oauth/token`);
  console.log(`📋 OAuth Discovery: ${BASE_URL}/.well-known/oauth-authorization-server`);
  console.log(`🪟 Authentication: Microsoft Entra ID with Windows account picker`);
  console.log(`📋 Protocol: Streamable HTTPS with OAuth 2 protection`);
  console.log(`🔒 M365 Auth: Direct Entra JWT validation (read-only)`);
});