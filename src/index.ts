#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { GraphClientManager } from './graph-client.js';
import { UserTools } from './tools/users.js';
import { GroupTools } from './tools/groups.js';
import { ApplicationTools } from './tools/applications.js';
import { DeviceTools } from './tools/devices.js';
import { ReportTools } from './tools/reports.js';
import { MCPError, ErrorCode } from './utils/error-handler.js';

// Initialize Graph client
const tenantId = process.env.ENTRA_TENANT_ID;
const clientId = process.env.ENTRA_CLIENT_ID;
const clientSecret = process.env.ENTRA_CLIENT_SECRET;

if (!tenantId || !clientId || !clientSecret) {
  console.error('Missing required environment variables:');
  console.error('  - ENTRA_TENANT_ID');
  console.error('  - ENTRA_CLIENT_ID');
  console.error('  - ENTRA_CLIENT_SECRET');
  process.exit(1);
}

const graphClient = new GraphClientManager({ tenantId, clientId, clientSecret });
const userTools = new UserTools(graphClient);
const groupTools = new GroupTools(graphClient);
const applicationTools = new ApplicationTools(graphClient);
const deviceTools = new DeviceTools(graphClient);
const reportTools = new ReportTools(graphClient);

// Create MCP server
const server = new Server(
  {
    name: 'entra-mcp-server',
    version: '1.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// List available tools
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      // User Tools
      {
        name: 'list_users',
        description: 'List users in Entra ID with optional filtering. Supports OData filtering syntax.',
        inputSchema: {
          type: 'object',
          properties: {
            top: {
              type: 'number',
              description: 'Maximum number of results to return (default: 100, max: 999)',
            },
            filter: {
              type: 'string',
              description: 'OData filter expression (e.g., "accountEnabled eq true")',
            },
            select: {
              type: 'array',
              items: { type: 'string' },
              description: 'Properties to select (e.g., ["id", "displayName", "mail"])',
            },
            orderBy: {
              type: 'string',
              description: 'Property to order by (e.g., "displayName desc")',
            },
          },
        },
      },
      {
        name: 'get_user',
        description: 'Get a specific user by ID or user principal name (UPN)',
        inputSchema: {
          type: 'object',
          properties: {
            userId: {
              type: 'string',
              description: 'User ID or user principal name (UPN)',
            },
            select: {
              type: 'array',
              items: { type: 'string' },
              description: 'Properties to select (e.g., ["id", "displayName", "mail", "signInActivity"])',
            },
          },
          required: ['userId'],
        },
      },
      {
        name: 'search_users',
        description: 'Search for users by display name, email, or user principal name',
        inputSchema: {
          type: 'object',
          properties: {
            searchTerm: {
              type: 'string',
              description: 'Search term to match against displayName, mail, or userPrincipalName',
            },
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 10)',
            },
          },
          required: ['searchTerm'],
        },
      },
      {
        name: 'get_inactive_users',
        description: 'Get users who have not signed in within the specified number of days',
        inputSchema: {
          type: 'object',
          properties: {
            days: {
              type: 'number',
              description: 'Number of days to look back for inactivity (default: 30)',
            },
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 100)',
            },
          },
        },
      },
      {
        name: 'get_users_mfa_status',
        description: 'Get users with their MFA/authentication method status',
        inputSchema: {
          type: 'object',
          properties: {
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 100)',
            },
          },
        },
      },
      {
        name: 'get_user_sign_ins',
        description: 'Get sign-in history for a specific user',
        inputSchema: {
          type: 'object',
          properties: {
            userId: {
              type: 'string',
              description: 'User ID or user principal name',
            },
            days: {
              type: 'number',
              description: 'Number of days to look back (default: 30)',
            },
          },
          required: ['userId'],
        },
      },
      // Group Tools
      {
        name: 'list_groups',
        description: 'List groups in Entra ID with optional filtering',
        inputSchema: {
          type: 'object',
          properties: {
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 100)',
            },
            filter: {
              type: 'string',
              description: 'OData filter expression',
            },
            select: {
              type: 'array',
              items: { type: 'string' },
              description: 'Properties to select',
            },
            orderBy: {
              type: 'string',
              description: 'Property to order by',
            },
          },
        },
      },
      {
        name: 'get_group_members',
        description: 'Get members of a specific group',
        inputSchema: {
          type: 'object',
          properties: {
            groupId: {
              type: 'string',
              description: 'Group ID',
            },
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 50)',
            },
          },
          required: ['groupId'],
        },
      },
      {
        name: 'get_user_groups',
        description: 'Get all groups a user is a member of',
        inputSchema: {
          type: 'object',
          properties: {
            userId: {
              type: 'string',
              description: 'User ID or user principal name',
            },
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 50)',
            },
          },
          required: ['userId'],
        },
      },
      {
        name: 'search_groups',
        description: 'Search for groups by display name or email',
        inputSchema: {
          type: 'object',
          properties: {
            searchTerm: {
              type: 'string',
              description: 'Search term',
            },
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 10)',
            },
          },
          required: ['searchTerm'],
        },
      },
      // Application Tools
      {
        name: 'list_applications',
        description: 'List all applications in Entra ID',
        inputSchema: {
          type: 'object',
          properties: {
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 100)',
            },
            filter: {
              type: 'string',
              description: 'OData filter expression',
            },
            select: {
              type: 'array',
              items: { type: 'string' },
              description: 'Properties to select',
            },
          },
        },
      },
      {
        name: 'list_service_principals',
        description: 'List service principals in Entra ID',
        inputSchema: {
          type: 'object',
          properties: {
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 100)',
            },
            filter: {
              type: 'string',
              description: 'OData filter expression',
            },
          },
        },
      },
      {
        name: 'search_applications',
        description: 'Search for applications by name or app ID',
        inputSchema: {
          type: 'object',
          properties: {
            searchTerm: {
              type: 'string',
              description: 'Search term',
            },
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 10)',
            },
          },
          required: ['searchTerm'],
        },
      },
      {
        name: 'get_application_permissions',
        description: 'Get permissions required by an application',
        inputSchema: {
          type: 'object',
          properties: {
            appId: {
              type: 'string',
              description: 'Application ID',
            },
          },
          required: ['appId'],
        },
      },
      // Device Tools
      {
        name: 'list_devices',
        description: 'List all devices in Entra ID',
        inputSchema: {
          type: 'object',
          properties: {
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 100)',
            },
            filter: {
              type: 'string',
              description: 'OData filter expression',
            },
          },
        },
      },
      {
        name: 'search_devices',
        description: 'Search for devices by display name, device ID, or operating system',
        inputSchema: {
          type: 'object',
          properties: {
            searchTerm: {
              type: 'string',
              description: 'Search term',
            },
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 10)',
            },
          },
          required: ['searchTerm'],
        },
      },
      {
        name: 'get_inactive_devices',
        description: 'Get devices that have not signed in within the specified number of days',
        inputSchema: {
          type: 'object',
          properties: {
            days: {
              type: 'number',
              description: 'Days of inactivity (default: 90)',
            },
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 100)',
            },
          },
        },
      },
      {
        name: 'get_user_devices',
        description: 'Get devices owned by a specific user',
        inputSchema: {
          type: 'object',
          properties: {
            userId: {
              type: 'string',
              description: 'User ID or user principal name',
            },
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 50)',
            },
          },
          required: ['userId'],
        },
      },
      // Report Tools
      {
        name: 'get_mfa_summary',
        description: 'Get MFA registration summary for the organization',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      },
      {
        name: 'get_user_registration_details',
        description: 'Get detailed user registration information for authentication methods',
        inputSchema: {
          type: 'object',
          properties: {
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 100)',
            },
          },
        },
      },
      {
        name: 'get_sign_ins_report',
        description: 'Get sign-in activity report for the organization',
        inputSchema: {
          type: 'object',
          properties: {
            days: {
              type: 'number',
              description: 'Number of days to look back (default: 30)',
            },
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 100)',
            },
          },
        },
      },
      {
        name: 'get_failed_sign_ins',
        description: 'Get failed sign-in attempts report',
        inputSchema: {
          type: 'object',
          properties: {
            days: {
              type: 'number',
              description: 'Number of days to look back (default: 30)',
            },
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 100)',
            },
          },
        },
      },
      {
        name: 'get_audit_logs',
        description: 'Get directory audit logs',
        inputSchema: {
          type: 'object',
          properties: {
            days: {
              type: 'number',
              description: 'Number of days to look back (default: 30)',
            },
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 100)',
            },
            category: {
              type: 'string',
              description: 'Audit log category to filter by',
            },
          },
        },
      },
      {
        name: 'get_risky_users',
        description: 'Get users flagged as risky by Identity Protection',
        inputSchema: {
          type: 'object',
          properties: {
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 50)',
            },
          },
        },
      },
      {
        name: 'get_users_with_admin_roles',
        description: 'Get all users with administrative role assignments',
        inputSchema: {
          type: 'object',
          properties: {
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 100)',
            },
          },
        },
      },
      {
        name: 'get_license_usage',
        description: 'Get license usage summary for subscribed SKUs',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      },
      {
        name: 'get_conditional_access_policies',
        description: 'List all conditional access policies',
        inputSchema: {
          type: 'object',
          properties: {
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 100)',
            },
          },
        },
      },
      {
        name: 'get_role_definitions',
        description: 'List all available role definitions in Entra ID',
        inputSchema: {
          type: 'object',
          properties: {
            top: {
              type: 'number',
              description: 'Maximum number of results (default: 100)',
            },
          },
        },
      },
    ],
  };
});

// Handle tool execution
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  const params = args || {} as Record<string, unknown>;

  try {
    switch (name) {
      // User Tools
      case 'list_users':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await userTools.listUsers(params),
                null,
                2
              ),
            },
          ],
        };

      case 'get_user':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await userTools.getUserById(String(params.userId), params.select as string[] | undefined),
                null,
                2
              ),
            },
          ],
        };

      case 'search_users':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await userTools.searchUsers(String(params.searchTerm), params.top as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      case 'get_inactive_users':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await userTools.getInactiveUsers(params.days as number | undefined, params.top as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      case 'get_users_mfa_status':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await userTools.getUsersWithMFAStatus(params.top as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      case 'get_user_sign_ins':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await userTools.getUserSignIns(String(params.userId), params.days as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      // Group Tools
      case 'list_groups':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await groupTools.listGroups(params),
                null,
                2
              ),
            },
          ],
        };

      case 'get_group_members':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await groupTools.getGroupMembers(String(params.groupId), params.top as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      case 'get_user_groups':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await groupTools.getUserGroups(String(params.userId), params.top as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      case 'search_groups':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await groupTools.searchGroups(String(params.searchTerm), params.top as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      // Application Tools
      case 'list_applications':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await applicationTools.listApplications(params),
                null,
                2
              ),
            },
          ],
        };

      case 'list_service_principals':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await applicationTools.listServicePrincipals(params),
                null,
                2
              ),
            },
          ],
        };

      case 'search_applications':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await applicationTools.searchApplications(String(params.searchTerm), params.top as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      case 'get_application_permissions':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await applicationTools.getApplicationPermissions(String(params.appId)),
                null,
                2
              ),
            },
          ],
        };

      // Device Tools
      case 'list_devices':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await deviceTools.listDevices(params),
                null,
                2
              ),
            },
          ],
        };

      case 'search_devices':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await deviceTools.searchDevices(String(params.searchTerm), params.top as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      case 'get_inactive_devices':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await deviceTools.getInactiveDevices(params.days as number | undefined, params.top as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      case 'get_user_devices':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await deviceTools.getUserDevices(String(params.userId), params.top as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      // Report Tools
      case 'get_mfa_summary':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(await reportTools.getMFASummary(), null, 2),
            },
          ],
        };

      case 'get_user_registration_details':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await reportTools.getUserRegistrationDetails(params.top as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      case 'get_sign_ins_report':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await reportTools.getSignInsReport(params.days as number | undefined, params.top as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      case 'get_failed_sign_ins':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await reportTools.getFailedSignInsReport(params.days as number | undefined, params.top as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      case 'get_audit_logs':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await reportTools.getAuditLogs(params.days as number | undefined, params.top as number | undefined, params.category as string | undefined),
                null,
                2
              ),
            },
          ],
        };

      case 'get_risky_users':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(await reportTools.getRiskyUsers(params.top as number | undefined), null, 2),
            },
          ],
        };

      case 'get_users_with_admin_roles':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await reportTools.getUsersWithAdminRoles(params.top as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      case 'get_license_usage':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(await reportTools.getLicenseUsageSummary(), null, 2),
            },
          ],
        };

      case 'get_conditional_access_policies':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await reportTools.getConditionalAccessPolicies(params.top as number | undefined),
                null,
                2
              ),
            },
          ],
        };

      case 'get_role_definitions':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(await reportTools.getRoleDefinitions(params.top as number | undefined), null, 2),
            },
          ],
        };

      default:
        throw new MCPError(`Unknown tool: ${name}`, ErrorCode.INVALID_PARAMETER);
    }
  } catch (error) {
    if (error instanceof MCPError) {
      throw error;
    }
    throw new MCPError(
      `Error executing tool ${name}: ${error instanceof Error ? error.message : String(error)}`,
      ErrorCode.INTERNAL_ERROR
    );
  }
});

// Start the server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Entra ID MCP Server running on stdio');
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
