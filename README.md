# Entra ID MCP Server

A Model Context Protocol (MCP) server for querying Microsoft Entra ID (formerly Azure AD) data using natural language. This server enables AI assistants to interact with Microsoft Graph API to retrieve information about users, groups, applications, devices, and administrative reports.

## Features

### User Management
- List and search users with filtering capabilities
- Get detailed user information by ID or principal name
- Find inactive users based on sign-in activity
- Check MFA and authentication method status
- Retrieve user sign-in history

### Group Management
- List and search groups
- Get group members and owners
- Find all groups a user belongs to
- Query nested group memberships

### Application Management
- List all applications and service principals
- Search applications by name or ID
- View application permissions and required access
- Get application owners

### Device Management
- List all devices
- Search devices by name or OS
- Find inactive devices
- Query devices owned by users

### Administrative Reports
- MFA registration summary and details
- Sign-in and failed sign-in reports
- Directory audit logs
- Risky users from Identity Protection
- Role assignments and definitions
- Users with administrative roles
- License usage summary
- Conditional access policies

## Prerequisites

1. **Microsoft Entra ID Tenant**
   - You need access to a Microsoft Entra ID tenant

2. **Application Registration**
   - Register a new application in your Entra ID tenant
   - Grant the appropriate Microsoft Graph API permissions:
     - `User.Read.All`
     - `Group.Read.All`
     - `Application.Read.All`
     - `Device.Read.All`
     - `AuditLog.Read.All`
     - `RoleManagement.Read.Directory`
     - `Report.Read.All`
     - `IdentityRiskyUser.Read.All`
   - Create a client secret
   - Grant admin consent for the permissions

3. **Node.js 18+**

## Installation

```bash
# Install dependencies
npm install

# Build the project
npm run build
```

## Configuration

1. Copy the example environment file:

```bash
cp .env.example .env
```

2. Edit `.env` and add your Entra ID credentials:

```env
ENTRA_TENANT_ID=your-tenant-id-here
ENTRA_CLIENT_ID=your-client-id-here
ENTRA_CLIENT_SECRET=your-client-secret-here
```

You can find these values in the Azure Portal under:
- Tenant ID: Entra ID > Overview > Properties
- Client ID: App registrations > Your app > Overview
- Client Secret: App registrations > Your app > Certificates & secrets

## Usage

### Running the Server

```bash
npm start
```

The server runs on stdio and communicates with MCP clients.

### Example Questions

The server enables AI assistants to answer questions like:

- "Which users didn't sign in last month?"
- "Is MFA enabled for all administrators?"
- "Show me all groups with 'Sales' in the name"
- "List all applications that require directory.read.all permission"
- "Find devices that haven't signed in for 90 days"
- "Who has the Global Administrator role?"
- "Get the failed sign-in attempts for the past week"

## Available Tools

### User Tools

| Tool | Description |
|------|-------------|
| `list_users` | List users with optional OData filtering |
| `get_user` | Get specific user by ID or UPN |
| `search_users` | Search users by name, email, or UPN |
| `get_inactive_users` | Find users inactive for N days |
| `get_users_mfa_status` | Get users with MFA status |
| `get_user_sign_ins` | Get sign-in history for a user |

### Group Tools

| Tool | Description |
|------|-------------|
| `list_groups` | List groups with filtering |
| `get_group_members` | Get members of a group |
| `get_user_groups` | Get all groups for a user |
| `search_groups` | Search groups by name |

### Application Tools

| Tool | Description |
|------|-------------|
| `list_applications` | List all applications |
| `list_service_principals` | List service principals |
| `search_applications` | Search applications |
| `get_application_permissions` | Get app permissions |

### Device Tools

| Tool | Description |
|------|-------------|
| `list_devices` | List all devices |
| `search_devices` | Search devices |
| `get_inactive_devices` | Find inactive devices |
| `get_user_devices` | Get user's devices |

### Report Tools

| Tool | Description |
|------|-------------|
| `get_mfa_summary` | Get MFA registration summary |
| `get_user_registration_details` | Get user auth method details |
| `get_sign_ins_report` | Get sign-in activity report |
| `get_failed_sign_ins` | Get failed sign-in report |
| `get_audit_logs` | Get directory audit logs |
| `get_risky_users` | Get risky users from Identity Protection |
| `get_users_with_admin_roles` | Get users with admin roles |
| `get_license_usage` | Get license usage summary |
| `get_conditional_access_policies` | List conditional access policies |
| `get_role_definitions` | List role definitions |

## MCP Client Configuration

To use this server with an MCP client (like Claude Desktop), add it to your MCP configuration:

```json
{
  "mcpServers": {
    "entra": {
      "command": "node",
      "args": ["path/to/entra-mcp-server/dist/index.js"],
      "env": {
        "ENTRA_TENANT_ID": "your-tenant-id",
        "ENTRA_CLIENT_ID": "your-client-id",
        "ENTRA_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

## Development

```bash
# Watch mode for development
npm run watch

# Build only
npm run build
```

## Security Considerations

1. **Store credentials securely**: Use environment variables or secure secret management
2. **Principle of least privilege**: Only grant the permissions you need
3. **Rotate secrets regularly**: Update client secrets periodically
4. **Monitor audit logs**: Track API usage through Entra ID audit logs
5. **Use Managed Identities**: Consider using managed identities for Azure deployments

## Error Handling

The server provides detailed error messages for:
- Missing or invalid credentials
- Insufficient permissions
- Invalid query parameters
- Microsoft Graph API errors

Errors are returned with codes for easier troubleshooting:
- `GRAPH_API_ERROR`: Microsoft Graph API returned an error
- `UNKNOWN_TOOL`: Requested tool doesn't exist
- `TOOL_EXECUTION_ERROR`: Error executing a tool

## Copyright

Copyright 2025 Chris Bunting <cbuntingde@gmail.com>

All rights reserved. This project was developed by Chris Bunting.

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## Support

For issues related to:
- **Microsoft Graph API**: Check the [Microsoft Graph documentation](https://docs.microsoft.com/graph/)
- **Entra ID**: Check the [Microsoft Entra ID documentation](https://docs.microsoft.com/entra/)
- **MCP Protocol**: Check the [Model Context Protocol documentation](https://modelcontextprotocol.io/)
