/**
 * Administrative and security reporting tools for Microsoft Graph API
 */

import { GraphClientManager } from '../graph-client.js';
import { handleGraphError } from '../utils/error-handler.js';
import {
  validateString,
  validateOptionalString,
  validateTop,
  validateDays,
  validateFilter,
  formatODataDate,
  getDateOffset,
  DEFAULT_TOP_MEDIUM,
  DEFAULT_TOP_LARGE,
} from '../utils/validation.js';
import { retryWithBackoff } from '../utils/retry.js';

export interface ReportOptions {
  days?: number;
  top?: number;
}

/**
 * Administrative and security reporting tools for Microsoft Graph API
 */
export class ReportTools {
  constructor(private graphClient: GraphClientManager) {}

  /**
   * Get MFA registration summary for the organization
   * @returns MFA summary data
   */
  async getMFASummary(): Promise<any> {
    return retryWithBackoff(async () => {
      try {
        const client = this.graphClient.getClient();

        // Get authentication methods overview
        const response = await client.api('/reports/authenticationMethods').get();

        return response;
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get detailed user registration information for authentication methods
   * @param top - Maximum number of results
   * @returns Array of user registration details
   */
  async getUserRegistrationDetails(top = DEFAULT_TOP_LARGE): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedTop = validateTop(top, DEFAULT_TOP_LARGE);
        const client = this.graphClient.getClient();

        const response = await client
          .api('/reports/credentialUserRegistrationDetails')
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get user authentication methods summary
   * @param top - Maximum number of results
   * @returns Array of user authentication method details
   */
  async getUserAuthMethodsSummary(top = DEFAULT_TOP_LARGE): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedTop = validateTop(top, DEFAULT_TOP_LARGE);
        const client = this.graphClient.getClient();

        const response = await client
          .api('/reports/authenticationMethodsUserRegistrationDetails')
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get sign-in activity report for the organization
   * @param days - Number of days to look back
   * @param top - Maximum number of results
   * @returns Array of sign-in records
   */
  async getSignInsReport(days = 30, top = DEFAULT_TOP_LARGE): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedDays = validateDays(days);
        const validatedTop = validateTop(top, DEFAULT_TOP_LARGE);
        const startDate = getDateOffset(validatedDays);
        const startDateStr = formatODataDate(startDate);

        const client = this.graphClient.getClient();

        const response = await client
          .api('/auditLogs/signIns')
          .filter(`createdDateTime ge ${startDateStr}`)
          .orderby('createdDateTime desc')
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get failed sign-in attempts report
   * @param days - Number of days to look back
   * @param top - Maximum number of results
   * @returns Array of failed sign-in records
   */
  async getFailedSignInsReport(days = 30, top = DEFAULT_TOP_LARGE): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedDays = validateDays(days);
        const validatedTop = validateTop(top, DEFAULT_TOP_LARGE);
        const startDate = getDateOffset(validatedDays);
        const startDateStr = formatODataDate(startDate);

        const client = this.graphClient.getClient();

        // Filter for failed sign-ins (status/errorCode ne 0 indicates failure)
        const response = await client
          .api('/auditLogs/signIns')
          .filter(`createdDateTime ge ${startDateStr} and status/errorCode ne 0`)
          .orderby('createdDateTime desc')
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get directory audit logs
   * @param days - Number of days to look back
   * @param top - Maximum number of results
   * @param category - Audit log category to filter by
   * @returns Array of audit log entries
   */
  async getAuditLogs(days = 30, top = DEFAULT_TOP_LARGE, category?: string): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedDays = validateDays(days);
        const validatedTop = validateTop(top, DEFAULT_TOP_LARGE);
        const validatedCategory = validateOptionalString(category, 'category');
        const startDate = getDateOffset(validatedDays);
        const startDateStr = formatODataDate(startDate);

        const client = this.graphClient.getClient();

        let filter = `activityDateTime ge ${startDateStr}`;
        if (validatedCategory) {
          // Escape the category to prevent injection
          const escapedCategory = validatedCategory.replace(/'/g, "''");
          filter += ` and category eq '${escapedCategory}'`;
        }

        const response = await client
          .api('/auditLogs/directoryAudits')
          .filter(filter)
          .orderby('activityDateTime desc')
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get users flagged as risky by Identity Protection
   * @param top - Maximum number of results
   * @returns Array of risky users
   */
  async getRiskyUsers(top = DEFAULT_TOP_MEDIUM): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedTop = validateTop(top, DEFAULT_TOP_MEDIUM);
        const client = this.graphClient.getClient();

        const response = await client
          .api('/identityProtection/riskyUsers')
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get all role assignments in the directory
   * @param top - Maximum number of results
   * @returns Array of role assignments
   */
  async getRoleAssignments(top = DEFAULT_TOP_LARGE): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedTop = validateTop(top, DEFAULT_TOP_LARGE);
        const client = this.graphClient.getClient();

        const response = await client
          .api('/roleManagement/directory/roleAssignments')
          .expand('principal,roleDefinition')
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get all available role definitions in Entra ID
   * @param top - Maximum number of results
   * @returns Array of role definitions
   */
  async getRoleDefinitions(top = DEFAULT_TOP_LARGE): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedTop = validateTop(top, DEFAULT_TOP_LARGE);
        const client = this.graphClient.getClient();

        const response = await client
          .api('/roleManagement/directory/roleDefinitions')
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get all users with administrative role assignments
   * @param top - Maximum number of results
   * @returns Array of role assignments with user details
   */
  async getUsersWithAdminRoles(top = DEFAULT_TOP_LARGE): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedTop = validateTop(top, DEFAULT_TOP_LARGE);
        const client = this.graphClient.getClient();

        const response = await client
          .api('/roleManagement/directory/roleAssignments')
          .expand("principal(\$select=id,displayName,userPrincipalName,mail),roleDefinition(\$select=id,displayName,description)")
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get license usage summary for subscribed SKUs
   * @returns Array of license information
   */
  async getLicenseUsageSummary(): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const client = this.graphClient.getClient();

        const response = await client.api('/subscribedSkus').get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * List all conditional access policies
   * @param top - Maximum number of results
   * @returns Array of conditional access policies
   */
  async getConditionalAccessPolicies(top = DEFAULT_TOP_LARGE): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedTop = validateTop(top, DEFAULT_TOP_LARGE);
        const client = this.graphClient.getClient();

        const response = await client
          .api('/identity/conditionalAccess/policies')
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }
}
