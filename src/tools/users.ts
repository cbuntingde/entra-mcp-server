/**
 * User management tools for Microsoft Graph API
 */

import { GraphClientManager } from '../graph-client.js';
import { handleGraphError } from '../utils/error-handler.js';
import {
  escapeODataString,
  validateString,
  validateOptionalString,
  validateTop,
  validateDays,
  validateStringArray,
  validateFilter,
  formatODataDate,
  getDateOffset,
  validateResponse,
  DEFAULT_TOP_SMALL,
  DEFAULT_TOP_MEDIUM,
  DEFAULT_TOP_LARGE,
} from '../utils/validation.js';
import { retryWithBackoff } from '../utils/retry.js';

export interface UserQueryOptions {
  top?: number;
  filter?: string;
  select?: string[];
  orderBy?: string;
}

export interface SignInActivityOptions {
  userId: string;
  days?: number;
}

/**
 * User tools for querying Microsoft Graph API
 */
export class UserTools {
  constructor(private graphClient: GraphClientManager) {}

  /**
   * List users with optional filtering
   * @param options - Query options
   * @returns Array of users
   */
  async listUsers(options: UserQueryOptions = {}): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const client = this.graphClient.getClient();
        let request = client.api('/users');

        // Validate and apply parameters
        const top = validateTop(options.top);
        if (top) request = request.top(top);

        const filter = validateFilter(options.filter);
        if (filter) request = request.filter(filter);

        const select = validateStringArray(options.select);
        if (select && select.length > 0) {
          request = request.select(select.join(','));
        }

        if (options.orderBy) request = request.orderby(options.orderBy);

        const response = await request.get();
        validateResponse(response);

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get a user by ID or user principal name
   * @param userId - User ID or user principal name
   * @param select - Properties to select
   * @returns User object
   */
  async getUserById(userId: string, select?: string[]): Promise<any> {
    return retryWithBackoff(async () => {
      try {
        const validatedUserId = validateString(userId, 'userId');
        const validatedSelect = validateStringArray(select);

        const client = this.graphClient.getClient();
        let request = client.api(`/users/${validatedUserId}`);

        if (validatedSelect && validatedSelect.length > 0) {
          request = request.select(validatedSelect.join(','));
        }

        return await request.get();
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get a user by user principal name (alias for getUserById)
   * @param upn - User principal name
   * @param select - Properties to select
   * @returns User object
   */
  async getUserByPrincipalName(upn: string, select?: string[]): Promise<any> {
    return this.getUserById(upn, select);
  }

  /**
   * Search for users by display name, email, or user principal name
   * @param searchTerm - Search term to match
   * @param top - Maximum number of results
   * @returns Array of matching users
   */
  async searchUsers(searchTerm: string, top = DEFAULT_TOP_SMALL): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedSearchTerm = validateString(searchTerm, 'searchTerm');
        const validatedTop = validateTop(top, DEFAULT_TOP_SMALL);
        const escapedTerm = escapeODataString(validatedSearchTerm);

        const client = this.graphClient.getClient();
        const filter = `startswith(displayName,'${escapedTerm}') or startswith(givenName,'${escapedTerm}') or startswith(surname,'${escapedTerm}') or startswith(mail,'${escapedTerm}') or startswith(userPrincipalName,'${escapedTerm}')`;

        const response = await client
          .api('/users')
          .filter(filter)
          .top(validatedTop)
          .get();

        validateResponse(response);
        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get users who haven't signed in within the specified number of days
   * @param days - Number of days to look back for inactivity
   * @param top - Maximum number of results
   * @returns Array of inactive users
   */
  async getInactiveUsers(days = 30, top = DEFAULT_TOP_LARGE): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedDays = validateDays(days);
        const validatedTop = validateTop(top, DEFAULT_TOP_LARGE);
        const startDate = getDateOffset(validatedDays);
        const startDateStr = formatODataDate(startDate);

        const client = this.graphClient.getClient();

        // Query for users who haven't signed in since the specified date
        // Note: signInActivity requires AuditLog.Read.All permission
        const response = await client
          .api('/users')
          .filter(`signInActivity/lastSignInDateTime le ${startDateStr}`)
          .select('id,displayName,mail,userPrincipalName,signInActivity')
          .top(validatedTop)
          .get();

        validateResponse(response);
        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get users with their MFA/authentication method status
   * @param top - Maximum number of results
   * @returns Array of users with authentication information
   */
  async getUsersWithMFAStatus(top = DEFAULT_TOP_LARGE): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedTop = validateTop(top, DEFAULT_TOP_LARGE);
        const client = this.graphClient.getClient();

        // Get users with authentication methods
        // Note: Expanding authentication requires UserAuthenticationMethod.Read.All permission
        const response = await client
          .api('/users')
          .select('id,displayName,mail,userPrincipalName')
          .expand('authentication($select=id,methodType,displayName)')
          .top(validatedTop)
          .get();

        validateResponse(response);
        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get sign-in history for a specific user
   * @param userId - User ID or user principal name
   * @param days - Number of days to look back
   * @returns Array of sign-in records
   */
  async getUserSignIns(userId: string, days = 30): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedUserId = validateString(userId, 'userId');
        const validatedDays = validateDays(days);
        const startDate = getDateOffset(validatedDays);
        const startDateStr = formatODataDate(startDate);

        const client = this.graphClient.getClient();

        const response = await client
          .api('/auditLogs/signIns')
          .filter(`userId eq '${validatedUserId}' and createdDateTime ge ${startDateStr}`)
          .orderby('createdDateTime desc')
          .get();

        validateResponse(response);
        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }
}
