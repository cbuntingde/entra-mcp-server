/**
 * Group management tools for Microsoft Graph API
 */

import { GraphClientManager } from '../graph-client.js';
import { handleGraphError } from '../utils/error-handler.js';
import {
  escapeODataString,
  validateString,
  validateTop,
  validateStringArray,
  validateFilter,
  DEFAULT_TOP_SMALL,
  DEFAULT_TOP_MEDIUM,
  DEFAULT_TOP_LARGE,
} from '../utils/validation.js';
import { retryWithBackoff } from '../utils/retry.js';

export interface GroupQueryOptions {
  top?: number;
  filter?: string;
  select?: string[];
  orderBy?: string;
}

/**
 * Group tools for querying Microsoft Graph API
 */
export class GroupTools {
  constructor(private graphClient: GraphClientManager) {}

  /**
   * List groups with optional filtering
   * @param options - Query options
   * @returns Array of groups
   */
  async listGroups(options: GroupQueryOptions = {}): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const client = this.graphClient.getClient();
        let request = client.api('/groups');

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
        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get a group by ID
   * @param groupId - Group ID
   * @param select - Properties to select
   * @returns Group object
   */
  async getGroupById(groupId: string, select?: string[]): Promise<any> {
    return retryWithBackoff(async () => {
      try {
        const validatedGroupId = validateString(groupId, 'groupId');
        const validatedSelect = validateStringArray(select);

        const client = this.graphClient.getClient();
        let request = client.api(`/groups/${validatedGroupId}`);

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
   * Get members of a specific group
   * @param groupId - Group ID
   * @param top - Maximum number of results
   * @returns Array of group members
   */
  async getGroupMembers(groupId: string, top = DEFAULT_TOP_MEDIUM): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedGroupId = validateString(groupId, 'groupId');
        const validatedTop = validateTop(top, DEFAULT_TOP_MEDIUM);

        const client = this.graphClient.getClient();
        const response = await client
          .api(`/groups/${validatedGroupId}/members`)
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get owners of a specific group
   * @param groupId - Group ID
   * @param top - Maximum number of results
   * @returns Array of group owners
   */
  async getGroupOwners(groupId: string, top = DEFAULT_TOP_MEDIUM): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedGroupId = validateString(groupId, 'groupId');
        const validatedTop = validateTop(top, DEFAULT_TOP_MEDIUM);

        const client = this.graphClient.getClient();
        const response = await client
          .api(`/groups/${validatedGroupId}/owners`)
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get all groups a user is a member of
   * @param userId - User ID or user principal name
   * @param top - Maximum number of results
   * @returns Array of groups
   */
  async getUserGroups(userId: string, top = DEFAULT_TOP_MEDIUM): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedUserId = validateString(userId, 'userId');
        const validatedTop = validateTop(top, DEFAULT_TOP_MEDIUM);

        const client = this.graphClient.getClient();
        const response = await client
          .api(`/users/${validatedUserId}/memberOf`)
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Search for groups by display name or email
   * @param searchTerm - Search term to match
   * @param top - Maximum number of results
   * @returns Array of matching groups
   */
  async searchGroups(searchTerm: string, top = DEFAULT_TOP_SMALL): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedSearchTerm = validateString(searchTerm, 'searchTerm');
        const validatedTop = validateTop(top, DEFAULT_TOP_SMALL);
        const escapedTerm = escapeODataString(validatedSearchTerm);

        const client = this.graphClient.getClient();
        const filter = `startswith(displayName,'${escapedTerm}') or startswith(mail,'${escapedTerm}')`;

        const response = await client
          .api('/groups')
          .filter(filter)
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get transitive members of a group (includes nested group members)
   * @param groupId - Group ID
   * @param top - Maximum number of results
   * @returns Array of all transitive members
   */
  async getGroupTransitiveMembers(groupId: string, top = DEFAULT_TOP_LARGE): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedGroupId = validateString(groupId, 'groupId');
        const validatedTop = validateTop(top, DEFAULT_TOP_LARGE);

        const client = this.graphClient.getClient();
        const response = await client
          .api(`/groups/${validatedGroupId}/transitiveMembers`)
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }
}
