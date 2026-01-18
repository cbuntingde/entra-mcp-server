/**
 * Application and service principal management tools for Microsoft Graph API
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

export interface ApplicationQueryOptions {
  top?: number;
  filter?: string;
  select?: string[];
  orderBy?: string;
}

/**
 * Application and service principal tools for querying Microsoft Graph API
 */
export class ApplicationTools {
  constructor(private graphClient: GraphClientManager) {}

  /**
   * List all applications in Entra ID
   * @param options - Query options
   * @returns Array of applications
   */
  async listApplications(options: ApplicationQueryOptions = {}): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const client = this.graphClient.getClient();
        let request = client.api('/applications');

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
   * Get an application by ID
   * @param appId - Application ID
   * @param select - Properties to select
   * @returns Application object
   */
  async getApplicationById(appId: string, select?: string[]): Promise<any> {
    return retryWithBackoff(async () => {
      try {
        const validatedAppId = validateString(appId, 'appId');
        const validatedSelect = validateStringArray(select);

        const client = this.graphClient.getClient();
        let request = client.api(`/applications/${validatedAppId}`);

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
   * List all service principals in Entra ID
   * @param options - Query options
   * @returns Array of service principals
   */
  async listServicePrincipals(options: ApplicationQueryOptions = {}): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const client = this.graphClient.getClient();
        let request = client.api('/servicePrincipals');

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
   * Get a service principal by ID
   * @param spId - Service principal ID
   * @param select - Properties to select
   * @returns Service principal object
   */
  async getServicePrincipalById(spId: string, select?: string[]): Promise<any> {
    return retryWithBackoff(async () => {
      try {
        const validatedSpId = validateString(spId, 'spId');
        const validatedSelect = validateStringArray(select);

        const client = this.graphClient.getClient();
        let request = client.api(`/servicePrincipals/${validatedSpId}`);

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
   * Get owners of an application
   * @param appId - Application ID
   * @param top - Maximum number of results
   * @returns Array of application owners
   */
  async getApplicationOwners(appId: string, top = DEFAULT_TOP_MEDIUM): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedAppId = validateString(appId, 'appId');
        const validatedTop = validateTop(top, DEFAULT_TOP_MEDIUM);

        const client = this.graphClient.getClient();
        const response = await client
          .api(`/applications/${validatedAppId}/owners`)
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Search for applications by display name or app ID
   * @param searchTerm - Search term to match
   * @param top - Maximum number of results
   * @returns Array of matching applications
   */
  async searchApplications(searchTerm: string, top = DEFAULT_TOP_SMALL): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedSearchTerm = validateString(searchTerm, 'searchTerm');
        const validatedTop = validateTop(top, DEFAULT_TOP_SMALL);
        const escapedTerm = escapeODataString(validatedSearchTerm);

        const client = this.graphClient.getClient();
        const filter = `startswith(displayName,'${escapedTerm}') or startswith(appId,'${escapedTerm}')`;

        const response = await client
          .api('/applications')
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
   * Get permissions required by an application
   * @param appId - Application ID
   * @returns Application with permission details
   */
  async getApplicationPermissions(appId: string): Promise<any> {
    return retryWithBackoff(async () => {
      try {
        const validatedAppId = validateString(appId, 'appId');

        const client = this.graphClient.getClient();
        const response = await client
          .api(`/applications/${validatedAppId}`)
          .select('id,displayName,appId,requiredResourceAccess,api,oauth2PermissionScopes')
          .get();

        return response;
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }
}
