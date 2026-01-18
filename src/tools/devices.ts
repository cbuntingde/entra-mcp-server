/**
 * Device management tools for Microsoft Graph API
 */

import { GraphClientManager } from '../graph-client.js';
import { handleGraphError } from '../utils/error-handler.js';
import {
  escapeODataString,
  validateString,
  validateTop,
  validateStringArray,
  validateFilter,
  validateDays,
  formatODataDate,
  getDateOffset,
  DEFAULT_TOP_SMALL,
  DEFAULT_TOP_MEDIUM,
  DEFAULT_TOP_LARGE,
} from '../utils/validation.js';
import { retryWithBackoff } from '../utils/retry.js';

export interface DeviceQueryOptions {
  top?: number;
  filter?: string;
  select?: string[];
  orderBy?: string;
}

/**
 * Device tools for querying Microsoft Graph API
 */
export class DeviceTools {
  constructor(private graphClient: GraphClientManager) {}

  /**
   * List all devices in Entra ID
   * @param options - Query options
   * @returns Array of devices
   */
  async listDevices(options: DeviceQueryOptions = {}): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const client = this.graphClient.getClient();
        let request = client.api('/devices');

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
   * Get a device by ID
   * @param deviceId - Device ID
   * @param select - Properties to select
   * @returns Device object
   */
  async getDeviceById(deviceId: string, select?: string[]): Promise<any> {
    return retryWithBackoff(async () => {
      try {
        const validatedDeviceId = validateString(deviceId, 'deviceId');
        const validatedSelect = validateStringArray(select);

        const client = this.graphClient.getClient();
        let request = client.api(`/devices/${validatedDeviceId}`);

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
   * Get registered owners of a device
   * @param deviceId - Device ID
   * @param top - Maximum number of results
   * @returns Array of registered owners
   */
  async getRegisteredOwners(deviceId: string, top = DEFAULT_TOP_MEDIUM): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedDeviceId = validateString(deviceId, 'deviceId');
        const validatedTop = validateTop(top, DEFAULT_TOP_MEDIUM);

        const client = this.graphClient.getClient();
        const response = await client
          .api(`/devices/${validatedDeviceId}/registeredOwners`)
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get registered users of a device
   * @param deviceId - Device ID
   * @param top - Maximum number of results
   * @returns Array of registered users
   */
  async getRegisteredUsers(deviceId: string, top = DEFAULT_TOP_MEDIUM): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedDeviceId = validateString(deviceId, 'deviceId');
        const validatedTop = validateTop(top, DEFAULT_TOP_MEDIUM);

        const client = this.graphClient.getClient();
        const response = await client
          .api(`/devices/${validatedDeviceId}/registeredUsers`)
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Get devices owned by a user
   * @param userId - User ID or user principal name
   * @param top - Maximum number of results
   * @returns Array of user's devices
   */
  async getUserDevices(userId: string, top = DEFAULT_TOP_MEDIUM): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedUserId = validateString(userId, 'userId');
        const validatedTop = validateTop(top, DEFAULT_TOP_MEDIUM);

        const client = this.graphClient.getClient();
        const response = await client
          .api(`/users/${validatedUserId}/ownedDevices`)
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }

  /**
   * Search for devices by display name, device ID, or operating system
   * @param searchTerm - Search term to match
   * @param top - Maximum number of results
   * @returns Array of matching devices
   */
  async searchDevices(searchTerm: string, top = DEFAULT_TOP_SMALL): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedSearchTerm = validateString(searchTerm, 'searchTerm');
        const validatedTop = validateTop(top, DEFAULT_TOP_SMALL);
        const escapedTerm = escapeODataString(validatedSearchTerm);

        const client = this.graphClient.getClient();
        const filter = `startswith(displayName,'${escapedTerm}') or startswith(deviceId,'${escapedTerm}') or startswith(operatingSystem,'${escapedTerm}')`;

        const response = await client
          .api('/devices')
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
   * Get devices that haven't signed in within the specified number of days
   * @param days - Number of days of inactivity
   * @param top - Maximum number of results
   * @returns Array of inactive devices
   */
  async getInactiveDevices(days = 90, top = DEFAULT_TOP_LARGE): Promise<any[]> {
    return retryWithBackoff(async () => {
      try {
        const validatedDays = validateDays(days);
        const validatedTop = validateTop(top, DEFAULT_TOP_LARGE);
        const startDate = getDateOffset(validatedDays);
        const startDateStr = formatODataDate(startDate);

        const client = this.graphClient.getClient();

        // Query for devices inactive since the specified date
        const response = await client
          .api('/devices')
          .filter(`approximateLastSignInDateTime le ${startDateStr}`)
          .select('id,displayName,deviceId,operatingSystem,operatingSystemVersion,approximateLastSignInDateTime')
          .top(validatedTop)
          .get();

        return response.value || [];
      } catch (error: any) {
        handleGraphError(error);
      }
    });
  }
}
