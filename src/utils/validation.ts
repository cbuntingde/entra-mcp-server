/**
 * Validation utilities for MCP server parameters
 */

import { MCPError, ErrorCode } from './error-handler.js';

// Constants
export const DEFAULT_TOP_SMALL = 10;
export const DEFAULT_TOP_MEDIUM = 50;
export const DEFAULT_TOP_LARGE = 100;
export const MAX_TOP = 999;
export const DEFAULT_DAYS = 30;
export const MAX_DAYS = 365;

/**
 * Escape single quotes in OData filter strings to prevent injection
 * @param input - The user input string to escape
 * @returns The escaped string safe for use in OData filters
 */
export function escapeODataString(input: string): string {
  if (typeof input !== 'string') {
    throw new MCPError('Input must be a string', ErrorCode.INVALID_PARAMETER);
  }
  // Replace single quotes with double single quotes (OData escaping)
  return input.replace(/'/g, "''");
}

/**
 * Validate and sanitize a string parameter
 * @param value - The value to validate
 * @param fieldName - The name of the field (for error messages)
 * @param allowEmpty - Whether empty strings are allowed (default: false)
 * @returns The validated string
 */
export function validateString(value: unknown, fieldName: string, allowEmpty = false): string {
  if (value === undefined || value === null) {
    throw new MCPError(`${fieldName} is required`, ErrorCode.INVALID_PARAMETER);
  }
  if (typeof value !== 'string') {
    throw new MCPError(`${fieldName} must be a string`, ErrorCode.INVALID_PARAMETER);
  }
  if (!allowEmpty && value.trim() === '') {
    throw new MCPError(`${fieldName} cannot be empty`, ErrorCode.INVALID_PARAMETER);
  }
  return value;
}

/**
 * Validate and sanitize an optional string parameter
 * @param value - The value to validate
 * @param fieldName - The name of the field (for error messages)
 * @returns The validated string or undefined
 */
export function validateOptionalString(value: unknown, fieldName: string): string | undefined {
  if (value === undefined || value === null) {
    return undefined;
  }
  return validateString(value, fieldName);
}

/**
 * Validate a numeric parameter (top/page size)
 * @param value - The value to validate
 * @param defaultValue - The default value if undefined
 * @param maxValue - The maximum allowed value (default: MAX_TOP)
 * @returns The validated number
 */
export function validateTop(value: unknown, defaultValue = DEFAULT_TOP_LARGE, maxValue = MAX_TOP): number {
  const fieldName = 'top';
  if (value === undefined || value === null) {
    return defaultValue;
  }
  if (typeof value !== 'number') {
    throw new MCPError(`${fieldName} must be a number`, ErrorCode.INVALID_PARAMETER);
  }
  if (!Number.isInteger(value)) {
    throw new MCPError(`${fieldName} must be an integer`, ErrorCode.INVALID_PARAMETER);
  }
  if (value < 1) {
    throw new MCPError(`${fieldName} must be at least 1`, ErrorCode.INVALID_PARAMETER);
  }
  if (value > maxValue) {
    throw new MCPError(`${fieldName} cannot exceed ${maxValue}`, ErrorCode.INVALID_PARAMETER);
  }
  return value;
}

/**
 * Validate a days parameter (for date calculations)
 * @param value - The value to validate
 * @param defaultValue - The default value if undefined
 * @returns The validated number
 */
export function validateDays(value: unknown, defaultValue = DEFAULT_DAYS): number {
  if (value === undefined || value === null) {
    return defaultValue;
  }
  if (typeof value !== 'number') {
    throw new MCPError('days must be a number', ErrorCode.INVALID_PARAMETER);
  }
  if (!Number.isInteger(value)) {
    throw new MCPError('days must be an integer', ErrorCode.INVALID_PARAMETER);
  }
  if (value < 1) {
    throw new MCPError('days must be at least 1', ErrorCode.INVALID_PARAMETER);
  }
  if (value > MAX_DAYS) {
    throw new MCPError(`days cannot exceed ${MAX_DAYS}`, ErrorCode.INVALID_PARAMETER);
  }
  return value;
}

/**
 * Validate a string array parameter (for $select)
 * @param value - The value to validate
 * @param fieldName - The name of the field (for error messages)
 * @returns The validated string array or undefined
 */
export function validateStringArray(value: unknown, fieldName = 'select'): string[] | undefined {
  if (value === undefined || value === null) {
    return undefined;
  }
  if (!Array.isArray(value)) {
    throw new MCPError(`${fieldName} must be an array`, ErrorCode.INVALID_PARAMETER);
  }
  if (!value.every((item) => typeof item === 'string')) {
    throw new MCPError(`All items in ${fieldName} must be strings`, ErrorCode.INVALID_PARAMETER);
  }
  if (value.length === 0) {
    return undefined;
  }
  return value as string[];
}

/**
 * Validate a filter parameter (OData filter expression)
 * @param value - The value to validate
 * @returns The validated filter string or undefined
 */
export function validateFilter(value: unknown): string | undefined {
  if (value === undefined || value === null) {
    return undefined;
  }
  if (typeof value !== 'string') {
    throw new MCPError('filter must be a string', ErrorCode.INVALID_PARAMETER);
  }
  if (value.trim() === '') {
    return undefined;
  }
  // Basic validation - check for potentially dangerous OData system queries
  const dangerous = ['$count', '$search', '$format', '$compute', '$apply'];
  const lowerValue = value.toLowerCase();
  for (const d of dangerous) {
    if (lowerValue.includes(d.toLowerCase())) {
      throw new MCPError(`filter cannot contain ${d}`, ErrorCode.INVALID_PARAMETER);
    }
  }
  return value;
}

/**
 * Format a date for OData filter queries
 * @param date - The date to format
 * @returns The ISO date string with proper quoting for OData
 */
export function formatODataDate(date: Date): string {
  return JSON.stringify(date.toISOString());
}

/**
 * Create a date offset from now
 * @param days - Number of days to subtract (positive = past, negative = future)
 * @returns The calculated date
 */
export function getDateOffset(days: number): Date {
  return new Date(Date.now() - days * 24 * 60 * 60 * 1000);
}

/**
 * Validate API response structure
 * @param response - The response object to validate
 * @param expectArray - Whether to expect an array in response.value
 * @returns The validated response
 */
export function validateResponse(response: unknown, expectArray = true): any {
  if (!response || typeof response !== 'object') {
    throw new MCPError('Invalid response from Graph API', ErrorCode.INVALID_RESPONSE);
  }
  if (expectArray) {
    if (!('value' in response) || !Array.isArray((response as any).value)) {
      throw new MCPError('Response does not contain expected value array', ErrorCode.INVALID_RESPONSE);
    }
  }
  return response;
}

/**
 * Safe type assertion for parameters from MCP protocol
 * @param value - The value to check
 * @param type - The expected type
 * @param fieldName - The field name for error messages
 * @returns The value if it matches the expected type
 */
export function assertType<T>(value: unknown, type: 'string' | 'number' | 'boolean' | 'object' | 'array', fieldName: string): T {
  if (value === undefined || value === null) {
    throw new MCPError(`${fieldName} is required`, ErrorCode.INVALID_PARAMETER);
  }

  let isValid = false;
  switch (type) {
    case 'string':
      isValid = typeof value === 'string';
      break;
    case 'number':
      isValid = typeof value === 'number' && !isNaN(value);
      break;
    case 'boolean':
      isValid = typeof value === 'boolean';
      break;
    case 'object':
      isValid = typeof value === 'object' && !Array.isArray(value);
      break;
    case 'array':
      isValid = Array.isArray(value);
      break;
  }

  if (!isValid) {
    throw new MCPError(`${fieldName} must be a ${type}`, ErrorCode.INVALID_PARAMETER);
  }

  return value as T;
}

/**
 * Extract and validate parameters from MCP request
 * @param params - The raw parameters object
 * @param schema - The parameter schema
 * @returns Validated and typed parameters
 */
export function extractParams<T extends Record<string, unknown>>(
  params: Record<string, unknown> | undefined,
  schema: {
    [K in keyof T]: {
      type: 'string' | 'number' | 'boolean' | 'array';
      required: boolean;
      default?: T[K];
    };
  }
): T {
  const result = {} as T;
  const safeParams = params || {};

  for (const [key, config] of Object.entries(schema)) {
    const value = safeParams[key];

    if (value === undefined || value === null) {
      if (config.required) {
        throw new MCPError(`Missing required parameter: ${key}`, ErrorCode.INVALID_PARAMETER);
      }
      result[key as keyof T] = config.default as T[keyof T];
      continue;
    }

    switch (config.type) {
      case 'string':
        result[key as keyof T] = assertType<string>(value, 'string', key) as T[keyof T];
        break;
      case 'number':
        result[key as keyof T] = assertType<number>(value, 'number', key) as T[keyof T];
        break;
      case 'boolean':
        result[key as keyof T] = assertType<boolean>(value, 'boolean', key) as T[keyof T];
        break;
      case 'array':
        result[key as keyof T] = assertType<unknown[]>(value, 'array', key) as T[keyof T];
        break;
    }
  }

  return result;
}
