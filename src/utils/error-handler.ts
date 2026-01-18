/**
 * Error codes for the MCP server
 */
export enum ErrorCode {
  // Server errors
  INTERNAL_ERROR = 'INTERNAL_ERROR',
  UNKNOWN_ERROR = 'UNKNOWN_ERROR',

  // Microsoft Graph API errors
  GRAPH_API_ERROR = 'GRAPH_API_ERROR',

  // Client input errors
  INVALID_PARAMETER = 'INVALID_PARAMETER',
  MISSING_PARAMETER = 'MISSING_PARAMETER',

  // Authentication/authorization errors
  UNAUTHORIZED = 'UNAUTHORIZED',
  FORBIDDEN = 'FORBIDDEN',

  // Resource errors
  NOT_FOUND = 'NOT_FOUND',
  INVALID_RESPONSE = 'INVALID_RESPONSE',

  // Rate limiting
  RATE_LIMITED = 'RATE_LIMITED',

  // Network errors
  NETWORK_ERROR = 'NETWORK_ERROR',
  TIMEOUT = 'TIMEOUT',
}

/**
 * Custom error class for MCP server errors
 */
export class MCPError extends Error {
  constructor(
    message: string,
    public code: ErrorCode = ErrorCode.INTERNAL_ERROR,
    public details?: Record<string, unknown>
  ) {
    super(message);
    this.name = 'MCPError';
    // Maintains proper stack trace for where our error was thrown (only available on V8)
    if (Error.captureStackTrace) {
      Error.captureStackTrace(this, MCPError);
    }
  }

  toJSON() {
    return {
      name: this.name,
      message: this.message,
      code: this.code,
      details: this.details,
    };
  }
}

/**
 * Map Microsoft Graph error codes to our error codes
 * @param graphCode - The Microsoft Graph error code
 * @returns The corresponding MCP error code
 */
function mapGraphErrorCode(graphCode: string): ErrorCode {
  // Authentication/authorization errors
  if (
    graphCode === 'Unauthorized' ||
    graphCode === 'AuthenticationFailed' ||
    graphCode === 'AuthenticationCanceled'
  ) {
    return ErrorCode.UNAUTHORIZED;
  }

  if (graphCode === 'Forbidden' || graphCode === 'AuthorizationFailed') {
    return ErrorCode.FORBIDDEN;
  }

  // Resource not found errors
  if (
    graphCode === 'Request_ResourceNotFound' ||
    graphCode === 'ResourceNotFound' ||
    graphCode === 'UserNotFound' ||
    graphCode === 'GroupNotFound' ||
    graphCode === 'ApplicationNotFound'
  ) {
    return ErrorCode.NOT_FOUND;
  }

  // Rate limiting errors
  if (
    graphCode === 'TooManyRequests' ||
    graphCode === 'RateLimitExceeded' ||
    graphCode === 'ThrottledRequest'
  ) {
    return ErrorCode.RATE_LIMITED;
  }

  // Invalid request errors
  if (
    graphCode === 'BadRequest' ||
    graphCode === 'InvalidRequest' ||
    graphCode === 'InvalidFilter' ||
    graphCode === 'InvalidQueryParameter'
  ) {
    return ErrorCode.INVALID_PARAMETER;
  }

  // Network/timeout errors
  if (graphCode === 'Timeout' || graphCode === 'RequestTimeout') {
    return ErrorCode.TIMEOUT;
  }

  // Default to graph API error for unmapped codes
  return ErrorCode.GRAPH_API_ERROR;
}

/**
 * Handle errors from Microsoft Graph API
 * @param error - The error object from Graph API
 * @throws MCPError with appropriate error code and message
 */
export function handleGraphError(error: any): never {
  // Check for Graph API error response
  if (error?.response?.error) {
    const graphError = error.response.error;
    const code = mapGraphErrorCode(graphError.code || '');

    // Extract additional details if available
    const details: Record<string, unknown> = {
      graphCode: graphError.code,
    };

    if (graphError.innerError) {
      details.innerError = graphError.innerError;
    }

    throw new MCPError(
      `Microsoft Graph API Error: ${graphError.message}`,
      code,
      details
    );
  }

  // Check for HTTP status code
  if (error?.response?.status) {
    const status = error.response.status;

    if (status === 401 || status === 403) {
      throw new MCPError(
        'Authentication or authorization failed',
        status === 401 ? ErrorCode.UNAUTHORIZED : ErrorCode.FORBIDDEN,
        { httpStatus: status }
      );
    }

    if (status === 404) {
      throw new MCPError(
        'Resource not found',
        ErrorCode.NOT_FOUND,
        { httpStatus: status }
      );
    }

    if (status === 429) {
      throw new MCPError(
        'Too many requests - rate limit exceeded',
        ErrorCode.RATE_LIMITED,
        { httpStatus: status }
      );
    }

    if (status >= 500) {
      throw new MCPError(
        'Microsoft Graph API server error',
        ErrorCode.GRAPH_API_ERROR,
        { httpStatus: status }
      );
    }
  }

  // Check for network errors
  if (error?.code === 'ECONNREFUSED' || error?.code === 'ENOTFOUND' || error?.code === 'ETIMEDOUT') {
    throw new MCPError(
      'Network error connecting to Microsoft Graph API',
      ErrorCode.NETWORK_ERROR,
      { originalCode: error.code }
    );
  }

  // Generic error
  throw new MCPError(
    error?.message || 'An unexpected error occurred',
    ErrorCode.UNKNOWN_ERROR,
    { originalError: error?.message || String(error) }
  );
}

/**
 * Wrap an async function with error handling
 * @param fn - The async function to wrap
 * @returns A function that catches and rethrows errors as MCPError
 */
export function withErrorHandling<T extends (...args: unknown[]) => Promise<any>>(
  fn: T
): T {
  return (async (...args: Parameters<T>) => {
    try {
      return await fn(...args);
    } catch (error) {
      if (error instanceof MCPError) {
        throw error;
      }
      handleGraphError(error);
    }
  }) as T;
}
