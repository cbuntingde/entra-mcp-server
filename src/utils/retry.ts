/**
 * Retry utility with exponential backoff for handling transient failures
 */

import { handleGraphError, MCPError, ErrorCode } from './error-handler.js';

export interface RetryOptions {
  /** Maximum number of retry attempts (default: 3) */
  maxRetries?: number;
  /** Base delay in milliseconds (default: 1000) */
  baseDelay?: number;
  /** Maximum delay in milliseconds (default: 30000) */
  maxDelay?: number;
  /** Whether to use jitter (default: true) */
  useJitter?: boolean;
  /** Retry on specific HTTP status codes */
  retryOnStatus?: number[];
  /** Custom retry condition */
  shouldRetry?: (error: any) => boolean;
}

const DEFAULT_RETRY_OPTIONS: Required<RetryOptions> = {
  maxRetries: 3,
  baseDelay: 1000,
  maxDelay: 30000,
  useJitter: true,
  retryOnStatus: [429, 500, 502, 503, 504],
  shouldRetry: () => false,
};

/**
 * Sleep for a specified number of milliseconds
 * @param ms - Milliseconds to sleep
 * @returns Promise that resolves after the delay
 */
function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Calculate delay with exponential backoff and optional jitter
 * @param attempt - The attempt number (0-indexed)
 * @param baseDelay - The base delay in milliseconds
 * @param maxDelay - The maximum delay in milliseconds
 * @param useJitter - Whether to add jitter
 * @returns The calculated delay in milliseconds
 */
function calculateDelay(
  attempt: number,
  baseDelay: number,
  maxDelay: number,
  useJitter: boolean
): number {
  // Exponential backoff: baseDelay * 2^attempt
  const exponentialDelay = baseDelay * Math.pow(2, attempt);

  // Add jitter if enabled (random +/- 25%)
  const jitter = useJitter ? (Math.random() - 0.5) * 0.5 * exponentialDelay : 0;

  // Cap at maxDelay
  return Math.min(exponentialDelay + jitter, maxDelay);
}

/**
 * Check if an error is retryable
 * @param error - The error to check
 * @param options - Retry options
 * @returns Whether the error is retryable
 */
function isRetryable(error: any, options: Required<RetryOptions>): boolean {
  // Check custom retry condition first
  if (options.shouldRetry?.(error)) {
    return true;
  }

  // Check for MCPError with retryable codes
  if (error instanceof MCPError) {
    if (error.code === ErrorCode.RATE_LIMITED || error.code === ErrorCode.TIMEOUT) {
      return true;
    }
  }

  // Check for HTTP status codes
  if (error?.response?.status) {
    return options.retryOnStatus.includes(error.response.status);
  }

  // Check for network errors
  if (
    error?.code === 'ECONNRESET' ||
    error?.code === 'ECONNREFUSED' ||
    error?.code === 'ETIMEDOUT' ||
    error?.code === 'ENOTFOUND' ||
    error?.code === 'EAI_AGAIN'
  ) {
    return true;
  }

  return false;
}

/**
 * Extract retry-after delay from error response
 * @param error - The error object
 * @returns The retry-after delay in milliseconds, or null if not present
 */
function getRetryAfterDelay(error: any): number | null {
  const retryAfter = error?.response?.headers?.['retry-after'];

  if (!retryAfter) {
    return null;
  }

  // Retry-After can be a number (seconds) or an HTTP-date
  if (typeof retryAfter === 'number') {
    return retryAfter * 1000; // Convert to milliseconds
  }

  if (typeof retryAfter === 'string') {
    const seconds = parseInt(retryAfter, 10);
    if (!isNaN(seconds)) {
      return seconds * 1000;
    }

    // Try parsing as HTTP-date
    const date = new Date(retryAfter);
    if (!isNaN(date.getTime())) {
      return Math.max(0, date.getTime() - Date.now());
    }
  }

  return null;
}

/**
 * Execute a function with retry logic and exponential backoff
 * @param fn - The async function to execute
 * @param options - Retry options
 * @returns The result of the function
 * @throws The last error if all retries are exhausted
 */
export async function retryWithBackoff<T>(
  fn: () => Promise<T>,
  options: RetryOptions = {}
): Promise<T> {
  const opts = { ...DEFAULT_RETRY_OPTIONS, ...options };

  let lastError: any;

  for (let attempt = 0; attempt <= opts.maxRetries; attempt++) {
    try {
      return await fn();
    } catch (error) {
      lastError = error;

      // Don't retry on the last attempt
      if (attempt === opts.maxRetries) {
        break;
      }

      // Check if error is retryable
      if (!isRetryable(error, opts)) {
        // Not retryable, throw immediately
        if (error instanceof MCPError) {
          throw error;
        }
        handleGraphError(error);
      }

      // Calculate delay
      let delay: number;

      // Check for retry-after header first
      const retryAfterDelay = getRetryAfterDelay(error);
      if (retryAfterDelay !== null) {
        delay = retryAfterDelay;
      } else {
        delay = calculateDelay(attempt, opts.baseDelay, opts.maxDelay, opts.useJitter);
      }

      // Log retry attempt (stderr to not interfere with MCP protocol)
      if (process.env.NODE_ENV !== 'production') {
        console.error(
          `[Retry] Attempt ${attempt + 1}/${opts.maxRetries + 1} failed. Retrying after ${Math.round(delay)}ms...`
        );
      }

      // Wait before retrying
      await sleep(delay);
    }
  }

  // All retries exhausted
  if (lastError instanceof MCPError) {
    throw lastError;
  }

  handleGraphError(lastError);
}

/**
 * Create a wrapped version of a function that includes retry logic
 * @param fn - The async function to wrap
 * @param options - Retry options
 * @returns A function that will retry on failure
 */
export function withRetry<T extends (...args: any[]) => Promise<any>>(
  fn: T,
  options: RetryOptions = {}
): T {
  return (async (...args: Parameters<T>) => {
    return retryWithBackoff(() => fn(...args), options);
  }) as T;
}
