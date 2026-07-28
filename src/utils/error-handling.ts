/**
 * Error Handling & Reliability Utilities
 * Provides comprehensive error handling, retry logic, and user-friendly messages
 */

/**
 * Custom error types for different SharePoint scenarios
 */
export class SharePointError extends Error {
  constructor(
    message: string,
    public readonly code: string,
    public readonly statusCode?: number,
    public readonly originalError?: Error
  ) {
    super(message);
    this.name = 'SharePointError';
  }
}

export class PermissionError extends SharePointError {
  constructor(message: string, statusCode?: number, originalError?: Error) {
    super(message, 'PERMISSION_DENIED', statusCode, originalError);
    this.name = 'PermissionError';
  }
}

export class AuthenticationError extends SharePointError {
  constructor(message: string, statusCode?: number, originalError?: Error) {
    super(message, 'AUTHENTICATION_FAILED', statusCode, originalError);
    this.name = 'AuthenticationError';
  }
}

export class NetworkError extends SharePointError {
  constructor(message: string, originalError?: Error) {
    super(message, 'NETWORK_ERROR', undefined, originalError);
    this.name = 'NetworkError';
  }
}

export class NotFoundError extends SharePointError {
  constructor(message: string, statusCode?: number, originalError?: Error) {
    super(message, 'NOT_FOUND', statusCode, originalError);
    this.name = 'NotFoundError';
  }
}

/**
 * Retry configuration options
 */
export interface RetryOptions {
  maxAttempts?: number;
  initialDelayMs?: number;
  maxDelayMs?: number;
  backoffMultiplier?: number;
  retryableStatusCodes?: number[];
  onRetry?: (attempt: number, error: Error) => void;
}

/**
 * Default retry configuration
 */
const DEFAULT_RETRY_OPTIONS: Required<RetryOptions> = {
  maxAttempts: 3,
  initialDelayMs: 1000,
  maxDelayMs: 10000,
  backoffMultiplier: 2,
  retryableStatusCodes: [408, 429, 500, 502, 503, 504],
  onRetry: () => {},
};

/**
 * Executes a function with exponential backoff retry logic
 */
export async function withRetry<T>(fn: () => Promise<T>, options: RetryOptions = {}): Promise<T> {
  const config = { ...DEFAULT_RETRY_OPTIONS, ...options };
  let lastError: Error | undefined;
  let delay = config.initialDelayMs;

  for (let attempt = 1; attempt <= config.maxAttempts; attempt++) {
    try {
      return await fn();
    } catch (error) {
      lastError = error instanceof Error ? error : new Error(String(error));

      // Check if error is retryable
      const isRetryable = isRetryableError(lastError, config.retryableStatusCodes);

      // Don't retry if not retryable or if this was the last attempt
      if (!isRetryable || attempt === config.maxAttempts) {
        throw lastError;
      }

      // Log retry attempt
      logDebug(`Retry attempt ${attempt}/${config.maxAttempts} after error:`, {
        errorName: lastError.name,
        errorMessage: lastError.message,
        delay,
      });

      // Call retry callback
      config.onRetry(attempt, lastError);

      // Wait before retrying
      await sleep(delay);

      // Calculate next delay with exponential backoff
      delay = Math.min(delay * config.backoffMultiplier, config.maxDelayMs);
    }
  }

  // This should never be reached, but TypeScript needs it
  throw lastError || new Error('Retry failed');
}

/**
 * Determines if an error is retryable
 */
function isRetryableError(error: Error, retryableStatusCodes: number[]): boolean {
  // Network errors are always retryable
  if (error instanceof NetworkError) {
    return true;
  }

  // SharePoint errors with retryable status codes
  if (error instanceof SharePointError && error.statusCode) {
    return retryableStatusCodes.includes(error.statusCode);
  }

  // Check for fetch network errors
  if (error.name === 'TypeError' && error.message.includes('fetch')) {
    return true;
  }

  return false;
}

/**
 * Sleep utility for retry delays
 */
function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Converts HTTP status codes to appropriate error types
 */
export function createErrorFromResponse(response: Response, message?: string): SharePointError {
  const status = response.status;
  const defaultMessage =
    message || `SharePoint API request failed: ${status} ${response.statusText}`;

  // Authentication errors
  if (status === 401) {
    return new AuthenticationError(
      message || 'Authentication required. Please sign in to SharePoint.',
      status
    );
  }

  // Permission errors
  if (status === 403) {
    return new PermissionError(
      message || 'You do not have permission to access this resource.',
      status
    );
  }

  // Not found errors
  if (status === 404) {
    return new NotFoundError(message || 'The requested resource was not found.', status);
  }

  // Rate limiting
  if (status === 429) {
    return new SharePointError(
      message || 'Too many requests. Please try again later.',
      'RATE_LIMITED',
      status
    );
  }

  // Server errors
  if (status >= 500) {
    return new SharePointError(
      message || 'SharePoint server error. Please try again later.',
      'SERVER_ERROR',
      status
    );
  }

  // Generic error
  return new SharePointError(defaultMessage, 'API_ERROR', status);
}

/**
 * Converts errors to user-friendly messages
 */
export function getUserFriendlyErrorMessage(error: Error | unknown): string {
  if (error instanceof PermissionError) {
    return 'You do not have permission to access this resource. Please contact your SharePoint administrator.';
  }

  if (error instanceof AuthenticationError) {
    return 'Authentication required. Please sign in to SharePoint and try again.';
  }

  if (error instanceof NotFoundError) {
    return 'The requested resource was not found. It may have been moved or deleted.';
  }

  if (error instanceof NetworkError) {
    return 'Network connection error. Please check your internet connection and try again.';
  }

  if (error instanceof SharePointError) {
    if (error.code === 'RATE_LIMITED') {
      return 'Too many requests. Please wait a moment and try again.';
    }
    if (error.code === 'SERVER_ERROR') {
      return 'SharePoint is experiencing issues. Please try again later.';
    }
    return error.message;
  }

  if (error instanceof Error) {
    // Generic error message without exposing sensitive details
    return 'An unexpected error occurred. Please try again.';
  }

  return 'An unknown error occurred. Please try again.';
}

/**
 * Sanitizes error for logging (removes sensitive data)
 */
export function sanitizeErrorForLogging(error: Error | unknown): Record<string, unknown> {
  if (error instanceof SharePointError) {
    return {
      name: error.name,
      code: error.code,
      message: error.message,
      statusCode: error.statusCode,
      // Don't log the original error stack trace which might contain sensitive URLs
    };
  }

  if (error instanceof Error) {
    return {
      name: error.name,
      message: error.message,
      // Don't log stack traces which might contain sensitive file paths
    };
  }

  return {
    type: typeof error,
    message: String(error),
  };
}

/**
 * Debug logging helper that sanitizes sensitive data
 */
export function logDebug(message: string, data?: Record<string, unknown>): void {
  // In development, log debug messages. Browser extensions don't have NODE_ENV.
  console.log(`[SharePoint DevTools] ${message}`, data);
}

/**
 * Error logging helper that sanitizes sensitive data
 */
export function logError(message: string, error: Error | unknown): void {
  const sanitized = sanitizeErrorForLogging(error);
  console.error(`[SharePoint DevTools] ${message}`, sanitized);
}

/**
 * Handles errors with fallback behavior
 */
export interface FallbackOptions<T> {
  fallbackValue: T;
  logError?: boolean;
  errorMessage?: string;
}

export async function withFallback<T>(
  fn: () => Promise<T>,
  options: FallbackOptions<T>
): Promise<T> {
  try {
    return await fn();
  } catch (error) {
    if (options.logError !== false) {
      logError(options.errorMessage || 'Operation failed, using fallback value', error);
    }
    return options.fallbackValue;
  }
}

/**
 * Safe async operation that won't throw
 */
export async function safely<T>(
  fn: () => Promise<T>,
  errorMessage?: string
): Promise<{ success: true; data: T } | { success: false; error: Error }> {
  try {
    const data = await fn();
    return { success: true, data };
  } catch (error) {
    const err = error instanceof Error ? error : new Error(String(error));
    if (errorMessage) {
      logError(errorMessage, err);
    }
    return { success: false, error: err };
  }
}
