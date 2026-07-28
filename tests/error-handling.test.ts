/**
 * Error Handling Tests
 * Validates error handling, retry logic, and error conversion
 */

import { describe, it, expect, jest } from '@jest/globals';
import {
  SharePointError,
  PermissionError,
  AuthenticationError,
  NetworkError,
  NotFoundError,
  withRetry,
  createErrorFromResponse,
  getUserFriendlyErrorMessage,
  sanitizeErrorForLogging,
  withFallback,
  safely,
} from '../src/utils/error-handling';

describe('Error Classes', () => {
  describe('SharePointError', () => {
    it('should create error with code and status', () => {
      const error = new SharePointError('Test error', 'TEST_CODE', 500);

      expect(error.message).toBe('Test error');
      expect(error.code).toBe('TEST_CODE');
      expect(error.statusCode).toBe(500);
      expect(error.name).toBe('SharePointError');
    });

    it('should accept original error', () => {
      const originalError = new Error('Original');
      const error = new SharePointError('Wrapped', 'CODE', 500, originalError);

      expect(error.originalError).toBe(originalError);
    });
  });

  describe('Specific Error Types', () => {
    it('should create PermissionError', () => {
      const error = new PermissionError('Access denied', 403);

      expect(error.name).toBe('PermissionError');
      expect(error.code).toBe('PERMISSION_DENIED');
      expect(error.statusCode).toBe(403);
    });

    it('should create AuthenticationError', () => {
      const error = new AuthenticationError('Not authenticated', 401);

      expect(error.name).toBe('AuthenticationError');
      expect(error.code).toBe('AUTHENTICATION_FAILED');
      expect(error.statusCode).toBe(401);
    });

    it('should create NetworkError', () => {
      const error = new NetworkError('Connection failed');

      expect(error.name).toBe('NetworkError');
      expect(error.code).toBe('NETWORK_ERROR');
      expect(error.statusCode).toBeUndefined();
    });

    it('should create NotFoundError', () => {
      const error = new NotFoundError('Resource not found', 404);

      expect(error.name).toBe('NotFoundError');
      expect(error.code).toBe('NOT_FOUND');
      expect(error.statusCode).toBe(404);
    });
  });
});

describe('withRetry', () => {
  // These tests use small real delays rather than fake timers: withRetry awaits a
  // setTimeout internally, so advancing fake timers requires the advancing call
  // itself to run on a real timer, which fake timers also intercept.
  const FAST_RETRY = { initialDelayMs: 5, maxDelayMs: 50 };

  it('should succeed on first attempt', async () => {
    const fn = jest.fn<() => Promise<string>>();
    fn.mockResolvedValue('success');

    const result = await withRetry(fn, { maxAttempts: 3, ...FAST_RETRY });

    expect(result).toBe('success');
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it('should retry on retryable errors', async () => {
    const fn = jest.fn<() => Promise<string>>();
    fn.mockRejectedValueOnce(new NetworkError('Network error'));
    fn.mockResolvedValue('success');

    const result = await withRetry(fn, { maxAttempts: 3, ...FAST_RETRY });

    expect(result).toBe('success');
    expect(fn).toHaveBeenCalledTimes(2);
  });

  it('should retry on specific status codes', async () => {
    const fn = jest.fn<() => Promise<string>>();
    fn.mockRejectedValueOnce(new SharePointError('Server error', 'ERROR', 503));
    fn.mockResolvedValue('success');

    const result = await withRetry(fn, {
      maxAttempts: 3,
      ...FAST_RETRY,
      retryableStatusCodes: [503],
    });

    expect(result).toBe('success');
    expect(fn).toHaveBeenCalledTimes(2);
  });

  it('should not retry when the status code is not retryable', async () => {
    const fn = jest.fn<() => Promise<string>>();
    fn.mockRejectedValue(new SharePointError('Teapot', 'ERROR', 418));

    await expect(
      withRetry(fn, { maxAttempts: 3, ...FAST_RETRY, retryableStatusCodes: [503] })
    ).rejects.toThrow(SharePointError);
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it('should not retry non-retryable errors', async () => {
    const fn = jest.fn<() => Promise<string>>();
    fn.mockRejectedValue(new PermissionError('Access denied', 403));

    await expect(withRetry(fn, { maxAttempts: 3, ...FAST_RETRY })).rejects.toThrow(PermissionError);
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it('should throw after max attempts', async () => {
    const fn = jest.fn<() => Promise<string>>();
    fn.mockRejectedValue(new NetworkError('Network error'));

    await expect(withRetry(fn, { maxAttempts: 2, ...FAST_RETRY })).rejects.toThrow(NetworkError);
    expect(fn).toHaveBeenCalledTimes(2);
  });

  it('should call onRetry callback', async () => {
    const onRetry = jest.fn();
    const fn = jest.fn<() => Promise<string>>();
    fn.mockRejectedValueOnce(new NetworkError('Error'));
    fn.mockResolvedValue('success');

    await withRetry(fn, { maxAttempts: 3, ...FAST_RETRY, onRetry });

    expect(onRetry).toHaveBeenCalledTimes(1);
    expect(onRetry).toHaveBeenCalledWith(1, expect.any(NetworkError));
  });

  it('should use exponential backoff', async () => {
    const gaps: number[] = [];
    let previous = Date.now();
    const fn = jest.fn<() => Promise<string>>();
    fn.mockRejectedValue(new NetworkError('Error'));

    await expect(
      withRetry(fn, {
        maxAttempts: 3,
        initialDelayMs: 30,
        maxDelayMs: 500,
        backoffMultiplier: 4,
        onRetry: () => {
          const now = Date.now();
          gaps.push(now - previous);
          previous = now;
        },
      })
    ).rejects.toThrow(NetworkError);

    expect(gaps.length).toBe(2);
    // First retry waits ~30ms, second waits ~120ms. Compare the observed gaps
    // rather than asserting exact timings, which are flaky under load.
    expect(gaps[1]).toBeGreaterThan(gaps[0]);
    expect(fn).toHaveBeenCalledTimes(3);
  });

  it('should cap the delay at maxDelayMs', async () => {
    const fn = jest.fn<() => Promise<string>>();
    fn.mockRejectedValue(new NetworkError('Error'));

    const started = Date.now();
    await expect(
      withRetry(fn, {
        maxAttempts: 3,
        initialDelayMs: 10,
        maxDelayMs: 15,
        backoffMultiplier: 100,
      })
    ).rejects.toThrow(NetworkError);

    // Without the cap the second wait would be 1000ms.
    expect(Date.now() - started).toBeLessThan(500);
    expect(fn).toHaveBeenCalledTimes(3);
  });
});

describe('createErrorFromResponse', () => {
  it('should create AuthenticationError for 401', () => {
    const response = { status: 401, statusText: 'Unauthorized' } as Response;
    const error = createErrorFromResponse(response);

    expect(error).toBeInstanceOf(AuthenticationError);
    expect(error.statusCode).toBe(401);
  });

  it('should create PermissionError for 403', () => {
    const response = { status: 403, statusText: 'Forbidden' } as Response;
    const error = createErrorFromResponse(response);

    expect(error).toBeInstanceOf(PermissionError);
    expect(error.statusCode).toBe(403);
  });

  it('should create NotFoundError for 404', () => {
    const response = { status: 404, statusText: 'Not Found' } as Response;
    const error = createErrorFromResponse(response);

    expect(error).toBeInstanceOf(NotFoundError);
    expect(error.statusCode).toBe(404);
  });

  it('should create rate limit error for 429', () => {
    const response = { status: 429, statusText: 'Too Many Requests' } as Response;
    const error = createErrorFromResponse(response);

    expect(error.code).toBe('RATE_LIMITED');
    expect(error.statusCode).toBe(429);
  });

  it('should create server error for 5xx', () => {
    const response = { status: 500, statusText: 'Internal Server Error' } as Response;
    const error = createErrorFromResponse(response);

    expect(error.code).toBe('SERVER_ERROR');
    expect(error.statusCode).toBe(500);
  });

  it('should create generic error for other status codes', () => {
    const response = { status: 400, statusText: 'Bad Request' } as Response;
    const error = createErrorFromResponse(response);

    expect(error.code).toBe('API_ERROR');
    expect(error.statusCode).toBe(400);
  });

  it('should use custom message if provided', () => {
    const response = { status: 500, statusText: 'Error' } as Response;
    const error = createErrorFromResponse(response, 'Custom error message');

    expect(error.message).toContain('Custom error message');
  });
});

describe('getUserFriendlyErrorMessage', () => {
  it('should return friendly message for PermissionError', () => {
    const error = new PermissionError('Access denied');
    const message = getUserFriendlyErrorMessage(error);

    expect(message).toContain('permission');
    expect(message).toContain('administrator');
  });

  it('should return friendly message for AuthenticationError', () => {
    const error = new AuthenticationError('Not authenticated');
    const message = getUserFriendlyErrorMessage(error);

    expect(message).toContain('Authentication');
    expect(message).toContain('sign in');
  });

  it('should return friendly message for NotFoundError', () => {
    const error = new NotFoundError('Not found');
    const message = getUserFriendlyErrorMessage(error);

    expect(message).toContain('not found');
  });

  it('should return friendly message for NetworkError', () => {
    const error = new NetworkError('Connection failed');
    const message = getUserFriendlyErrorMessage(error);

    expect(message).toContain('Network');
    expect(message).toContain('connection');
  });

  it('should return friendly message for rate limit errors', () => {
    const error = new SharePointError('Too many requests', 'RATE_LIMITED', 429);
    const message = getUserFriendlyErrorMessage(error);

    expect(message).toContain('Too many requests');
  });

  it('should return generic message for unknown errors', () => {
    const error = new Error('Unknown error');
    const message = getUserFriendlyErrorMessage(error);

    expect(message).toContain('unexpected error');
  });
});

describe('sanitizeErrorForLogging', () => {
  it('should sanitize SharePointError', () => {
    const error = new SharePointError('Test error', 'TEST_CODE', 500);
    const sanitized = sanitizeErrorForLogging(error);

    expect(sanitized).toEqual({
      name: 'SharePointError',
      code: 'TEST_CODE',
      message: 'Test error',
      statusCode: 500,
    });
  });

  it('should sanitize generic Error', () => {
    const error = new Error('Generic error');
    const sanitized = sanitizeErrorForLogging(error);

    expect(sanitized).toEqual({
      name: 'Error',
      message: 'Generic error',
    });
  });

  it('should handle non-Error objects', () => {
    const sanitized = sanitizeErrorForLogging('string error');

    expect(sanitized).toEqual({
      type: 'string',
      message: 'string error',
    });
  });
});

describe('withFallback', () => {
  it('should return result on success', async () => {
    const fn = jest.fn<() => Promise<string>>();
    fn.mockResolvedValue('success');
    const result = await withFallback(fn, { fallbackValue: 'fallback' });

    expect(result).toBe('success');
  });

  it('should return fallback value on error', async () => {
    const fn = jest.fn<() => Promise<string>>();
    fn.mockRejectedValue(new Error('Failed'));
    const result = await withFallback(fn, { fallbackValue: 'fallback' });

    expect(result).toBe('fallback');
  });

  it('should not log error when logError is false', async () => {
    const fn = jest.fn<() => Promise<string>>();
    fn.mockRejectedValue(new Error('Failed'));
    const consoleSpy = jest.spyOn(console, 'error');

    await withFallback(fn, { fallbackValue: 'fallback', logError: false });

    expect(consoleSpy).not.toHaveBeenCalled();
  });
});

describe('safely', () => {
  it('should return success result', async () => {
    const fn = jest.fn<() => Promise<string>>();
    fn.mockResolvedValue('data');
    const result = await safely(fn);

    expect(result).toEqual({ success: true, data: 'data' });
  });

  it('should return error result', async () => {
    const error = new Error('Failed');
    const fn = jest.fn<() => Promise<string>>();
    fn.mockRejectedValue(error);
    const result = await safely(fn);

    expect(result).toEqual({ success: false, error });
  });

  it('should convert non-Error to Error', async () => {
    const fn = jest.fn<() => Promise<string>>();
    fn.mockRejectedValue('string error');
    const result = await safely(fn);

    expect(result.success).toBe(false);
    if (result.success === false) {
      expect(result.error).toBeInstanceOf(Error);
      expect(result.error.message).toBe('string error');
    }
  });

  it('should log error when message provided', async () => {
    const fn = jest.fn<() => Promise<string>>();
    fn.mockRejectedValue(new Error('Failed'));
    const consoleSpy = jest.spyOn(console, 'error');

    await safely(fn, 'Operation failed');

    expect(consoleSpy).toHaveBeenCalled();
  });
});
