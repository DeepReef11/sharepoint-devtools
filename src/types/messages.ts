/**
 * Message types for communication between content scripts and background service worker
 */

export enum MessageType {
  // Content script -> Background
  SHAREPOINT_DETECTED = 'SHAREPOINT_DETECTED',
  GET_CONTEXT = 'GET_CONTEXT',
  OPEN_QUICKNAV = 'OPEN_QUICKNAV',

  // Background -> Content script
  CONTEXT_RESPONSE = 'CONTEXT_RESPONSE',
  ERROR = 'ERROR',
}

/**
 * Base message interface
 */
export interface BaseMessage {
  type: MessageType;
  timestamp: number;
}

/**
 * SharePoint detection message
 */
export interface SharePointDetectedMessage extends BaseMessage {
  type: MessageType.SHAREPOINT_DETECTED;
  data: {
    url: string;
    isSharePoint: boolean;
    isModern: boolean;
  };
}

/**
 * Context request message
 */
export interface GetContextMessage extends BaseMessage {
  type: MessageType.GET_CONTEXT;
}

/**
 * Open QuickNav modal message
 */
export interface OpenQuickNavMessage extends BaseMessage {
  type: MessageType.OPEN_QUICKNAV;
}

/**
 * Context response message
 */
export interface ContextResponseMessage extends BaseMessage {
  type: MessageType.CONTEXT_RESPONSE;
  data: {
    siteUrl?: string;
    webUrl?: string;
    listId?: string;
    isSharePoint: boolean;
    isModern: boolean;
  };
}

/**
 * Error message
 */
export interface ErrorMessage extends BaseMessage {
  type: MessageType.ERROR;
  error: {
    message: string;
    code?: string;
  };
}

/**
 * Union type of all possible messages
 */
export type ExtensionMessage =
  | SharePointDetectedMessage
  | GetContextMessage
  | OpenQuickNavMessage
  | ContextResponseMessage
  | ErrorMessage;

/**
 * Message response type
 */
export interface MessageResponse<T = unknown> {
  success: boolean;
  data?: T;
  error?: string;
}
