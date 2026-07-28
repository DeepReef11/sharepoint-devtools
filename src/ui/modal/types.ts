/**
 * Type definitions for the Modal UI Component
 */

export interface ModalConfig {
  title?: string;
  placeholder?: string;
  darkMode?: boolean;
  onClose?: () => void;
  onSearch?: (query: string) => void;
  onSelect?: (item: ResultItem) => void;
}

export interface ResultItem {
  id: string;
  title: string;
  description?: string;
  url?: string;
  category: string;
  metadata?: Record<string, unknown>;
}

export interface CategoryGroup {
  category: string;
  items: ResultItem[];
}

export type ModalState = 'idle' | 'loading' | 'error' | 'success' | 'empty';

export interface ModalStateData {
  state: ModalState;
  message?: string;
  error?: Error;
}

export interface ErrorDisplayOptions {
  title?: string;
  message: string;
  details?: string;
  canRetry?: boolean;
  onRetry?: () => void;
}
