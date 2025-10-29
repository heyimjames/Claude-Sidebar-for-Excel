import type Anthropic from '@anthropic-ai/sdk';

// File attachment for UI tracking (images and documents)
export interface ImageAttachment {
  id: string;
  type: 'base64' | 'url';
  data: string; // base64 data (without prefix) or URL
  mediaType: 'image/jpeg' | 'image/png' | 'image/gif' | 'image/webp' | 'application/pdf';
  previewUrl?: string; // data URL for UI preview (images only)
  name?: string;
  fileType: 'image' | 'document';
}

// Content blocks matching Anthropic's API format
export interface TextContent {
  type: 'text';
  text: string;
}

export interface ImageContent {
  type: 'image';
  source: {
    type: 'base64' | 'url';
    media_type: 'image/jpeg' | 'image/png' | 'image/gif' | 'image/webp';
    data?: string; // for base64
    url?: string; // for url
  };
}

export interface DocumentContent {
  type: 'document';
  source: {
    type: 'base64';
    media_type: 'application/pdf';
    data: string;
  };
}

export type MessageContent = string | Array<TextContent | ImageContent | DocumentContent>;

export interface ChatMessage {
  id: string;
  role: 'user' | 'assistant';
  content: MessageContent;
  attachments?: ImageAttachment[]; // For UI state tracking
  isStreaming?: boolean;
  isAnimating?: boolean;
}

export interface ExcelTool {
  name: string;
  description: string;
  input_schema: {
    type: 'object';
    properties: Record<string, any>;
    required?: string[];
  };
}

export type AnthropicTool = Anthropic.Tool;

export interface ToolExecutionResult {
  success: boolean;
  data?: any;
  error?: string;
}
