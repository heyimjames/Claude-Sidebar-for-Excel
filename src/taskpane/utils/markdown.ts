// Utility functions for markdown processing

/**
 * Detects the language from a code fence
 * e.g., "```typescript" returns "typescript"
 */
export function detectLanguage(codeBlockInfo: string): string {
  const match = /^(\w+)/.exec(codeBlockInfo || '');
  return match ? match[1] : 'text';
}

/**
 * Checks if a code block is inline or multiline
 */
export function isInlineCode(code: string): boolean {
  return !code.includes('\n');
}

/**
 * Formats code for display
 */
export function formatCode(code: string): string {
  return code.replace(/\n$/, ''); // Remove trailing newline
}
