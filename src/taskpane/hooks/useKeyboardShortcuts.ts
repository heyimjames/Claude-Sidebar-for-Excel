import { useEffect } from 'react';

interface ShortcutConfig {
  key: string;
  metaKey?: boolean; // Cmd on Mac, Ctrl on Windows
  ctrlKey?: boolean;
  shiftKey?: boolean;
  altKey?: boolean;
  callback: () => void;
  preventDefault?: boolean;
}

export function useKeyboardShortcuts(shortcuts: ShortcutConfig[]) {
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Ignore shortcuts when typing in input fields, textareas, or contenteditable elements
      const target = e.target as HTMLElement;
      const isTyping =
        target.tagName === 'INPUT' ||
        target.tagName === 'TEXTAREA' ||
        target.isContentEditable;

      // Allow Cmd/Ctrl shortcuts even when typing, but ignore other shortcuts
      if (isTyping && !e.metaKey && !e.ctrlKey) {
        return;
      }

      for (const shortcut of shortcuts) {
        const metaMatch = shortcut.metaKey === undefined || shortcut.metaKey === e.metaKey;
        const ctrlMatch = shortcut.ctrlKey === undefined || shortcut.ctrlKey === e.ctrlKey;
        const shiftMatch = shortcut.shiftKey === undefined || shortcut.shiftKey === e.shiftKey;
        const altMatch = shortcut.altKey === undefined || shortcut.altKey === e.altKey;
        const keyMatch = shortcut.key.toLowerCase() === e.key.toLowerCase();

        if (metaMatch && ctrlMatch && shiftMatch && altMatch && keyMatch) {
          if (shortcut.preventDefault !== false) {
            e.preventDefault();
          }
          shortcut.callback();
          break;
        }
      }
    };

    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [shortcuts]);
}

// Helper to detect OS
export function isMac() {
  return navigator.platform.toUpperCase().indexOf('MAC') >= 0;
}

// Format shortcut for display
export function formatShortcut(shortcut: ShortcutConfig): string {
  const parts: string[] = [];
  if (shortcut.metaKey) parts.push(isMac() ? '⌘' : 'Ctrl');
  if (shortcut.ctrlKey) parts.push('Ctrl');
  if (shortcut.altKey) parts.push(isMac() ? '⌥' : 'Alt');
  if (shortcut.shiftKey) parts.push('⇧');
  parts.push(shortcut.key.toUpperCase());
  return parts.join('+');
}
