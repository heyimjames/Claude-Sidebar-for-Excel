import { Dialog, DialogSurface, DialogTitle, DialogBody } from '@fluentui/react-components';
import { isMac } from '../hooks/useKeyboardShortcuts';
import '../styles/shortcuts-help.css';

interface ShortcutsHelpProps {
  open: boolean;
  onClose: () => void;
}

export default function ShortcutsHelp({ open, onClose }: ShortcutsHelpProps) {
  const shortcuts = [
    { keys: `${isMac() ? '⌘' : 'Ctrl'}+Enter`, action: 'Send message' },
    { keys: `${isMac() ? '⌘' : 'Ctrl'}+K`, action: 'Focus message input' },
    { keys: '/', action: 'Show quick commands' },
    { keys: `${isMac() ? '⌘' : 'Ctrl'}+L`, action: 'Clear chat history' },
    { keys: 'Shift+Enter', action: 'New line in message' },
    { keys: 'Escape', action: 'Clear input / Close palette' },
    { keys: 'Shift+?', action: 'Show this help' },
    { keys: '↑ / ↓', action: 'Navigate commands' },
    { keys: 'Enter', action: 'Select command' },
  ];

  return (
    <Dialog open={open} onOpenChange={(_, data) => !data.open && onClose()}>
      <DialogSurface>
        <DialogTitle>Keyboard Shortcuts</DialogTitle>
        <DialogBody>
          <div className="shortcuts-list">
            {shortcuts.map((shortcut, index) => (
              <div key={index} className="shortcut-item">
                <kbd className="shortcut-keys">{shortcut.keys}</kbd>
                <span className="shortcut-action">{shortcut.action}</span>
              </div>
            ))}
          </div>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
