import { useState, useEffect } from 'react';
import { commands, Command } from '../lib/commands';
import '../styles/command-palette.css';

interface CommandPaletteProps {
  query: string; // The text after '/'
  onSelect: (command: Command) => void;
  onClose: () => void;
  position: { top: number; left: number };
}

export default function CommandPalette({ query, onSelect, onClose, position }: CommandPaletteProps) {
  const [selectedIndex, setSelectedIndex] = useState(0);

  const filteredCommands = commands.filter(
    (cmd) =>
      cmd.label.toLowerCase().includes(query.toLowerCase()) ||
      cmd.description.toLowerCase().includes(query.toLowerCase())
  );

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setSelectedIndex((prev) => (prev + 1) % filteredCommands.length);
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        setSelectedIndex((prev) => (prev - 1 + filteredCommands.length) % filteredCommands.length);
      } else if (e.key === 'Enter') {
        e.preventDefault();
        if (filteredCommands[selectedIndex]) {
          onSelect(filteredCommands[selectedIndex]);
        }
      } else if (e.key === 'Escape') {
        e.preventDefault();
        onClose();
      }
    };

    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [selectedIndex, filteredCommands, onSelect, onClose]);

  if (filteredCommands.length === 0) return null;

  return (
    <div className="command-palette" style={{ top: position.top, left: position.left }}>
      {filteredCommands.map((cmd, index) => (
        <button
          key={cmd.id}
          className={`command-item ${index === selectedIndex ? 'selected' : ''}`}
          onClick={() => onSelect(cmd)}
          onMouseEnter={() => setSelectedIndex(index)}
          aria-label={`${cmd.label} - ${cmd.description}`}
        >
          {cmd.icon && <span className="command-icon">{cmd.icon}</span>}
          <div className="command-details">
            <div className="command-label">{cmd.label}</div>
            <div className="command-description">{cmd.description}</div>
          </div>
        </button>
      ))}
    </div>
  );
}
