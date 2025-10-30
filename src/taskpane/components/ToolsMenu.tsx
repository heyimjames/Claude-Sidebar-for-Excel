import { useState, useRef, useEffect } from 'react';
import { MoreVertical24Regular, Checkmark24Regular } from '@fluentui/react-icons';
import type { ChatMessage } from '../lib/types';
import '../styles/tools-menu.css';

interface FreezeState {
  topRowFrozen: boolean;
  firstColumnFrozen: boolean;
}

interface ToolsMenuProps {
  messages?: ChatMessage[];
}

export default function ToolsMenu({ messages = [] }: ToolsMenuProps) {
  const [isOpen, setIsOpen] = useState(false);
  const [freezeState, setFreezeState] = useState<FreezeState>({
    topRowFrozen: false,
    firstColumnFrozen: false,
  });
  const menuRef = useRef<HTMLDivElement>(null);

  // Check freeze state when menu opens
  const checkFreezeState = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const freezePanes = sheet.freezePanes;
        freezePanes.load(['rows', 'columns']);
        await context.sync();

        setFreezeState({
          topRowFrozen: freezePanes.rows > 0,
          firstColumnFrozen: freezePanes.columns > 0,
        });
      });
    } catch (error) {
      console.error('Error checking freeze state:', error);
    }
  };

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };

    if (isOpen) {
      checkFreezeState();
      document.addEventListener('mousedown', handleClickOutside);
    }

    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isOpen]);

  const handleAutoFit = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getUsedRange();
        range.format.autofitColumns();
        range.format.autofitRows();
        await context.sync();
      });
    } catch (error) {
      console.error('AutoFit error:', error);
      alert('Error auto-fitting columns');
    }
  };

  const handleFormatAsTable = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address');
        await context.sync();

        const table = context.workbook.tables.add(range, true);
        table.style = 'TableStyleMedium2';
        await context.sync();
      });
    } catch (error) {
      console.error('Format as table error:', error);
      alert('Error formatting as table. Make sure your selection has headers.');
    }
  };

  const handleClearFormatting = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.clear();
        range.format.font.color = 'Black';
        range.format.font.bold = false;
        range.format.font.italic = false;
        range.format.borders.getItem('EdgeTop').style = 'None';
        range.format.borders.getItem('EdgeBottom').style = 'None';
        range.format.borders.getItem('EdgeLeft').style = 'None';
        range.format.borders.getItem('EdgeRight').style = 'None';
        await context.sync();
      });
    } catch (error) {
      console.error('Clear formatting error:', error);
      alert('Error clearing formatting');
    }
  };

  const handleFreezeTopRow = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        if (freezeState.topRowFrozen) {
          sheet.freezePanes.unfreeze();
        } else {
          sheet.freezePanes.freezeRows(1);
        }
        await context.sync();
      });
      await checkFreezeState();
    } catch (error) {
      console.error('Freeze top row error:', error);
      alert('Error toggling freeze top row');
    }
  };

  const handleFreezeFirstColumn = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        if (freezeState.firstColumnFrozen) {
          sheet.freezePanes.unfreeze();
        } else {
          sheet.freezePanes.freezeColumns(1);
        }
        await context.sync();
      });
      await checkFreezeState();
    } catch (error) {
      console.error('Freeze first column error:', error);
      alert('Error toggling freeze first column');
    }
  };

  const handleUnfreeze = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.freezePanes.unfreeze();
        await context.sync();
      });
      await checkFreezeState();
    } catch (error) {
      console.error('Unfreeze error:', error);
      alert('Error unfreezing panes');
    }
  };

  const handleSortAscending = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address');
        await context.sync();

        range.sort.apply([{ key: 0, ascending: true }]);
        await context.sync();
      });
    } catch (error) {
      console.error('Sort ascending error:', error);
      alert('Error sorting data');
    }
  };

  const handleSortDescending = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address');
        await context.sync();

        range.sort.apply([{ key: 0, ascending: false }]);
        await context.sync();
      });
    } catch (error) {
      console.error('Sort descending error:', error);
      alert('Error sorting data');
    }
  };

  const handleInsertRow = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.insert(Excel.InsertShiftDirection.down);
        await context.sync();
      });
    } catch (error) {
      console.error('Insert row error:', error);
      alert('Error inserting row');
    }
  };

  const handleInsertColumn = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.insert(Excel.InsertShiftDirection.right);
        await context.sync();
      });
    } catch (error) {
      console.error('Insert column error:', error);
      alert('Error inserting column');
    }
  };

  const handleDeleteRows = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.delete(Excel.DeleteShiftDirection.up);
        await context.sync();
      });
    } catch (error) {
      console.error('Delete rows error:', error);
      alert('Error deleting rows');
    }
  };

  const handleDeleteColumns = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.delete(Excel.DeleteShiftDirection.left);
        await context.sync();
      });
    } catch (error) {
      console.error('Delete columns error:', error);
      alert('Error deleting columns');
    }
  };

  const handleSelectAll = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getUsedRange();
        range.select();
        await context.sync();
      });
    } catch (error) {
      console.error('Select all error:', error);
      alert('Error selecting all data');
    }
  };

  const handleGoToA1 = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange('A1');
        range.select();
        await context.sync();
      });
    } catch (error) {
      console.error('Go to A1 error:', error);
      alert('Error navigating to A1');
    }
  };

  const handleCopyEntireChat = async () => {
    try {
      if (messages.length === 0) {
        alert('No chat history to copy');
        return;
      }

      // Format all messages as text
      const chatText = messages
        .map((msg) => {
          const role = msg.role === 'user' ? 'User' : 'Claude';
          const content = typeof msg.content === 'string'
            ? msg.content
            : msg.content.filter((block) => block.type === 'text').map((block) => block.text).join('\n');
          return `${role}:\n${content}\n`;
        })
        .join('\n---\n\n');

      // Copy to clipboard
      await navigator.clipboard.writeText(chatText);
      alert('Chat copied to clipboard!');
      setIsOpen(false);
    } catch (error) {
      console.error('Copy chat error:', error);
      alert('Error copying chat to clipboard');
    }
  };

  const tools = [
    {
      label: 'Copy Entire Chat to Clipboard',
      action: handleCopyEntireChat,
      description: 'Copy all chat messages to clipboard',
      active: false,
    },
    { divider: true },
    {
      label: 'AutoFit Columns & Rows',
      action: handleAutoFit,
      description: 'Auto-fit all columns and rows to content',
      active: false,
    },
    {
      label: 'Format as Table',
      action: handleFormatAsTable,
      description: 'Convert selection to a table',
      active: false,
    },
    {
      label: 'Clear Formatting',
      action: handleClearFormatting,
      description: 'Remove all formatting from selection',
      active: false,
    },
    { divider: true },
    {
      label: 'Insert Row Above',
      action: handleInsertRow,
      description: 'Insert new row above selection',
      active: false,
    },
    {
      label: 'Insert Column Left',
      action: handleInsertColumn,
      description: 'Insert new column to the left',
      active: false,
    },
    {
      label: 'Delete Selected Rows',
      action: handleDeleteRows,
      description: 'Delete the selected rows',
      active: false,
    },
    {
      label: 'Delete Selected Columns',
      action: handleDeleteColumns,
      description: 'Delete the selected columns',
      active: false,
    },
    { divider: true },
    {
      label: 'Freeze Top Row',
      action: handleFreezeTopRow,
      description: freezeState.topRowFrozen ? 'Unfreeze top row' : 'Keep top row visible when scrolling',
      active: freezeState.topRowFrozen,
    },
    {
      label: 'Freeze First Column',
      action: handleFreezeFirstColumn,
      description: freezeState.firstColumnFrozen ? 'Unfreeze first column' : 'Keep first column visible when scrolling',
      active: freezeState.firstColumnFrozen,
    },
    {
      label: 'Unfreeze Panes',
      action: handleUnfreeze,
      description: 'Remove all freeze panes',
      active: false,
    },
    { divider: true },
    {
      label: 'Sort Ascending',
      action: handleSortAscending,
      description: 'Sort selection A to Z',
      active: false,
    },
    {
      label: 'Sort Descending',
      action: handleSortDescending,
      description: 'Sort selection Z to A',
      active: false,
    },
    { divider: true },
    {
      label: 'Select All Data',
      action: handleSelectAll,
      description: 'Select all used cells in worksheet',
      active: false,
    },
    {
      label: 'Go to Cell A1',
      action: handleGoToA1,
      description: 'Jump to the top-left cell',
      active: false,
    },
  ];

  return (
    <div className="tools-menu" ref={menuRef}>
      <button
        className="icon-button"
        onClick={() => setIsOpen(!isOpen)}
        aria-label="Tools menu"
        aria-expanded={isOpen}
        title="Quick Tools"
        type="button"
      >
        <MoreVertical24Regular style={{ width: '14px', height: '14px' }} />
      </button>

      {isOpen && (
        <div className="tools-dropdown">
          {tools.map((tool, index) => {
            if ('divider' in tool) {
              return <div key={`divider-${index}`} className="tools-divider" />;
            }
            return (
              <button
                key={tool.label}
                className={`tools-item ${tool.active ? 'active' : ''}`}
                onClick={tool.action}
                type="button"
              >
                <div className="tools-item-content">
                  <div className="tools-item-text">
                    <div className="tools-item-label">{tool.label}</div>
                    <div className="tools-item-description">{tool.description}</div>
                  </div>
                  {tool.active && (
                    <Checkmark24Regular className="tools-item-checkmark" />
                  )}
                </div>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}
