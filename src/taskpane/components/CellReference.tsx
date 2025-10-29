import React from 'react';
import { Badge } from '@fluentui/react-components';
import '../styles/cell-reference.css';

interface CellReferenceProps {
  reference: string;
  onClick?: () => void;
}

export default function CellReference({ reference, onClick }: CellReferenceProps) {
  const handleClick = async () => {
    if (!onClick && window.Excel) {
      // Default behavior: select the referenced cell
      try {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const range = sheet.getRange(reference);
          range.select();

          // Briefly highlight the range
          range.format.fill.color = '#FFE6CC';
          await context.sync();

          // Clear highlight after 1 second
          setTimeout(async () => {
            await Excel.run(async (ctx) => {
              const s = ctx.workbook.worksheets.getActiveWorksheet();
              const r = s.getRange(reference);
              r.format.fill.clear();
              await ctx.sync();
            }).catch(() => {
              // Ignore errors if range is no longer available
            });
          }, 1000);
        });
      } catch (error) {
        console.error('Error selecting cell:', error);
      }
    } else if (onClick) {
      onClick();
    }
  };

  return (
    <Badge
      appearance="tint"
      size="small"
      className="cell-reference"
      onClick={handleClick}
      style={{ cursor: 'pointer' }}
      role="button"
      aria-label={`Cell reference ${reference}, click to select`}
    >
      📍 {reference}
    </Badge>
  );
}

// Utility function to detect cell references in text
export function detectCellReferences(text: string): {
  segments: Array<{ type: 'text' | 'cell'; content: string }>;
} {
  // Regex to match Excel cell references (e.g., A1, B2:C10, $A$1, Sheet1!A1)
  const cellReferenceRegex = /\b([A-Z]{1,3}[0-9]{1,7}(?::[A-Z]{1,3}[0-9]{1,7})?|(?:[A-Za-z0-9_]+!)?[A-Z]{1,3}[0-9]{1,7}(?::[A-Z]{1,3}[0-9]{1,7})?|\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?)\b/g;

  const segments: Array<{ type: 'text' | 'cell'; content: string }> = [];
  let lastIndex = 0;

  const matches = Array.from(text.matchAll(cellReferenceRegex));

  matches.forEach((match) => {
    const matchIndex = match.index!;

    // Add text before the match
    if (matchIndex > lastIndex) {
      segments.push({
        type: 'text',
        content: text.substring(lastIndex, matchIndex),
      });
    }

    // Add the cell reference
    segments.push({
      type: 'cell',
      content: match[0],
    });

    lastIndex = matchIndex + match[0].length;
  });

  // Add remaining text
  if (lastIndex < text.length) {
    segments.push({
      type: 'text',
      content: text.substring(lastIndex),
    });
  }

  return { segments };
}
