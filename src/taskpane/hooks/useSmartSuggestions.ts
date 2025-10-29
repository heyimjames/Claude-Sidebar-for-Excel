import { useMemo } from 'react';
import type { ExcelContext } from './useExcelContext';

export interface Suggestion {
  id: string;
  label: string;
  icon: string;
  prompt: string;
  description: string;
}

export function useSmartSuggestions(context: ExcelContext): Suggestion[] {
  return useMemo(() => {
    if (!context.hasData || !context.values) {
      return [];
    }

    const suggestions: Suggestion[] = [];
    const totalCells = context.rowCount * context.columnCount;

    // Analyze data types
    const hasNumbers = context.values.some((row) =>
      row.some((cell) => typeof cell === 'number')
    );
    const hasDates = context.values.some((row) =>
      row.some((cell) => cell instanceof Date || isDateString(cell))
    );
    const hasText = context.values.some((row) =>
      row.some((cell) => typeof cell === 'string' && cell.trim() !== '')
    );
    const hasFormulas =
      context.formulas &&
      context.formulas.some((row) => row.some((formula) => typeof formula === 'string' && formula.startsWith('=')));

    // Multi-row/column data
    const isTable = context.rowCount > 1 && context.columnCount > 1;
    const isSingleColumn = context.columnCount === 1 && context.rowCount > 1;
    const isSingleRow = context.rowCount === 1 && context.columnCount > 1;

    // Suggestions for numeric data
    if (hasNumbers) {
      if (isSingleColumn || isSingleRow) {
        suggestions.push({
          id: 'sum',
          label: 'Sum',
          icon: '∑',
          prompt: 'Calculate the sum of these values',
          description: 'Add all selected numbers',
        });
        suggestions.push({
          id: 'average',
          label: 'Average',
          icon: '≈',
          prompt: 'Calculate the average of these values',
          description: 'Find the mean',
        });
      }

      if (isTable) {
        suggestions.push({
          id: 'chart',
          label: 'Create Chart',
          icon: '📊',
          prompt: 'Create a chart from this data',
          description: 'Visualize your data',
        });
      }

      if (totalCells >= 10) {
        suggestions.push({
          id: 'analyze',
          label: 'Analyze',
          icon: '🔍',
          prompt: 'Analyze this data and provide insights',
          description: 'Get AI insights',
        });
      }
    }

    // Suggestions for dates
    if (hasDates && isTable) {
      suggestions.push({
        id: 'timeline',
        label: 'Timeline',
        icon: '📅',
        prompt: 'Create a timeline visualization of this data',
        description: 'Show data over time',
      });
    }

    // Suggestions for tables
    if (isTable) {
      suggestions.push({
        id: 'sort',
        label: 'Sort',
        icon: '⇅',
        prompt: 'Sort this data',
        description: 'Order rows',
      });

      if (context.rowCount > 5 && context.columnCount >= 2) {
        suggestions.push({
          id: 'pivot',
          label: 'Pivot Table',
          icon: '📋',
          prompt: 'Create a pivot table from this data',
          description: 'Summarize and group',
        });
      }

      suggestions.push({
        id: 'format',
        label: 'Format',
        icon: '🎨',
        prompt: 'Format this table to look professional',
        description: 'Apply table styling',
      });
    }

    // Suggestions for any data
    if (context.hasData) {
      if (totalCells > 20) {
        suggestions.push({
          id: 'summarize',
          label: 'Summarize',
          icon: '📝',
          prompt: 'Summarize what this data shows',
          description: 'Get a quick overview',
        });
      }

      if (hasFormulas) {
        suggestions.push({
          id: 'explain-formulas',
          label: 'Explain Formulas',
          icon: 'ƒx',
          prompt: 'Explain the formulas in this selection',
          description: 'Understand the calculations',
        });
      }
    }

    // Limit to 4 suggestions to avoid clutter
    return suggestions.slice(0, 4);
  }, [context]);
}

// Helper function to detect date strings
function isDateString(value: any): boolean {
  if (typeof value !== 'string') return false;
  const datePatterns = [
    /^\d{1,2}\/\d{1,2}\/\d{2,4}$/, // MM/DD/YYYY or M/D/YY
    /^\d{4}-\d{2}-\d{2}$/, // YYYY-MM-DD
    /^\d{1,2}-\d{1,2}-\d{2,4}$/, // MM-DD-YYYY or M-D-YY
  ];
  return datePatterns.some((pattern) => pattern.test(value));
}
