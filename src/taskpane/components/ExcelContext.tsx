import React from 'react';
import { Badge } from '@fluentui/react-components';
import { Table24Regular, Document24Regular } from '@fluentui/react-icons';
import type { ExcelContext as ExcelContextType } from '../hooks/useExcelContext';
import '../styles/excel-context.css';

interface ExcelContextProps {
  context: ExcelContextType;
  isLoading: boolean;
}

export default function ExcelContext({ context, isLoading }: ExcelContextProps) {
  if (isLoading) {
    return (
      <div className="excel-context loading" role="status" aria-label="Loading Excel context">
        <div className="excel-context-skeleton" />
      </div>
    );
  }

  if (!context.address) {
    return null;
  }

  const formatCellValue = (value: any): string => {
    if (value === null || value === undefined || value === '') return '—';
    if (typeof value === 'number') {
      return value.toLocaleString();
    }
    return String(value);
  };

  const getPreviewData = (): string => {
    const totalCells = context.rowCount * context.columnCount;

    // For extremely large selections where we didn't load any data
    if (!context.values || context.values.length === 0) {
      if (totalCells > 50000) {
        return `Very large selection (${totalCells.toLocaleString()} cells) - data not loaded to preserve performance`;
      }
      return 'No data';
    }

    const isLargeSelection = totalCells > 100;

    // Show first few cells as preview
    const maxPreview = 3;
    const firstRow = context.values[0];
    if (!firstRow) return 'No data';

    const preview = firstRow.slice(0, maxPreview).map(formatCellValue).join(', ');
    const hasMore = firstRow.length > maxPreview || context.values.length > 1;

    if (isLargeSelection) {
      return `${preview}... (preview of ${totalCells.toLocaleString()} cells)`;
    }

    return preview + (hasMore ? '...' : '');
  };

  const getCellInfo = (): string => {
    if (context.rowCount === 1 && context.columnCount === 1) {
      return '1 cell';
    }
    return `${context.rowCount} × ${context.columnCount} cells`;
  };

  return (
    <div className="excel-context" role="complementary" aria-label="Current Excel selection">
      <div className="excel-context-header">
        <div className="excel-context-icon">
          <Table24Regular />
        </div>
        <div className="excel-context-details">
          <div className="excel-context-title">
            <span className="excel-context-range">{context.address}</span>
            <Badge appearance="tint" size="small">
              {getCellInfo()}
            </Badge>
          </div>
          <div className="excel-context-subtitle">
            <Document24Regular className="sheet-icon" />
            <span>{context.sheetName}</span>
          </div>
        </div>
      </div>

      {context.hasData && (
        <div className="excel-context-preview">
          <div className="excel-context-preview-label">Preview:</div>
          <div className="excel-context-preview-content">{getPreviewData()}</div>
        </div>
      )}

      {context.hasData &&
        context.formulas &&
        context.formulas.some((row) => row.some((f) => typeof f === 'string' && f.startsWith('='))) && (
          <Badge appearance="outline" size="small" className="excel-context-badge">
            Contains formulas
          </Badge>
        )}
    </div>
  );
}
