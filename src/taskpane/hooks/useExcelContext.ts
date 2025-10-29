import { useState, useEffect, useCallback } from 'react';

export interface ExcelContext {
  address: string;
  sheetName: string;
  values: any[][] | null;
  formulas: string[][] | null;
  rowCount: number;
  columnCount: number;
  hasData: boolean;
}

export function useExcelContext() {
  const [context, setContext] = useState<ExcelContext>({
    address: '',
    sheetName: '',
    values: null,
    formulas: null,
    rowCount: 0,
    columnCount: 0,
    hasData: false,
  });

  const [isLoading, setIsLoading] = useState(false);

  const updateContext = useCallback(async () => {
    if (!window.Excel) {
      console.log('[ExcelContext] Excel API not available');
      return;
    }

    setIsLoading(true);

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = context.workbook.getSelectedRange();

        // First, load just the basic info
        sheet.load('name');
        range.load(['address', 'rowCount', 'columnCount']);

        await context.sync();

        const totalCells = range.rowCount * range.columnCount;
        console.log(`[ExcelContext] Selection: ${range.address}, Total cells: ${totalCells.toLocaleString()}`);

        // Very aggressive limits to prevent crashes
        const MAX_CELLS_FOR_FULL_LOAD = 100;
        const MAX_CELLS_FOR_PREVIEW = 50000; // Don't even try to load data beyond this

        let values: any[][] | null = null;
        let formulas: string[][] | null = null;
        let hasData = false;

        if (totalCells > MAX_CELLS_FOR_PREVIEW) {
          // For extremely large selections, don't load any data at all
          console.log(`[ExcelContext] Selection too large (${totalCells.toLocaleString()} cells), skipping data load`);
          hasData = true; // Assume it has data
        } else if (totalCells <= MAX_CELLS_FOR_FULL_LOAD) {
          // Load all data for small selections
          console.log('[ExcelContext] Loading full selection data');
          range.load(['values', 'formulas']);
          await context.sync();

          values = range.values;
          formulas = range.formulas;
          hasData =
            range.values.length > 0 &&
            range.values.some((row) => row.some((cell) => cell !== null && cell !== ''));
        } else {
          // For medium selections, only load a small preview
          const previewRows = Math.min(3, range.rowCount);
          const previewCols = Math.min(5, range.columnCount);
          console.log(`[ExcelContext] Loading preview (${previewRows}x${previewCols})`);

          try {
            const previewRange = range.getCell(0, 0).getResizedRange(previewRows - 1, previewCols - 1);
            previewRange.load(['values', 'formulas']);
            await context.sync();

            values = previewRange.values;
            formulas = previewRange.formulas;
            hasData = true;
          } catch (previewError) {
            console.error('[ExcelContext] Error loading preview:', previewError);
            // Even preview failed, just skip data loading
            hasData = true;
          }
        }

        // Update state with new context
        const newContext = {
          address: range.address,
          sheetName: sheet.name,
          values,
          formulas,
          rowCount: range.rowCount,
          columnCount: range.columnCount,
          hasData,
        };

        console.log('[ExcelContext] Context updated successfully');
        setContext(newContext);
      });
    } catch (error) {
      console.error('[ExcelContext] Error updating Excel context:', error);
      // Don't clear the context on error, keep the last valid state
      // This prevents the UI from disappearing
    } finally {
      setIsLoading(false);
    }
  }, []);

  // Update context on mount and when selection changes
  useEffect(() => {
    if (!window.Excel) {
      console.log('[ExcelContext] Excel API not available, skipping event setup');
      return;
    }

    // Initial load
    updateContext();

    // Set up event handler
    let handlerId: string | null = null;

    const setupEventHandler = async () => {
      try {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const eventResult = sheet.onSelectionChanged.add(() => {
            // Use a small delay to debounce rapid selection changes
            setTimeout(updateContext, 100);
          });
          await context.sync();
          handlerId = eventResult.value;
          console.log('[ExcelContext] Selection change listener registered');
        });
      } catch (error) {
        console.error('[ExcelContext] Error setting up selection change listener:', error);
      }
    };

    setupEventHandler();

    // Cleanup
    return () => {
      if (handlerId) {
        Excel.run(async (context) => {
          try {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.onSelectionChanged.remove(handlerId!);
            await context.sync();
            console.log('[ExcelContext] Selection change listener removed');
          } catch (error) {
            console.error('[ExcelContext] Error removing selection change listener:', error);
          }
        });
      }
    };
  }, [updateContext]);

  const highlightRange = useCallback(async (address: string, color: string = '#FFE6CC') => {
    if (!window.Excel) return;

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);
        range.format.fill.color = color;
        await context.sync();
      });
    } catch (error) {
      console.error('Error highlighting range:', error);
    }
  }, []);

  const clearHighlight = useCallback(async (address: string) => {
    if (!window.Excel) return;

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);
        range.format.fill.clear();
        await context.sync();
      });
    } catch (error) {
      console.error('Error clearing highlight:', error);
    }
  }, []);

  const selectRange = useCallback(async (address: string) => {
    if (!window.Excel) return;

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);
        range.select();
        await context.sync();
      });
    } catch (error) {
      console.error('Error selecting range:', error);
    }
  }, []);

  return {
    context,
    isLoading,
    updateContext,
    highlightRange,
    clearHighlight,
    selectRange,
  };
}
