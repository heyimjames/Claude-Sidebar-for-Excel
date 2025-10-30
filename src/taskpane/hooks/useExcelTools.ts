import { useCallback } from 'react';
import type { ToolExecutionResult } from '../lib/types';
import { excelTools } from '../lib/excel-tools';

/* global Excel */

export function useExcelTools() {
  const executeTool = useCallback(async (toolName: string, input: any): Promise<ToolExecutionResult> => {
    try {
      // Handle web search separately (doesn't require Excel context)
      if (toolName === 'web_search') {
        // The Anthropic API with extended thinking will handle the actual search
        // We just need to acknowledge the tool use
        return {
          success: true,
          data: {
            query: input.query,
            note: 'Web search executed. Results integrated into response.',
          },
        };
      }

      return await Excel.run(async (context) => {
        switch (toolName) {
          case 'read_range': {
            const sheet = input.worksheet
              ? context.workbook.worksheets.getItem(input.worksheet)
              : context.workbook.worksheets.getActiveWorksheet();

            const range = sheet.getRange(input.range);
            range.load('values, formulas, numberFormat, address');
            await context.sync();

            return {
              success: true,
              data: {
                address: range.address,
                values: range.values,
                formulas: range.formulas,
                formats: range.numberFormat,
              },
            };
          }

          case 'write_range': {
            const sheet = input.worksheet
              ? context.workbook.worksheets.getItem(input.worksheet)
              : context.workbook.worksheets.getActiveWorksheet();

            const range = sheet.getRange(input.range);
            range.values = input.values;
            await context.sync();

            return {
              success: true,
              data: { range: input.range, rowsWritten: input.values.length },
            };
          }

          case 'get_selection': {
            const range = context.workbook.getSelectedRange();
            range.load('address, values, formulas, rowCount, columnCount');
            await context.sync();

            return {
              success: true,
              data: {
                address: range.address,
                values: range.values,
                formulas: range.formulas,
                rowCount: range.rowCount,
                columnCount: range.columnCount,
              },
            };
          }

          case 'get_workbook_info': {
            const sheets = context.workbook.worksheets;
            sheets.load('items/name, items/position');
            const activeSheet = context.workbook.worksheets.getActiveWorksheet();
            activeSheet.load('name');
            await context.sync();

            const sheetNames = sheets.items.map(s => s.name);

            return {
              success: true,
              data: {
                worksheets: sheetNames,
                activeWorksheet: activeSheet.name,
              },
            };
          }

          case 'create_table': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const tableName = input.tableName || `Table${Date.now()}`;
            const hasHeaders = input.hasHeaders !== false;

            const table = sheet.tables.add(input.range, hasHeaders);
            table.name = tableName;
            table.getHeaderRowRange().format.fill.color = '#4472C4';
            table.getHeaderRowRange().format.font.color = 'white';
            table.getHeaderRowRange().format.font.bold = true;
            await context.sync();

            return {
              success: true,
              data: { tableName, range: input.range },
            };
          }

          case 'create_chart': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const dataRange = sheet.getRange(input.dataRange);
            const chart = sheet.charts.add(input.chartType, dataRange, Excel.ChartSeriesBy.auto);

            if (input.title) {
              chart.title.text = input.title;
            }

            chart.setPosition('F2');
            await context.sync();

            return {
              success: true,
              data: { chartType: input.chartType, title: input.title },
            };
          }

          case 'apply_formula': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);

            // Load range dimensions
            range.load('rowCount, columnCount');
            await context.sync();

            // Create 2D array filled with the formula for each cell
            // Excel will auto-adjust relative references (e.g., A1 -> A2, A3, etc.)
            const formulas: string[][] = [];
            for (let row = 0; row < range.rowCount; row++) {
              const rowFormulas: string[] = [];
              for (let col = 0; col < range.columnCount; col++) {
                rowFormulas.push(input.formula);
              }
              formulas.push(rowFormulas);
            }

            range.formulas = formulas;
            await context.sync();

            return {
              success: true,
              data: {
                range: input.range,
                formula: input.formula,
                cellsAffected: range.rowCount * range.columnCount
              },
            };
          }

          case 'format_range': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);

            if (input.format.numberFormat) {
              range.numberFormat = [[input.format.numberFormat]];
            }

            if (input.format.fontBold !== undefined) {
              range.format.font.bold = input.format.fontBold;
            }

            if (input.format.fontSize) {
              range.format.font.size = input.format.fontSize;
            }

            if (input.format.fillColor) {
              range.format.fill.color = input.format.fillColor;
            }

            if (input.format.fontColor) {
              range.format.font.color = input.format.fontColor;
            }

            await context.sync();

            return {
              success: true,
              data: { range: input.range, format: input.format },
            };
          }

          case 'insert_rows': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRangeByIndexes(input.index, 0, input.count, 1);
            range.insert(Excel.InsertShiftDirection.down);
            await context.sync();

            return {
              success: true,
              data: { index: input.index, count: input.count },
            };
          }

          case 'delete_rows': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRangeByIndexes(input.index, 0, input.count, 1000);
            range.delete(Excel.DeleteShiftDirection.up);
            await context.sync();

            return {
              success: true,
              data: { index: input.index, count: input.count },
            };
          }

          case 'sort_range': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);
            const hasHeaders = input.hasHeaders !== false;
            const ascending = input.ascending !== false;

            range.load('columnIndex');
            await context.sync();

            const sortColumn = input.sortOn.charCodeAt(0) - 65;
            const sortFields = [{
              key: sortColumn,
              ascending: ascending,
            }];

            range.sort.apply(sortFields, hasHeaders);
            await context.sync();

            return {
              success: true,
              data: { range: input.range, sortedBy: input.sortOn },
            };
          }

          case 'create_pivot_table': {
            const sourceSheet = context.workbook.worksheets.getActiveWorksheet();
            const sourceRange = sourceSheet.getRange(input.sourceRange);

            // Create or get destination worksheet
            const destSheetName = input.destinationSheet || 'Pivot Table';
            let destSheet: Excel.Worksheet;

            try {
              destSheet = context.workbook.worksheets.getItem(destSheetName);
            } catch {
              destSheet = context.workbook.worksheets.add(destSheetName);
            }

            // Create the pivot table
            const pivotTable = context.workbook.pivotTables.add(
              'PivotTable1',
              sourceRange,
              destSheet.getRange('A1')
            );

            // Add row fields
            for (const field of input.rowFields || []) {
              const hierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(field));
              hierarchy.fields.getItem(field).sortByLabels(Excel.SortBy.ascending);
            }

            // Add column fields (optional)
            if (input.columnFields && input.columnFields.length > 0) {
              for (const field of input.columnFields) {
                pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(field));
              }
            }

            // Add data fields
            for (const dataField of input.dataFields || []) {
              const pivotField = pivotTable.dataHierarchies.add(
                pivotTable.hierarchies.getItem(dataField.field)
              );

              // Set aggregation function
              const summarizeBy = dataField.function || 'Sum';
              pivotField.summarizeBy = Excel.AggregationFunction[summarizeBy as keyof typeof Excel.AggregationFunction];
            }

            await context.sync();

            return {
              success: true,
              data: {
                pivotTableName: 'PivotTable1',
                destinationSheet: destSheetName,
                rowFields: input.rowFields,
                dataFields: input.dataFields,
              },
            };
          }

          case 'apply_conditional_formatting': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);

            if (input.type === 'cellValue' && input.rule) {
              const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
              const cellValueFormat = conditionalFormat.cellValue;

              // Map operator to Excel format
              const operatorMap: { [key: string]: any } = {
                greaterThan: Excel.ConditionalCellValueOperator.greaterThan,
                lessThan: Excel.ConditionalCellValueOperator.lessThan,
                equalTo: Excel.ConditionalCellValueOperator.equalTo,
                notEqualTo: Excel.ConditionalCellValueOperator.notEqualTo,
                between: Excel.ConditionalCellValueOperator.between,
              };

              cellValueFormat.rule = {
                formula1: input.rule.value?.toString(),
                operator: operatorMap[input.rule.operator || 'greaterThan'],
              };

              if (input.rule.color) {
                cellValueFormat.format.fill.color = input.rule.color;
              }
            } else if (input.type === 'colorScale') {
              range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
            } else if (input.type === 'dataBar') {
              range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
            } else if (input.type === 'iconSet') {
              range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
            }

            await context.sync();

            return {
              success: true,
              data: { range: input.range, type: input.type },
            };
          }

          case 'add_data_validation': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);

            if (input.type === 'list') {
              range.dataValidation.rule = {
                list: {
                  inCellDropDown: true,
                  source: input.source,
                },
              };
            } else if (input.type === 'wholeNumber' || input.type === 'decimal') {
              const operatorMap: { [key: string]: any } = {
                between: Excel.DataValidationOperator.between,
                notBetween: Excel.DataValidationOperator.notBetween,
                equalTo: Excel.DataValidationOperator.equalTo,
                notEqualTo: Excel.DataValidationOperator.notEqualTo,
                greaterThan: Excel.DataValidationOperator.greaterThan,
                lessThan: Excel.DataValidationOperator.lessThan,
              };

              range.dataValidation.rule = {
                wholeNumber: {
                  formula1: input.source,
                  operator: operatorMap[input.operator || 'greaterThan'],
                },
              };
            } else if (input.type === 'date') {
              range.dataValidation.rule = {
                date: {
                  formula1: input.source,
                  operator: Excel.DataValidationOperator.greaterThan,
                },
              };
            }

            if (input.errorMessage) {
              range.dataValidation.errorAlert = {
                message: input.errorMessage,
                showAlert: true,
                style: Excel.DataValidationAlertStyle.stop,
                title: 'Invalid Data',
              };
            }

            if (input.allowBlank !== undefined) {
              range.dataValidation.ignoreBlanks = input.allowBlank;
            }

            await context.sync();

            return {
              success: true,
              data: { range: input.range, type: input.type },
            };
          }

          case 'find_replace': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const searchRange = input.range ? sheet.getRange(input.range) : sheet.getUsedRange();

            // Load the range values
            searchRange.load('values, rowCount, columnCount, address');
            await context.sync();

            // Perform find and replace on the values
            let replacedCount = 0;
            const newValues = searchRange.values.map((row) =>
              row.map((cell) => {
                const cellValue = String(cell ?? '');
                let newValue = cellValue;

                if (input.matchEntireCell) {
                  // Match entire cell
                  if (input.matchCase) {
                    if (cellValue === input.find) {
                      newValue = input.replace;
                      replacedCount++;
                    }
                  } else {
                    if (cellValue.toLowerCase() === input.find.toLowerCase()) {
                      newValue = input.replace;
                      replacedCount++;
                    }
                  }
                } else {
                  // Match substring
                  if (input.matchCase) {
                    if (cellValue.includes(input.find)) {
                      newValue = cellValue.split(input.find).join(input.replace);
                      replacedCount++;
                    }
                  } else {
                    const regex = new RegExp(input.find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
                    if (regex.test(cellValue)) {
                      newValue = cellValue.replace(regex, input.replace);
                      replacedCount++;
                    }
                  }
                }

                return newValue;
              })
            );

            // Write the updated values back
            searchRange.values = newValues;
            await context.sync();

            return {
              success: true,
              data: {
                replaced: replacedCount,
                find: input.find,
                replace: input.replace,
                range: searchRange.address
              },
            };
          }

          case 'apply_autofilter': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);

            if (input.remove) {
              range.autoFilter.remove();
            } else {
              range.autoFilter.apply(range);
            }

            await context.sync();

            return {
              success: true,
              data: { range: input.range, action: input.remove ? 'removed' : 'applied' },
            };
          }

          case 'manage_worksheet': {
            if (input.action === 'create') {
              const newSheet = context.workbook.worksheets.add(input.name);
              newSheet.activate();
              await context.sync();

              return {
                success: true,
                data: { action: 'create', name: input.name },
              };
            } else if (input.action === 'delete') {
              const sheet = context.workbook.worksheets.getItem(input.name!);
              sheet.delete();
              await context.sync();

              return {
                success: true,
                data: { action: 'delete', name: input.name },
              };
            } else if (input.action === 'rename') {
              const sheet = context.workbook.worksheets.getItem(input.name!);
              sheet.name = input.newName!;
              await context.sync();

              return {
                success: true,
                data: { action: 'rename', oldName: input.name, newName: input.newName },
              };
            } else if (input.action === 'move') {
              const sheet = context.workbook.worksheets.getItem(input.name!);
              sheet.position = input.position!;
              await context.sync();

              return {
                success: true,
                data: { action: 'move', name: input.name, position: input.position },
              };
            }

            return {
              success: false,
              error: 'Invalid worksheet action',
            };
          }

          case 'clear_range': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);

            if (input.clearType === 'contents') {
              range.clear(Excel.ClearApplyTo.contents);
            } else if (input.clearType === 'formats') {
              range.clear(Excel.ClearApplyTo.formats);
            } else if (input.clearType === 'all') {
              range.clear(Excel.ClearApplyTo.all);
            }

            await context.sync();

            return {
              success: true,
              data: { range: input.range, clearType: input.clearType },
            };
          }

          case 'add_comment': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const cell = sheet.getRange(input.cell);

            const comment = cell.comment;
            comment.content = input.comment;
            if (input.author) {
              comment.authorName = input.author;
            }

            await context.sync();

            return {
              success: true,
              data: { cell: input.cell, comment: input.comment },
            };
          }

          case 'autofit_columns': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);

            if (input.direction === 'columns') {
              range.format.autofitColumns();
            } else {
              range.format.autofitRows();
            }

            await context.sync();

            return {
              success: true,
              data: { range: input.range, direction: input.direction },
            };
          }

          case 'create_named_range': {
            context.workbook.names.add(input.name, input.range);
            await context.sync();

            return {
              success: true,
              data: { name: input.name, range: input.range },
            };
          }

          case 'copy_range': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const sourceRange = sheet.getRange(input.sourceRange);
            const destRange = sheet.getRange(input.destinationRange);

            // Load dimensions to check if we're copying a single cell to multiple cells
            sourceRange.load('rowCount, columnCount');
            destRange.load('rowCount, columnCount');
            await context.sync();

            const isSingleCellSource = sourceRange.rowCount === 1 && sourceRange.columnCount === 1;
            const isMultiCellDest = destRange.rowCount > 1 || destRange.columnCount > 1;

            if (input.copyType === 'all' || !input.copyType) {
              destRange.copyFrom(sourceRange, Excel.RangeCopyType.all);
            } else if (input.copyType === 'values') {
              sourceRange.load('values');
              await context.sync();

              if (isSingleCellSource && isMultiCellDest) {
                // Fill entire destination with single source value
                const value = sourceRange.values[0][0];
                const filledValues: any[][] = [];
                for (let row = 0; row < destRange.rowCount; row++) {
                  const rowValues: any[] = [];
                  for (let col = 0; col < destRange.columnCount; col++) {
                    rowValues.push(value);
                  }
                  filledValues.push(rowValues);
                }
                destRange.values = filledValues;
              } else {
                destRange.values = sourceRange.values;
              }
            } else if (input.copyType === 'formulas') {
              sourceRange.load('formulas');
              await context.sync();

              if (isSingleCellSource && isMultiCellDest) {
                // Fill entire destination with formula (Excel auto-adjusts references)
                const formula = sourceRange.formulas[0][0];
                const filledFormulas: string[][] = [];
                for (let row = 0; row < destRange.rowCount; row++) {
                  const rowFormulas: string[] = [];
                  for (let col = 0; col < destRange.columnCount; col++) {
                    rowFormulas.push(formula);
                  }
                  filledFormulas.push(rowFormulas);
                }
                destRange.formulas = filledFormulas;
              } else {
                destRange.formulas = sourceRange.formulas;
              }
            } else if (input.copyType === 'formats') {
              destRange.copyFrom(sourceRange, Excel.RangeCopyType.formats);
            }

            await context.sync();

            return {
              success: true,
              data: {
                sourceRange: input.sourceRange,
                destinationRange: input.destinationRange,
                copyType: input.copyType,
                cellsCopied: destRange.rowCount * destRange.columnCount,
              },
            };
          }

          case 'apply_borders': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);

            const styleMap: { [key: string]: any } = {
              thin: Excel.BorderLineStyle.continuous,
              medium: Excel.BorderLineStyle.continuous,
              thick: Excel.BorderLineStyle.continuous,
              double: Excel.BorderLineStyle.double,
            };

            const weightMap: { [key: string]: any } = {
              thin: Excel.BorderWeight.thin,
              medium: Excel.BorderWeight.medium,
              thick: Excel.BorderWeight.thick,
              double: Excel.BorderWeight.medium,
            };

            const style = styleMap[input.style || 'thin'];
            const weight = weightMap[input.style || 'thin'];
            const color = input.color || '#000000';

            if (input.borderType === 'all') {
              range.format.borders.getItem('EdgeTop').style = style;
              range.format.borders.getItem('EdgeTop').weight = weight;
              range.format.borders.getItem('EdgeTop').color = color;

              range.format.borders.getItem('EdgeBottom').style = style;
              range.format.borders.getItem('EdgeBottom').weight = weight;
              range.format.borders.getItem('EdgeBottom').color = color;

              range.format.borders.getItem('EdgeLeft').style = style;
              range.format.borders.getItem('EdgeLeft').weight = weight;
              range.format.borders.getItem('EdgeLeft').color = color;

              range.format.borders.getItem('EdgeRight').style = style;
              range.format.borders.getItem('EdgeRight').weight = weight;
              range.format.borders.getItem('EdgeRight').color = color;

              range.format.borders.getItem('InsideHorizontal').style = style;
              range.format.borders.getItem('InsideHorizontal').weight = weight;
              range.format.borders.getItem('InsideHorizontal').color = color;

              range.format.borders.getItem('InsideVertical').style = style;
              range.format.borders.getItem('InsideVertical').weight = weight;
              range.format.borders.getItem('InsideVertical').color = color;
            } else if (input.borderType === 'outline') {
              range.format.borders.getItem('EdgeTop').style = style;
              range.format.borders.getItem('EdgeTop').weight = weight;
              range.format.borders.getItem('EdgeTop').color = color;

              range.format.borders.getItem('EdgeBottom').style = style;
              range.format.borders.getItem('EdgeBottom').weight = weight;
              range.format.borders.getItem('EdgeBottom').color = color;

              range.format.borders.getItem('EdgeLeft').style = style;
              range.format.borders.getItem('EdgeLeft').weight = weight;
              range.format.borders.getItem('EdgeLeft').color = color;

              range.format.borders.getItem('EdgeRight').style = style;
              range.format.borders.getItem('EdgeRight').weight = weight;
              range.format.borders.getItem('EdgeRight').color = color;
            } else {
              const borderNames: { [key: string]: string } = {
                top: 'EdgeTop',
                bottom: 'EdgeBottom',
                left: 'EdgeLeft',
                right: 'EdgeRight',
              };

              const borderName = borderNames[input.borderType];
              range.format.borders.getItem(borderName).style = style;
              range.format.borders.getItem(borderName).weight = weight;
              range.format.borders.getItem(borderName).color = color;
            }

            await context.sync();

            return {
              success: true,
              data: { range: input.range, borderType: input.borderType },
            };
          }

          case 'protect_range': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);

            range.format.protection.locked = input.protect;

            if (input.protect) {
              sheet.protection.protect();
            }

            await context.sync();

            return {
              success: true,
              data: { range: input.range, protected: input.protect },
            };
          }

          case 'freeze_panes': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();

            if (input.type === 'unfreeze') {
              sheet.freezePanes.unfreeze();
            } else if (input.type === 'rows' && input.cell) {
              const range = sheet.getRange(input.cell);
              sheet.freezePanes.freezeRows(range.rowIndex);
            } else if (input.type === 'columns' && input.cell) {
              const range = sheet.getRange(input.cell);
              sheet.freezePanes.freezeColumns(range.columnIndex);
            } else if (input.type === 'both' && input.cell) {
              const range = sheet.getRange(input.cell);
              sheet.freezePanes.freezeAt(range);
            }

            await context.sync();

            return {
              success: true,
              data: { type: input.type, cell: input.cell },
            };
          }

          case 'merge_cells': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);

            if (input.merge) {
              if (input.across) {
                range.merge(false); // Merge across (each row separately)
              } else {
                range.merge(true); // Merge all cells
              }
            } else {
              range.unmerge();
            }

            await context.sync();

            return {
              success: true,
              data: { range: input.range, merged: input.merge },
            };
          }

          case 'remove_duplicates': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);

            const hasHeaders = input.hasHeaders !== false;
            const columnIndices = input.columnIndices || [];

            // Use removeDuplicates API
            const result = range.removeDuplicates(columnIndices, hasHeaders);
            result.load('removed, uniqueRemaining');
            await context.sync();

            return {
              success: true,
              data: {
                range: input.range,
                removed: result.removed,
                remaining: result.uniqueRemaining,
              },
            };
          }

          case 'transpose_range': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const sourceRange = sheet.getRange(input.sourceRange);

            // Load source data
            sourceRange.load('values, rowCount, columnCount');
            await context.sync();

            // Transpose the values
            const transposed: any[][] = [];
            for (let col = 0; col < sourceRange.columnCount; col++) {
              const row: any[] = [];
              for (let rowIdx = 0; rowIdx < sourceRange.rowCount; rowIdx++) {
                row.push(sourceRange.values[rowIdx][col]);
              }
              transposed.push(row);
            }

            // Write transposed data to destination
            const destRange = sheet.getRange(input.destinationCell);
            const targetRange = destRange.getResizedRange(transposed.length - 1, transposed[0].length - 1);
            targetRange.values = transposed;

            await context.sync();

            return {
              success: true,
              data: {
                sourceRange: input.sourceRange,
                destinationCell: input.destinationCell,
                rows: transposed.length,
                columns: transposed[0].length,
              },
            };
          }

          case 'text_to_columns': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);

            // Load current values
            range.load('values, rowCount, columnCount');
            await context.sync();

            // Split each cell by delimiter
            const splitData: any[][] = [];
            let maxColumns = 0;

            for (let row = 0; row < range.rowCount; row++) {
              const rowData: any[] = [];
              for (let col = 0; col < range.columnCount; col++) {
                const cellValue = String(range.values[row][col] || '');
                const parts = cellValue.split(input.delimiter);
                rowData.push(...parts);
                maxColumns = Math.max(maxColumns, rowData.length);
              }
              splitData.push(rowData);
            }

            // Pad rows to same length
            splitData.forEach((row) => {
              while (row.length < maxColumns) {
                row.push('');
              }
            });

            // Write split data back
            const targetRange = range.getResizedRange(splitData.length - 1, maxColumns - 1);
            targetRange.values = splitData;

            await context.sync();

            return {
              success: true,
              data: {
                range: input.range,
                delimiter: input.delimiter,
                columns: maxColumns,
              },
            };
          }

          case 'hide_unhide': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);

            if (input.type === 'rows') {
              range.rowHidden = input.hide;
            } else if (input.type === 'columns') {
              range.columnHidden = input.hide;
            }

            await context.sync();

            return {
              success: true,
              data: { range: input.range, type: input.type, hidden: input.hide },
            };
          }

          case 'add_sparkline': {
            // Note: Excel JavaScript API doesn't directly support sparklines
            // We'll use a formula-based approach with SPARKLINE function (if available)
            // or create a small embedded chart as alternative

            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const targetCell = sheet.getRange(input.targetCell);

            // Try to use SPARKLINE function (available in some Excel versions)
            let formula = '';
            if (input.type === 'line') {
              formula = `=SPARKLINE(${input.dataRange})`;
            } else if (input.type === 'column') {
              formula = `=SPARKLINE(${input.dataRange},"column")`;
            } else if (input.type === 'winLoss') {
              formula = `=SPARKLINE(${input.dataRange},"winloss")`;
            }

            // Since SPARKLINE isn't universally available, we'll create a tiny embedded chart
            try {
              const dataRange = sheet.getRange(input.dataRange);
              const chartType =
                input.type === 'column' ? Excel.ChartType.columnClustered : Excel.ChartType.line;

              const chart = sheet.charts.add(chartType, dataRange, Excel.ChartSeriesBy.auto);
              chart.height = 50;
              chart.width = 100;
              chart.top = targetCell.top;
              chart.left = targetCell.left;
              chart.title.text = '';
              chart.legend.visible = false;

              await context.sync();

              return {
                success: true,
                data: { targetCell: input.targetCell, dataRange: input.dataRange, type: input.type },
              };
            } catch (error) {
              // Fallback: just add a note
              targetCell.values = [[`Sparkline: ${input.dataRange}`]];
              await context.sync();

              return {
                success: true,
                data: {
                  targetCell: input.targetCell,
                  note: 'Created reference (sparklines require chart objects)',
                },
              };
            }
          }

          case 'add_hyperlink': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const cell = sheet.getRange(input.cell);

            // Set display text
            if (input.displayText) {
              cell.values = [[input.displayText]];
            }

            // Add hyperlink
            cell.hyperlink = {
              address: input.url,
              textToDisplay: input.displayText || input.url,
            };

            await context.sync();

            return {
              success: true,
              data: { cell: input.cell, url: input.url },
            };
          }

          case 'calculate_statistics': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);

            // Load values
            range.load('values');
            await context.sync();

            // Calculate statistics
            const values: number[] = [];
            range.values.forEach((row) => {
              row.forEach((cell) => {
                if (typeof cell === 'number') {
                  values.push(cell);
                }
              });
            });

            if (values.length === 0) {
              return {
                success: false,
                error: 'No numeric values found in range',
              };
            }

            values.sort((a, b) => a - b);

            const sum = values.reduce((a, b) => a + b, 0);
            const mean = sum / values.length;
            const min = values[0];
            const max = values[values.length - 1];

            // Median
            const mid = Math.floor(values.length / 2);
            const median = values.length % 2 === 0 ? (values[mid - 1] + values[mid]) / 2 : values[mid];

            // Standard deviation
            const variance = values.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / values.length;
            const stdDev = Math.sqrt(variance);

            return {
              success: true,
              data: {
                range: input.range,
                count: values.length,
                sum,
                mean,
                median,
                min,
                max,
                stdDev,
              },
            };
          }

          case 'set_alignment': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);

            if (input.horizontal) {
              const alignmentMap: { [key: string]: any } = {
                left: Excel.HorizontalAlignment.left,
                center: Excel.HorizontalAlignment.center,
                right: Excel.HorizontalAlignment.right,
                justify: Excel.HorizontalAlignment.justify,
              };
              range.format.horizontalAlignment = alignmentMap[input.horizontal];
            }

            if (input.vertical) {
              const alignmentMap: { [key: string]: any } = {
                top: Excel.VerticalAlignment.top,
                middle: Excel.VerticalAlignment.center,
                bottom: Excel.VerticalAlignment.bottom,
              };
              range.format.verticalAlignment = alignmentMap[input.vertical];
            }

            if (input.wrapText !== undefined) {
              range.format.wrapText = input.wrapText;
            }

            if (input.indent !== undefined) {
              range.format.indentLevel = input.indent;
            }

            await context.sync();

            return {
              success: true,
              data: { range: input.range, alignment: { horizontal: input.horizontal, vertical: input.vertical } },
            };
          }

          case 'check_duplicates': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);
            range.load('values');
            await context.sync();

            const duplicates: any[] = [];
            const toleranceDays = input.toleranceDays || 0;
            const newDate = new Date(input.newEntry.date);
            const newAmount = input.newEntry.amount;
            const newMerchant = (input.newEntry.merchant || '').toLowerCase();

            for (let row = 0; row < range.values.length; row++) {
              const rowData = range.values[row];
              const existingDate = new Date(rowData[input.dateColumn]);
              const existingAmount = rowData[input.amountColumn];
              const existingMerchant = String(rowData[input.merchantColumn] || '').toLowerCase();

              // Check date tolerance
              const daysDiff = Math.abs((newDate.getTime() - existingDate.getTime()) / (1000 * 60 * 60 * 24));
              const dateMatch = daysDiff <= toleranceDays;

              // Check amount match (exact)
              const amountMatch = Math.abs(existingAmount - newAmount) < 0.01;

              // Check merchant similarity (partial match)
              const merchantMatch = newMerchant.includes(existingMerchant) || existingMerchant.includes(newMerchant);

              if (dateMatch && amountMatch && merchantMatch) {
                duplicates.push({
                  row: row + 2, // +2 for 1-based indexing and header row
                  date: existingDate.toISOString().split('T')[0],
                  merchant: rowData[input.merchantColumn],
                  amount: existingAmount,
                });
              }
            }

            return {
              success: true,
              data: {
                hasDuplicates: duplicates.length > 0,
                duplicates,
                message:
                  duplicates.length > 0
                    ? `Found ${duplicates.length} potential duplicate(s)`
                    : 'No duplicates found',
              },
            };
          }

          case 'convert_currency': {
            // Calculate converted amount
            const convertedAmount = input.amount * input.exchangeRate;

            return {
              success: true,
              data: {
                originalAmount: input.amount,
                originalCurrency: input.fromCurrency,
                convertedAmount: Math.round(convertedAmount * 100) / 100,
                targetCurrency: input.toCurrency,
                exchangeRate: input.exchangeRate,
                date: input.date,
                formatted: `${input.fromCurrency} ${input.amount} = ${input.toCurrency} ${Math.round(convertedAmount * 100) / 100} (Rate: ${input.exchangeRate})`,
              },
            };
          }

          case 'generate_expense_summary': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const dataRange = sheet.getRange(input.dataRange);
            dataRange.load('values');
            await context.sync();

            // Calculate statistics
            let total = 0;
            const categoryTotals: { [key: string]: number } = {};
            const dates: Date[] = [];
            let count = 0;

            for (let row = 1; row < dataRange.values.length; row++) {
              // Skip header
              const rowData = dataRange.values[row];
              const amount = rowData[input.amountColumn];
              const date = new Date(rowData[input.dateColumn]);

              if (typeof amount === 'number') {
                total += amount;
                count++;
                dates.push(date);

                if (input.categoryColumn !== undefined) {
                  const category = String(rowData[input.categoryColumn] || 'Uncategorized');
                  categoryTotals[category] = (categoryTotals[category] || 0) + amount;
                }
              }
            }

            // Sort dates
            dates.sort((a, b) => a.getTime() - b.getTime());
            const startDate = dates[0]?.toISOString().split('T')[0] || 'N/A';
            const endDate = dates[dates.length - 1]?.toISOString().split('T')[0] || 'N/A';
            const average = count > 0 ? total / count : 0;

            // Build summary data
            const summaryData: any[][] = [
              ['EXPENSE SUMMARY REPORT', ''],
              ['', ''],
              ['Date Range', `${startDate} to ${endDate}`],
              ['Total Expenses', total],
              ['Number of Expenses', count],
              ['Average Expense', Math.round(average * 100) / 100],
              ['', ''],
            ];

            if (Object.keys(categoryTotals).length > 0) {
              summaryData.push(['BREAKDOWN BY CATEGORY', '']);
              for (const [category, amount] of Object.entries(categoryTotals)) {
                summaryData.push([category, Math.round(amount * 100) / 100]);
              }
            }

            // Write summary to sheet
            const outputRange = sheet.getRange(input.outputCell);
            const targetRange = outputRange.getResizedRange(
              summaryData.length - 1,
              summaryData[0].length - 1
            );
            targetRange.values = summaryData;

            // Format the summary
            const titleCell = sheet.getRange(input.outputCell);
            titleCell.format.font.bold = true;
            titleCell.format.font.size = 14;
            titleCell.format.fill.color = '#4472C4';
            titleCell.format.font.color = 'white';

            await context.sync();

            return {
              success: true,
              data: {
                total,
                count,
                average: Math.round(average * 100) / 100,
                dateRange: { start: startDate, end: endDate },
                categories: categoryTotals,
                summaryLocation: input.outputCell,
              },
            };
          }

          case 'export_to_csv': {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(input.range);
            range.load('values');
            await context.sync();

            const includeHeaders = input.includeHeaders !== false;
            const rows = range.values;
            let csv = '';

            for (let i = 0; i < rows.length; i++) {
              if (i === 0 && !includeHeaders) continue;

              const row = rows[i];
              const csvRow = row
                .map((cell: any) => {
                  // Escape quotes and wrap in quotes if contains comma, quote, or newline
                  const cellStr = String(cell ?? '');
                  if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
                    return `"${cellStr.replace(/"/g, '""')}"`;
                  }
                  return cellStr;
                })
                .join(',');

              csv += csvRow + '\n';
            }

            return {
              success: true,
              data: {
                csv,
                rowCount: rows.length - (includeHeaders ? 0 : 1),
                message: `Exported ${rows.length} rows to CSV format. Copy the CSV data from the response.`,
              },
            };
          }

          default:
            return {
              success: false,
              error: `Unknown tool: ${toolName}`,
            };
        }
      });
    } catch (error: any) {
      console.error(`Error executing tool ${toolName}:`, error);
      return {
        success: false,
        error: error.message || 'An error occurred',
      };
    }
  }, []);

  return {
    tools: excelTools,
    executeTool,
  };
}
