import type { ExcelTool } from './types';

export const excelTools: ExcelTool[] = [
  {
    name: 'web_search',
    description: 'Search the web for current information, historical data, exchange rates, prices, or any real-time information. Use this when you need up-to-date information that you don\'t have in your training data, such as current exchange rates, recent events, pricing information, or to verify facts.',
    input_schema: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'The search query. Be specific and include dates when searching for historical information (e.g., "EUR to USD exchange rate on 2024-01-15")',
        },
      },
      required: ['query'],
    },
  },
  {
    name: 'read_range',
    description: 'Read values from a range of cells in Excel. Returns the values, formulas, and number formats.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Cell range in A1 notation (e.g., "A1:B10" or "A1")',
        },
        worksheet: {
          type: 'string',
          description: 'Worksheet name (optional, uses active sheet if not specified)',
        },
      },
      required: ['range'],
    },
  },
  {
    name: 'write_range',
    description: 'Write values to a range of cells in Excel. Can write text, numbers, or formulas.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Starting cell in A1 notation (e.g., "A1")',
        },
        values: {
          type: 'array',
          description: '2D array of values to write. Each row is an array of cell values.',
          items: {
            type: 'array',
          },
        },
        worksheet: {
          type: 'string',
          description: 'Worksheet name (optional, uses active sheet if not specified)',
        },
      },
      required: ['range', 'values'],
    },
  },
  {
    name: 'get_selection',
    description: 'Get the currently selected cells in Excel, including their address, values, and formulas.',
    input_schema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'get_workbook_info',
    description: 'Get information about the current workbook, including all worksheet names and the active sheet.',
    input_schema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'create_table',
    description: 'Create a formatted table from a range of data in Excel.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Cell range for the table (e.g., "A1:C10")',
        },
        tableName: {
          type: 'string',
          description: 'Name for the table',
        },
        hasHeaders: {
          type: 'boolean',
          description: 'Whether the first row contains headers (default: true)',
        },
      },
      required: ['range'],
    },
  },
  {
    name: 'create_chart',
    description: 'Create a chart from data in Excel.',
    input_schema: {
      type: 'object',
      properties: {
        dataRange: {
          type: 'string',
          description: 'Range containing the data to chart',
        },
        chartType: {
          type: 'string',
          description: 'Type of chart to create',
          enum: ['ColumnClustered', 'ColumnStacked', 'Line', 'LineMarkers', 'Pie', 'BarClustered', 'Area', 'XYScatter'],
        },
        title: {
          type: 'string',
          description: 'Chart title',
        },
      },
      required: ['dataRange', 'chartType'],
    },
  },
  {
    name: 'apply_formula',
    description: 'Apply a formula to a cell or range in Excel.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Cell or range to apply the formula to',
        },
        formula: {
          type: 'string',
          description: 'Excel formula (e.g., "=SUM(A1:A10)")',
        },
      },
      required: ['range', 'formula'],
    },
  },
  {
    name: 'format_range',
    description: 'Apply formatting to a range of cells.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Cell range to format',
        },
        format: {
          type: 'object',
          description: 'Formatting options',
          properties: {
            numberFormat: {
              type: 'string',
              description: 'Number format (e.g., "0.00", "$#,##0.00", "m/d/yyyy")',
            },
            fontBold: {
              type: 'boolean',
              description: 'Make text bold',
            },
            fontSize: {
              type: 'number',
              description: 'Font size in points',
            },
            fillColor: {
              type: 'string',
              description: 'Background color (e.g., "#FF0000" or "red")',
            },
            fontColor: {
              type: 'string',
              description: 'Text color',
            },
          },
        },
      },
      required: ['range', 'format'],
    },
  },
  {
    name: 'insert_rows',
    description: 'Insert new rows into the worksheet.',
    input_schema: {
      type: 'object',
      properties: {
        index: {
          type: 'number',
          description: 'Row index where to insert (0-based)',
        },
        count: {
          type: 'number',
          description: 'Number of rows to insert',
        },
      },
      required: ['index', 'count'],
    },
  },
  {
    name: 'delete_rows',
    description: 'Delete rows from the worksheet.',
    input_schema: {
      type: 'object',
      properties: {
        index: {
          type: 'number',
          description: 'Starting row index (0-based)',
        },
        count: {
          type: 'number',
          description: 'Number of rows to delete',
        },
      },
      required: ['index', 'count'],
    },
  },
  {
    name: 'sort_range',
    description: 'Sort a range of cells by one or more columns.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to sort',
        },
        sortOn: {
          type: 'string',
          description: 'Column letter to sort by (e.g., "A")',
        },
        ascending: {
          type: 'boolean',
          description: 'Sort in ascending order (default: true)',
        },
        hasHeaders: {
          type: 'boolean',
          description: 'Whether the range has headers (default: true)',
        },
      },
      required: ['range', 'sortOn'],
    },
  },
  {
    name: 'create_pivot_table',
    description: 'Create a pivot table from a data range in Excel. The pivot table will be placed on a new worksheet.',
    input_schema: {
      type: 'object',
      properties: {
        sourceRange: {
          type: 'string',
          description: 'Source data range for the pivot table (e.g., "A1:D100")',
        },
        destinationSheet: {
          type: 'string',
          description: 'Name for the new worksheet where the pivot table will be created (optional)',
        },
        rowFields: {
          type: 'array',
          description: 'Column names or indices to use as row fields',
          items: {
            type: 'string',
          },
        },
        columnFields: {
          type: 'array',
          description: 'Column names or indices to use as column fields (optional)',
          items: {
            type: 'string',
          },
        },
        dataFields: {
          type: 'array',
          description: 'Fields to aggregate in the values area',
          items: {
            type: 'object',
            properties: {
              field: {
                type: 'string',
                description: 'Column name or index',
              },
              function: {
                type: 'string',
                description: 'Aggregation function',
                enum: ['Sum', 'Count', 'Average', 'Min', 'Max'],
              },
            },
          },
        },
      },
      required: ['sourceRange', 'rowFields', 'dataFields'],
    },
  },
  {
    name: 'apply_conditional_formatting',
    description: 'Apply conditional formatting to a range based on rules (color scales, data bars, or custom rules).',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Cell range to format',
        },
        type: {
          type: 'string',
          description: 'Type of conditional formatting',
          enum: ['colorScale', 'dataBar', 'cellValue', 'iconSet'],
        },
        rule: {
          type: 'object',
          description: 'Formatting rule configuration',
          properties: {
            operator: {
              type: 'string',
              description: 'Comparison operator for cellValue type',
              enum: ['greaterThan', 'lessThan', 'between', 'equalTo', 'notEqualTo'],
            },
            value: {
              type: 'number',
              description: 'Value to compare against',
            },
            color: {
              type: 'string',
              description: 'Color to apply (hex code)',
            },
          },
        },
      },
      required: ['range', 'type'],
    },
  },
  {
    name: 'add_data_validation',
    description: 'Add data validation rules to cells (dropdown lists, date validation, number ranges).',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Cell range to validate',
        },
        type: {
          type: 'string',
          description: 'Type of validation',
          enum: ['list', 'wholeNumber', 'decimal', 'date', 'textLength'],
        },
        operator: {
          type: 'string',
          description: 'Comparison operator',
          enum: ['between', 'notBetween', 'equalTo', 'notEqualTo', 'greaterThan', 'lessThan'],
        },
        source: {
          type: 'string',
          description: 'For list type: comma-separated values or range. For others: comparison value',
        },
        allowBlank: {
          type: 'boolean',
          description: 'Allow blank cells (default: true)',
        },
        errorMessage: {
          type: 'string',
          description: 'Error message to show when validation fails',
        },
      },
      required: ['range', 'type'],
    },
  },
  {
    name: 'find_replace',
    description: 'Find and replace text or values in a range. Useful for formatting changes like converting commas to periods in numbers (e.g., "23,6" to "23.6") or vice versa, standardizing text formats, or bulk text corrections.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to search in (optional, searches active sheet if not specified)',
        },
        find: {
          type: 'string',
          description: 'Text to find',
        },
        replace: {
          type: 'string',
          description: 'Text to replace with',
        },
        matchCase: {
          type: 'boolean',
          description: 'Match case (default: false)',
        },
        matchEntireCell: {
          type: 'boolean',
          description: 'Match entire cell contents (default: false)',
        },
      },
      required: ['find', 'replace'],
    },
  },
  {
    name: 'apply_autofilter',
    description: 'Apply or remove AutoFilter on a range to enable filtering.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to apply filter to',
        },
        remove: {
          type: 'boolean',
          description: 'Remove existing filter (default: false)',
        },
      },
      required: ['range'],
    },
  },
  {
    name: 'manage_worksheet',
    description: 'Create, delete, rename, or reorder worksheets.',
    input_schema: {
      type: 'object',
      properties: {
        action: {
          type: 'string',
          description: 'Action to perform',
          enum: ['create', 'delete', 'rename', 'move'],
        },
        name: {
          type: 'string',
          description: 'Worksheet name (for create/delete/rename)',
        },
        newName: {
          type: 'string',
          description: 'New name (for rename action)',
        },
        position: {
          type: 'number',
          description: 'Position to move to (for move action, 0-based)',
        },
      },
      required: ['action'],
    },
  },
  {
    name: 'clear_range',
    description: 'Clear contents, formatting, or both from a range.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to clear',
        },
        clearType: {
          type: 'string',
          description: 'What to clear',
          enum: ['contents', 'formats', 'all'],
        },
      },
      required: ['range', 'clearType'],
    },
  },
  {
    name: 'add_comment',
    description: 'Add a comment or note to a cell.',
    input_schema: {
      type: 'object',
      properties: {
        cell: {
          type: 'string',
          description: 'Cell address (e.g., "A1")',
        },
        comment: {
          type: 'string',
          description: 'Comment text',
        },
        author: {
          type: 'string',
          description: 'Comment author name (optional)',
        },
      },
      required: ['cell', 'comment'],
    },
  },
  {
    name: 'autofit_columns',
    description: 'Automatically resize columns or rows to fit content.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to autofit',
        },
        direction: {
          type: 'string',
          description: 'Resize columns or rows',
          enum: ['columns', 'rows'],
        },
      },
      required: ['range', 'direction'],
    },
  },
  {
    name: 'create_named_range',
    description: 'Create a named range for easier formula references.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to name',
        },
        name: {
          type: 'string',
          description: 'Name for the range (must start with letter, no spaces)',
        },
      },
      required: ['range', 'name'],
    },
  },
  {
    name: 'copy_range',
    description: 'Copy data, formulas, or formatting from one range to another.',
    input_schema: {
      type: 'object',
      properties: {
        sourceRange: {
          type: 'string',
          description: 'Source range to copy from',
        },
        destinationRange: {
          type: 'string',
          description: 'Destination range to paste to',
        },
        copyType: {
          type: 'string',
          description: 'What to copy',
          enum: ['all', 'values', 'formulas', 'formats'],
        },
      },
      required: ['sourceRange', 'destinationRange'],
    },
  },
  {
    name: 'apply_borders',
    description: 'Apply borders to cells.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to apply borders to',
        },
        borderType: {
          type: 'string',
          description: 'Type of border',
          enum: ['all', 'outline', 'top', 'bottom', 'left', 'right'],
        },
        style: {
          type: 'string',
          description: 'Border style',
          enum: ['thin', 'medium', 'thick', 'double'],
        },
        color: {
          type: 'string',
          description: 'Border color (hex code, default: black)',
        },
      },
      required: ['range', 'borderType'],
    },
  },
  {
    name: 'protect_range',
    description: 'Protect or unprotect a range from editing.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to protect/unprotect',
        },
        protect: {
          type: 'boolean',
          description: 'True to protect, false to unprotect',
        },
      },
      required: ['range', 'protect'],
    },
  },
  {
    name: 'freeze_panes',
    description: 'Freeze rows and/or columns to keep them visible while scrolling.',
    input_schema: {
      type: 'object',
      properties: {
        cell: {
          type: 'string',
          description: 'Cell where freeze starts (e.g., "B2" freezes first row and first column)',
        },
        type: {
          type: 'string',
          description: 'Type of freeze',
          enum: ['rows', 'columns', 'both', 'unfreeze'],
        },
      },
      required: ['type'],
    },
  },
  {
    name: 'merge_cells',
    description: 'Merge or unmerge cells in a range.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to merge/unmerge',
        },
        merge: {
          type: 'boolean',
          description: 'True to merge, false to unmerge',
        },
        across: {
          type: 'boolean',
          description: 'If true, merge each row separately (default: false)',
        },
      },
      required: ['range', 'merge'],
    },
  },
  {
    name: 'remove_duplicates',
    description: 'Remove duplicate rows from a range based on specified columns.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to check for duplicates',
        },
        columnIndices: {
          type: 'array',
          description: 'Column indices to check (0-based). If not specified, checks all columns.',
          items: {
            type: 'number',
          },
        },
        hasHeaders: {
          type: 'boolean',
          description: 'Whether first row contains headers (default: true)',
        },
      },
      required: ['range'],
    },
  },
  {
    name: 'transpose_range',
    description: 'Transpose a range (flip rows to columns and vice versa).',
    input_schema: {
      type: 'object',
      properties: {
        sourceRange: {
          type: 'string',
          description: 'Source range to transpose',
        },
        destinationCell: {
          type: 'string',
          description: 'Top-left cell where transposed data should be placed',
        },
      },
      required: ['sourceRange', 'destinationCell'],
    },
  },
  {
    name: 'text_to_columns',
    description: 'Split text in cells by a delimiter into multiple columns.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range containing text to split',
        },
        delimiter: {
          type: 'string',
          description: 'Delimiter to split by (e.g., ",", " ", "|")',
        },
      },
      required: ['range', 'delimiter'],
    },
  },
  {
    name: 'hide_unhide',
    description: 'Hide or unhide rows or columns.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range of rows or columns to hide/unhide',
        },
        type: {
          type: 'string',
          description: 'What to hide/unhide',
          enum: ['rows', 'columns'],
        },
        hide: {
          type: 'boolean',
          description: 'True to hide, false to unhide',
        },
      },
      required: ['range', 'type', 'hide'],
    },
  },
  {
    name: 'add_sparkline',
    description: 'Add a sparkline (mini chart) to a cell.',
    input_schema: {
      type: 'object',
      properties: {
        dataRange: {
          type: 'string',
          description: 'Range containing data for the sparkline',
        },
        targetCell: {
          type: 'string',
          description: 'Cell where sparkline should be placed',
        },
        type: {
          type: 'string',
          description: 'Type of sparkline',
          enum: ['line', 'column', 'winLoss'],
        },
      },
      required: ['dataRange', 'targetCell', 'type'],
    },
  },
  {
    name: 'add_hyperlink',
    description: 'Add a hyperlink to a cell.',
    input_schema: {
      type: 'object',
      properties: {
        cell: {
          type: 'string',
          description: 'Cell to add hyperlink to',
        },
        url: {
          type: 'string',
          description: 'URL or email address (use mailto: for email)',
        },
        displayText: {
          type: 'string',
          description: 'Text to display (optional, uses URL if not specified)',
        },
      },
      required: ['cell', 'url'],
    },
  },
  {
    name: 'calculate_statistics',
    description: 'Calculate statistics (mean, median, min, max, std dev, count) for a range.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to analyze',
        },
      },
      required: ['range'],
    },
  },
  {
    name: 'set_alignment',
    description: 'Set text alignment for a range.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to align',
        },
        horizontal: {
          type: 'string',
          description: 'Horizontal alignment',
          enum: ['left', 'center', 'right', 'justify'],
        },
        vertical: {
          type: 'string',
          description: 'Vertical alignment',
          enum: ['top', 'middle', 'bottom'],
        },
        wrapText: {
          type: 'boolean',
          description: 'Enable text wrapping',
        },
        indent: {
          type: 'number',
          description: 'Indentation level (0-15)',
        },
      },
      required: ['range'],
    },
  },
  {
    name: 'check_duplicates',
    description: 'Check for potential duplicate expenses before adding new entries. Scans existing data for similar entries based on date, merchant/description, and amount. Returns warnings if potential duplicates are found.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to check for duplicates (e.g., "A2:E100")',
        },
        newEntry: {
          type: 'object',
          description: 'The new entry to check against existing data',
          properties: {
            date: {
              type: 'string',
              description: 'Date of the expense',
            },
            merchant: {
              type: 'string',
              description: 'Merchant or description',
            },
            amount: {
              type: 'number',
              description: 'Amount of the expense',
            },
          },
        },
        dateColumn: {
          type: 'number',
          description: 'Column index for date (0-based, e.g., 0 for column A)',
        },
        merchantColumn: {
          type: 'number',
          description: 'Column index for merchant/description',
        },
        amountColumn: {
          type: 'number',
          description: 'Column index for amount',
        },
        toleranceDays: {
          type: 'number',
          description: 'Number of days tolerance for date matching (default: 0)',
        },
      },
      required: ['range', 'newEntry', 'dateColumn', 'merchantColumn', 'amountColumn'],
    },
  },
  {
    name: 'convert_currency',
    description: 'Convert an amount from one currency to another using current or historical exchange rates. Use web_search first to get the exchange rate, then use this to perform the conversion and format the result.',
    input_schema: {
      type: 'object',
      properties: {
        amount: {
          type: 'number',
          description: 'Amount to convert',
        },
        fromCurrency: {
          type: 'string',
          description: 'Source currency code (e.g., "EUR", "GBP", "JPY")',
        },
        toCurrency: {
          type: 'string',
          description: 'Target currency code (e.g., "USD")',
        },
        exchangeRate: {
          type: 'number',
          description: 'Exchange rate to use for conversion',
        },
        date: {
          type: 'string',
          description: 'Date for the exchange rate (for record keeping)',
        },
      },
      required: ['amount', 'fromCurrency', 'toCurrency', 'exchangeRate'],
    },
  },
  {
    name: 'generate_expense_summary',
    description: 'Generate a comprehensive summary report of expenses from a data range. Creates totals, category breakdowns, date ranges, and statistics.',
    input_schema: {
      type: 'object',
      properties: {
        dataRange: {
          type: 'string',
          description: 'Range containing expense data to summarize',
        },
        dateColumn: {
          type: 'number',
          description: 'Column index for dates (0-based)',
        },
        amountColumn: {
          type: 'number',
          description: 'Column index for amounts (0-based)',
        },
        categoryColumn: {
          type: 'number',
          description: 'Column index for categories (0-based, optional)',
        },
        outputCell: {
          type: 'string',
          description: 'Starting cell for the summary report (e.g., "A1")',
        },
        summaryType: {
          type: 'string',
          description: 'Type of summary to generate',
          enum: ['monthly', 'quarterly', 'yearly', 'all'],
        },
      },
      required: ['dataRange', 'dateColumn', 'amountColumn', 'outputCell'],
    },
  },
  {
    name: 'export_to_csv',
    description: 'Export a range of data to CSV format. Returns the CSV content as text that can be saved or copied.',
    input_schema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range to export (e.g., "A1:E100")',
        },
        includeHeaders: {
          type: 'boolean',
          description: 'Whether to include the first row as headers (default: true)',
        },
      },
      required: ['range'],
    },
  },
];
