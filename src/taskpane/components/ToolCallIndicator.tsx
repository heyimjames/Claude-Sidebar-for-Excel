import React from 'react';
import {
  Globe24Regular,
  DataPie24Regular,
  ArrowSwap24Regular,
  DocumentCopy24Regular,
  Checkmark24Regular,
  TableSimpleMultiple24Regular,
} from '@fluentui/react-icons';
import '../styles/tool-indicator.css';

export interface ToolCall {
  id: string;
  name: string;
  status: 'running' | 'completed';
}

interface ToolCallIndicatorProps {
  toolCalls: ToolCall[];
}

const TOOL_CONFIG: {
  [key: string]: {
    icon: React.ComponentType<any>;
    label: string;
    color: string;
  };
} = {
  web_search: {
    icon: Globe24Regular,
    label: 'Searching the web',
    color: '#0078D4',
  },
  check_duplicates: {
    icon: DocumentCopy24Regular,
    label: 'Checking for duplicates',
    color: '#D83B01',
  },
  convert_currency: {
    icon: ArrowSwap24Regular,
    label: 'Converting currency',
    color: '#107C10',
  },
  generate_expense_summary: {
    icon: DataPie24Regular,
    label: 'Generating summary',
    color: '#5C2D91',
  },
  export_to_csv: {
    icon: TableSimpleMultiple24Regular,
    label: 'Exporting data',
    color: '#008272',
  },
  // Excel tools
  read_range: {
    icon: TableSimpleMultiple24Regular,
    label: 'Reading data',
    color: '#0078D4',
  },
  write_range: {
    icon: TableSimpleMultiple24Regular,
    label: 'Writing data',
    color: '#107C10',
  },
  add_row: {
    icon: TableSimpleMultiple24Regular,
    label: 'Adding row',
    color: '#107C10',
  },
  delete_row: {
    icon: TableSimpleMultiple24Regular,
    label: 'Deleting row',
    color: '#D83B01',
  },
  format_cells: {
    icon: TableSimpleMultiple24Regular,
    label: 'Formatting cells',
    color: '#5C2D91',
  },
  create_chart: {
    icon: DataPie24Regular,
    label: 'Creating chart',
    color: '#0078D4',
  },
  apply_formula: {
    icon: TableSimpleMultiple24Regular,
    label: 'Applying formula',
    color: '#008272',
  },
};

export function ToolCallIndicator({ toolCalls }: ToolCallIndicatorProps) {
  const activeTools = toolCalls.filter((t) => t.status === 'running');

  if (activeTools.length === 0) return null;

  return (
    <div className="tool-indicator-container">
      {activeTools.map((tool) => {
        const config = TOOL_CONFIG[tool.name] || {
          icon: TableSimpleMultiple24Regular,
          label: `Using ${tool.name}`,
          color: '#605E5C',
        };
        const IconComponent = config.icon;

        return (
          <div key={tool.id} className="tool-indicator" style={{ '--tool-color': config.color } as any}>
            <div className="tool-icon shimmer">
              <IconComponent />
            </div>
            <span className="tool-label">{config.label}...</span>
          </div>
        );
      })}
    </div>
  );
}
