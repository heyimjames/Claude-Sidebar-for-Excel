export interface Command {
  id: string;
  label: string;
  description: string;
  icon: string;
  template: string; // The text to insert
  category: 'analysis' | 'chart' | 'format' | 'data';
}

export const commands: Command[] = [
  {
    id: 'analyze',
    label: '/analyze',
    description: 'Analyze selected data',
    icon: '',
    template: 'Analyze the data in ',
    category: 'analysis'
  },
  {
    id: 'summarize',
    label: '/summarize',
    description: 'Summarize data',
    icon: '',
    template: 'Summarize the data in ',
    category: 'analysis'
  },
  {
    id: 'explain',
    label: '/explain',
    description: 'Explain formulas or data',
    icon: '',
    template: 'Explain ',
    category: 'analysis'
  },
  {
    id: 'chart',
    label: '/chart',
    description: 'Create a chart from data',
    icon: '',
    template: 'Create a chart from ',
    category: 'chart'
  },
  {
    id: 'format',
    label: '/format',
    description: 'Format cells',
    icon: '',
    template: 'Format the range ',
    category: 'format'
  },
  {
    id: 'conditional',
    label: '/conditional',
    description: 'Apply conditional formatting',
    icon: '',
    template: 'Apply conditional formatting to ',
    category: 'format'
  },
  {
    id: 'formula',
    label: '/formula',
    description: 'Create a formula',
    icon: '',
    template: 'Create a formula that ',
    category: 'data'
  },
  {
    id: 'vlookup',
    label: '/vlookup',
    description: 'Create VLOOKUP formula',
    icon: '',
    template: 'Create a VLOOKUP formula to ',
    category: 'data'
  },
  {
    id: 'table',
    label: '/table',
    description: 'Convert range to table',
    icon: '',
    template: 'Convert the range to a table in ',
    category: 'data'
  },
  {
    id: 'pivot',
    label: '/pivot',
    description: 'Create pivot table',
    icon: '',
    template: 'Create a pivot table from ',
    category: 'data'
  },
  {
    id: 'sort',
    label: '/sort',
    description: 'Sort data',
    icon: '',
    template: 'Sort the data in ',
    category: 'data'
  },
  {
    id: 'filter',
    label: '/filter',
    description: 'Filter data',
    icon: '',
    template: 'Filter the data in ',
    category: 'data'
  },
  {
    id: 'transpose',
    label: '/transpose',
    description: 'Transpose rows and columns',
    icon: '',
    template: 'Transpose the data in ',
    category: 'data'
  },
  {
    id: 'duplicate',
    label: '/duplicate',
    description: 'Find or remove duplicates',
    icon: '',
    template: 'Find duplicates in ',
    category: 'data'
  },
  {
    id: 'clean',
    label: '/clean',
    description: 'Clean and normalize data',
    icon: '',
    template: 'Clean the data in ',
    category: 'data'
  },
  {
    id: 'calculate',
    label: '/calculate',
    description: 'Perform calculations',
    icon: '',
    template: 'Calculate ',
    category: 'data'
  },
  {
    id: 'validate',
    label: '/validate',
    description: 'Add data validation',
    icon: '',
    template: 'Add data validation to ',
    category: 'data'
  },
  {
    id: 'merge',
    label: '/merge',
    description: 'Merge cells',
    icon: '',
    template: 'Merge cells in ',
    category: 'format'
  }
];
