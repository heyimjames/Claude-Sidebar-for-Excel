import { useState, useEffect } from 'react';

const EXAMPLE_PROMPTS = [
  'Sum all values in column B',
  'Create a pie chart from this data',
  'Format this table professionally',
  'Find duplicates in this selection',
  'Calculate the average of these numbers',
  'Sort this data by date',
  'Remove empty rows',
  'Convert these currencies to USD',
  'Highlight cells greater than 100',
  'Create a monthly sales report',
  'Analyze trends in this data',
  'Generate a pivot table',
  'Find the top 10 values',
  'Calculate year-over-year growth',
  'Merge cells with same values',
  'Split names into first and last',
  'Count unique values',
  'Create a line chart showing trends',
  'Format numbers as currency',
  'Add percentage formatting',
  'Calculate the median',
  'Find cells containing errors',
  'Create conditional formatting rules',
  'Extract dates from text',
  'Generate a summary table',
  'Compare two columns for differences',
  'Create a bar chart',
  'Calculate running total',
  'Filter data by criteria',
  'Add borders to this table',
  'Create a heat map',
  'Calculate compound annual growth rate',
  'Find and replace text',
  'Generate random numbers',
  'Calculate standard deviation',
];

export function useAnimatedPlaceholder() {
  const [placeholder, setPlaceholder] = useState('');
  const [currentPromptIndex, setCurrentPromptIndex] = useState(0);
  const [wordIndex, setWordIndex] = useState(0);
  const [opacity, setOpacity] = useState(1);
  const [isTransitioning, setIsTransitioning] = useState(false);

  useEffect(() => {
    const currentPrompt = EXAMPLE_PROMPTS[currentPromptIndex];
    const words = currentPrompt.split(' ');

    if (isTransitioning) {
      // During transition, clear text and move to next prompt
      const timeout = setTimeout(() => {
        setPlaceholder('');
        setWordIndex(0);
        setIsTransitioning(false);
        setOpacity(1);
      }, 500); // Match fade out duration

      return () => clearTimeout(timeout);
    }

    // Typing animation - word by word
    if (wordIndex < words.length) {
      const timeout = setTimeout(() => {
        setPlaceholder((prev) => {
          const nextWord = words[wordIndex];
          return prev ? `${prev} ${nextWord}` : nextWord;
        });
        setWordIndex((prev) => prev + 1);
      }, 150); // 150ms between words

      return () => clearTimeout(timeout);
    } else {
      // Wait 2 seconds, then fade out and transition
      const timeout = setTimeout(() => {
        setOpacity(0);
        setIsTransitioning(true);
        setCurrentPromptIndex((prev) => (prev + 1) % EXAMPLE_PROMPTS.length);
      }, 2000);

      return () => clearTimeout(timeout);
    }
  }, [wordIndex, currentPromptIndex, isTransitioning]);

  return { text: placeholder, opacity };
}
