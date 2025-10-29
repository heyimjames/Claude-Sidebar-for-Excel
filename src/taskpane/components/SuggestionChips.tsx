import React from 'react';
import type { Suggestion } from '../hooks/useSmartSuggestions';
import '../styles/message-input.css';

interface SuggestionChipsProps {
  suggestions: Suggestion[];
  onSuggestionClick: (prompt: string) => void;
}

export default function SuggestionChips({ suggestions, onSuggestionClick }: SuggestionChipsProps) {
  if (suggestions.length === 0) {
    return null;
  }

  return (
    <div className="suggestion-chips" role="group" aria-label="Smart suggestions">
      {suggestions.map((suggestion) => (
        <button
          key={suggestion.id}
          className="suggestion-chip"
          onClick={() => onSuggestionClick(suggestion.prompt)}
          title={suggestion.description}
          aria-label={`${suggestion.label}: ${suggestion.description}`}
        >
          <span className="suggestion-icon" aria-hidden="true">
            {suggestion.icon}
          </span>
          <span className="suggestion-label">{suggestion.label}</span>
        </button>
      ))}
    </div>
  );
}
