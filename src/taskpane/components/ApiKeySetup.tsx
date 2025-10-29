import { useState } from 'react';
import { Button, Input, Field } from '@fluentui/react-components';
import '../styles/api-key-setup.css';

interface ApiKeySetupProps {
  onSave: (apiKey: string) => void;
}

export default function ApiKeySetup({ onSave }: ApiKeySetupProps) {
  const [apiKey, setApiKey] = useState('');
  const [error, setError] = useState('');

  const handleSubmit = () => {
    if (!apiKey.trim()) {
      setError('Please enter an API key');
      return;
    }

    if (!apiKey.startsWith('sk-ant-')) {
      setError('Invalid API key format. Should start with sk-ant-');
      return;
    }

    onSave(apiKey.trim());
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      handleSubmit();
    }
  };

  return (
    <div className="api-key-setup">
      <div className="setup-content">
        <div className="setup-header">
          <div className="setup-icon">
            <svg width="48" height="48" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M24 4L4 14L24 24L44 14L24 4Z" stroke="currentColor" strokeWidth="2.5" strokeLinejoin="round"/>
              <path d="M4 34L24 44L44 34" stroke="currentColor" strokeWidth="2.5" strokeLinejoin="round"/>
              <path d="M4 24L24 34L44 24" stroke="currentColor" strokeWidth="2.5" strokeLinejoin="round"/>
            </svg>
          </div>
          <h1>Welcome to Claude for Excel</h1>
          <p>Get started by entering your Anthropic API key</p>
        </div>

        <div className="setup-form">
          <Field
            label="Anthropic API Key"
            validationMessage={error}
            validationState={error ? 'error' : undefined}
          >
            <Input
              type="password"
              placeholder="sk-ant-..."
              value={apiKey}
              onChange={(_, data) => {
                setApiKey(data.value);
                setError('');
              }}
              onKeyDown={handleKeyDown}
              size="large"
            />
          </Field>

          <Button
            appearance="primary"
            onClick={handleSubmit}
            disabled={!apiKey.trim()}
            size="large"
            className="submit-button"
          >
            Get Started
          </Button>

          <div className="setup-help">
            <p className="help-text">
              Don't have an API key?{' '}
              <a href="https://console.anthropic.com" target="_blank" rel="noopener noreferrer">
                Get one from Anthropic
              </a>
            </p>
            <p className="help-note">
              Your API key is stored securely in your Excel workbook settings and is never shared.
            </p>
          </div>
        </div>

        <div className="setup-features">
          <h2>What you can do with Claude</h2>
          <div className="features-grid">
            <div className="feature-card">
              <div className="feature-icon">📊</div>
              <h3>Analyze Data</h3>
              <p>Understand patterns and trends in your spreadsheet</p>
            </div>
            <div className="feature-card">
              <div className="feature-icon">✏️</div>
              <h3>Edit Content</h3>
              <p>Update cells, apply formulas, and format data</p>
            </div>
            <div className="feature-card">
              <div className="feature-icon">📈</div>
              <h3>Create Charts</h3>
              <p>Generate visualizations from your data</p>
            </div>
            <div className="feature-card">
              <div className="feature-icon">🔍</div>
              <h3>Ask Questions</h3>
              <p>Get insights and explanations about your data</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
