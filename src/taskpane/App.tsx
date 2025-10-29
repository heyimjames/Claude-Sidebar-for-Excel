import { useState, useEffect } from 'react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import ChatInterface from './components/ChatInterface';
import ApiKeySetup from './components/ApiKeySetup';
import { ErrorBoundary } from './components/ErrorBoundary';

/* global Office */

export default function App() {
  const [apiKey, setApiKey] = useState<string | null>(null);
  const [isReady, setIsReady] = useState(false);

  useEffect(() => {
    // Load API key from Office settings
    try {
      const savedKey = Office.context.document.settings.get('anthropic_api_key');
      if (savedKey) {
        setApiKey(savedKey as string);
      }
    } catch (error) {
      console.error('Error loading API key:', error);
    }
    setIsReady(true);
  }, []);

  const handleApiKeySave = async (key: string) => {
    try {
      setApiKey(key);
      Office.context.document.settings.set('anthropic_api_key', key);
      await new Promise<void>((resolve, reject) => {
        Office.context.document.settings.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error('Failed to save API key'));
          }
        });
      });
    } catch (error) {
      console.error('Error saving API key:', error);
    }
  };

  if (!isReady) {
    return (
      <FluentProvider theme={webLightTheme}>
        <div style={{
          height: '100vh',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          background: '#fafafa'
        }}>
          <div style={{ textAlign: 'center' }}>
            <div style={{ fontSize: '24px', marginBottom: '8px' }}>⏳</div>
            <div>Loading...</div>
          </div>
        </div>
      </FluentProvider>
    );
  }

  return (
    <FluentProvider theme={webLightTheme}>
      <ErrorBoundary>
        {apiKey ? (
          <ChatInterface apiKey={apiKey} />
        ) : (
          <ApiKeySetup onSave={handleApiKeySave} />
        )}
      </ErrorBoundary>
    </FluentProvider>
  );
}
