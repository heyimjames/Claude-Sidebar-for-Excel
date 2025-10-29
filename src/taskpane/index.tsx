import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import './styles/global.css';

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const container = document.getElementById('root');
    if (container) {
      const root = ReactDOM.createRoot(container);
      root.render(
        <React.StrictMode>
          <App />
        </React.StrictMode>
      );
    }
  }
});
