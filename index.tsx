import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import { getFontFaceCSS } from './fonts';
import './index.css';

// Inject custom font CSS into the page
const fontCSS = getFontFaceCSS();
const styleSheet = document.createElement('style');
styleSheet.textContent = fontCSS;
document.head.appendChild(styleSheet);

const rootElement = document.getElementById('root');
if (!rootElement) {
  throw new Error("Could not find root element to mount to");
}

const root = ReactDOM.createRoot(rootElement);
root.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);