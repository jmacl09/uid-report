import React from "react";
import ReactDOM from "react-dom/client";
import "./index.css";
import App from "./App";
import reportWebVitals from "./reportWebVitals";
import { initializeIcons } from '@fluentui/react';
// Initialize Fluent UI icons once at app startup
initializeIcons();

const root = ReactDOM.createRoot(
  document.getElementById("root") as HTMLElement
);

// Apply persisted theme classes before rendering to avoid flash
const applyInitialTheme = () => {
  try {
    const rootEl = document.documentElement;
    const storedTheme = localStorage.getItem('appTheme') || 'dark';
    const animations = localStorage.getItem('appAnimations');
    const compact = localStorage.getItem('appCompact');

    if (storedTheme === 'light') rootEl.classList.add('light-theme');
    else rootEl.classList.remove('light-theme');

    if (animations === 'false') rootEl.classList.add('no-animations');
    else rootEl.classList.remove('no-animations');

    if (compact === 'true') rootEl.classList.add('compact-mode');
    else rootEl.classList.remove('compact-mode');
  } catch {
    // ignore
  }
};

applyInitialTheme();

root.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

reportWebVitals();
