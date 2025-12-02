import React from "react";
import ReactDOM from "react-dom/client";
import "./index.css";
import App from "./App";
import reportWebVitals from "./reportWebVitals";
import { initializeIcons } from '@fluentui/react';
import { appInsights } from './telemetry';
// Initialize Fluent UI icons once at app startup
initializeIcons();

const root = ReactDOM.createRoot(
  document.getElementById("root") as HTMLElement
);

// Apply persisted theme classes before rendering to avoid flash
const applyInitialTheme = () => {
  try {
    const rootEl = document.documentElement;
    const loggedInEmail = localStorage.getItem('loggedInEmail') || '';
    const themeKey = loggedInEmail ? `appTheme_${loggedInEmail}` : 'appTheme';
    const storedTheme = localStorage.getItem(themeKey) || localStorage.getItem('appTheme') || 'dark';
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

// If authentication info is present in sessionStorage, set the authenticated user context
try {
  const raw = sessionStorage.getItem('clientPrincipal');
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      const userId = parsed?.userId || parsed?.user_id || parsed?.user?.id || parsed?.id || parsed?.sub || null;
      if (userId) appInsights.setAuthenticatedUserContext(String(userId));
    } catch {}
  }
} catch {}

root.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

reportWebVitals();
