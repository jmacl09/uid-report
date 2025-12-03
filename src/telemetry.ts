import React from 'react';
import { ApplicationInsights } from '@microsoft/applicationinsights-web';

// Allow the key to be provided at runtime by the Static Web App (window) or at build-time
const INSTRUMENTATION_KEY = (typeof window !== 'undefined' && (window as any).REACT_APP_APPINSIGHTS_INSTRUMENTATIONKEY) || process.env.REACT_APP_APPINSIGHTS_INSTRUMENTATIONKEY || '';

let ai: any = null;
let appInsights: any = null;

if (INSTRUMENTATION_KEY) {
  const config = {
    instrumentationKey: INSTRUMENTATION_KEY,
    enableAutoRouteTracking: true,
    enableCorsCorrelation: true,
    enableRequestHeaderTracking: true,
    enableResponseHeaderTracking: true,
    enableAutoExceptionTracking: true,
    enableUnhandledPromiseRejectionTracking: true,
  };

  ai = new ApplicationInsights({ config });
  ai.loadAppInsights();

  // Expose appInsights instance
  appInsights = ai;
} else {
  // No instrumentation key provided — expose a safe no-op stub so callers can call methods without checks
  const noop = () => {};
  appInsights = {
    trackEvent: noop,
    trackException: noop,
    trackTrace: noop,
    trackMetric: noop,
    setAuthenticatedUserContext: noop,
    context: { user: { id: null } },
  } as any;
}

// Track global window errors
if (typeof window !== 'undefined') {
  try {
    window.addEventListener('error', (ev: ErrorEvent) => {
      try {
        const err = ev.error || new Error(ev.message || 'window error');
        appInsights.trackException({ error: err, severityLevel: 3 });
      } catch {}
    });

    window.addEventListener('unhandledrejection', (ev: PromiseRejectionEvent) => {
      try {
        const reason = (ev.reason instanceof Error) ? ev.reason : new Error(String(ev.reason));
        appInsights.trackException({ error: reason, severityLevel: 3 });
      } catch {}
    });
  } catch {}
}

// Lightweight React Error Boundary that reports to Application Insights
export class ErrorBoundary extends React.Component<any, { hasError: boolean }> {
  constructor(props: any) {
    super(props);
    this.state = { hasError: false };
  }

  static getDerivedStateFromError() {
    return { hasError: true };
  }

  componentDidCatch(error: any, info: any) {
    try {
      appInsights.trackException({ error: error instanceof Error ? error : new Error(String(error)), severityLevel: 3 });
      appInsights.trackTrace({ message: 'React ErrorBoundary', properties: { info: JSON.stringify(info || {}) } });
    } catch {}
  }

  render() {
    if (this.state.hasError) {
      return this.props.fallback ?? null;
    }
    return this.props.children;
  }
}

// Helpers for performance tracking
export const trackRenderMetric = (name: string, value: number, properties?: { [k: string]: any }) => {
  try { appInsights.trackMetric({ name, average: value }, properties); } catch {}
};

export { appInsights };
export default appInsights;
