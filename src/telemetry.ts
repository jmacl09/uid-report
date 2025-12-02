import React from 'react';
import { ApplicationInsights } from '@microsoft/applicationinsights-web';

const INSTRUMENTATION_KEY = process.env.REACT_APP_APPINSIGHTS_INSTRUMENTATIONKEY || '';

const config = {
  instrumentationKey: INSTRUMENTATION_KEY,
  enableAutoRouteTracking: true,
  enableCorsCorrelation: true,
  enableRequestHeaderTracking: true,
  enableResponseHeaderTracking: true,
  enableAutoExceptionTracking: true,
  enableUnhandledPromiseRejectionTracking: true,
};

const ai = new ApplicationInsights({ config });
ai.loadAppInsights();

// Expose appInsights instance
export const appInsights = ai;

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

export default appInsights;
