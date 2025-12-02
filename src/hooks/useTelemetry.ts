import { useEffect } from 'react';
import { useLocation } from 'react-router-dom';
import { appInsights } from '../telemetry';

type TelemetryApi = {
  trackClick: (elementId: string, value?: any) => void;
  trackInput: (elementId: string, value?: any) => void;
  trackComponent: (eventType: string, extra?: Record<string, any>) => void;
};

const getUserId = (): string | null => {
  try {
    const raw = sessionStorage.getItem('clientPrincipal') || localStorage.getItem('clientPrincipal');
    if (!raw) return appInsights?.context?.user?.id || null;
    try { const parsed = JSON.parse(raw); return parsed?.userId || parsed?.user?.id || appInsights?.context?.user?.id || null; } catch { return raw || appInsights?.context?.user?.id || null; }
  } catch { return appInsights?.context?.user?.id || null; }
};

export default function useTelemetry(componentName: string): TelemetryApi {
  const location = useLocation();
  const userId = getUserId() || '';

  const track = (eventType: string, props?: Record<string, any>) => {
    try {
      const name = `${componentName}.${eventType}`;
      const properties = {
        userId: userId || undefined,
        route: location?.pathname || undefined,
        ...props,
      } as Record<string, any>;
      appInsights.trackEvent({ name }, properties);
      appInsights.trackTrace({ message: name }, properties);
    } catch {}
  };

  useEffect(() => {
    try { track('mounted'); } catch {}
    return () => { try { track('unmounted'); } catch {} };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return {
    trackClick: (elementId: string, value?: any) => track(`click:${elementId}`, { elementId, value }),
    trackInput: (elementId: string, value?: any) => track(`input:${elementId}`, { elementId, value }),
    trackComponent: (eventType: string, extra?: Record<string, any>) => track(eventType, extra),
  };
}
