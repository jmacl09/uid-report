import { appInsights } from '../telemetry';

const getUserId = (): string | null => {
  try {
    const raw = sessionStorage.getItem('clientPrincipal') || localStorage.getItem('clientPrincipal');
    if (!raw) return appInsights?.context?.user?.id || null;
    try {
      const parsed = JSON.parse(raw);
      return (parsed?.userId || parsed?.user_id || parsed?.user?.id) || appInsights?.context?.user?.id || null;
    } catch {
      return raw || appInsights?.context?.user?.id || null;
    }
  } catch {
    return appInsights?.context?.user?.id || null;
  }
};

export async function apiFetch(input: RequestInfo, init?: RequestInit): Promise<Response> {
  const start = Date.now();
  const method = (init && init.method) || (typeof input === 'string' ? 'GET' : (input as Request).method) || 'GET';
  const url = typeof input === 'string' ? input : (input as Request).url;
  const userId = getUserId() || '';

  try {
    const resp = await fetch(input, init);
    const duration = Date.now() - start;
    // try to read content-length header first
    let respSize: number | null = null;
    try {
      const cl = resp.headers.get('content-length');
      if (cl) respSize = parseInt(cl, 10);
      else {
        // fallback: clone and measure
        try {
          const buf = await resp.clone().arrayBuffer();
          respSize = buf.byteLength;
        } catch {}
      }
    } catch {}

    try {
      appInsights.trackEvent({ name: 'ApiCall' }, {
        url: String(url),
        method: String(method),
        status: String(resp.status),
        durationMs: String(duration),
        success: resp.ok ? 'true' : 'false',
        responseSize: respSize !== null ? String(respSize) : undefined,
        userId: userId || undefined,
      });

      // also track dependency for richer telemetry (use trackDependencyData)
      try {
        // trackDependencyData expects a dependency telemetry payload
        appInsights.trackDependencyData({
          id: undefined as any,
          name: String(method),
          resultCode: String(resp.status),
          duration: duration,
          success: !!resp.ok,
          properties: {
            url: String(url),
            responseSize: respSize !== null ? String(respSize) : undefined,
            userId: userId || undefined,
          },
          target: String(url),
          type: 'HTTP',
        } as any);
      } catch {}
    } catch {}

    return resp;
  } catch (err: any) {
    const duration = Date.now() - start;
    try {
      appInsights.trackEvent({ name: 'ApiCall' }, {
        url: String(url),
        method: String(method),
        status: 'NETWORK_ERROR',
        durationMs: String(duration),
        success: 'false',
        userId: userId || undefined,
      });
      appInsights.trackException({ error: err instanceof Error ? err : new Error(String(err)) });
    } catch {}
    throw err;
  }
}

export default apiFetch;
