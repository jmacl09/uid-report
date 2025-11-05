// API base configuration.
// - Local dev: set REACT_APP_API_BASE=http://localhost:7071/api
// - Production (Azure Static Web Apps): default '/api' will be proxied to Functions via staticwebapp.config.json.
//   This avoids cross-origin calls and CORS issues.
export const API_BASE: string = (process.env.REACT_APP_API_BASE as string) || "/api";

export type { };

