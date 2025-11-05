// API base configuration.
// - Local dev: set REACT_APP_API_BASE=http://localhost:7071/api
// - Production: if no env override, prefer SWA proxy '/api';
//   When hosted at optical360.net, fall back to the deployed Functions URL to avoid 404s.
const envBase = process.env.REACT_APP_API_BASE as string | undefined;
let computedBase = envBase || "/api";
try {
	if (!envBase && typeof window !== 'undefined') {
		const host = window.location.hostname.toLowerCase();
		if (host.endsWith('optical360.net')) {
			computedBase = 'https://optical360v2-ffa9ewbfafdvfyd8.westeurope-01.azurewebsites.net/api';
		}
	}
} catch {}

export const API_BASE: string = computedBase;

export type { };

