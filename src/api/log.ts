import { apiFetch } from "./http";

export interface LogPayload {
  email: string;
  action: string;
  metadata?: any;
}

export async function logAction(email: string, action: string, metadata?: any): Promise<void> {
  if (!email || !action) return;

  try {
    await apiFetch("/api/log", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ email, action, metadata } as LogPayload)
    });
  } catch {
    // Swallow logging errors so UX is never blocked
  }
}
