import { apiFetch } from "./http";

export interface LogPayload {
  email: string;
  action: string;
  category: string;
  metadata?: any;
}

export async function logAction(email: string, action: string, metadata?: any): Promise<void> {
  if (!action) return;

  const safeEmail =
    email ||
    localStorage.getItem("loggedInEmail") ||
    sessionStorage.getItem("loggedInEmail") ||
    "UnknownUser";

  try {
    await apiFetch("/api/log", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        email: safeEmail,
        action,
        category: "ActivityLog",
        metadata
      } as LogPayload)
    });
  } catch {}
}
