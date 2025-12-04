import { apiFetch } from "./http";

export interface LogPayload {
  // Raw fields we always send
  email: string;
  action: string;
  category: string;
  title: string;
  description: string;
  owner: string;
  savedAt: string;
  // Optional extra metadata blob (not persisted as first-class columns)
  metadata?: any;
}

export async function logAction(email: string, action: string, metadata?: any): Promise<void> {
  if (!action) return;

  const safeEmail =
    email ||
    localStorage.getItem("loggedInEmail") ||
    sessionStorage.getItem("loggedInEmail") ||
    "UnknownUser";

  const nowIso = new Date().toISOString();

  const metaObj = (metadata && typeof metadata === "object") ? metadata : undefined;

  const title =
    (metaObj && typeof metaObj.title === "string" && metaObj.title.trim()) ||
    action;

  const description =
    (metaObj && typeof metaObj.description === "string" && metaObj.description.trim()) ||
    (metaObj ? JSON.stringify(metaObj) : "");

  const owner =
    (metaObj && typeof metaObj.owner === "string" && metaObj.owner.trim()) ||
    safeEmail;

  const payload: LogPayload = {
    email: safeEmail,
    action,
    category: "ActivityLog",
    title,
    description,
    owner,
    savedAt: nowIso,
    metadata,
  };

  try {
    await apiFetch("/api/log", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });
  } catch {}
}
