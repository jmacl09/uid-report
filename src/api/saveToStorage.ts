import { API_BASE } from "./config";

export type StorageCategory =
  | "Comments"
  | "Notes"
  | "Projects"
  | "Troubleshooting"
  | "Calendar"
  | "Status"
  | "Suggestions";

export interface SaveInput {
  category: StorageCategory;
  uid?: string;
  title?: string;
  description: string;
  owner: string;
  endpoint?: string;
  timestamp?: string | Date;
  rowKey?: string;
  extras?: Record<string, any>;
}

export async function saveToStorage(input: SaveInput): Promise<any> {
  const category = input.category;
  const endpoint = input.endpoint || "HttpTrigger1";
  const url = `${API_BASE}/${endpoint}?category=${category.toLowerCase()}`;

  const body: any = {
    category,
    title: input.title || "",
    description: input.description,
    owner: input.owner,
  };

  if (input.rowKey) body.rowKey = input.rowKey;

  if (input.timestamp) {
    // Ensure we don't call toISOString on an invalid Date (throws RangeError)
    if (input.timestamp instanceof Date) {
      const d = input.timestamp as Date;
      if (!isNaN(d.getTime())) {
        body.timestamp = d.toISOString();
      } else {
        // fallback to current time if provided Date is invalid
        body.timestamp = new Date().toISOString();
      }
    } else {
      // Try to interpret string-like timestamps as Dates; otherwise use the raw string
      const parsed = new Date(String(input.timestamp));
      if (!isNaN(parsed.getTime())) {
        body.timestamp = parsed.toISOString();
      } else {
        body.timestamp = String(input.timestamp);
      }
    }
  }

  if (input.extras) body.extras = input.extras;

  // If caller provided a uid, include it in the payload for backends that require it
  if (input.uid) body.uid = input.uid;
  if (category === "Troubleshooting") {
    if (!input.uid) throw new Error("Troubleshooting requires UID");
    body.uid = input.uid;
  }

  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Save failed ${res.status}: ${text}`);
  }

  // Try to return parsed JSON, fall back to raw text
  try {
    return await res.json();
  } catch {
    return await res.text();
  }
}
