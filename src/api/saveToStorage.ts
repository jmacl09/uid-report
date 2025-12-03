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
    body.timestamp =
      input.timestamp instanceof Date
        ? input.timestamp.toISOString()
        : String(input.timestamp);
  }

  if (input.extras) body.extras = input.extras;

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
    throw new Error(`Save failed ${res.status}`);
  }

  return res.json();
}
