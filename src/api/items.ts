import { API_BASE } from "./config";

export type NoteEntity = {
  partitionKey: string;
  rowKey: string;
  UID?: string;
  Category?: string;
  Comment?: string;
  User?: string;
  Title?: string;
  CreatedAt?: string;
  timestamp?: string;
  [key: string]: any;
};

export type GetItemsResponse = {
  ok: boolean;
  uid: string;
  category: string;
  items: NoteEntity[];
};

export async function getNotesForUid(uid: string, signal?: AbortSignal): Promise<NoteEntity[]> {
  const url = `${API_BASE}/GetItems?uid=${encodeURIComponent(uid)}&category=Notes`;
  const res = await fetch(url, { method: 'GET', signal });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`GetItems failed ${res.status}: ${text}`);
  }
  const data = (await res.json()) as GetItemsResponse;
  if (!data?.ok) throw new Error(`GetItems returned not ok`);
  return Array.isArray(data.items) ? data.items : [];
}

export async function deleteNote(partitionKey: string, rowKey: string, signal?: AbortSignal): Promise<void> {
  const url = `${API_BASE}/DeleteItem`;
  const body = { category: 'Notes', partitionKey, rowKey };
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
    signal,
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`DeleteItem failed ${res.status}: ${text}`);
  }
}
