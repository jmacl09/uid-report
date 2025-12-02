import React, { useEffect, useState } from "react";
import { API_BASE } from "../api/config";

interface Props {
  uid: string;
}

type CommentItem = {
  rowKey?: string;
  description: string;
  owner: string;
  savedAt: string;
  title?: string;
  category?: string;
};

const storageKeyFor = (uid: string) => `uid-comments:${uid}`;

const UIDComments: React.FC<Props> = ({ uid }) => {
  const [comment, setComment] = useState("");
  const [owner, setOwner] = useState("Josh Maclean");

  const [saving, setSaving] = useState(false);
  const [lastSaved, setLastSaved] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const [items, setItems] = useState<CommentItem[]>([]);
  const [editingRowKey, setEditingRowKey] = useState<string | null>(null);
  const [editingText, setEditingText] = useState("");
  const [editSaving, setEditSaving] = useState(false);
  const [deletingKey, setDeletingKey] = useState<string | null>(null);

  const persistItems = (data: CommentItem[]) => {
    try { localStorage.setItem(storageKeyFor(uid), JSON.stringify(data)); } catch {}
  };

  /** Load from backend */
  const refresh = async () => {
    try {
      const resp = await fetch(`${API_BASE}/comments?uid=${encodeURIComponent(uid)}`);
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const data = await resp.json();

      const mapped: CommentItem[] = (data || []).map((r: any) => ({
        rowKey: r.rowKey,
        description: r.description,
        owner: r.owner,
        category: r.category || "Comments",
        title: r.title || "General comment",
        savedAt: r.savedAt || new Date().toISOString(),
      }));

      setItems(prev => {
        const next = mapped.sort((a, b) => (a.savedAt < b.savedAt ? 1 : -1));
        persistItems(next);
        return next;
      });
    } catch (e) {
      console.warn("[UIDComments] Failed to fetch", e);
    }
  };

  /** Initial load */
  useEffect(() => {
    let cancelled = false;

    const load = async () => {
      try {
        const raw = localStorage.getItem(storageKeyFor(uid));
        if (raw) {
          const arr = JSON.parse(raw);
          if (!cancelled && Array.isArray(arr)) setItems(arr);
        }
      } catch {}

      if (!cancelled) await refresh();
    };

    load();
    return () => { cancelled = true; };
  }, [uid]);

  /** Save new comment */
  const handleSave = async () => {
    if (!comment.trim()) return;
    setSaving(true);
    setError(null);

    const newComment: CommentItem = {
      description: comment.trim(),
      owner: owner.trim() || "Unknown",
      savedAt: new Date().toISOString(),
      category: "Comments",
      title: "General comment",
    };

    // Optimistic local update
    setItems(prev => {
      const next = [newComment, ...prev];
      persistItems(next);
      return next;
    });

    setComment("");

    try {
      const resp = await fetch(`${API_BASE}/comments`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          uid,
          description: newComment.description,
          owner: newComment.owner,
          category: "Comments",
          title: "General comment",
        })
      });

      if (!resp.ok) {
        throw new Error(`HTTP ${resp.status}`);
      }

      const data = await resp.json();

      // Replace optimistic entry with official entry
      if (data?.rowKey) {
        setItems(prev => {
          const copy = [...prev];
          copy[0] = {
            ...copy[0],
            rowKey: data.rowKey,
            savedAt: data.savedAt || copy[0].savedAt,
          };
          persistItems(copy);
          return copy;
        });
      }

      setLastSaved(new Date().toLocaleTimeString());
      setTimeout(refresh, 1200);

    } catch (e: any) {
      setError(e?.message || String(e));
    } finally {
      setSaving(false);
    }
  };

  /** Editing */
  const startEdit = (item: CommentItem) => {
    if (!item.rowKey) {
      setError("This comment is still syncing — cannot edit yet.");
      return;
    }
    setError(null);
    setEditingRowKey(item.rowKey);
    setEditingText(item.description);
  };

  const cancelEdit = () => {
    if (editSaving) return;
    setEditingRowKey(null);
    setEditingText("");
  };

  const saveEdit = async () => {
    if (!editingRowKey) return;

    const text = editingText.trim();
    if (!text) return;

    const target = items.find(i => i.rowKey === editingRowKey);
    if (!target) return;

    setEditSaving(true);
    setError(null);

    let snapshot: CommentItem[] = [];

    setItems(prev => {
      snapshot = prev;
      const next = prev.map(it => it.rowKey === editingRowKey ? { ...it, description: text } : it);
      persistItems(next);
      return next;
    });

    try {
      const resp = await fetch(`${API_BASE}/comments`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          uid,
          rowKey: editingRowKey,
          description: text,
          owner: target.owner,
          title: target.title,
          category: target.category,
        })
      });

      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);

      await refresh();
      setLastSaved(new Date().toLocaleTimeString());

    } catch (e: any) {
      setItems(snapshot);
      setError(e?.message || String(e));
    } finally {
      setEditSaving(false);
      setEditingRowKey(null);
      setEditingText("");
    }
  };

  /** Delete */
  const handleDelete = async (item: CommentItem, index: number) => {
    const key = item.rowKey || `${item.savedAt}-${index}`;
    setDeletingKey(key);

    let snapshot: CommentItem[] = [];

    setItems(prev => {
      snapshot = prev;
      const next = prev.filter((_, i) => i !== index);
      persistItems(next);
      return next;
    });

    try {
      if (item.rowKey) {
        await fetch(`${API_BASE}/comments?uid=${encodeURIComponent(uid)}&rowKey=${encodeURIComponent(item.rowKey)}`, {
          method: "DELETE"
        });
      }
      await refresh();
    } catch (e: any) {
      setItems(snapshot);
      setError(e?.message || String(e));
    } finally {
      setDeletingKey(null);
    }
  };

  /** UI */
  const handleKeyDown: React.KeyboardEventHandler<HTMLTextAreaElement> = (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      void handleSave();
    }
  };

  return (
    <div style={{ display: "flex", flexDirection: 'column', gap: 8, alignItems: "flex-start", maxWidth: 500 }}>
      <label style={{ width: '100%' }}>
        <span style={{ fontWeight: 600 }}>Comment</span>
        <textarea
          value={comment}
          onChange={(e) => setComment(e.target.value)}
          onKeyDown={handleKeyDown}
          rows={3}
          placeholder="Add a comment about this UID"
          disabled={saving}
          style={{ width: '100%', marginTop: 4, resize: 'vertical', fontFamily: 'inherit', fontSize: 14 }}
        />
      </label>

      <label>
        <span style={{ fontWeight: 600 }}>Owner:&nbsp;</span>
        <input
          type="text"
          value={owner}
          onChange={(e) => setOwner(e.target.value)}
          disabled={saving}
          style={{ fontSize: 14 }}
        />
      </label>

      <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <button onClick={handleSave} disabled={saving || !comment.trim()}>
          {saving ? "Saving…" : "Save Comment"}
        </button>

        {lastSaved && !error && (
          <span style={{ fontSize: 12, color: "#2d7" }}>Last saved {lastSaved}</span>
        )}

        {error && (
          <span style={{ fontSize: 12, color: "#d33" }}>Error: {error}</span>
        )}
      </div>

      {items.length > 0 && (
        <div style={{ marginTop: 8, width: '100%' }}>
          <div style={{ fontWeight: 600, marginBottom: 4 }}>Recent comments</div>
          <ul style={{ paddingLeft: 16, margin: 0, display: 'flex', flexDirection: 'column', gap: 6 }}>
            {items.map((it, idx) => {
              const localKey = it.rowKey || `${it.savedAt}-${idx}`;
              const isEditing = editingRowKey === it.rowKey;
              const isDeleting = deletingKey === localKey;
              const editDisabled =
                !it.rowKey || (editingRowKey && editingRowKey !== it.rowKey) || editSaving || isDeleting;

              return (
                <li key={localKey} style={{ listStyle: 'disc' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>

                    {isEditing ? (
                      <>
                        <textarea
                          value={editingText}
                          onChange={(e) => setEditingText(e.target.value)}
                          rows={3}
                          disabled={editSaving}
                          style={{ width: '100%', resize: 'vertical', fontFamily: 'inherit', fontSize: 14 }}
                        />

                        <div style={{ display: 'flex', gap: 8 }}>
                          <button onClick={saveEdit} disabled={editSaving || !editingText.trim()}>
                            {editSaving ? "Saving…" : "Save"}
                          </button>

                          <button onClick={cancelEdit} disabled={editSaving}>
                            Cancel
                          </button>
                        </div>
                      </>
                    ) : (
                      <>
                        <div style={{ whiteSpace: 'pre-wrap' }}>{it.description}</div>
                        <div style={{ fontSize: 12, color: "#666" }}>
                          by {it.owner} — {new Date(it.savedAt).toLocaleString()}
                        </div>
                      </>
                    )}

                    <div style={{ display: 'flex', gap: 8 }}>
                      <button onClick={() => startEdit(it)} disabled={editDisabled}>
                        Edit
                      </button>
                      <button onClick={() => handleDelete(it, idx)} disabled={isDeleting}>
                        {isDeleting ? "Deleting…" : "Delete"}
                      </button>
                    </div>

                  </div>
                </li>
              );
            })}
          </ul>
        </div>
      )}
    </div>
  );
};

export default UIDComments;
