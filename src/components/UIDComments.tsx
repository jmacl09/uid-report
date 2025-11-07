import React, { useEffect, useState } from "react";
import { saveToStorage, SaveError, type StorageCategory } from "../api/saveToStorage";
import { getCommentsForUid, deleteNote as deleteTableEntity } from "../api/items";

interface Props {
  uid: string;
}

type CommentItem = {
  description: string;
  owner: string;
  savedAt: string;
  title?: string;
  category?: string;
  rowKey?: string;
};

const STORAGE_CATEGORIES = ["Comments", "Notes", "Projects", "Troubleshooting", "Calendar"] as const;
const toStorageCategory = (value?: string | null): StorageCategory => {
  if (!value) return "Comments";
  const normalized = value.trim().toLowerCase();
  const match = STORAGE_CATEGORIES.find(cat => cat.toLowerCase() === normalized);
  return match ?? "Comments";
};

const storageKeyFor = (uid: string) => `uid-comments:${uid}`;
const COMMENTS_ENDPOINT =
  "https://optical360v2-ffa9ewbfafdvfyd8.westeurope-01.azurewebsites.net/api/HttpTrigger1";

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

  const mergeComments = (existing: CommentItem[], incoming: CommentItem[]) => {
    const seen = new Set(existing.map(i => i.rowKey || i.savedAt + i.owner + i.description));
    const merged = [...existing];
    for (const c of incoming) {
      const key = c.rowKey || c.savedAt + c.owner + c.description;
      if (!seen.has(key)) merged.push(c);
    }
    merged.sort((a, b) => (a.savedAt < b.savedAt ? 1 : -1));
    return merged;
  };

  const refresh = async () => {
    try {
      const remote = await getCommentsForUid(uid, COMMENTS_ENDPOINT);
      if (Array.isArray(remote)) {
        const mapped: CommentItem[] = remote.map((r: any) => ({
          description: r.description || r.Description || '',
          owner: r.owner || r.Owner || 'Unknown',
          savedAt: r.savedAt || r.timestamp || r.rowKey || new Date().toISOString(),
          title: r.title || r.Title || 'General comment',
          category: r.category || r.Category || 'Comments',
          rowKey: r.rowKey,
        }));
        setItems(prev => {
          const next = mergeComments(prev, mapped);
          persistItems(next);
          return next;
        });
      }
    } catch (e) {
      console.warn('[UIDComments] Failed to fetch remote comments', e);
    }
  };

  useEffect(() => {
    setEditingRowKey(null);
    setEditingText("");
    setDeletingKey(null);
    let cancelled = false;
    const load = async () => {
      try {
        const raw = localStorage.getItem(storageKeyFor(uid));
        if (raw) {
          const parsed = JSON.parse(raw) as CommentItem[];
          if (!cancelled && Array.isArray(parsed)) setItems(parsed);
        }
      } catch {}
      if (!cancelled) await refresh();
    };
    void load();
    return () => { cancelled = true; };
  }, [uid]);

  const handleSave = async () => {
    if (!comment.trim()) return;
    setSaving(true);
    setError(null);
    const newComment: CommentItem = {
      description: comment.trim(),
      owner: owner.trim() || 'Unknown',
      savedAt: new Date().toISOString(),
      title: 'General comment',
      category: 'Comments',
    };

    // Show it immediately
    setItems(prev => {
      const next = [newComment, ...prev];
      persistItems(next);
      return next;
    });
    setComment("");

    try {
      const result = await saveToStorage({
        endpoint: COMMENTS_ENDPOINT,
        category: "Comments",
        uid,
        title: "General comment",
        description: newComment.description,
        owner: newComment.owner,
      });

      // Parse returned entity data if present
      try {
        const parsed = JSON.parse(result);
        const entity = parsed?.entity ?? parsed?.Entity;
        if (entity?.RowKey) {
          newComment.rowKey = entity.RowKey;
          newComment.savedAt = entity.Timestamp || newComment.savedAt;
        }
      } catch {}

      setLastSaved(new Date().toLocaleTimeString());

      // 🔧 Delay refresh slightly so Function result is persisted in Table Storage
      setTimeout(() => refresh(), 2000);
    } catch (e: any) {
      const se: SaveError | undefined = e instanceof SaveError ? e : undefined;
      setError(se?.body || se?.message || String(e));
      console.error("Failed to save comment:", e);
    } finally {
      setSaving(false);
    }
  };

  const handleKeyDown: React.KeyboardEventHandler<HTMLTextAreaElement> = (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      void handleSave();
    }
  };

  const startEdit = (item: CommentItem) => {
    if (!item.rowKey) {
      setError("Please wait for this comment to finish syncing before editing.");
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
    const nextText = editingText.trim();
    if (!nextText) return;
    const target = items.find(it => it.rowKey === editingRowKey);
    if (!target) return;
    setEditSaving(true);
    setError(null);
    let snapshot: CommentItem[] = [];
    setItems(prev => {
      snapshot = prev;
      const next = prev.map(it => it.rowKey === editingRowKey ? { ...it, description: nextText } : it);
      persistItems(next);
      return next;
    });
    try {
      const result = await saveToStorage({
        endpoint: COMMENTS_ENDPOINT,
        category: toStorageCategory(target.category),
        uid,
        title: target.title || "General comment",
        description: nextText,
        owner: target.owner,
        rowKey: editingRowKey,
      });
      try {
        const parsed = JSON.parse(result);
        const entity = parsed?.entity ?? parsed?.Entity;
        if (entity?.RowKey) {
          setItems(prev => {
            const next = prev.map(it => {
              if (it.rowKey !== entity.RowKey) return it;
              return {
                ...it,
                description: entity.Description || entity.description || nextText,
                owner: entity.Owner || entity.owner || it.owner,
                savedAt: entity.Timestamp || entity.savedAt || new Date().toISOString(),
                title: entity.Title || entity.title || it.title,
              };
            });
            persistItems(next);
            return next;
          });
        }
      } catch {}
      setLastSaved(new Date().toLocaleTimeString());
      await refresh();
    } catch (e: any) {
      setItems(() => {
        persistItems(snapshot);
        return snapshot;
      });
      const se: SaveError | undefined = e instanceof SaveError ? e : undefined;
      setError(se?.body || se?.message || String(e));
    } finally {
      setEditSaving(false);
      setEditingRowKey(null);
      setEditingText("");
    }
  };

  const handleDelete = async (item: CommentItem, index: number) => {
    setError(null);
    const localKey = item.rowKey || `${item.savedAt}-${index}`;
    setDeletingKey(localKey);
    if (editingRowKey && editingRowKey === item.rowKey) {
      setEditingRowKey(null);
      setEditingText("");
    }
    let snapshot: CommentItem[] = [];
    setItems(prev => {
      snapshot = prev;
      const next = prev.filter((_, idx) => idx !== index);
      persistItems(next);
      return next;
    });
    try {
      if (item.rowKey) {
        await deleteTableEntity(`UID_${uid}`, item.rowKey, COMMENTS_ENDPOINT);
      }
      await refresh();
    } catch (e: any) {
      setItems(() => {
        persistItems(snapshot);
        return snapshot;
      });
      setError(e?.message || String(e));
    } finally {
      setDeletingKey(null);
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
          style={{ width: '100%', resize: 'vertical', fontFamily: 'inherit', fontSize: 14, marginTop: 4 }}
          placeholder="Add a comment about this UID"
          disabled={saving}
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
          {saving ? 'Saving...' : 'Save Comment'}
        </button>
        {lastSaved && !error && (
          <span style={{ fontSize: 12, color: '#2d7' }}>Last saved {lastSaved}</span>
        )}
        {error && (
          <span style={{ fontSize: 12, color: '#d33' }}>Error: {error}</span>
        )}
      </div>
      {items.length > 0 && (
        <div style={{ marginTop: 8, width: '100%' }}>
          <div style={{ fontWeight: 600, marginBottom: 4 }}>Recent comments</div>
          <ul style={{ paddingLeft: 16, margin: 0, display: 'flex', flexDirection: 'column', gap: 6 }}>
            {items.map((it, idx) => {
              const localKey = it.rowKey || `${it.savedAt}-${idx}`;
              const isEditing = !!it.rowKey && editingRowKey === it.rowKey;
              const isDeleting = deletingKey === localKey;
              const editDisabled = !it.rowKey || (!!editingRowKey && editingRowKey !== it.rowKey) || editSaving || isDeleting;
              const deleteDisabled = isDeleting || (editingRowKey === it.rowKey && editSaving);
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
                          <button type="button" onClick={saveEdit} disabled={editSaving || !editingText.trim()}>
                            {editSaving ? 'Saving...' : 'Save'}
                          </button>
                          <button type="button" onClick={cancelEdit} disabled={editSaving}>
                            Cancel
                          </button>
                        </div>
                      </>
                    ) : (
                      <>
                        <div style={{ whiteSpace: 'pre-wrap' }}>{it.description}</div>
                        <div style={{ fontSize: 12, color: '#666' }}>
                          by {it.owner} - {new Date(it.savedAt).toLocaleString()}
                        </div>
                      </>
                    )}
                    <div style={{ display: 'flex', gap: 8 }}>
                      <button
                        type="button"
                        onClick={() => startEdit(it)}
                        disabled={editDisabled}
                        title={it.rowKey ? 'Edit comment' : 'Awaiting server sync before editing'}
                      >
                        Edit
                      </button>
                      <button
                        type="button"
                        onClick={() => handleDelete(it, idx)}
                        disabled={deleteDisabled}
                      >
                        {isDeleting ? 'Deleting...' : 'Delete'}
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
