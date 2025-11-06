import React, { useEffect, useState } from "react";
import { saveToStorage, SaveError } from "../api/saveToStorage";
import { getCommentsForUid } from "../api/items";

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
          try { localStorage.setItem(storageKeyFor(uid), JSON.stringify(next)); } catch {}
          return next;
        });
      }
    } catch (e) {
      console.warn('[UIDComments] Failed to fetch remote comments', e);
    }
  };

  useEffect(() => {
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
      try { localStorage.setItem(storageKeyFor(uid), JSON.stringify(next)); } catch {}
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
          {saving ? 'Saving…' : 'Save Comment'}
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
            {items.map((it, idx) => (
              <li key={it.rowKey || `${it.savedAt}-${idx}`} style={{ listStyle: 'disc' }}>
                <div style={{ whiteSpace: 'pre-wrap' }}>{it.description}</div>
                <div style={{ fontSize: 12, color: '#666' }}>
                  by {it.owner} • {new Date(it.savedAt).toLocaleString()}
                </div>
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
};

export default UIDComments;
