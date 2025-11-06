import React, { useEffect, useState } from "react";
import { saveToStorage, SaveError } from "../api/saveToStorage";
import { getCommentsForUid } from "../api/items";

interface Props {
  uid: string;
}

type CommentItem = {
  description: string;
  owner: string;
  savedAt: string; // ISO or human readable
  title?: string;
  category?: string;
  rowKey?: string;
};

const storageKeyFor = (uid: string) => `uid-comments:${uid}`;
const COMMENTS_ENDPOINT =
  "https://optical360v2-ffa9ewbfafdvfyd8.westeurope-01.azurewebsites.net/api/HttpTrigger1";

const UIDComments: React.FC<Props> = ({ uid }) => {
  const [comment, setComment] = useState("This UID was validated by NOC");
  const [owner, setOwner] = useState("Josh Maclean");
  const [saving, setSaving] = useState(false);
  const [lastSaved, setLastSaved] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [items, setItems] = useState<CommentItem[]>([]);

  const refresh = async () => {
    try {
      const remote = await getCommentsForUid(uid, COMMENTS_ENDPOINT);
      if (Array.isArray(remote)) {
        const mapped: CommentItem[] = remote.map((r: any) => ({
          description: r.description || r.Description || '',
          owner: r.owner || r.Owner || 'Unknown',
          savedAt: r.savedAt || r.rowKey || new Date().toISOString(),
          title: r.title || r.Title || 'General comment',
          category: r.category || r.Category || 'Comments',
          rowKey: r.rowKey,
        }));
        mapped.sort((a, b) => (a.savedAt < b.savedAt ? 1 : a.savedAt > b.savedAt ? -1 : 0));
        setItems(mapped);
        try { localStorage.setItem(storageKeyFor(uid), JSON.stringify(mapped)); } catch {}
      }
    } catch (e) {
      // eslint-disable-next-line no-console
      console.warn('[UIDComments] Failed to fetch remote comments', e);
    }
  };

  // Load any locally cached comments for this UID, then refresh from server
  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      try {
        const raw = localStorage.getItem(storageKeyFor(uid));
        if (raw) {
          const parsed = JSON.parse(raw) as CommentItem[];
          if (!cancelled && Array.isArray(parsed)) setItems(parsed);
        }
      } catch { /* ignore */ }
      if (!cancelled) await refresh();
    };
    void load();
    return () => { cancelled = true; };
  }, [uid]);

  const persist = (next: CommentItem[]) => {
    setItems(next);
    try { localStorage.setItem(storageKeyFor(uid), JSON.stringify(next)); } catch {}
  };

  const handleSave = async () => {
    if (!comment.trim()) return;
    setSaving(true);
    setError(null);
    try {
      const result = await saveToStorage({
        // Call the deployed Azure Function directly (cross-origin)
        endpoint: COMMENTS_ENDPOINT,
        category: "Comments",             // domain specific categorization
        uid,
        title: "General comment",         // required by backend
        description: comment.trim(),
        owner: owner.trim() || 'Unknown',
      });
      // Try to parse response to extract saved entity details
      let savedAt = new Date().toISOString();
      let rowKey: string | undefined;
      try {
        const parsed = JSON.parse(result);
        const entity = parsed?.entity ?? parsed?.Entity;
        if (entity) {
          savedAt = entity.savedAt || entity.SavedAt || entity.rowKey || savedAt;
          rowKey = entity.rowKey || entity.RowKey;
        }
      } catch { /* non-JSON response, ignore */ }

      const nextItem: CommentItem = {
        description: comment.trim(),
        owner: owner.trim() || 'Unknown',
        savedAt,
        title: 'General comment',
        category: 'Comments',
        rowKey,
      };
      // Optimistic update using functional state to avoid stale closure
      setItems((prev) => {
        const next = [nextItem, ...prev];
        try { localStorage.setItem(storageKeyFor(uid), JSON.stringify(next)); } catch {}
        return next;
      });
      setComment("");
      setLastSaved(new Date().toLocaleTimeString());
      console.log(`[save] Comment saved for UID ${uid}:`, nextItem);

      // Refresh from server for authoritative IDs and ordering
      await refresh();
    } catch (e: any) {
      const se: SaveError | undefined = e instanceof SaveError ? e : undefined;
      if (se?.status && se.status >= 500) {
        console.error("Server error while saving comment:", se?.body || se?.message);
        setError(`Server error (${se.status}): ${se.body || se.message}`);
      } else {
        console.error("Failed to save comment:", se?.body || se?.message || e);
        setError(se?.body || se?.message || String(e));
      }
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
          rows={4}
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
