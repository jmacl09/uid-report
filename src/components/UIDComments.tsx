import React, { useState } from "react";
import { saveToStorage, SaveError } from "../api/saveToStorage";

interface Props {
  uid: string;
}

const UIDComments: React.FC<Props> = ({ uid }) => {
  const [comment, setComment] = useState("This UID was validated by NOC");
  const [owner, setOwner] = useState("Josh Maclean");
  const [saving, setSaving] = useState(false);
  const [lastSaved, setLastSaved] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const handleSave = async () => {
    if (!comment.trim()) return;
    setSaving(true);
    setError(null);
    try {
      const result = await saveToStorage({
        // Call the deployed Azure Function directly (cross-origin)
        endpoint: 'https://optical360v2-ffa9ewbfafdvfyd8.westeurope-01.azurewebsites.net/api/HttpTrigger1',
        category: "Comments",             // domain specific categorization
        uid,
        title: "General comment",         // required by backend
        description: comment.trim(),
        owner: owner.trim() || 'Unknown',
      });
      console.log(`[save] Comment saved for UID ${uid}:`, result);
      setLastSaved(new Date().toLocaleTimeString());
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

  return (
    <div style={{ display: "flex", flexDirection: 'column', gap: 8, alignItems: "flex-start", maxWidth: 500 }}>
      <label style={{ width: '100%' }}>
        <span style={{ fontWeight: 600 }}>Comment</span>
        <textarea
          value={comment}
          onChange={(e) => setComment(e.target.value)}
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
    </div>
  );
};

export default UIDComments;
