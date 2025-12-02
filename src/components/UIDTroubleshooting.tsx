// UIDTroubleshooting.tsx
import React, { useState } from "react";
import { saveToStorage, SaveError } from "../api/saveToStorage";

interface Props {
  uid: string;
  /** Optional LinkKey to associate this entry with a specific link row */
  linkKey?: string;
}

const UIDTroubleshooting: React.FC<Props> = ({ uid, linkKey }) => {
  const [saving, setSaving] = useState(false);
  const [lastSaved, setLastSaved] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [description, setDescription] = useState<string>("");

  const handleSave = async () => {
    setSaving(true);
    setError(null);
    try {
      const extras: Record<string, unknown> = {
        TableName: "Troubleshooting",
        tableName: "Troubleshooting",
        targetTable: "Troubleshooting",
      };
      if (linkKey) extras.LinkKey = linkKey;

      const result = await saveToStorage({
        category: "Troubleshooting",
        uid,
        title: "Troubleshooting Entry",
        description: description || "",
        owner: "Unknown",
        extras,
      });

      // Parse the saved entity (rowKey, timestamp, etc.)
      try {
        const parsed = JSON.parse(result);
        const entity = parsed?.entity ?? parsed?.Entity;
        if (entity?.Timestamp) {
          setLastSaved(new Date(entity.Timestamp).toLocaleTimeString());
        } else {
          setLastSaved(new Date().toLocaleTimeString());
        }
      } catch {
        setLastSaved(new Date().toLocaleTimeString());
      }

      console.log(`[Troubleshooting] Saved for UID ${uid}:`, result);
    } catch (e: any) {
      const se: SaveError | undefined = e instanceof SaveError ? e : undefined;
      const msg = se?.body || se?.message || String(e);
      setError(msg);
      console.error("Troubleshooting save failed:", msg);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        gap: 8,
        alignItems: "flex-start",
        width: "100%",
        maxWidth: 380,
      }}
    >
      <textarea
        placeholder="Enter troubleshooting notes for this UID (or optional LinkKey mapping)..."
        value={description}
        onChange={(e) => setDescription(e.target.value)}
        rows={4}
        style={{
          width: "100%",
          padding: "6px 8px",
          borderRadius: 4,
          border: '1px solid rgba(166,183,198,0.10)',
          background: 'transparent',
          color: '#d0e7ff',
        }}
      />

      <button
        onClick={handleSave}
        disabled={saving}
        style={{
          padding: "6px 12px",
          borderRadius: 6,
          cursor: saving ? "not-allowed" : "pointer",
        }}
      >
        {saving ? "Saving..." : "Save Troubleshooting"}
      </button>

      {lastSaved && !error && (
        <span style={{ fontSize: 12, color: "#2d7" }}>
          Last saved {lastSaved}
        </span>
      )}

      {error && (
        <span style={{ fontSize: 12, color: "#d33" }}>
          Error: {error}
        </span>
      )}
    </div>
  );
};

export default UIDTroubleshooting;
