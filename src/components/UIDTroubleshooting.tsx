// UIDTroubleshooting.tsx
import React, { useState } from "react";
import { saveToStorage, SaveError } from "../api/saveToStorage";

interface Props {
  uid: string;
}

const UIDTroubleshooting: React.FC<Props> = ({ uid }) => {
  const [saving, setSaving] = useState(false);
  const [lastSaved, setLastSaved] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const handleSave = async () => {
    setSaving(true);
    setError(null);
    try {
      const result = await saveToStorage({
        category: "Troubleshooting",
        uid,
        title: "Fiber flap investigation",
        description: "Investigating intermittent LOS alarms on span RE-0083",
        owner: "Josh Maclean",
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
      <button
        onClick={handleSave}
        disabled={saving}
        style={{
          padding: "6px 12px",
          borderRadius: 6,
          cursor: saving ? "not-allowed" : "pointer",
        }}
      >
        {saving ? "Saving..." : "Save Troubleshooting (example)"}
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
