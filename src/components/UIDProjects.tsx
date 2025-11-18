import React from "react";
import { saveToStorage } from "../api/saveToStorage";

interface Props {
  uid: string;
}

const UIDProjects: React.FC<Props> = ({ uid }) => {
  // Save an example project into the local projects cache (no server call).
  const handleSave = () => {
    try {
      const key = 'uidLocalProjects';
      const raw = localStorage.getItem(key);
      const arr = raw ? JSON.parse(raw) : [];
      const id = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
      // Try to obtain a full snapshot from the parent page (UIDLookup exposes
      // a small helper on window). Fall back to a minimal shape when not
      // available.
      let dataSnapshot: any = { sourceUids: [uid].filter(Boolean) };
      try {
        const fn = (window as any).getCurrentViewSnapshot;
        if (typeof fn === 'function') {
          const snap = fn();
          if (snap && typeof snap === 'object') dataSnapshot = snap;
        }
      } catch {}

      const proj = {
        id,
        name: `Local Project ${id}`,
        createdAt: Date.now(),
        data: dataSnapshot,
        owners: ["local"],
        section: undefined,
        notes: undefined,
        __local: true,
      } as any;
      arr.unshift(proj);
      try { localStorage.setItem(key, JSON.stringify(arr)); } catch {}
      console.log(`[local-save] Project saved locally for UID ${uid}:`, proj);
      // Optionally notify other parts of the app via storage event; consumers read from localStorage on mount
      try { window.dispatchEvent(new Event('storage')); } catch {}
      // Informal UX feedback
      alert('Project saved locally to Projects cache.');
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error('Failed to save project locally', err);
      alert('Failed to save project locally');
    }
  };

  return (
    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
      <button className="sleek-btn repo accent-cta" onClick={handleSave} title="Save project to local Projects cache">Save Project (local)</button>
    </div>
  );
};

export default UIDProjects;
