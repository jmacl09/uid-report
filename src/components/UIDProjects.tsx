import React from "react";
import { saveToStorage } from "../api/saveToStorage";

interface Props {
  uid: string;
}

const UIDProjects: React.FC<Props> = ({ uid }) => {
  const handleSave = async () => {
    try {
      const result = await saveToStorage({
        category: "Projects",
        uid,
        title: "Dark Fiber Audit",
        description: "Audit of newly provisioned fiber links",
        owner: "Josh Maclean",
      });
      console.log(`[save] Projects saved for UID ${uid}:`, result);
    } catch (e: any) {
      if (e?.status && e.status >= 500) {
        console.error("Server error while saving project:", e?.body || e?.message);
      } else {
        console.error("Failed to save project:", e?.body || e?.message || e);
      }
    }
  };

  return (
    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
      <button onClick={handleSave}>Save Project (example)</button>
    </div>
  );
};

export default UIDProjects;
