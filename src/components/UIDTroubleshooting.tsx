import React from "react";
import { saveToStorage } from "../api/saveToStorage";

interface Props {
  uid: string;
}

const UIDTroubleshooting: React.FC<Props> = ({ uid }) => {
  const handleSave = async () => {
    try {
      const result = await saveToStorage({
        category: "Troubleshooting",
        uid,
        title: "Fiber flap investigation",
        description: "Investigating intermittent LOS alarms on span RE-0083",
        owner: "Josh Maclean",
      });
      console.log(`[save] Troubleshooting saved for UID ${uid}:`, result);
    } catch (e: any) {
      if (e?.status && e.status >= 500) {
        console.error("Server error while saving troubleshooting:", e?.body || e?.message);
      } else {
        console.error("Failed to save troubleshooting:", e?.body || e?.message || e);
      }
    }
  };

  return (
    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
      <button onClick={handleSave}>Save Troubleshooting (example)</button>
    </div>
  );
};

export default UIDTroubleshooting;
