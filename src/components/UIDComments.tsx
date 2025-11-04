import React from "react";
import { saveToStorage } from "../api/saveToStorage";

interface Props {
  uid: string;
}

const UIDComments: React.FC<Props> = ({ uid }) => {
  const handleSave = async () => {
    try {
      const result = await saveToStorage({
        category: "Comments",
        uid,
        title: "General comment",
        description: "This UID was validated by NOC",
        owner: "Josh Maclean",
      });
      console.log(`[save] Comments saved for UID ${uid}:`, result);
      // optional: toast.success("Saved comment")
    } catch (e: any) {
      if (e?.status && e.status >= 500) {
        console.error("Server error while saving comment:", e?.body || e?.message);
        // optional: toast.error("Server error while saving comment")
      } else {
        console.error("Failed to save comment:", e?.body || e?.message || e);
      }
    }
  };

  return (
    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
      <button onClick={handleSave}>Save Comment (example)</button>
    </div>
  );
};

export default UIDComments;
