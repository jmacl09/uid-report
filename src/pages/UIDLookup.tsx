import React, { useState, useEffect } from "react";
import { saveToStorage } from "../api/saveToStorage";
import { getNotesForUid, deleteNote as deleteNoteApi, NoteEntity } from "../api/items";
import { useLocation, useNavigate } from "react-router-dom";
import {
  Stack,
  Text,
  IconButton,
  TextField,
  PrimaryButton,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  MessageBar,
  MessageBarType,
} from '@fluentui/react';
import ThemedProgressBar from "../components/ThemedProgressBar";
import UIDSummaryPanel from "../components/UIDSummaryPanel";
import UIDStatusPanel from "../components/UIDStatusPanel";
import CapacityCircle from "../components/CapacityCircle";
import deriveLineForC0 from "../data/mappedlines";
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

// Note type used across the component for UID notes
type Note = { id: string; uid: string; authorEmail?: string; authorAlias?: string; text: string; ts: number; _pk?: string; _rk?: string };

// Map a table entity from Notes table to local Note type
const mapEntityToNote = (uidKey: string, e: NoteEntity): Note => {
  const authorAlias = (e.User || e.user || e.Owner || (e as any)?.OwnerName || (e as any)?.owner || '').toString() || undefined;
  const rowKey = (e.rowKey || (e as any)?.RowKey || '').toString();
  const partitionKey = (e.partitionKey || (e as any)?.PartitionKey || '').toString();
  const savedAt =
    (e.savedAt || e.timestamp || (e as any)?.Timestamp || rowKey || new Date().toISOString()) as string;
  const ts = (() => {
    const d = Date.parse(savedAt);
    return Number.isFinite(d) ? d : Date.now();
  })();
  const authorEmail = (e.authorEmail || (e as any)?.AuthorEmail || (e as any)?.Email || '').toString() || undefined;
  return {
    id: rowKey || `${Date.now()}`,
    uid: uidKey,
    authorAlias,
    authorEmail,
    text: String(e.Comment || e.comment || e.Description || e.description || e.Title || ''),
    ts,
    _pk: partitionKey,
    _rk: rowKey,
  };
};

// Pure helpers moved to module scope to avoid react-hooks exhaustive-deps issues
const niceWorkflowStatus = (raw?: any): string => {
  const t = String(raw ?? '').trim();
  if (!t) return '';
  const isCancelled = /cancel|cancelled|canceled/i.test(t);
  const isDecom = /decom/i.test(t);
  const isFinished = /wffinished|wf finished|finished/i.test(t);
  const isInProgress = /inprogress|in progress|in-progress|running/i.test(t);
  return isCancelled
    ? 'WF Cancelled'
    : isDecom
    ? 'DECOM'
    : isFinished
    ? 'WF Finished'
    : isInProgress
    ? 'WF In Progress'
    : t;
};

const getWFStatusFor = (src: any, uidKey?: string | null): string => {
  try {
    const map: Record<string, string> | undefined = (src as any)?.__WFStatusByUid;
    const u = (uidKey || '').toString();
    if (u && map && map[u]) {
      return niceWorkflowStatus(map[u]);
    }
    return niceWorkflowStatus(src?.KQLData?.WorkflowStatus);
  } catch {
    return niceWorkflowStatus(src?.KQLData?.WorkflowStatus);
  }
};

export default function UIDLookup() {
  const location = useLocation();
  const navigate = useNavigate();

  const [uid, setUid] = useState<string>("");
  const [data, setData] = useState<any>(null);
  const [loading, setLoading] = useState<boolean>(false);
  const [history, setHistory] = useState<string[]>(() => {
    try {
      const raw = localStorage.getItem("uidHistory");
      return raw ? JSON.parse(raw) : [];
    } catch {
      return [];
    }
  });
  const [lastSearched, setLastSearched] = useState<string>("");
  const [error, setError] = useState<string | null>(null);
  // Inline validation error for UID input (shown under the search box)
  const [uidError, setUidError] = useState<string | null>(null);
  // Loading bar state
  const [progressVisible, setProgressVisible] = useState<boolean>(false);
  const [progressComplete, setProgressComplete] = useState<boolean>(false);
  // Adaptive expected duration for the loading bar (EMA over previous runs)
  const EXPECTED_MS_KEY = 'uidLookupExpectedMs';
  const DEFAULT_EXPECTED_MS = 30000;
  const [expectedMsEstimate, setExpectedMsEstimate] = useState<number>(() => {
    try {
      const raw = localStorage.getItem(EXPECTED_MS_KEY);
      const n = raw ? Number(raw) : NaN;
      return Number.isFinite(n) && n > 0 ? n : DEFAULT_EXPECTED_MS;
    } catch {
      return DEFAULT_EXPECTED_MS;
    }
  });
  const updateExpectedMs = (lastDurationMs: number) => {
    // Exponential moving average with light smoothing
    const alpha = 0.35; // weight for the latest observation
    const prev = expectedMsEstimate || DEFAULT_EXPECTED_MS;
    // Clamp observed duration to reasonable bounds to avoid wild swings
    const observed = Math.min(Math.max(lastDurationMs, 4000), 90000);
    const next = Math.round(alpha * observed + (1 - alpha) * prev);
    setExpectedMsEstimate(next);
    try { localStorage.setItem(EXPECTED_MS_KEY, String(next)); } catch {}
  };
  const firstTableRef = React.useRef<HTMLDivElement | null>(null);
  // Absolute placement for capacity circle (as before)
  const [capacityLeft, setCapacityLeft] = useState<number | null>(null);
  const [capacityTop, setCapacityTop] = useState<number | null>(null);
  const summaryContainerRef = React.useRef<HTMLDivElement | null>(null);
  const [summaryShift, setSummaryShift] = useState<number>(0);
  const [isWide, setIsWide] = useState<boolean>(() => {
    try { return typeof window !== 'undefined' ? window.innerWidth >= 1400 : false; } catch { return false; }
  });
  const [showCancelDialog, setShowCancelDialog] = useState<boolean>(false);
  const [cancelDialogTitle, setCancelDialogTitle] = useState<string>("WF Cancelled");
  const [cancelDialogMsg, setCancelDialogMsg] = useState<string>("");
  const [cancelDialogLink, setCancelDialogLink] = useState<string | null>(null);
  const [lastPromptUid, setLastPromptUid] = useState<string | null>(null);
  // Projects state (localStorage-persisted)
  type Snapshot = {
    sourceUids: string[];
    AExpansions?: any;
    ZExpansions?: any;
    KQLData?: any;
    OLSLinks: any[];
    AssociatedUIDs: any[];
    GDCOTickets: any[];
    MGFXA: any[];
    MGFXZ: any[];
  };
  type Project = {
    id: string;
  name: string; // Computed title e.g., SLS-12345_OSL22 ↔ SVG20, fallback to UID
    createdAt: number;
    data: Snapshot;
    owners?: string[]; // optional display of owners, each shown on its own line
    section?: string; // optional grouping section (e.g., a person's name)
    pinned?: boolean; // optional pin to top
    notes?: Record<string, Note[]>; // notes keyed by UID
    urgent?: boolean; // optional urgent tag
  };
  const [projects, setProjects] = useState<Project[]>(() => {
    try {
      const raw = localStorage.getItem("uidProjects");
      const arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr : [];
    } catch { return []; }
  });
  const [activeProjectId, setActiveProjectId] = useState<string | null>(null);
  const [projectFilter, setProjectFilter] = useState<string>("");
  // Rail filter and viewer selection
  const COLLAPSED_SECTIONS_KEY = 'uidCollapsedSections';
  const VIEWER_SECTION_KEY = 'uidProjectsViewerSection';
  const [viewerSection, setViewerSection] = useState<string | null>(() => {
    try { return localStorage.getItem(VIEWER_SECTION_KEY) || null; } catch { return null; }
  });
  const [collapsedSections, setCollapsedSections] = useState<string[]>(() => {
    try {
      const raw = localStorage.getItem(COLLAPSED_SECTIONS_KEY);
      const arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr.filter(Boolean) : [];
    } catch { return []; }
  });
  useEffect(() => { try { localStorage.setItem(COLLAPSED_SECTIONS_KEY, JSON.stringify(collapsedSections)); } catch {} }, [collapsedSections]);
  useEffect(() => {
    try {
      if (viewerSection) localStorage.setItem(VIEWER_SECTION_KEY, viewerSection);
      else localStorage.removeItem(VIEWER_SECTION_KEY);
    } catch {}
  }, [viewerSection]);

  // If the persisted viewerSection no longer has any projects, fall back to showing all
  useEffect(() => {
    if (!viewerSection) return;
    try {
      const any = projects.some(p => (p.section || '') === viewerSection);
      if (!any) setViewerSection(null);
    } catch {}
  }, [projects, viewerSection]);
  // Modal for actions (projects and sections)
  type ModalKind = 'rename' | 'owners' | 'section' | 'new-section' | 'delete-section' | 'rename-section' | 'move-section' | 'delete-project' | 'create-project' | 'confirm-merge';
  const [modalType, setModalType] = useState<ModalKind | null>(null);
  const [modalProjectId, setModalProjectId] = useState<string | null>(null);
  const [modalValue, setModalValue] = useState<string>("");
  const [modalSection, setModalSection] = useState<string | null>(null);
  // Create Project modal state
  const [createSectionChoice, setCreateSectionChoice] = useState<string>("");
  const [createNewSection, setCreateNewSection] = useState<string>("");
  const [createError, setCreateError] = useState<string | null>(null);
  // Drag & drop state
  const [hoveredSection, setHoveredSection] = useState<string | null>(null);
  const [dragProjectId, setDragProjectId] = useState<string | null>(null);
  const [dropTargetSection, setDropTargetSection] = useState<string | null>(null);
  const [dropProjectId, setDropProjectId] = useState<string | null>(null);
  // Project sections (user-defined group names)
  const [sections, setSections] = useState<string[]>(() => {
    try {
      const raw = localStorage.getItem('uidProjectSections');
      const arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr.filter(Boolean) : [];
    } catch { return []; }
  });
  useEffect(() => {
    try { localStorage.setItem('uidProjectSections', JSON.stringify(sections)); } catch {}
  }, [sections]);
  // Notes/chatbox state
  const [notes, setNotes] = useState<Note[]>([]);
  const [noteText, setNoteText] = useState<string>("");
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editingText, setEditingText] = useState<string>("");
  const [deletingNoteId, setDeletingNoteId] = useState<string | null>(null);
  // Project notes compose state
  const [projNoteText, setProjNoteText] = useState<string>("");
  const [projTargetUid, setProjTargetUid] = useState<string | null>(null);
  // Projects rail collapsed state (persisted)
  const RAIL_KEY = 'uidProjectsRailCollapsed';
  const [railCollapsed, setRailCollapsed] = useState<boolean>(() => {
    try { return localStorage.getItem(RAIL_KEY) === '1'; } catch { return false; }
  });
  useEffect(() => {
    try { localStorage.setItem(RAIL_KEY, railCollapsed ? '1' : '0'); } catch {}
  }, [railCollapsed]);
  // Projects rail width (resizable)
  const RAIL_WIDTH_KEY = 'uidProjectsRailWidth';
  const [railWidth, setRailWidth] = useState<number>(() => {
    try {
      const raw = localStorage.getItem(RAIL_WIDTH_KEY);
      const n = raw ? Number(raw) : NaN;
      return Number.isFinite(n) && n >= 160 && n <= 600 ? n : 260;
    } catch { return 260; }
  });
  useEffect(() => {
    try { localStorage.setItem(RAIL_WIDTH_KEY, String(railWidth)); } catch {}
  }, [railWidth]);
  const railDragRef = React.useRef<{ startX: number; startW: number } | null>(null);
  const onRailDragStart = (e: React.MouseEvent<HTMLDivElement>) => {
    e.preventDefault();
    railDragRef.current = { startX: e.clientX, startW: railWidth };
    const onMove = (ev: MouseEvent) => {
      const ctx = railDragRef.current; if (!ctx) return;
      const dx = ev.clientX - ctx.startX;
      const next = Math.max(160, Math.min(600, Math.round(ctx.startW + dx)));
      setRailWidth(next);
    };
    const onUp = () => {
      railDragRef.current = null;
      window.removeEventListener('mousemove', onMove as any);
      window.removeEventListener('mouseup', onUp as any);
    };
    window.addEventListener('mousemove', onMove as any);
    window.addEventListener('mouseup', onUp as any);
  };
  const onRailKeyResize = (e: React.KeyboardEvent<HTMLDivElement>) => {
    if (railCollapsed) return;
    const step = (e.shiftKey ? 40 : 10);
    if (e.key === 'ArrowLeft') { e.preventDefault(); setRailWidth(w => Math.max(160, w - step)); }
    if (e.key === 'ArrowRight') { e.preventDefault(); setRailWidth(w => Math.min(600, w + step)); }
  };
  const getEmail = () => {
    try { return localStorage.getItem('loggedInEmail') || ''; } catch { return ''; }
  };
  const getAlias = (email?: string | null) => {
    const e = (email || '').trim();
    if (!e) return '';
    const at = e.indexOf('@');
    return at > 0 ? e.slice(0, at) : e;
  };
  // Removed unused myAlias
  // Paste handler: if 11 digits are pasted into the UID box, auto-search
  const handleUidPaste = (e: React.ClipboardEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    try {
      const raw = e.clipboardData.getData('text') || '';
      const cleaned = raw.replace(/\D/g, '').slice(0, 11);
      if (!cleaned) return; // let default paste happen
      // Override default paste to ensure sanitized content only once
      e.preventDefault();
      setUid(cleaned);
      setUidError(() => {
        if (!cleaned) return null;
        return cleaned.length === 11 ? null : 'Invalid UID. It must contain exactly 11 numbers.';
      });
      if (cleaned.length === 11) {
        handleSearch(cleaned);
      }
    } catch {
      // if anything goes wrong, allow default paste behavior
    }
  };
  
  // Reusable inline copy icon + transient message component (shows message next to the icon)
  const CopyIconInline = ({ onCopy, message }: { onCopy: () => void; message?: string }) => {
    const [visible, setVisible] = useState(false);
    const timer = React.useRef<number | null>(null);
    const handle = () => {
      try { onCopy(); } catch {}
      setVisible(true);
      if (timer.current) window.clearTimeout(timer.current);
      timer.current = window.setTimeout(() => setVisible(false), 1600) as unknown as number;
    };
    return (
      <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
        <IconButton
          iconProps={{ iconName: 'Copy' }}
          title="Copy"
          ariaLabel="Copy"
          onClick={handle}
          styles={{ root: { transform: 'scale(0.7)', transformOrigin: 'center', padding: 0, height: 20, minWidth: 20 } }}
        />
        {visible && (
          <span style={{ background: '#e6fff1', border: '1px solid #9fe9b8', color: '#033a16', padding: '1px 6px', borderRadius: 4, fontSize: 11, lineHeight: 1, display: 'inline-block' }}>
            {message ?? 'Copied'}
          </span>
        )}
      </span>
    );
  };
  useEffect(() => {
    localStorage.setItem("uidHistory", JSON.stringify(history.slice(0, 10)));
  }, [history]);

  // persist projects whenever they change
  useEffect(() => {
    try { localStorage.setItem('uidProjects', JSON.stringify(projects)); } catch {}
  }, [projects]);

  // Load notes when the current UID changes (server data only)
  useEffect(() => {
    const keyUid = lastSearched || '';
    if (!keyUid) { setNotes([]); return; }
    let cancelled = false;
    (async () => {
      try {
        const items = await getNotesForUid(keyUid, NOTES_ENDPOINT);
        if (cancelled) return;
        const mapped: Note[] = items.map(e => mapEntityToNote(keyUid, e)).sort((a,b)=>b.ts-a.ts);
        setNotes(mapped);
      } catch (err) {
        setNotes([]);
      }
    })();
    return () => { cancelled = true; };
  }, [lastSearched]);
  const addNote = () => {
    const uidKey = lastSearched || '';
    if (!uidKey) return;
    const text = noteText.trim();
    if (!text) return;
    const email = getEmail();
    const alias = getAlias(email);
    const n: Note = { id: `${Date.now()}-${Math.random().toString(36).slice(2,8)}`, uid: uidKey, authorEmail: email || undefined, authorAlias: alias || undefined, text, ts: Date.now() };
    // Optimistic local append using functional state to avoid stale closures
    setNotes(prev => [n, ...prev]);
    setNoteText('');
    // Also add to any saved projects that contain this UID
    setProjects(prev => prev.map(p => {
      const src = p?.data?.sourceUids || [];
      if (!src.includes(uidKey)) return p;
      const map = { ...(p.notes || {}) } as Record<string, Note[]>;
      const existing = map[uidKey] || [];
      // de-dup by id
      if (!existing.find(x => x.id === n.id)) map[uidKey] = [n, ...existing];
      return { ...p, notes: map };
    }));
    // Fire-and-forget server save of this comment to the Projects table via Function app
    // Title kept compact; description holds full note text
    try {
      void saveToStorage({
        endpoint: NOTES_ENDPOINT,
        category: "Notes",
        uid: uidKey,
        title: "UID Comment",
        description: text,
        owner: alias || email || "",
      }).then(async (resultText) => {
        let savedNote: Note | null = null;
        try {
          const parsed = JSON.parse(resultText);
          const entity = (parsed?.entity || parsed?.Entity) as NoteEntity | undefined;
          if (entity) {
            savedNote = mapEntityToNote(uidKey, entity);
            // Replace the optimistic entry with the canonical server entity
            setNotes(prev => {
              const idx = prev.findIndex(entry => entry.id === n.id);
              if (idx === -1) return prev;
              const next = [...prev];
              next[idx] = savedNote as Note;
              return next;
            });
          }
        } catch {
          savedNote = null;
        }

        // Delay refresh slightly so the Function has time to persist to Table Storage
        await new Promise<void>(resolve => setTimeout(resolve, 1500));

        try {
          const items = await getNotesForUid(uidKey, NOTES_ENDPOINT);
          const mapped: Note[] = items.map(e => mapEntityToNote(uidKey, e)).sort((a,b)=>b.ts-a.ts);
          setNotes(prev => {
            const remoteIds = new Set(mapped.map(item => item.id));
            const leftovers = prev.filter(item => !remoteIds.has(item.id));
            const next = [...mapped, ...leftovers];
            return next;
          });
        } catch {
          // Keep optimistic notes if refresh fails; they'll sync on next load
        }
      }).catch((e) => {
        // eslint-disable-next-line no-console
        console.warn("Server-side save failed (comment kept locally):", e?.body || e?.message || e);
      });
    } catch {}
  };
  const canModify = (n: Note) => {
    const email = getEmail();
    return !!email && email === (n.authorEmail || '');
  };
  const removeNote = async (id: string) => {
    const uidKey = lastSearched || '';
    if (!uidKey) return;
    const target = notes.find(n => n.id === id);
    if (!target) return;
    const pk = target._pk || `UID_${uidKey}`;
    setDeletingNoteId(id);
    try {
      if (target._rk) {
        await deleteNoteApi(pk, target._rk, NOTES_ENDPOINT);
      }
      // Refresh notes from server after successful delete
      const items = await getNotesForUid(uidKey, NOTES_ENDPOINT);
      const mapped: Note[] = items.map(e => mapEntityToNote(uidKey, e)).sort((a,b)=>b.ts-a.ts);
      setNotes(mapped);
    } catch (err) {
      // On failure, keep current list but log
      // eslint-disable-next-line no-console
      console.warn('Failed to delete note from server:', err);
      // Optimistic local removal as fallback
      const next = notes.filter(n => n.id !== id);
      setNotes(next);
    } finally {
      setDeletingNoteId(current => (current === id ? null : current));
    }
  };
  const startEdit = (n: Note) => {
    setEditingId(n.id);
    setEditingText(n.text);
  };
  const saveEdit = () => {
    const uidKey = lastSearched || '';
    if (!uidKey || !editingId) return;
    const text = editingText.trim();
    const next = notes.map(n => n.id === editingId ? { ...n, text } : n);
    setNotes(next);
    setEditingId(null);
    setEditingText('');
  };
  const cancelEdit = () => { setEditingId(null); setEditingText(''); };
  useEffect(() => {
    localStorage.setItem("uidHistory", JSON.stringify(history.slice(0, 10)));
  }, [history]);

  // Reset to landing view when sidebar forces a reset param
  useEffect(() => {
    const params = new URLSearchParams(location.search);
    if (params.has("reset")) {
      setUid("");
      setLastSearched("");
      setData(null);
      setError(null);
      setLoading(false);
      navigate("/uid", { replace: true });
    }
  }, [location.search, navigate]);

  // If the URL contains a uid param, perform the search on mount / when it changes
  useEffect(() => {
    const params = new URLSearchParams(location.search);
    const uidParam = params.get("uid");
    if (uidParam && uidParam !== lastSearched) {
      handleSearch(uidParam);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [location.search]);

  // Show a prompt if WF is cancelled or DECOM for this UID (only once per UID)
  useEffect(() => {
    if (!data || !lastSearched) return;
    if (lastPromptUid === lastSearched) return; // already prompted for this UID
    const raw = String(getWFStatusFor(data, lastSearched) || '').trim();
    const isCancelled = /cancel|cancelled|canceled/i.test(raw);
    const isDecom = /decom/i.test(raw);
    if (isCancelled || isDecom) {
      const jobId = data?.KQLData?.JobId;
      const link = jobId ? `https://azcis.trafficmanager.net/Public/NetworkingOptical/JobDetails/${jobId}` : null;
      setCancelDialogTitle(isCancelled ? 'WF Cancelled' : 'DECOM');
      setCancelDialogMsg(isCancelled ? 'This workflow has been cancelled. Please check the job in CIS below to confirm.' : 'This workflow appears to be decommissioned.');
      setCancelDialogLink(link);
      setShowCancelDialog(true);
      setLastPromptUid(lastSearched);
    }
  }, [data, lastSearched, lastPromptUid]);

  // Helper to get the dataset currently being viewed (live or project snapshot)
  const getViewData = () => {
    const p = projects.find(p => p.id === activeProjectId) || null;
    return p ? p.data : data;
  };
  const getActiveProject = () => projects.find(p => p.id === activeProjectId) || null;
  // Reset project note target when active project changes
  useEffect(() => {
    if (!activeProjectId) { setProjTargetUid(null); setProjNoteText(''); return; }
    const ap = projects.find(p => p.id === activeProjectId) || null;
    const first = (ap?.data?.sourceUids || [])[0] || null;
    setProjTargetUid(first);
    setProjNoteText('');
  }, [activeProjectId, projects]);
  const addProjectNote = () => {
    const ap = getActiveProject();
    if (!ap) return;
    const uidKey = (projTargetUid || ap.data?.sourceUids?.[0] || '').toString();
    if (!uidKey) return;
    const text = projNoteText.trim();
    if (!text) return;
    const email = getEmail();
    const alias = getAlias(email);
    const n: Note = { id: `${Date.now()}-${Math.random().toString(36).slice(2,8)}`, uid: uidKey, authorEmail: email || undefined, authorAlias: alias || undefined, text, ts: Date.now() };
    setProjects(prev => prev.map(p => {
      if (p.id !== ap.id) return p;
      const map = { ...(p.notes || {}) } as Record<string, Note[]>;
      const arr = map[uidKey] || [];
      map[uidKey] = [n, ...arr];
      return { ...p, notes: map };
    }));
    setProjNoteText('');
    // Fire-and-forget server save for project notes as well
    try {
      void saveToStorage({
        category: "Notes",
        uid: uidKey,
        title: "Project Comment",
        description: text,
        owner: alias || email || "",
      }).catch((e) => {
        // eslint-disable-next-line no-console
        console.warn("Server-side save failed (project comment kept locally):", e?.body || e?.message || e);
      });
    } catch {}
  };

  const naturalSort = (a: string, b: string) =>
    a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });

  // Normalize WorkflowStatus strings: now using module-scope helpers above

  // Associated UIDs view filter: show only In Progress by default
  const [showAllAssociatedWF, setShowAllAssociatedWF] = useState<boolean>(false);
  // Track which view (UID or Project) we've auto-applied the default for, so user toggles aren't overridden
  const [associatedWFViewKey, setAssociatedWFViewKey] = useState<string | null>(null);

  // When changing to a new view (new UID or project), auto-toggle: if there are no In-Progress rows, show all by default
  useEffect(() => {
    try {
      const viewKey = activeProjectId ? `project:${activeProjectId}` : (lastSearched ? `uid:${lastSearched}` : null);
      const current = activeProjectId
        ? (projects.find(p => p.id === activeProjectId)?.data || null)
        : data;
      if (!viewKey || !current) return;
      if (associatedWFViewKey === viewKey) return; // already applied for this view
      const rows: any[] = Array.isArray((current as any).AssociatedUIDs) ? (current as any).AssociatedUIDs : [];
      const wfMap: Record<string, string> | undefined = (current as any).__WFStatusByUid;
      const hasInProgress = rows.some((r: any) => {
        const uid = r?.UID ?? r?.Uid ?? r?.uid ?? '';
        const wf = niceWorkflowStatus(wfMap?.[String(uid)]) || '';
        return /in\s*-?\s*progress|running/i.test(wf);
      });
      setShowAllAssociatedWF(!hasInProgress);
      setAssociatedWFViewKey(viewKey);
    } catch {
      const viewKey = activeProjectId ? `project:${activeProjectId}` : (lastSearched ? `uid:${lastSearched}` : null);
      if (viewKey) setAssociatedWFViewKey(viewKey);
      setShowAllAssociatedWF(true);
    }
  }, [activeProjectId, lastSearched, data, projects, associatedWFViewKey]);

  // computeCapacity: derive display strings from Link Summary rows, summing mixed per-link speeds
  const computeCapacity = (links: any[] | undefined, increment?: string | number | null, deviceAFallback?: string | null) => {
    if (!links || !links.length) return { main: '?', sub: '0 links', count: 0 };
    const incNum = increment != null && increment !== '' && !isNaN(Number(increment)) ? Number(increment) : null;
    const parseSpeedFromText = (s: any): number | null => {
      const t = String(s ?? '').toLowerCase();
      if (!t) return null;
      if (/four\s*hundred|\b400\b/.test(t)) return 400;
      if (/\bhundred\b|\b100\b/.test(t)) return 100;
      if (/ten\s*g|10\s*g|\b10g\b|tengig/.test(t)) return 10;
      const m = t.match(/(\d+)\s*g/);
      if (m) return parseInt(m[1], 10);
      return null;
    };
    const perLinkGb = (row: any): number | null => {
      // Prefer parsing from port names per row
      const fromAPort = parseSpeedFromText(row?.APort);
      if (fromAPort != null) return fromAPort;
      const fromZPort = parseSpeedFromText(row?.ZPort);
      if (fromZPort != null) return fromZPort;
      // Fallback: try devices in the row
      const fromADev = parseSpeedFromText(row?.ADevice || row?.['A Device'] || row?.DeviceA || row?.['Device A']);
      if (fromADev != null) return fromADev;
      const fromZDev = parseSpeedFromText(row?.ZDevice || row?.['Z Device'] || row?.DeviceZ || row?.['Device Z']);
      if (fromZDev != null) return fromZDev;
      // Fallback: global increment
      if (incNum != null) return incNum;
      // Fallback: global deviceA hint
      const fromGlobalDev = parseSpeedFromText(deviceAFallback);
      if (fromGlobalDev != null) return fromGlobalDev;
      return null;
    };
    let totalGb = 0;
    let known = 0;
    for (const r of links) {
      const gb = perLinkGb(r);
      if (gb != null) { totalGb += gb; known++; }
    }
    const linkCount = links.length;
    const totalDisplay = known > 0 ? `${totalGb}G` : '?';
    // Build a compact distribution text like "2x400G + 1x100G"
    const bucket = new Map<number, number>();
    for (const r of links) {
      const gb = perLinkGb(r);
      if (gb != null) bucket.set(gb, (bucket.get(gb) || 0) + 1);
    }
    const dist = Array.from(bucket.entries()).sort((a,b)=>b[0]-a[0]).map(([gb,c])=>`${c}x${gb}G`).join(' + ');
    return { main: totalDisplay, sub: `${linkCount} link${linkCount===1?'':'s'}`, count: linkCount, distribution: dist } as { main: string; sub: string; count: number; distribution?: string };
  };

  // Derive expected delivery date for a project from its UIDs (earliest non-empty date)
  const getProjectExpectedDelivery = (p: Project): string | null => {
    try {
      const uids: string[] = Array.isArray(p?.data?.sourceUids) ? p.data.sourceUids : [];
      const dates: string[] = [];
      for (const u of uids) {
        const raw = localStorage.getItem(`uidStatus:${u}`);
        if (!raw) continue;
        const parsed = JSON.parse(raw);
        const d: string | null = parsed?.expectedDeliveryDate ?? null;
        if (d) dates.push(d);
      }
      if (!dates.length) return null;
      // pick min date (earliest)
      dates.sort();
      return dates[0];
    } catch { return null; }
  };

  // ---- Project helpers ----
  const sanitizeArrays = (obj: any) => JSON.parse(JSON.stringify(obj ?? {}));
  const stripSide = (rows: any[]) => (rows || []).map(({ Side, ...keep }) => keep);
  const buildSnapshotFrom = (src: any, srcUid: string): Snapshot => {
    const snap: Snapshot = {
      sourceUids: [srcUid].filter(Boolean),
      AExpansions: sanitizeArrays(src?.AExpansions),
      ZExpansions: sanitizeArrays(src?.ZExpansions),
      KQLData: sanitizeArrays(src?.KQLData),
      OLSLinks: Array.isArray(src?.OLSLinks) ? sanitizeArrays(src.OLSLinks) : [],
      AssociatedUIDs: Array.isArray(src?.AssociatedUIDs) ? sanitizeArrays(src.AssociatedUIDs) : [],
      GDCOTickets: Array.isArray(src?.GDCOTickets) ? sanitizeArrays(src.GDCOTickets) : [],
      MGFXA: Array.isArray(src?.MGFXA) ? stripSide(sanitizeArrays(src.MGFXA)) : [],
      MGFXZ: Array.isArray(src?.MGFXZ) ? stripSide(sanitizeArrays(src.MGFXZ)) : [],
    };
    return snap;
  };
  const dedupMerge = (arrA: any[], arrB: any[]) => {
    const seen = new Set<string>();
    const push = (acc: any[], item: any) => {
      const key = JSON.stringify(item);
      if (!seen.has(key)) { seen.add(key); acc.push(item); }
      return acc;
    };
    const acc: any[] = [];
    (arrA || []).forEach(i => push(acc, i));
    (arrB || []).forEach(i => push(acc, i));
    return acc;
  };
  const mergeSnapshots = (base: Snapshot, add: Snapshot): Snapshot => {
    return {
      sourceUids: Array.from(new Set([...(base.sourceUids||[]), ...(add.sourceUids||[])])),
      // Keep the first non-empty value for details; fallback to added if base empty
      AExpansions: base.AExpansions && Object.keys(base.AExpansions).length ? base.AExpansions : add.AExpansions,
      ZExpansions: base.ZExpansions && Object.keys(base.ZExpansions).length ? base.ZExpansions : add.ZExpansions,
      KQLData: base.KQLData && Object.keys(base.KQLData).length ? base.KQLData : add.KQLData,
      OLSLinks: dedupMerge(base.OLSLinks, add.OLSLinks),
      AssociatedUIDs: dedupMerge(base.AssociatedUIDs, add.AssociatedUIDs),
      GDCOTickets: dedupMerge(base.GDCOTickets, add.GDCOTickets),
      MGFXA: dedupMerge(base.MGFXA, add.MGFXA),
      MGFXZ: dedupMerge(base.MGFXZ, add.MGFXZ),
    };
  };
  const getFirstSites = (src: any, uidKey?: string): { a?: string|null; z?: string|null } => {
    try {
      // 1) Prefer Associated UIDs row matching the entered UID
      const rows: any[] = Array.isArray(src?.AssociatedUIDs) ? src.AssociatedUIDs : [];
      if (rows.length) {
        const match = uidKey ? rows.find(r => String(r?.UID || r?.Uid || r?.uid || '') === String(uidKey)) : null;
        const r = match || rows[0] || {};
        const siteA = r['Site A'] ?? r['SiteA'] ?? r['A Site'] ?? r['ASite'] ?? r['Site'] ?? null;
        const siteZ = r['Site Z'] ?? r['SiteZ'] ?? r['Z Site'] ?? r['ZSite'] ?? null;
        if (siteA || siteZ) return { a: siteA || null, z: siteZ || null };
      }
      // 2) Fallback to A/ZExpansions if Associated UIDs are unavailable or blank
      const a = src?.AExpansions?.DCLocation || src?.AExpansions?.Site || null;
      const z = src?.ZExpansions?.DCLocation || src?.ZExpansions?.Site || null;
      if (a || z) return { a, z };
      return { a: null, z: null };
    } catch { return { a: null, z: null }; }
  };

  // Get SRLGID with fallback: prefer AExpansions.SRLGID; else use Associated UIDs (match current UID if possible)
  const getSrlgIdFrom = (src: any, uidKey?: string): string | null => {
    try {
      const fromA = src?.AExpansions?.SRLGID || src?.AExpansions?.SrlgId || src?.AExpansions?.SRLGId;
      if (fromA) return String(fromA);
      const rows: any[] = Array.isArray(src?.AssociatedUIDs) ? src.AssociatedUIDs : [];
      if (!rows.length) return null;
      const match = uidKey ? rows.find(r => String(r?.UID || r?.Uid || r?.uid || '') === String(uidKey)) : null;
      const r = match || rows[0] || {};
      const val = r['SrlgId'] ?? r['SRLGID'] ?? r['SrlgID'] ?? r['SRLGId'] ?? r['srlgid'] ?? r['Srlg Id'] ?? r['SRLG Id'];
      const s = (val != null) ? String(val).trim() : '';
      return s || null;
    } catch {
      return null;
    }
  };

  // Get SRLG with fallback: prefer AExpansions.SRLG; else use KQLData.SRLG
  const getSrlgFrom = (src: any): string | null => {
    try {
      const a = src?.AExpansions?.SRLG ?? src?.AExpansions?.Srlg;
      const k = src?.KQLData?.SRLG ?? src?.KQLData?.Srlg;
      const val = (a != null && String(a).trim()) ? String(a).trim() : (k != null && String(k).trim() ? String(k).trim() : '');
      return val || null;
    } catch { return null; }
  };
  const computeProjectTitle = (src: any, uidKey: string): string => {
    try {
      const sols = getSolutionIds(src).map(formatSolutionId).filter(Boolean);
      const sol = sols[0] || '';
      const sites = getFirstSites(src, uidKey);
      const a = (sites.a || '').toString().trim();
      const z = (sites.z || '').toString().trim();
  if (sol && a && z) return `${sol}_${a} ↔ ${z}`;
      if (sol && (a || z)) return `${sol}_${a || z}`;
      if (sol) return sol;
  if (a && z) return `${a} ↔ ${z}`;
      return uidKey;
    } catch { return uidKey; }
  };
  // Derive a project "Type" label from Associated UIDs Type column; default to "Standard"
  const getProjectType = (p: Project): string => {
    try {
      const rows: any[] = Array.isArray(p?.data?.AssociatedUIDs) ? p.data.AssociatedUIDs : [];
      const vals = rows
        .map((r) => {
          const v = (r && (r['Type'] ?? r['type'] ?? r['TYPE'])) as any;
          return String(v ?? '').trim();
        })
        .filter(Boolean);
      if (vals.length) {
        // Prefer a value that contains Owned/Hybrid; fallback to first
        const preferred = vals.find((s) => /owned|hybrid/i.test(s)) || vals[0];
        const t = preferred.replace(/_/g, '-').replace(/\s+/g, ' ').trim();
        if (/^owned$/i.test(t)) return 'Owned-OLS';
        if (/^hybrid$/i.test(t)) return 'Hybrid-OLS';
        // Normalize "Owned OLS" to "Owned-OLS"
        const norm = t.replace(/\b(Owned|Hybrid)\b\s*[-_ ]?\s*\b(OLS)\b/i, (_m, a, _b) => `${a}-OLS`);
        return norm;
      }
    } catch {}
    return 'Standard';
  };
  const createProjectFromCurrent = () => {
    if (!data || !lastSearched) return;
    // Require selecting or creating a section
    setCreateSectionChoice("");
    setCreateNewSection("");
    setCreateError(null);
    setModalType('create-project');
  };
  const addCurrentToProject = (targetId: string) => {
    if (!data || !targetId || !lastSearched) return;
    const p = projects.find(pp => pp.id === targetId);
    if (!p) return;
    // Compare SolutionID(s); if both sides have values and they differ, prompt
    try {
      const curS = (getSolutionIds(data) || []).map(formatSolutionId).filter(Boolean);
      const projS = (getSolutionIds(p.data) || []).map(formatSolutionId).filter(Boolean);
      const sameSet = (() => {
        if (!curS.length || !projS.length) return true; // only warn on definite mismatch
        const a = Array.from(new Set(curS)).sort().join('|');
        const b = Array.from(new Set(projS)).sort().join('|');
        return a === b;
      })();
      if (!sameSet) {
        setModalProjectId(targetId);
        setModalType('confirm-merge');
        return;
      }
    } catch {}
    // No warning needed, merge immediately
    setProjects(prev => prev.map(pp => {
      if (pp.id !== targetId) return pp;
      const merged = mergeSnapshots(pp.data, buildSnapshotFrom(data, lastSearched));
      return { ...pp, data: merged, notes: { ...(pp.notes || {}) } };
    }));
    setActiveProjectId(targetId);
  };
  // removed unused deleteProject (deletion flows via requestDeleteProject + modal)
  const requestDeleteProject = (id: string) => {
    setModalProjectId(id);
    setModalType('delete-project');
  };
  const renameProject = (id: string) => {
    const p = projects.find(x => x.id === id); if (!p) return;
    setModalProjectId(id);
    setModalValue(p.name || '');
    setModalType('rename');
  };
  const editOwners = (id: string) => {
    const p = projects.find(x => x.id === id); if (!p) return;
    setModalProjectId(id);
    setModalValue((p.owners || []).join('\n'));
    setModalType('owners');
  };
  // removed unused assignSection
  const togglePin = (id: string) => {
    setProjects(prev => prev.map(x => x.id === id ? { ...x, pinned: !x.pinned } : x));
  };
  const toggleUrgent = (id: string) => {
    setProjects(prev => prev.map(x => x.id === id ? { ...x, urgent: !x.urgent } : x));
  };
  const addSection = () => {
    setModalProjectId(null);
    setModalSection(null);
    setModalValue('');
    setModalType('new-section');
  };
  const closeModal = () => { setModalType(null); setModalProjectId(null); setModalValue(''); };
  const saveModal = () => {
    const value = (modalValue || '').trim();
    if (!modalType) return;
    if (modalType === 'confirm-merge' && modalProjectId) {
      // Proceed with merge now that user confirmed
      const targetId = modalProjectId;
      if (!data || !lastSearched) { closeModal(); return; }
      setProjects(prev => prev.map(pp => {
        if (pp.id !== targetId) return pp;
        const merged = mergeSnapshots(pp.data, buildSnapshotFrom(data, lastSearched));
        return { ...pp, data: merged, notes: { ...(pp.notes || {}) } };
      }));
      setActiveProjectId(targetId);
      closeModal();
      return;
    }
    if (modalType === 'create-project') {
      if (!data || !lastSearched) { closeModal(); return; }
      const chosen = (createSectionChoice || '').trim();
      const newName = (createNewSection || '').trim();
      let finalSection: string | undefined = undefined;
      if (newName) {
        if (!sections.includes(newName)) setSections([...sections, newName]);
        finalSection = newName;
      } else if (chosen) {
        // Archives means unassigned
        finalSection = (chosen === 'Archives') ? undefined : chosen;
      }
      if (!newName && !chosen) {
        setCreateError('Please choose a section or enter a new section name.');
        return;
      }
      const id = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
      // Seed notes for the current UID if any
      const notesMap: Record<string, Note[]> = {};
      try { if (notes && notes.length) notesMap[lastSearched] = [...notes]; } catch {}
      const proj: Project = {
        id,
        name: computeProjectTitle(data, lastSearched),
        createdAt: Date.now(),
        data: buildSnapshotFrom(data, lastSearched),
        section: finalSection,
        notes: Object.keys(notesMap).length ? notesMap : undefined,
      };
      setProjects((prev) => [proj, ...prev]);
      setActiveProjectId(id);
      // reset create-project state
      setCreateSectionChoice('');
      setCreateNewSection('');
      setCreateError(null);
      closeModal();
      return;
    }
    if (modalType === 'rename' && modalProjectId) {
      if (!value) { closeModal(); return; }
      setProjects(prev => prev.map(x => x.id === modalProjectId ? { ...x, name: value } : x));
    } else if (modalType === 'owners' && modalProjectId) {
      const owners = value
        .split(/\n|,/) // comma or newline
        .map(s => s.trim())
        .filter(Boolean);
      setProjects(prev => prev.map(x => x.id === modalProjectId ? { ...x, owners } : x));
    } else if (modalType === 'section' && modalProjectId) {
      const section = value;
      if (section && !sections.includes(section)) setSections([...sections, section]);
      setProjects(prev => prev.map(x => x.id === modalProjectId ? { ...x, section: section || undefined } : x));
    } else if (modalType === 'new-section') {
      if (value && !sections.includes(value)) setSections([...sections, value]);
    } else if (modalType === 'delete-section' && modalSection) {
      // Remove section and unassign projects
      const s = modalSection;
      setSections(prev => prev.filter(x => x !== s));
      setProjects(prev => prev.map(p => p.section === s ? { ...p, section: undefined } : p));
    } else if (modalType === 'rename-section' && modalSection) {
      const oldName = modalSection;
      const newName = value;
      if (!newName) { closeModal(); return; }
      // Update sections list (replace or move to end if duplicates)
      setSections(prev => {
        const list = prev.filter(x => x !== oldName);
        if (!list.includes(newName)) list.push(newName);
        return list;
      });
      setProjects(prev => prev.map(p => p.section === oldName ? { ...p, section: newName } : p));
    } else if (modalType === 'move-section' && dropTargetSection && dropProjectId) {
      const target = dropTargetSection;
      if (target && !sections.includes(target)) setSections([...sections, target]);
      setProjects(prev => prev.map(p => p.id === dropProjectId ? { ...p, section: target || undefined } : p));
    } else if (modalType === 'delete-project' && modalProjectId) {
      const id = modalProjectId;
      setProjects(prev => prev.filter(p => p.id !== id));
      if (activeProjectId === id) setActiveProjectId(null);
    }
    closeModal();
  };
  const requestDeleteSection = (sec: string) => {
    setModalType('delete-section');
    setModalSection(sec);
    setModalValue('');
  };
  const requestRenameSection = (sec: string) => {
    setModalType('rename-section');
    setModalSection(sec);
    setModalValue(sec);
  };

  const handleSearch = async (searchUid?: string) => {
    const query = (searchUid || uid || "").toString();
    if (!/^\d{11}$/.test(query)) {
      setUidError('Invalid UID. It must contain exactly 11 numbers.');
      return;
    }
    // If currently viewing a saved project, exit project view for a fresh live search
    if (activeProjectId) setActiveProjectId(null);
    
    setUid(query);
    setLoading(true);
    setProgressVisible(true);
    setProgressComplete(false);
    setError(null);
  setData(null);

  // mark start time for adaptive timing
  const t0 = Date.now();

  // Direct Logic App call (no local proxy)
  const directUrl = `https://fibertools-dsavavdcfdgnh2cm.westeurope-01.azurewebsites.net/api/fiberflow/triggers/When_an_HTTP_request_is_received/invoke?api-version=2022-05-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=8KqIymphhOqUAlnd7UGwLRaxP0ot5ZH30b7jWCEUedQ&UID=${encodeURIComponent(query)}`;

    try {
      // Helper to verify JSON response
      const isJson = (r: Response) => /application\/json/i.test(r.headers.get('content-type') || '');

      const res = await fetch(directUrl, { redirect: 'follow' });
      if (!(res.ok && isJson(res))) {
        const text = await res.text().catch(() => '');
        const statusPart = `HTTP ${res.status}`;
        const bodyPart = text ? `: ${text.slice(0, 220)}` : '';
        throw new Error(statusPart + bodyPart);
      }
  const result = await res.json();
  // Parse and attach AllWorkflowStatus (stringified or array)
  try {
    let wfList: any[] = [];
    if (typeof result.AllWorkflowStatus === 'string') {
      const s = result.AllWorkflowStatus.trim();
      if (s.startsWith('[')) {
        wfList = JSON.parse(s);
      }
    } else if (Array.isArray(result.AllWorkflowStatus)) {
      wfList = result.AllWorkflowStatus;
    }
    if (Array.isArray(wfList) && wfList.length) {
      const wfMap: Record<string, string> = {};
      for (const it of wfList) {
        const uid = String(it?.Uid ?? it?.UID ?? it?.uid ?? '').trim();
        if (!uid) continue;
        wfMap[uid] = niceWorkflowStatus(it?.WorkflowStatus);
      }
      try { Object.defineProperty(result, '__AllWorkflowStatus', { value: wfList, enumerable: false }); } catch { (result as any).__AllWorkflowStatus = wfList; }
      try { Object.defineProperty(result, '__WFStatusByUid', { value: wfMap, enumerable: false }); } catch { (result as any).__WFStatusByUid = wfMap; }
    }
  } catch { /* ignore parse errors */ }
  // Stable sorts for consistent UI
  result.OLSLinks?.sort((a: any, b: any) => naturalSort(a.APort, b.APort));
  result.MGFXA?.sort((a: any, b: any) => naturalSort(a.XOMT, b.XOMT));
  result.MGFXZ?.sort((a: any, b: any) => naturalSort(a.XOMT, b.XOMT));
      if (Array.isArray(result.AssociatedUIDs)) {
        // Sort Associated UIDs in ascending order by UID (numeric-aware)
        result.AssociatedUIDs.sort((a: any, b: any) => {
          const uidA = String(a?.UID || a?.Uid || a?.uid || "");
          const uidB = String(b?.UID || b?.Uid || b?.uid || "");
          return uidA.localeCompare(uidB, undefined, { numeric: true });
        });
      }

      setData(result);
      setLastSearched(query);
      if (!history.includes(query)) setHistory([query, ...history]);
    } catch (err: any) {
      // Normalize error messages for clarity
      const msg = String(err?.message || err);
      if (/Failed to fetch/i.test(msg)) {
        setError("Network or CORS error: request did not reach the server.");
      } else {
        setError(msg);
      }
    } finally {
      setLoading(false);
      // Update adaptive estimate with actual duration
      const dt = Date.now() - t0;
      updateExpectedMs(dt);
      // Signal progress bar to accelerate to completion; it will hide itself via onDone
      setProgressComplete(true);
    }
  };

  const pad = (text: string, width: number) => {
    text = text == null ? "" : String(text);
    return text.padEnd(width, " ");
  };

  const copyTableText = (title: string, rows: Record<string, any>[], headers: string[]) => {
    if (!rows?.length) return;
    const colWidths = headers.map((h, i) =>
      Math.max(h.length, ...rows.map((r) => String(Object.values(r)[i] ?? "").length)) + 2
    );

    let output = `${title}\n`;
    output += headers.map((h, i) => pad(h, colWidths[i])).join("") + "\n";
    output += "-".repeat(colWidths.reduce((a, b) => a + b, 0)) + "\n";

    for (const r of rows) {
      const vals = Object.values(r);
      output += vals.map((v, i) => pad(v, colWidths[i])).join("") + "\n";
    }

  navigator.clipboard.writeText(output.trimEnd());
  };

  const formatTableText = (
    title: string,
    rows: Record<string, any>[] | undefined,
    headers: string[]
  ): string => {
    if (!rows || !rows.length) return "";
    const colWidths = headers.map((h, i) =>
      Math.max(h.length, ...rows.map((r) => String(Object.values(r)[i] ?? "").length)) + 2
    );
    let out = `${title}\n`;
    out += headers.map((h, i) => pad(h, colWidths[i])).join("") + "\n";
    out += "-".repeat(colWidths.reduce((a, b) => a + b, 0)) + "\n";
    for (const r of rows) {
      const vals = Object.values(r);
      out += vals.map((v, i) => pad(String(v ?? ""), colWidths[i])).join("") + "\n";
    }
    return out.trimEnd() + "\n\n";
  };

  // Build full plain‑text export of all sections
  const buildAllText = (): string => {
    const dataNow = getViewData();
    if (!dataNow) return "";
    let text = "";

    // Project UIDs (if viewing a saved project snapshot)
    try {
      const srcUids: string[] = Array.isArray((dataNow as any)?.sourceUids) ? (dataNow as any).sourceUids : [];
      if (srcUids.length) {
        const headers = ["UID"];
        const rows = srcUids.map(u => ({ UID: String(u) }));
        text += formatTableText("Project UIDs", rows as any, headers);
      }
    } catch {}

    // Details
    try {
      const rawStatus = String(getWFStatusFor(dataNow, lastSearched) || '').trim();
      const isCancelled = /cancel|cancelled|canceled/i.test(rawStatus);
      const isDecom = /decom/i.test(rawStatus);
      const statusDisplay = isCancelled ? 'WF Cancelled' : isDecom ? 'DECOM' : (rawStatus || '—');
  const jobId = dataNow?.KQLData?.JobId;
      const cisLink = jobId ? `https://azcis.trafficmanager.net/Public/NetworkingOptical/JobDetails/${jobId}` : '';

      const detailsHeaders = ["SRLGID", "SRLG", "SolutionID", "Status", "CIS Workflow"];
      const detailsRows = [
        {
          SRLGID: String(getSrlgIdFrom(dataNow, lastSearched) ?? ""),
          SRLG: String(getSrlgFrom(dataNow) ?? ""),
          SolutionID: (() => getSolutionIds(dataNow).map(formatSolutionId).join(', '))(),
          Status: statusDisplay,
          "CIS Workflow": cisLink,
        },
      ].map((r) => Object.values(r).reduce((acc: any, v: any, i: number) => ({ ...acc, [detailsHeaders[i]]: v }), {}));
      text += formatTableText("Details", detailsRows as any, detailsHeaders);
      // Tools / quick links (A/Z WAN checker + Deployment Validator)
      try {
        const aWan = String(dataNow?.AExpansions?.AUrl || '');
        const aDeploy = String(dataNow?.AExpansions?.AOpticalUrl || '');
        const zWan = String(dataNow?.ZExpansions?.ZUrl || '');
        const zDeploy = String(dataNow?.ZExpansions?.ZOpticalUrl || '');
        const toolsLines: string[] = [];
        if (aWan) toolsLines.push(`A WAN Checker: ${aWan}`);
        if (aDeploy) toolsLines.push(`A Deployment Validator: ${aDeploy}`);
        if (zWan) toolsLines.push(`Z WAN Checker: ${zWan}`);
        if (zDeploy) toolsLines.push(`Z Deployment Validator: ${zDeploy}`);
        if (toolsLines.length) {
          text += `Tools\n` + toolsLines.join('\n') + '\n\n';
        }
      } catch {}
    } catch {}

    // Link Summary
    text += formatTableText(
      "Link Summary",
  dataNow.OLSLinks,
      [
        "A Device",
        "A Port",
        "Z Device",
        "Z Port",
        "A Optical Device",
        "A Optical Port",
        "Z Optical Device",
        "Z Optical Port",
        "Wirecheck",
      ]
    );

    // Associated UIDs (with Workflow Status if available)
    try {
      const rows: any[] = Array.isArray((dataNow as any).AssociatedUIDs) ? (dataNow as any).AssociatedUIDs : [];
      const wfMap: Record<string, string> | undefined = (dataNow as any).__WFStatusByUid;
      const mapped = rows.map((r: any) => {
        const uid = r?.UID ?? r?.Uid ?? r?.uid ?? '';
        const srlg = r?.SrlgId ?? r?.SRLGID ?? r?.SrlgID ?? r?.srlgid ?? '';
        const action = r?.Action ?? r?.action ?? '';
        const type = r?.Type ?? r?.type ?? '';
        const aDev = r['A Device'] ?? r['Device A'] ?? r?.ADevice ?? r?.DeviceA ?? '';
        const zDev = r['Z Device'] ?? r['Device Z'] ?? r?.ZDevice ?? r?.DeviceZ ?? '';
        const siteA = r['Site A'] ?? r?.ASite ?? r?.SiteA ?? r?.Site ?? '';
        const siteZ = r['Site Z'] ?? r?.ZSite ?? r?.SiteZ ?? '';
        const wf = niceWorkflowStatus(wfMap?.[String(uid)]) || '';
        return {
          UID: uid,
          SrlgId: srlg,
          Action: action,
          Type: type,
          'Device A': aDev,
          'Device Z': zDev,
          'Site A': siteA,
          'Site Z': siteZ,
          'WF Status': wf,
        };
      });
      mapped.sort((a: any, b: any) => {
        const aInProg = /in\s*-?\s*progress/i.test(String(a['WF Status']));
        const bInProg = /in\s*-?\s*progress/i.test(String(b['WF Status']));
        if (aInProg !== bInProg) return aInProg ? -1 : 1;
        return String(a.UID).localeCompare(String(b.UID), undefined, { numeric: true });
      });
      text += formatTableText(
        "Associated UIDs",
        mapped as any,
        ["UID", "SrlgId", "Action", "Type", "Device A", "Device Z", "Site A", "Site Z", "WF Status"]
      );
    } catch {
      text += formatTableText(
        "Associated UIDs",
        (dataNow as any).AssociatedUIDs,
        ["UID", "SrlgId", "Action", "Type", "Device A", "Device Z", "Site A", "Site Z"]
      );
    }

    // GDCO Tickets
  text += formatTableText("GDCO Tickets", dataNow.GDCOTickets, ["Ticket Id", "DC Code", "Title", "State", "Assigned To", "Link"]);

    // MGFX A/Z with derived Line column (and without SKU column)
    const mgfxHeaders = [
      "XOMT",
      "C0 Device",
      "C0 Port",
      "Line",
      "M0 Device",
      "M0 Port",
      "C0 DIFF",
      "M0 DIFF",
    ];
    const mapMgfx = (rows: any[]) =>
      (rows || []).map((r: any) => {
        const row: any = { ...(r || {}) };
        const xomt = row["XOMT"] ?? row["xomt"] ?? "";
        const c0Dev = row["C0 Device"] ?? row["C0Device"] ?? row["C0_Device"] ?? "";
        const c0Port = row["C0 Port"] ?? row["C0Port"] ?? row["C0_Port"] ?? "";
        const sku = row["StartHardwareSku"] ?? row["HardwareSku"] ?? row["SKU"] ?? "";
        const line = deriveLineForC0(String(sku || ""), String(c0Port || ""));
        const m0Dev = row["M0 Device"] ?? row["M0Device"] ?? row["M0_Device"] ?? "";
        const m0Port = row["M0 Port"] ?? row["M0Port"] ?? row["M0_Port"] ?? "";
        const c0Diff = row["C0 DIFF"] ?? row["C0_DIFF"] ?? row["C0Diff"] ?? "";
        const m0Diff = row["M0 DIFF"] ?? row["M0_DIFF"] ?? row["M0Diff"] ?? "";
        return {
          "XOMT": xomt,
          "C0 Device": c0Dev,
          "C0 Port": c0Port,
          "Line": line ?? "",
          "M0 Device": m0Dev,
          "M0 Port": m0Port,
          "C0 DIFF": c0Diff,
          "M0 DIFF": m0Diff,
        };
      });
    const mgfxA = mapMgfx(dataNow.MGFXA || []);
    const mgfxZ = mapMgfx(dataNow.MGFXZ || []);
    text += formatTableText("MGFX A-Side", mgfxA, mgfxHeaders);
    text += formatTableText("MGFX Z-Side", mgfxZ, mgfxHeaders);

    return text.trimEnd();
  };

  const copyAll = async () => {
    const text = buildAllText();
    if (!text) return;
  await navigator.clipboard.writeText(text);
  };

  const exportOneNote = async () => {
    const text = buildAllText();
    if (text) {
      try { await navigator.clipboard.writeText(text); } catch {}
    }
    // Open OneNote (web quick note). Content is on clipboard for immediate paste; no alerts shown.
    // If the Windows app is registered, this deep link may open it on some systems:
    // window.location.href = 'onenote:';
    window.open("https://www.onenote.com/quicknote?auth=1", "_blank");
  };

  const exportExcel = () => {
    const dataNow = getViewData();
    if (!dataNow || !(uid || (projects.find(p=>p.id===activeProjectId)?.name))) return;
    const wb = XLSX.utils.book_new();
    // include Details and Tools sheets as well
    const detailsRow = [
      {
        SRLGID: String(getSrlgIdFrom(dataNow, lastSearched) ?? ""),
        SRLG: String(getSrlgFrom(dataNow) ?? ""),
        SolutionID: getSolutionIds(dataNow).map(formatSolutionId).join(', '),
        Status: String(getWFStatusFor(dataNow, lastSearched) || ""),
        CIS_Workflow: dataNow?.KQLData?.JobId ? `https://azcis.trafficmanager.net/Public/NetworkingOptical/JobDetails/${dataNow?.KQLData?.JobId}` : "",
      },
    ];

    const toolsRows = [
      {
        Tool: 'A WAN Checker',
        URL: String(dataNow?.AExpansions?.AUrl || ""),
      },
      {
        Tool: 'A Deployment Validator',
        URL: String(dataNow?.AExpansions?.AOpticalUrl || ""),
      },
      {
        Tool: 'Z WAN Checker',
        URL: String(dataNow?.ZExpansions?.ZUrl || ""),
      },
      {
        Tool: 'Z Deployment Validator',
        URL: String(dataNow?.ZExpansions?.ZOpticalUrl || ""),
      },
    ].filter(r => r.URL);

    // Build Associated UIDs rows with WF Status if available
    const associatedRows = (() => {
      try {
        const rows: any[] = Array.isArray((dataNow as any).AssociatedUIDs) ? (dataNow as any).AssociatedUIDs : [];
        const wfMap: Record<string, string> | undefined = (dataNow as any).__WFStatusByUid;
        const mapped = rows.map((r: any) => {
          const uid = r?.UID ?? r?.Uid ?? r?.uid ?? '';
          const srlg = r?.SrlgId ?? r?.SRLGID ?? r?.SrlgID ?? r?.srlgid ?? '';
          const action = r?.Action ?? r?.action ?? '';
          const type = r?.Type ?? r?.type ?? '';
          const aDev = r['A Device'] ?? r['Device A'] ?? r?.ADevice ?? r?.DeviceA ?? '';
          const zDev = r['Z Device'] ?? r['Device Z'] ?? r?.ZDevice ?? r?.DeviceZ ?? '';
          const siteA = r['Site A'] ?? r?.ASite ?? r?.SiteA ?? r?.Site ?? '';
          const siteZ = r['Site Z'] ?? r?.ZSite ?? r?.SiteZ ?? '';
          const wf = niceWorkflowStatus(wfMap?.[String(uid)]) || '';
          return {
            UID: uid,
            SrlgId: srlg,
            Action: action,
            Type: type,
            'Device A': aDev,
            'Device Z': zDev,
            'Site A': siteA,
            'Site Z': siteZ,
            'WF Status': wf,
          };
        });
        mapped.sort((a: any, b: any) => {
          const aInProg = /in\s*-?\s*progress/i.test(String(a['WF Status']));
          const bInProg = /in\s*-?\s*progress/i.test(String(b['WF Status']));
          if (aInProg !== bInProg) return aInProg ? -1 : 1;
          return String(a.UID).localeCompare(String(b.UID), undefined, { numeric: true });
        });
        return mapped;
      } catch { return (dataNow as any).AssociatedUIDs; }
    })();

    const sections = {
      ...(Array.isArray((dataNow as any)?.sourceUids) && (dataNow as any).sourceUids.length ? {
        "Project UIDs": ((dataNow as any).sourceUids as any[]).map(u => ({ UID: String(u) }))
      } : {}),
      "Details": detailsRow,
      "Tools": toolsRows,
  "Link Summary": dataNow.OLSLinks,
  "Associated UIDs": associatedRows,
      "GDCO Tickets": dataNow.GDCOTickets,
      "MGFX A-Side": dataNow.MGFXA,
      "MGFX Z-Side": dataNow.MGFXZ,
    } as Record<string, any[]>;
    for (const [title, rows] of Object.entries(sections)) {
      if (!Array.isArray(rows) || !rows.length) continue;
      const ws = XLSX.utils.json_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, title.slice(0, 31));
    }
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const rawName = activeProjectId ? `Project_${projects.find(p=>p.id===activeProjectId)?.name || 'UID_Project'}` : `UID_Report_${uid}`;
    const safeName = rawName.replace(/[<>:"/\\|?*]+/g, '_').slice(0, 150);
    saveAs(blob, `${safeName}.xlsx`);
  };

  const Table = ({ title, headers, rows, highlightUid, headerRight }: any) => {
    // Determine keys from first row to ensure consistent ordering and sorting (safe fallback)
    const keys = rows && rows[0] ? Object.keys(rows[0]) : [];

    const [sortKey, setSortKey] = useState<string | null>(null);
    const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");

    const effectiveHeaders = headers && headers.length === keys.length ? headers : keys;

    // Compute sortedRows as a hook (must run unconditionally before any return)
    const sortedRows = React.useMemo(() => {
      if (!sortKey) return rows;
      return [...rows].sort((a: any, b: any) => {
        const va = a[sortKey];
        const vb = b[sortKey];
        if (va == null && vb == null) return 0;
        if (va == null) return sortDir === "asc" ? -1 : 1;
        if (vb == null) return sortDir === "asc" ? 1 : -1;
        const aStr = String(va);
        const bStr = String(vb);
        // numeric-aware compare
        const cmp = aStr.localeCompare(bStr, undefined, { numeric: true, sensitivity: "base" });
        return sortDir === "asc" ? cmp : -cmp;
      });
    }, [rows, sortKey, sortDir]);

  // Keep rendering header even when there are no rows, so controls (like filters) remain accessible
  const noRows = !rows || rows.length === 0;

    

    const toggleSort = (k: string) => {
      if (sortKey === k) setSortDir(sortDir === "asc" ? "desc" : "asc");
      else {
        setSortKey(k);
        setSortDir("asc");
      }
    };

    const findLinkForRow = (row: any) => {
      const linkKey = Object.keys(row).find((k) => k.toLowerCase().includes("link") || k.toLowerCase().includes("url"));
      const fromKey = linkKey ? row[linkKey] : null;
      if (fromKey) return fromKey;
      // allow a non-enumerable hidden link so we can keep Ticket Id clickable without rendering a Link column
      const hidden = (row as any).__hiddenLink;
      return hidden || null;
    };

    const isLinkSummary = title === 'Link Summary';
    const isScrollCandidate = title === 'GDCO Tickets' || title === 'Associated UIDs';
    const shouldScroll = isScrollCandidate && Array.isArray(rows) && rows.length > 10;

    return (
      <div className="table-container">
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text className="section-title">{title}</Text>
          <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
            {headerRight}
            <CopyIconInline onCopy={() => copyTableText(title, rows, effectiveHeaders)} message="Table copied" />
          </div>
        </Stack>
        {noRows ? (
          <div style={{ padding: '8px 0', color: '#a6b7c6' }}>No rows to display.</div>
        ) : (
          <div style={shouldScroll ? { maxHeight: 360, overflowY: 'auto', marginTop: 4 } : undefined}>
        <table className="data-table">
          <thead>
            <tr>
              {effectiveHeaders.map((h: string, i: number) => {
                const k = keys[i] ?? h;
                const active = sortKey === k;
                const isStatusMini = isLinkSummary && (/admin|oper|state/i.test(String(k)) || /admin|state/i.test(String(h)));
                return (
                  <th
                    key={i}
                    onClick={() => toggleSort(k)}
                    style={{
                      cursor: 'pointer',
                      userSelect: 'none',
                      textAlign: isStatusMini ? 'center' : undefined,
                      width: isStatusMini ? 24 : undefined,
                      minWidth: isStatusMini ? 24 : undefined,
                    }}
                  >
                    <span>{h}</span>
                    <span style={{ marginLeft: 6 }}>{active ? (sortDir === 'asc' ? '▲' : '▼') : '↕'}</span>
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {sortedRows.map((row: any, i: number) => {
              const uidKey = keys.find((k) => k.toLowerCase() === 'uid');
              const uidVal = uidKey ? row[uidKey] : undefined;
              const highlight = highlightUid && String(uidVal ?? '') === highlightUid;
              return (
                <tr key={i} className={highlight ? 'highlight-row' : ''}>
                  {keys.map((key: string, j: number) => {
                    const val = row[key];

                    // Render Admin/Oper arrow indicators compactly
                    if (isLinkSummary && /admin|oper|state/i.test(String(key))) {
                      const v = String(val ?? '').trim();
                      const isUp = v === '1' || v.toLowerCase() === 'up' || v === 'true';
                      const isDown = v === '0' || v.toLowerCase() === 'down' || v === 'false';
                      return (
                        <td key={j} style={{ textAlign: 'center', width: 24, minWidth: 24 }} title={isUp ? 'Up' : isDown ? 'Down' : String(val ?? '')}>
                          <span style={{ color: isUp ? '#107c10' : isDown ? '#d13438' : '#a6b7c6', fontWeight: 800, fontSize: 12, lineHeight: '14px' }}>
                            {isUp ? '▲' : isDown ? '▼' : ''}
                          </span>
                        </td>
                      );
                    }

                    // If column is a link-like field, show Open + Copy (include Wirecheck as link)
                    {
                      const keyLower = String(key).toLowerCase();
                      const headerLower = String(effectiveHeaders[j] || '').toLowerCase();
                      const looksLikeLink = ['workflow', 'diff', 'ticketlink', 'url', 'link', 'wirecheck'].some(s => keyLower.includes(s) || headerLower.includes(s));
                      if (looksLikeLink) {
                        const link = val;
                        return (
                          <td key={j}>
                            {link ? (
                              <>
                                <button className="open-btn" onClick={() => window.open(link, '_blank')}>Open</button>
                                <CopyIconInline onCopy={() => { navigator.clipboard.writeText(String(link)); }} message="Link copied" />
                              </>
                            ) : null}
                          </td>
                        );
                      }
                    }

                    // Special: Associated/Project UIDs clicking behavior
                    if ((title === 'Associated UIDs' || title === 'Project UIDs') && key.toLowerCase() === 'uid') {
                      const v = val;
                      return (
                        <td key={j}>
                                <span
                                  className="uid-click"
                                  onClick={() => {
                                    // Always open the clicked UID in a new tab so the current results remain
                                    const url = `${window.location.pathname}?uid=${encodeURIComponent(String(v))}`;
                                    window.open(url, '_blank');
                                  }}
                                  title={`Search UID ${v}`}
                                >
                                  {v}
                                </span>
                        </td>
                      );
                    }

                    // Special: Colored WF Status badge in Associated UIDs
                    if (title === 'Associated UIDs' && (String(key).toLowerCase() === 'wf status' || String(effectiveHeaders[j]).toLowerCase() === 'wf status')) {
                      const s = String(val ?? '').trim();
                      const isCancelled = /cancel|cancelled|canceled/i.test(s);
                      const isDecom = /decom/i.test(s);
                      const isFinished = /wf\s*finished|finished/i.test(s);
                      const isInProgress = /in\s*-?\s*progress|running/i.test(s);
                      const display = s || '—';
                      if (isFinished) {
                        return (
                          <td key={j} style={{ textAlign: 'center' }}>
                            <span
                              className="wf-finished-badge wf-finished-pulse"
                              style={{
                                color: '#00c853',
                                fontWeight: 900,
                                fontSize: 12,
                                padding: '2px 8px',
                                borderRadius: 10,
                                border: '1px solid rgba(0,200,83,0.45)'
                              }}
                            >
                              {display}
                            </span>
                          </td>
                        );
                      }
                      if (isInProgress) {
                        return (
                          <td key={j} style={{ textAlign: 'center' }}>
                            <span
                              className="wf-inprogress-badge wf-inprogress-pulse"
                              style={{
                                color: '#50b3ff',
                                fontWeight: 800,
                                fontSize: 11,
                                padding: '1px 6px',
                                borderRadius: 10,
                                border: '1px solid rgba(80,179,255,0.28)'
                              }}
                            >
                              {display}
                            </span>
                          </td>
                        );
                      }
                      const color = (isCancelled || isDecom) ? '#d13438' : '#a6b7c6';
                      const border = (isCancelled || isDecom) ? '1px solid rgba(209,52,56,0.45)' : '1px solid rgba(166,183,198,0.35)';
                      return (
                        <td key={j} style={{ textAlign: 'center' }}>
                          <span style={{ color, fontWeight: 700, fontSize: 12, padding: '2px 8px', borderRadius: 10, border }}>{display}</span>
                        </td>
                      );
                    }

                    // If this is a Ticket Id cell, try to hyperlink to the ticket URL if available
                    if (String(key).toLowerCase().includes('ticket') || String(effectiveHeaders[j]).toLowerCase().includes('ticket')) {
                      const link = findLinkForRow(row);
                      if (link) {
                        return (
                          <td key={j}>
                            <a className="uid-click" href={String(link)} target="_blank" rel="noopener noreferrer">{val}</a>
                            {title !== 'GDCO Tickets' && (
                              <button className="open-btn" onClick={() => window.open(String(link), '_blank')}>Open</button>
                            )}
                          </td>
                        );
                      }
                    }

                    // Default cell
                    return <td key={j}>{val}</td>;
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
          </div>
        )}
      </div>
    );
  };

  const isInitialView = !lastSearched && !loading && !data && !activeProjectId;

  // removed unused flipped animation state

  // Absolute positioning removed in favor of a responsive flex row

  // flip the circle every 3s
  // removed flip interval effect

  // Position circle: align its LEFT edge to the first table's LEFT edge; align TOP to Summary top
  useEffect(() => {
    const computeCircle = () => {
      const viewData = getViewData();
      if (!viewData) return;
      const tableEl = firstTableRef.current as HTMLElement | null;
      const summaryEl = summaryContainerRef.current as HTMLElement | null;
      if (!tableEl) return;
      const mainEl = (summaryEl || tableEl).closest('.main') as HTMLElement | null;
      if (!mainEl) return;
      const tableRect = tableEl.getBoundingClientRect();
      const mainRect = mainEl.getBoundingClientRect();
      const left = Math.max(12, tableRect.left - mainRect.left);
      // Align top to the summary container if present; fallback to a small offset
      let top = 12;
      if (summaryEl) {
        const sRect = summaryEl.getBoundingClientRect();
        top = Math.max(12, sRect.top - mainRect.top);
      }
      setCapacityLeft(left);
      setCapacityTop(top);
    };
    computeCircle();
    window.addEventListener('resize', computeCircle);
    return () => window.removeEventListener('resize', computeCircle);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data, activeProjectId]);

  // Ensure AI Summary starts AFTER the circle: add a dynamic left margin if needed
  useEffect(() => {
    const onResize = () => {
      try { setIsWide(window.innerWidth >= 1400); } catch {}
    };
    const adjustSummary = () => {
      const CIRCLE = 140; const GAP = 16;
      if (!isWide) {
        // On non-wide layouts, keep a constant offset so the card never sits under the circle
        setSummaryShift(CIRCLE + GAP);
        return;
      }
      const el = summaryContainerRef.current as HTMLElement | null;
      const mainEl = el?.closest('.main') as HTMLElement | null;
      if (!el || !mainEl) { setSummaryShift(CIRCLE + GAP); return; }
      const sRect = el.getBoundingClientRect();
      const mRect = mainEl.getBoundingClientRect();
      const desiredLeft = (capacityLeft ?? 12) + CIRCLE + GAP;
      const currentLeft = sRect.left - mRect.left;
      const delta = Math.max(0, Math.round(desiredLeft - currentLeft));
      setSummaryShift(delta);
    };
    adjustSummary();
    window.addEventListener('resize', adjustSummary);
    window.addEventListener('resize', onResize);
    return () => {
      window.removeEventListener('resize', adjustSummary);
      window.removeEventListener('resize', onResize);
    };
  }, [capacityLeft, data, activeProjectId, lastSearched, isWide]);

  // compute capacity values for the reusable component
  const viewData = getViewData();
  const capacity = (() => {
    if (!viewData) return null;
    const linksArr: any[] = Array.isArray(viewData.OLSLinks) ? viewData.OLSLinks : [];
    // If no link rows, synthesize a single placeholder when KQLData provides devices, so the circle shows Increment capacity
    const hasKdDevices = !!(viewData?.KQLData?.DeviceA || viewData?.KQLData?.DeviceZ);
    const effectiveLinks = linksArr.length ? linksArr : (hasKdDevices ? [{}] : []);
    return computeCapacity(effectiveLinks, viewData?.KQLData?.Increment, viewData?.KQLData?.DeviceA);
  })();

  // Extract SolutionID(s): prefer KQLData.SolutionId; fallback to any solutionId-like fields in KQLData/OLSLinks
  const getSolutionIds = (src: any): string[] => {
    try {
      const out = new Set<string>();
      if (!src) return [];
      const normKey = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, "");
      const tryAdd = (v: any) => {
        if (v == null) return;
        const s = String(v).trim();
        if (s) out.add(s);
      };
      const kql = src?.KQLData || {};
      for (const [k, v] of Object.entries(kql)) {
        const nk = normKey(k);
        if (nk === 'solutionid') tryAdd(v);
        if (Array.isArray(v) && nk === 'solutionids') (v as any[]).forEach(tryAdd);
        if (!Array.isArray(v) && /solution/i.test(k) && /id/i.test(k)) tryAdd(v);
      }
      const rows: any[] = Array.isArray(src?.OLSLinks) ? src.OLSLinks : [];
      for (const r of rows) {
        for (const [k, v] of Object.entries(r || {})) {
          const nk = normKey(k);
          if (nk === 'solutionid' || (/solution/i.test(k) && /id/i.test(k))) tryAdd(v);
        }
      }
      return Array.from(out);
    } catch {
      return [];
    }
  };
  const formatSolutionId = (s: string) => {
    const t = String(s || '').trim();
    if (!t) return '';
    return /^sls-/i.test(t) ? `SLS-${t.slice(4)}` : `SLS-${t}`;
  };
  const solutionIdDisplay = (viewData ? getSolutionIds(viewData) : [])
    .map(formatSolutionId)
    .filter(Boolean)
    .join(', ');

  // Troubleshooting section component (collapsible, interactive per-link tracking)
  const TroubleshootingSection: React.FC<{ contextKey: string; rows: any[] }> = ({ contextKey, rows }) => {
    const STORE_KEY = `${contextKey}:troubles`;
    const COLLAPSE_KEY = `${contextKey}:troublesCollapsed`;
  type TItem = { note?: string; notes?: Array<{ id: string; text: string }>; color?: string; done?: boolean };
    const [map, setMap] = useState<Record<string, TItem>>(() => {
      try { const raw = localStorage.getItem(STORE_KEY); const obj = raw ? JSON.parse(raw) : {}; return obj && typeof obj === 'object' ? obj : {}; } catch { return {}; }
    });
    const [collapsed, setCollapsed] = useState<boolean>(() => {
      try { const raw = localStorage.getItem(COLLAPSE_KEY); return raw == null ? true : raw === '1'; } catch { return true; }
    });
  useEffect(() => { try { localStorage.setItem(STORE_KEY, JSON.stringify(map)); } catch {} }, [map, STORE_KEY]);
  useEffect(() => { try { localStorage.setItem(COLLAPSE_KEY, collapsed ? '1' : '0'); } catch {} }, [collapsed, COLLAPSE_KEY]);

    const normalize = (r: any) => {
      const aDev = r["ADevice"] ?? r["A Device"] ?? r["DeviceA"] ?? r["Device A"] ?? '';
      const aPort = r["APort"] ?? r["A Port"] ?? r["PortA"] ?? r["Port A"] ?? '';
      const aOptDev = r["AOpticalDevice"] ?? r["A Optical Device"] ?? '';
      const aOptPort = r["AOpticalPort"] ?? r["A Optical Port"] ?? '';
      const zDev = r["ZDevice"] ?? r["Z Device"] ?? r["DeviceZ"] ?? r["Device Z"] ?? '';
      const zPort = r["ZPort"] ?? r["Z Port"] ?? r["PortZ"] ?? r["Port Z"] ?? '';
      const zOptDev = r["ZOpticalDevice"] ?? r["Z Optical Device"] ?? '';
      const zOptPort = r["ZOpticalPort"] ?? r["Z Optical Port"] ?? '';
      return { aDev, aPort, aOptDev, aOptPort, zDev, zPort, zOptDev, zOptPort };
    };
    const keyFor = (r: any) => {
      const n = normalize(r);
      return `${n.aDev}|${n.aPort}|${n.zDev}|${n.zPort}`;
    };
    const setField = (id: string, patch: Partial<TItem>) => {
      setMap(prev => ({ ...prev, [id]: { ...(prev[id] || {}), ...patch } }));
    };
    const clearRow = (id: string) => {
      setMap(prev => { const next = { ...prev }; delete next[id]; return next; });
    };

    const colorStyle = (c?: string): React.CSSProperties => {
      if (!c) return {};
      const bg = c === 'yellow' ? '#3a3a00' : c === 'orange' ? '#442a00' : c === 'red' ? '#4d1f1f' : c === 'blue' ? '#0d2a4d' : c === 'purple' ? '#3a1f4d' : '';
      const border = c === 'yellow' ? '#b3a100' : c === 'orange' ? '#b36b00' : c === 'red' ? '#b33a3a' : c === 'blue' ? '#3b7bd6' : c === 'purple' ? '#9159c1' : '';
      return bg ? { background: bg, border: `1px solid ${border}` } : {};
    };

    if (!Array.isArray(rows) || rows.length === 0) return null;
    return (
      <div className="notes-card" style={{ marginTop: 12 }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text className="section-title">Troubleshooting</Text>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <button
              className="sleek-btn"
              style={{ padding: '4px 10px', fontSize: 12, background: '#2b2b2b', color: '#e6f6ff', border: '1px solid #3a4a5e' }}
              onClick={() => setCollapsed(c => !c)}
            >
              {collapsed ? 'Expand' : 'Collapse'}
            </button>
          </div>
        </Stack>
        {!collapsed && (
          <div style={{ marginTop: 8, display: 'flex', flexDirection: 'column', gap: 8 }}>
            {rows.map((r: any, idx: number) => {
              const id = keyFor(r);
              const n = normalize(r);
              const item = map[id] || {};
              const done = !!item.done;
              const rowStyle: React.CSSProperties = done
                ? { background: '#0f3d24', border: '1px solid #2e7d32' }
                : colorStyle(item.color);
              return (
                <div key={id || idx} className="table-container" style={{ padding: 8, ...rowStyle }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
                    <div style={{ lineHeight: 1.3 }}>
                      <div>
                        <b style={{ color: '#cfe7ff' }}>{n.aDev}</b>
                        <span style={{ opacity: 0.6 }}> · </span>
                        <span>{n.aPort}</span>
                        <span style={{ margin: '0 10px', opacity: 0.8 }}>⇄</span>
                        <b style={{ color: '#cfe7ff' }}>{n.zDev}</b>
                        <span style={{ opacity: 0.6 }}> · </span>
                        <span>{n.zPort}</span>
                      </div>
                      <div style={{ fontSize: 12, color: '#e6f6ff', marginTop: 4 }}>
                        <span style={{ fontWeight: 700, color: '#9fd1ff' }}>A Optical:</span>
                        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, marginLeft: 6 }}>
                          <span style={{ background: '#15324f', border: '1px solid #295a86', borderRadius: 10, padding: '1px 8px' }}>{n.aOptDev || '—'}</span>
                          {n.aOptPort ? <span style={{ opacity: 0.85 }}>{n.aOptPort}</span> : null}
                        </span>
                        <span style={{ margin: '0 10px', opacity: 0.55 }}>|</span>
                        <span style={{ fontWeight: 700, color: '#9fd1ff' }}>Z Optical:</span>
                        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, marginLeft: 6 }}>
                          <span style={{ background: '#15324f', border: '1px solid #295a86', borderRadius: 10, padding: '1px 8px' }}>{n.zOptDev || '—'}</span>
                          {n.zOptPort ? <span style={{ opacity: 0.85 }}>{n.zOptPort}</span> : null}
                        </span>
                      </div>
                    </div>
                    {/* Middle: output area for added notes (multiple) */}
                    <div style={{ flex: 1, minWidth: 200, display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                      {(() => {
                        // Back-compat: if legacy single note exists, show it as a chip
                        const notesArr: Array<{ id: string; text: string }> =
                          (item.notes && Array.isArray(item.notes))
                            ? item.notes
                            : (item.note ? [{ id: `legacy-${idx}`, text: item.note }] : []);
                        return notesArr.map((nObj) => (
                          <span
                            key={nObj.id}
                            style={{
                              position: 'relative',
                              background: '#1f2d3a',
                              border: '1px solid #335c8a',
                              color: '#e8f0ff',
                              padding: '4px 22px 4px 10px',
                              borderRadius: 12,
                              fontSize: 13,
                              fontWeight: 600,
                              maxWidth: 520,
                              overflow: 'hidden',
                              textOverflow: 'ellipsis',
                              whiteSpace: 'nowrap',
                              boxShadow: 'inset 0 0 6px rgba(255,255,255,0.03)'
                            }}
                            title={nObj.text}
                          >
                            {nObj.text}
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                const existing = (item.notes && Array.isArray(item.notes)) ? item.notes : (item.note ? [{ id: `legacy-${idx}`, text: item.note }] : []);
                                const filtered = existing.filter(n => n.id !== nObj.id);
                                setField(id, { notes: filtered, note: undefined });
                              }}
                              aria-label="Delete note"
                              title="Delete note"
                              style={{
                                position: 'absolute',
                                top: 2,
                                right: 2,
                                width: 16,
                                height: 16,
                                borderRadius: 999,
                                background: 'rgba(58,74,94,0.9)',
                                color: '#e8f0ff',
                                border: '1px solid #2b3a4e',
                                cursor: 'pointer',
                                fontSize: 11,
                                lineHeight: '14px',
                                display: 'inline-flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                zIndex: 1
                              }}
                            >
                              ×
                            </button>
                          </span>
                        ));
                      })()}
                    </div>
                    {/* Right: input + color / done / clear */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 320, justifyContent: 'flex-end' }}>
                      <input
                        className="projects-filter-input"
                        placeholder="Add note… (Enter)"
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') {
                            const val = (e.currentTarget as HTMLInputElement).value.trim();
                            if (val) {
                              const newNote = { id: `${Date.now()}-${Math.random().toString(36).slice(2,8)}`, text: val };
                              const existing = (item.notes && Array.isArray(item.notes)) ? item.notes : (item.note ? [{ id: `legacy-${idx}`, text: item.note }] : []);
                              setField(id, { notes: [...existing, newNote], note: undefined });
                              (e.currentTarget as HTMLInputElement).value = '';
                            }
                          }
                        }}
                        style={{ minWidth: 160, width: 180 }}
                        title="Type a note and press Enter"
                      />
                      <select
                        className="sleek-select"
                        value={item.color || ''}
                        onChange={(e) => setField(id, { color: e.target.value || undefined })}
                        title="Highlight color"
                      >
                        <option value="">No highlight</option>
                        <option value="yellow">Yellow</option>
                        <option value="orange">Orange</option>
                        <option value="red">Red</option>
                        <option value="blue">Blue</option>
                        <option value="purple">Purple</option>
                      </select>
                      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }} title="Mark as complete">
                        <input type="checkbox" checked={!!item.done} onChange={(e) => setField(id, { done: e.target.checked })} />
                        <span>Done</span>
                      </label>
                      <button className="sleek-btn" style={{ padding: '4px 10px', fontSize: 12, background: '#444' }} onClick={() => clearRow(id)}>Clear</button>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    );
  };

  // Timestamp helpers
  const getTimestamp = (obj: any): string | null => {
    if (!obj) return null;
    // 1) Top-level or KQLData
    const direct = obj.TIMESTAMP || obj.Timestamp || obj.timestamp || obj?.KQLData?.TIMESTAMP || obj?.KQLData?.Timestamp || obj?.KQLData?.timestamp;
    if (direct) return String(direct);
    // 2) Common collections (e.g., some sources attach TIMESTAMP per row)
    const collections = [obj.OLSLinks, obj.AssociatedUIDs, obj.MGFXA, obj.MGFXZ, obj.GDCOTickets];
    for (const coll of collections) {
      if (Array.isArray(coll) && coll.length) {
        const candidate = coll.find((r: any) => r?.TIMESTAMP || r?.Timestamp || r?.timestamp) || coll[0];
        const val = candidate?.TIMESTAMP || candidate?.Timestamp || candidate?.timestamp;
        if (val) return String(val);
      }
    }
    return null;
  };
  const formatTimestamp = (ts: string | null | undefined): string | null => {
    if (!ts) return null;
    try {
      const d = new Date(ts);
      if (isNaN(d.getTime())) return ts;
      return d.toLocaleString();
    } catch { return ts; }
  };

  return (
  <Stack className="main" style={{ position: 'relative' }}>
      <Dialog
        hidden={!showCancelDialog}
        onDismiss={() => setShowCancelDialog(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: cancelDialogTitle,
        }}
        modalProps={{ isBlocking: true }}
      >
        <div style={{ textAlign: 'center', padding: '12px 20px', fontSize: 14 }}>{cancelDialogMsg}</div>
        <DialogFooter>
          {cancelDialogLink && (
            <PrimaryButton onClick={() => window.open(cancelDialogLink!, '_blank')} text="Open CIS WF" />
          )}
          <DefaultButton onClick={() => setShowCancelDialog(false)} text="Dismiss" />
        </DialogFooter>
      </Dialog>
      {viewData && (
        <div className="uid-top-inline">
          <div
            ref={summaryContainerRef}
            className="table-container combined-horizontal"
            style={{ marginLeft: summaryShift, position: 'relative', zIndex: 2 }}
          >
            <div className="combined-inner">
              <UIDSummaryPanel
                data={viewData}
                currentUid={(() => {
                  if (activeProjectId) {
                    const ap = getActiveProject();
                    const first = (ap?.data?.sourceUids || [])[0] || null;
                    return first || (lastSearched || null);
                  }
                  return lastSearched || null;
                })()}
                bare
              />
              <UIDStatusPanel uid={lastSearched || null} data={viewData} bare />
            </div>
          </div>
        </div>
      )}

      {/* Restore capacity circle to original absolute placement (aligned to Details table) */}
      {viewData && (
        <div
          style={{
            position: 'absolute',
            left: capacityLeft ?? 40,
            top: capacityTop ?? 8,
            pointerEvents: 'none',
            filter: 'drop-shadow(0 0 2px rgba(0,120,212,0.12))',
            zIndex: 50,
          }}
        >
          <CapacityCircle main={capacity?.main ?? '?'} size={140} />
        </div>
      )}
      {/* Removed the second Links circle per request; CapacityCircle remains */}
      {isInitialView ? (
      <div className="vso-form-container glow" style={{ width: "80%", maxWidth: 800 }}>
        <div className="banner-title">
          <span className="title-text">UID Assistant</span>
          <span className="title-sub">The Ultimate UID Assistant Tool</span>
        </div>

  <div style={{ display: "flex", gap: 10, alignItems: "center", justifyContent: "center" }}>
          <TextField
            placeholder="Enter UID (e.g., 20190610163)"
            value={uid}
            onChange={(_e, v) => {
              const cleaned = (v ?? "").replace(/\D/g, "").slice(0, 11);
              setUid(cleaned);
              setUidError(() => {
                if (!cleaned) return null;
                return cleaned.length === 11 ? null : 'Invalid UID. It must contain exactly 11 numbers.';
              });
            }}
            className="input-field"
            inputMode="numeric"
            pattern="[0-9]*"
            onPaste={handleUidPaste}
            onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); handleSearch(); } }}
          />
          <PrimaryButton
            text="Search"
            disabled={loading}
            onClick={() => handleSearch()}
            className="search-btn"
            style={{ marginLeft: 20 }}
          />
        </div>

        {uidError && (
          <div style={{ display: 'flex', justifyContent: 'center', marginTop: 6 }}>
            <div className="uid-inline-error" aria-live="polite" style={{ width: 220 }}>{uidError}</div>
          </div>
        )}

        {history.length === 0 && (
          <div style={{ marginTop: 8, textAlign: "center", fontSize: 12, color: "#aaa" }}>
            First time here?{' '}
            <span className="uid-click" onClick={() => handleSearch('20190610161')}>
              Try now
            </span>
          </div>
        )}

        {lastSearched && (
          <Text className="last-searched" style={{ marginTop: 6 }}>
            Last searched:{" "}
            <span className="uid-click" onClick={() => handleSearch(lastSearched)}>
              {lastSearched}
            </span>
          </Text>
        )}

        {history.length > 0 && (
          <div style={{ marginTop: 6, color: "#aaa", fontSize: 12 }}>
            Recent: {history.slice(0, 5).map((h, i) => (
              <span
                key={h}
                className="uid-click"
                style={{ marginLeft: i === 0 ? 0 : 10 }}
                onClick={() => handleSearch(h)}
              >
                {h}
              </span>
            ))}
          </div>
        )}


      </div>
      ) : null}

      {progressVisible && (
        <ThemedProgressBar
          active={progressVisible}
          complete={progressComplete}
          expectedMs={expectedMsEstimate}
          label="Fetching data…"
          onDone={() => setProgressVisible(false)}
          style={{ marginTop: 6 }}
        />
      )}

      <div className="last-searched-gap" />

      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}

      {viewData && (
        <>
          {/* Projects toolbar */}
          <div className="projects-toolbar">
            <div className="toolbar-left" />
            <div className="toolbar-center">
              {/* Centered UID search */}
              <div style={{ display: 'inline-flex', gap: 8, alignItems: 'center' }}>
                <div style={{ width: 220 }}>
                  <TextField
                    placeholder="Enter UID (e.g., 20190610163)"
                    value={uid}
                    onChange={(_e, v) => {
                      const cleaned = (v ?? "").replace(/\D/g, "").slice(0, 11);
                      setUid(cleaned);
                      setUidError(() => {
                        if (!cleaned) return null;
                        return cleaned.length === 11 ? null : 'Invalid UID. It must contain exactly 11 numbers.';
                      });
                    }}
                    className="input-field"
                    inputMode="numeric"
                    pattern="[0-9]*"
                    styles={{ fieldGroup: { width: 220 } }}
                    onPaste={handleUidPaste}
                    onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); handleSearch(); } }}
                  />
                  {uidError && (
                    <div className="uid-inline-error" style={{ marginTop: 4 }} aria-live="polite">{uidError}</div>
                  )}
                </div>
                <PrimaryButton
                  text="Search"
                  disabled={loading}
                  onClick={() => handleSearch()}
                  className="search-btn"
                />
              </div>
            </div>
            <div className="toolbar-right">
              {/* Only allow creating/adding when viewing live results */}
              {!activeProjectId && data && (
                <>
                  <button className="sleek-btn repo" onClick={createProjectFromCurrent}>Create Project</button>
                  <button className="sleek-btn repo" onClick={addSection} title="Create a personal section">New Section</button>
                  {projects.length > 0 && (
                    <>
                      <select
                        className="sleek-select"
                        onChange={(e) => {
                          const id = e.target.value; if (id) addCurrentToProject(id);
                          // reset selection to placeholder
                          e.currentTarget.selectedIndex = 0;
                        }}
                      >
                        <option value="">Add to project…</option>
                        {projects.map((p) => (
                          <option key={p.id} value={p.id}>{p.name}{p.section ? `  •  ${p.section}` : ''}</option>
                        ))}
                      </select>
                    </>
                  )}
                </>
              )}
              {activeProjectId && (
                <>
                  <span style={{ color: '#a6b7c6', fontSize: 12 }}>Viewing project:</span>
                  {(() => {
                    const p = projects.find(pp=>pp.id===activeProjectId);
                    const typeLabel = p ? getProjectType(p) : null;
                    const typeCls = typeLabel ? (typeLabel.toLowerCase().includes('hybrid') ? 'hybrid' : typeLabel.toLowerCase().includes('owned') ? 'owned' : 'standard') : 'standard';
                    return (
                      <>
                        <span className="uid-click" onClick={() => setActiveProjectId(null)} title="Exit project view">{p?.name}</span>
                        {typeLabel && (<span className={`proj-type-badge ${typeCls}`} style={{ marginLeft: 6 }}>{typeLabel}</span>)}
                        {p?.urgent ? (<span className="proj-urgent-badge" style={{ marginLeft: 6 }}>Urgent</span>) : null}
                      </>
                    );
                  })()}
                  <button className="sleek-btn" style={{ background:'#444' }} onClick={() => setActiveProjectId(null)}>Exit</button>
                </>
              )}
            </div>
          </div>
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginBottom: 8 }}>
            <CopyIconInline onCopy={copyAll} message="All sections copied" />
            <IconButton iconProps={{ iconName: 'ExcelLogo' }} title="Export to Excel" ariaLabel="Export to Excel" onClick={exportExcel} />
            <IconButton iconProps={{ iconName: 'OneNoteLogo' }} title="Export to OneNote" ariaLabel="Export to OneNote" onClick={exportOneNote} />
            <IconButton iconProps={{ iconName: 'Info' }} title="Tip: Use Copy All to capture this report, then paste in OneNote." ariaLabel="Tip" styles={{ root: { transform: 'scale(0.9)', opacity: 0.7 } }} />
          </div>

          {/* Project UIDs (always shown for saved projects) — single line with " | " separators */}
          {activeProjectId && (() => {
            const ap = getActiveProject();
            const uids: string[] = Array.from(new Set(ap?.data?.sourceUids || [])).filter(Boolean);
            const listText = uids.join(' | ');
            return (
              <div className="table-container" style={{ marginBottom: 12 }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                  <Text className="section-title">Project UIDs</Text>
                  <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
                    <CopyIconInline onCopy={() => { if (listText) navigator.clipboard.writeText(listText); }} message="UIDs copied" />
                  </div>
                </Stack>
                <div style={{ padding: '6px 0' }}>
                  {uids.length > 0 ? (
                    <div style={{ whiteSpace: 'nowrap', overflowX: 'auto' }}>
                      {uids.map((u, i) => (
                        <React.Fragment key={u}>
                          <span
                            className="uid-click"
                            onClick={() => {
                              const url = `${window.location.pathname}?uid=${encodeURIComponent(String(u))}`;
                              window.open(url, '_blank');
                            }}
                            title={`Open ${u}`}
                          >
                            {u}
                          </span>
                          {i < uids.length - 1 && <span style={{ opacity: 0.6, margin: '0 6px' }}> | </span>}
                        </React.Fragment>
                      ))}
                    </div>
                  ) : (
                    <div style={{ color: '#a6b7c6' }}>No UIDs</div>
                  )}
                </div>
              </div>
            );
          })()}
          
          {/* Details Section */}
          <div className="table-container details-fit" ref={firstTableRef}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Text className="section-title">Details</Text>
            </Stack>
            <table className="data-table details-table">
              <thead>
                <tr>
                  <th>SRLGID</th>
                  <th>SRLG</th>
                  <th>SolutionID</th>
                  <th style={{ textAlign: 'center' }}>Status</th>
                  <th>CIS Workflow</th>
                  <th>Repository</th>
                  <th>Fiber Planner</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td>{getSrlgIdFrom(viewData, lastSearched) || ''}</td>
                  <td>{getSrlgFrom(viewData) || ''}</td>
                  <td>{solutionIdDisplay || '—'}</td>
                  <td style={{ textAlign: 'center' }}>
                    {(() => {
                      const raw = String(getWFStatusFor(viewData, lastSearched) || '').trim();
                      const isCancelled = /cancel|cancelled|canceled/i.test(raw);
                      const isDecom = /decom/i.test(raw);
                      const isFinished = /wffinished|wf finished|finished/i.test(raw);
                      const isInProgress = /inprogress|in progress|in-progress|running/i.test(raw);
                      const display = isCancelled
                        ? 'WF Cancelled'
                        : isDecom
                        ? 'DECOM'
                        : isFinished
                        ? 'WF Finished'
                        : isInProgress
                        ? 'WF In Progress'
                        : (raw || '—');
                      if (isFinished) {
                        return (
                          <span
                            className="wf-finished-badge wf-finished-pulse"
                            style={{
                              color: '#00c853',
                              fontWeight: 900,
                              fontSize: 13,
                              padding: '4px 8px',
                              borderRadius: 10,
                              border: '1px solid rgba(0,200,83,0.45)'
                            }}
                          >
                            {display}
                          </span>
                        );
                      }
                      if (isInProgress) {
                        return (
                          <span
                            className="wf-inprogress-badge wf-inprogress-pulse"
                            style={{
                              color: '#50b3ff',
                              fontWeight: 900,
                              fontSize: 13,
                              padding: '4px 8px',
                              borderRadius: 10,
                              border: '1px solid rgba(80,179,255,0.28)'
                            }}
                          >
                            {display}
                          </span>
                        );
                      }
                      const color = (isCancelled || isDecom) ? '#d13438' : '#107c10';
                      return <span style={{ color, fontWeight: 600 }}>{display}</span>;
                    })()}
                  </td>
                  <td>
                    {(() => {
                      const jobId = viewData?.KQLData?.JobId;
                      const link = jobId ? `https://azcis.trafficmanager.net/Public/NetworkingOptical/JobDetails/${jobId}` : null;
                      return link ? (
                        <>
                          <button
                            className="sleek-btn repo"
                            onClick={() => window.open(link, '_blank')}
                          >
                            Open Workflow
                          </button>
                          <CopyIconInline onCopy={() => { navigator.clipboard.writeText(String(link)); }} message="Link copied" />
                        </>
                      ) : null;
                    })()}
                  </td>
                  <td>
                    {(() => {
                      const repoUid = (lastSearched || (Array.isArray((viewData as any)?.sourceUids) ? (viewData as any).sourceUids[0] : '')) as string;
                      const repoLink = repoUid ? `https://microsoft.sharepoint.com/teams/WAN-Capacity/Shared%20Documents/Cabling/${encodeURIComponent(repoUid)}` : null;
                      return repoLink ? (
                        <>
                          <button
                            className="sleek-btn repo"
                            onClick={() => window.open(repoLink, "_blank")}
                          >
                            WAN Capacity Repository
                          </button>
                          <CopyIconInline onCopy={() => { navigator.clipboard.writeText(String(repoLink)); }} message="Link copied" />
                        </>
                      ) : null;
                    })()}
                  </td>
                  <td>
                    {(() => {
                      const sites = getFirstSites(viewData, lastSearched || undefined);
                      const a = (sites.a || '').toString().trim();
                      const z = (sites.z || '').toString().trim();
                      const label = a && z ? `${a} ↔ ${z} KMZ Route` : a ? `${a} KMZ Route` : z ? `${z} KMZ Route` : 'KMZ Route';
                      const srlgId = getSrlgIdFrom(viewData, lastSearched);
                      if (!srlgId) return null;
                      const fp = `https://fiberplanner.cloudg.is/?srlg=${encodeURIComponent(String(srlgId))}`;
                      return (
                        <>
                          <button
                            className="sleek-btn kmz"
                            onClick={() => window.open(fp, "_blank")}
                          >
                            {label}
                          </button>
                          <CopyIconInline onCopy={() => { navigator.clipboard.writeText(fp); }} message="Link copied" />
                        </>
                      );
                    })()}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* WAN Buttons */}
          <div className="button-header-align-left">
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <Text className="side-label">A Side:</Text>
                  {viewData?.AExpansions?.AUrl && (
                    <>
                      <button
                        className="sleek-btn wan"
                        onClick={() => window.open(viewData?.AExpansions?.AUrl, "_blank")}
                      >
                        WAN Checker
                      </button>
                      <CopyIconInline onCopy={() => { navigator.clipboard.writeText(String(viewData?.AExpansions?.AUrl)); }} message="Link copied" />
                    </>
                  )}
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  {viewData?.AExpansions?.AOpticalUrl && (
                    <>
                      <button
                        className="sleek-btn optical"
                        onClick={() => window.open(viewData?.AExpansions?.AOpticalUrl, "_blank")}
                      >
                        Deployment Validator
                      </button>
                      <CopyIconInline onCopy={() => { navigator.clipboard.writeText(String(viewData?.AExpansions?.AOpticalUrl)); }} message="Link copied" />
                    </>
                  )}
                </div>
              </div>

              <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                <Text className="side-label">Z Side</Text>
                <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  {viewData?.ZExpansions?.ZUrl && (
                    <>
                      <button
                        className="sleek-btn wan"
                        onClick={() => window.open(viewData?.ZExpansions?.ZUrl, "_blank")}
                      >
                        WAN Checker
                      </button>
                      <CopyIconInline onCopy={() => { navigator.clipboard.writeText(String(viewData?.ZExpansions?.ZUrl)); }} message="Link copied" />
                    </>
                  )}
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  {viewData?.ZExpansions?.ZOpticalUrl && (
                    <>
                      <button
                        className="sleek-btn optical"
                        onClick={() => window.open(viewData?.ZExpansions?.ZOpticalUrl, "_blank")}
                      >
                        Deployment Validator
                      </button>
                      <CopyIconInline onCopy={() => { navigator.clipboard.writeText(String(viewData?.ZExpansions?.ZOpticalUrl)); }} message="Link copied" />
                    </>
                  )}
                </div>
              </div>
            </div>
          </div>

          {/* Tables */}
          <Table
            title="Link Summary"
            headers={[
              "A Device",
              "A Port",
              "Admin",
              "State",
              "A Optical Device",
              "A Optical Port",
              "Z Device",
              "Z Port",
              "Admin",
              "State",
              "Z Optical Device",
              "Z Optical Port",
              "Wirecheck",
            ]}
            rows={(() => {
              const links: any[] = Array.isArray(viewData.OLSLinks) ? viewData.OLSLinks : [];
              // If there are no link rows, synthesize a single fallback row from KQLData
              if (!links.length) {
                const kd = (viewData as any)?.KQLData || {};
                const aDev = String(kd?.DeviceA ?? '').trim();
                const zDev = String(kd?.DeviceZ ?? '').trim();
                // Do NOT use LagA/LagZ for ports; ports and LAGs are not the same. Leave ports blank.
                const aPort = '';
                const zPort = '';
                if (aDev || zDev) {
                  return [
                    {
                      "A Device": aDev,
                      "A Port": aPort,
                      "A Admin": '',
                      "A Oper": '',
                      "A Optical Device": '',
                      "A Optical Port": '',
                      "Z Device": zDev,
                      "Z Port": zPort,
                      "Z Admin": '',
                      "Z Oper": '',
                      "Z Optical Device": '',
                      "Z Optical Port": '',
                      "Wirecheck": '',
                    },
                  ];
                }
                // No fallback data available
                return [];
              }

              return links.map((r: any) => {
              // Map directly from canonical keys you provided, with safe fallbacks
              const aDevRaw = r["ADevice"] ?? r["A Device"] ?? r["A_Device"] ?? r["DeviceA"] ?? r["Device A"] ?? "";
              const aPort = r["APort"] ?? r["A Port"] ?? r["A_Port"] ?? r["PortA"] ?? r["Port A"] ?? "";
              const zDevRaw = r["ZDevice"] ?? r["Z Device"] ?? r["Z_Device"] ?? r["DeviceZ"] ?? r["Device Z"] ?? "";
              const zPort = r["ZPort"] ?? r["Z Port"] ?? r["Z_Port"] ?? r["PortZ"] ?? r["Port Z"] ?? "";
              const aOptDev = r["AOpticalDevice"] ?? r["A Optical Device"] ?? r["A_Optical_Device"] ?? r["A OpticalDevice"] ?? r["A Optical"] ?? "";
              const aOptPort = r["AOpticalPort"] ?? r["A Optical Port"] ?? r["A_Optical_Port"] ?? r["A OpticalPort"] ?? "";
              const zOptDev = r["ZOpticalDevice"] ?? r["Z Optical Device"] ?? r["Z_Optical_Device"] ?? r["Z OpticalDevice"] ?? r["Z Optical"] ?? "";
              const zOptPort = r["ZOpticalPort"] ?? r["Z Optical Port"] ?? r["Z_Optical_Port"] ?? r["Z OpticalPort"] ?? "";
              const workflow = r["Workflow"] ?? r["workflow"] ?? r["Link"] ?? r["link"] ?? r["URL"] ?? r["Url"] ?? "";

              // Fallback to KQLData DeviceA/DeviceZ only if per-row device fields are blank
              const aDev = (String(aDevRaw ?? '').trim() || String(viewData?.KQLData?.DeviceA ?? '').trim());
              const zDev = (String(zDevRaw ?? '').trim() || String(viewData?.KQLData?.DeviceZ ?? '').trim());

              // Admin/Oper status for A/Z sides (support multiple possible key names; fallback to global AdminStatus/OperStatus)
              const aAdmin = r["AAdminStatus"] ?? r["AdminStatusA"] ?? r["AdminStatus_A"] ?? r["A_AdminStatus"] ?? r["A AdminStatus"] ?? r["AdminStatus"] ?? '';
              const aOper = r["AOperStatus"] ?? r["OperStatusA"] ?? r["OperStatus_A"] ?? r["A_OperStatus"] ?? r["A OperStatus"] ?? r["OperStatus"] ?? '';
              const zAdmin = r["ZAdminStatus"] ?? r["AdminStatusZ"] ?? r["AdminStatus_Z"] ?? r["Z_AdminStatus"] ?? r["Z AdminStatus"] ?? r["AdminStatus"] ?? '';
              const zOper = r["ZOperStatus"] ?? r["OperStatusZ"] ?? r["OperStatus_Z"] ?? r["Z_OperStatus"] ?? r["Z OperStatus"] ?? r["OperStatus"] ?? '';

              // Return object with keys matching header order; no swaps
              return {
                "A Device": aDev,
                "A Port": aPort,
                "A Admin": aAdmin,
                "A Oper": aOper,
                "A Optical Device": aOptDev,
                "A Optical Port": aOptPort,
                "Z Device": zDev,
                "Z Port": zPort,
                "Z Admin": zAdmin,
                "Z Oper": zOper,
                "Z Optical Device": zOptDev,
                "Z Optical Port": zOptPort,
                "Wirecheck": workflow,
              };
              });
            })()}
            headerRight={(() => {
              const ts = formatTimestamp(getTimestamp(viewData));
              return ts ? (
                <span style={{ color: '#a6b7c6', fontSize: 12 }}>Last refresh: <b style={{ color: '#d0e7ff' }}>{ts}</b></span>
              ) : null;
            })()}
          />

          <Stack horizontal tokens={{ childrenGap: 20 }} styles={{ root: { width: '100%' } }} className="equal-tables-row">
            <Table
              title="Associated UIDs"
              headers={[
                "UID",
                "SrlgId",
                "Action",
                "Type",
                "Device A",
                "Device Z",
                "Site A",
                "Site Z",
                "WF Status",
              ]}
              rows={(() => {
                const rows = Array.isArray(viewData.AssociatedUIDs) ? viewData.AssociatedUIDs : [];
                const wfMap: Record<string, string> | undefined = (viewData as any).__WFStatusByUid;
                const mapped = rows.map((r: any) => {
                  const uid = r?.UID ?? r?.Uid ?? r?.uid ?? '';
                  const srlg = r?.SrlgId ?? r?.SRLGID ?? r?.SrlgID ?? r?.srlgid ?? '';
                  const action = r?.Action ?? r?.action ?? '';
                  const type = r?.Type ?? r?.type ?? '';
                  const aDev = r['A Device'] ?? r['Device A'] ?? r?.ADevice ?? r?.DeviceA ?? '';
                  const zDev = r['Z Device'] ?? r['Device Z'] ?? r?.ZDevice ?? r?.DeviceZ ?? '';
                  const siteA = r['Site A'] ?? r?.ASite ?? r?.SiteA ?? r?.Site ?? '';
                  const siteZ = r['Site Z'] ?? r?.ZSite ?? r?.SiteZ ?? '';
                  const wf = niceWorkflowStatus(wfMap?.[String(uid)]) || '';
                  return {
                    UID: uid,
                    SrlgId: srlg,
                    Action: action,
                    Type: type,
                    'Device A': aDev,
                    'Device Z': zDev,
                    'Site A': siteA,
                    'Site Z': siteZ,
                    'WF Status': wf,
                  };
                });
                // Apply default filter: only show In Progress unless toggled to show all
                const base = showAllAssociatedWF
                  ? mapped
                  : mapped.filter((r: any) => /in\s*-?\s*progress|running/i.test(String(r['WF Status'])));
                // Always show In Progress rows at the top by default
                base.sort((a: any, b: any) => {
                  const aInProg = /in\s*-?\s*progress/i.test(String(a['WF Status']));
                  const bInProg = /in\s*-?\s*progress/i.test(String(b['WF Status']));
                  if (aInProg !== bInProg) return aInProg ? -1 : 1;
                  // tie-breaker: by UID (numeric-aware)
                  return String(a.UID).localeCompare(String(b.UID), undefined, { numeric: true });
                });
                return base;
              })()}
              headerRight={(
                <button
                  className="sleek-btn repo"
                  onClick={() => setShowAllAssociatedWF(v => !v)}
                  title={showAllAssociatedWF ? 'Show only In Progress' : 'Show all UIDs'}
                >
                  {showAllAssociatedWF ? 'Show In Progress only' : 'Show All'}
                </button>
              )}
              highlightUid={uid}
            />
            <Table
              title="GDCO Tickets"
              headers={["Ticket Id", "DC Code", "Title", "State", "Assigned To"]}
              rows={(() => {
                const rows = Array.isArray(viewData.GDCOTickets) ? viewData.GDCOTickets : [];
                // Remove any rows that are effectively empty (no visible fields populated)
                const nonEmpty = rows.filter((r: any) => {
                  if (!r || typeof r !== 'object') return false;
                  const visibleFields = [
                    'Ticket Id','TicketId','TicketID',
                    'DC Code','DCCode','Datacenter','DC',
                    'Title',
                    'State','Status',
                    'Assigned To','AssignedTo','Owner'
                  ];
                  return visibleFields.some((k) => {
                    const v = (r as any)[k];
                    return v != null && String(v).trim() !== '';
                  });
                });
                return nonEmpty.map((r: any) => {
                  const { Link, link, URL, Url, TicketLink, ticketLink, ...rest } = r || {};
                  const l = r?.Link || r?.link || r?.URL || r?.Url || r?.TicketLink || r?.ticketLink || null;
                  const obj: any = { ...rest };
                  if (l) {
                    try { Object.defineProperty(obj, '__hiddenLink', { value: l, enumerable: false }); } catch {}
                  }
                  return obj;
                });
              })()}
            />
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 20 }} styles={{ root: { width: '100%' } }}>
            <Table
              title="MGFX A-Side"
              headers={[
                "XOMT",
                "C0 Device",
                "C0 Port",
                "Line",
                "M0 Device",
                "M0 Port",
                "C0 DIFF",
                "M0 DIFF",
              ]}
              rows={(viewData.MGFXA || []).map((r: any) => {
                const row: any = { ...(r || {}) };
                const xomt = row["XOMT"] ?? row["xomt"] ?? "";
                const c0Dev = row["C0 Device"] ?? row["C0Device"] ?? row["C0_Device"] ?? "";
                const c0Port = row["C0 Port"] ?? row["C0Port"] ?? row["C0_Port"] ?? "";
                const sku = row["StartHardwareSku"] ?? row["HardwareSku"] ?? row["SKU"] ?? "";
                const line = deriveLineForC0(String(sku || ""), String(c0Port || ""));
                const m0Dev = row["M0 Device"] ?? row["M0Device"] ?? row["M0_Device"] ?? "";
                const m0Port = row["M0 Port"] ?? row["M0Port"] ?? row["M0_Port"] ?? "";
                const c0Diff = row["C0 DIFF"] ?? row["C0_DIFF"] ?? row["C0Diff"] ?? "";
                const m0Diff = row["M0 DIFF"] ?? row["M0_DIFF"] ?? row["M0Diff"] ?? "";

                // Return object with keys in desired order; omit SKU
                return {
                  "XOMT": xomt,
                  "C0 Device": c0Dev,
                  "C0 Port": c0Port,
                  "Line": line ?? "",
                  "M0 Device": m0Dev,
                  "M0 Port": m0Port,
                  "C0 DIFF": c0Diff,
                  "M0 DIFF": m0Diff,
                };
              })}
            />
            <Table
              title="MGFX Z-Side"
              headers={[
                "XOMT",
                "C0 Device",
                "C0 Port",
                "Line",
                "M0 Device",
                "M0 Port",
                "C0 DIFF",
                "M0 DIFF",
              ]}
              rows={(viewData.MGFXZ || []).map((r: any) => {
                const row: any = { ...(r || {}) };
                const xomt = row["XOMT"] ?? row["xomt"] ?? "";
                const c0Dev = row["C0 Device"] ?? row["C0Device"] ?? row["C0_Device"] ?? "";
                const c0Port = row["C0 Port"] ?? row["C0Port"] ?? row["C0_Port"] ?? "";
                const sku = row["StartHardwareSku"] ?? row["HardwareSku"] ?? row["SKU"] ?? "";
                const line = deriveLineForC0(String(sku || ""), String(c0Port || ""));
                const m0Dev = row["M0 Device"] ?? row["M0Device"] ?? row["M0_Device"] ?? "";
                const m0Port = row["M0 Port"] ?? row["M0Port"] ?? row["M0_Port"] ?? "";
                const c0Diff = row["C0 DIFF"] ?? row["C0_DIFF"] ?? row["C0Diff"] ?? "";
                const m0Diff = row["M0 DIFF"] ?? row["M0_DIFF"] ?? row["M0Diff"] ?? "";

                return {
                  "XOMT": xomt,
                  "C0 Device": c0Dev,
                  "C0 Port": c0Port,
                  "Line": line ?? "",
                  "M0 Device": m0Dev,
                  "M0 Port": m0Port,
                  "C0 DIFF": c0Diff,
                  "M0 DIFF": m0Diff,
                };
              })}
            />
          </Stack>

          {/* Notes / Chatbox (per UID) */}
          {lastSearched && !activeProjectId && (
            <div className="notes-card">
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Text className="section-title">Notes</Text>
              </Stack>

              <div className="notes-input-row">
                <textarea
                  className="notes-textarea"
                  placeholder={"Add a note for this UID..."}
                  value={noteText}
                  onChange={(e) => setNoteText(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && !e.shiftKey) {
                      e.preventDefault();
                      addNote();
                    }
                  }}
                  rows={3}
                />
                <div className="notes-user-hint">
                  {(() => {
                    const email = getEmail();
                    const alias = getAlias(email);
                    return email ? (
                      <span>Posting as <b style={{ color: '#c9ffd8' }}>{alias}</b></span>
                    ) : (
                      <span style={{ color: '#a6b7c6' }}>Sign in to post with your alias.</span>
                    );
                  })()}
                </div>
              </div>

              <div className="notes-list">
                {notes.length === 0 ? (
                  <div className="note-empty">No notes yet for this UID.</div>
                ) : (
                  notes.map((n) => (
                    <div key={n.id} className="note-item">
                      <div className="note-header">
                        <div className="note-meta">
                          <span className="note-alias">{n.authorAlias || 'guest'}</span>
                          {n.authorEmail && <span className="note-email">@{(n.authorEmail.split('@')[1] || '').split('.')[0]}</span>}
                          <span className="note-dot">·</span>
                          <span className="note-time">{new Date(n.ts).toLocaleString()}</span>
                        </div>
                        <div className="note-controls" style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                          {canModify(n) && (
                            editingId === n.id ? (
                              <>
                                <button className="note-btn save" onClick={saveEdit}>Save</button>
                                <button className="note-btn" onClick={cancelEdit}>Cancel</button>
                              </>
                            ) : (
                              <>
                                <button className="note-btn" onClick={() => startEdit(n)}>Edit</button>
                              </>
                            )
                          )}
                          <button
                            className="note-btn danger"
                            onClick={() => removeNote(n.id)}
                            disabled={!n._rk || deletingNoteId === n.id}
                            title={n._rk ? undefined : 'Still syncing to server; try again in a moment'}
                          >
                            {deletingNoteId === n.id ? 'Deleting...' : 'Delete'}
                          </button>
                        </div>
                      </div>
                      <div className="note-body">
                        {editingId === n.id ? (
                          <textarea
                            className="notes-textarea inline-edit"
                            rows={3}
                            value={editingText}
                            onChange={(e) => setEditingText(e.target.value)}
                          />
                        ) : (
                          <div className="note-text">{n.text}</div>
                        )}
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          )}

          {/* Troubleshooting (below UID notes) */}
          {lastSearched && !activeProjectId && (
            <TroubleshootingSection contextKey={`uid:${lastSearched}`} rows={Array.isArray(viewData?.OLSLinks) ? viewData.OLSLinks : []} />
          )}

          {/* Project Notes (when viewing a saved project) */}
          {activeProjectId && (() => {
            const ap = getActiveProject();
            if (!ap) return null;
            const uids: string[] = Array.from(new Set(ap.data?.sourceUids || []));
            const amap: Record<string, Note[]> = ap.notes || {} as any;
            const allNotes: Note[] = uids.flatMap(uid => (amap[uid] || []).map(n => ({ ...n, uid })));
            allNotes.sort((a, b) => b.ts - a.ts);
            return (
              <div className="notes-card">
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                  <Text className="section-title">Notes</Text>
                </Stack>

                <div className="notes-input-row">
                  <textarea
                    className="notes-textarea"
                    placeholder={uids.length ? `Add a note to UID ${projTargetUid || uids[0]}...` : 'No UIDs in this project yet'}
                    value={projNoteText}
                    onChange={(e) => setProjNoteText(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' && !e.shiftKey) {
                        e.preventDefault();
                        addProjectNote();
                      }
                    }}
                    rows={3}
                    disabled={!uids.length}
                  />
                  <div className="notes-user-hint">
                    <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                      <span style={{ color: '#a6b7c6' }}>Target UID:</span>
                      <select className="sleek-select" value={projTargetUid || ''} onChange={(e) => setProjTargetUid(e.target.value || null)}>
                        {uids.map(u => <option key={u} value={u}>{u}</option>)}
                      </select>
                      {(() => { const email = getEmail(); const alias = getAlias(email); return email ? (<span>Posting as <b style={{ color: '#c9ffd8' }}>{alias}</b></span>) : (<span style={{ color: '#a6b7c6' }}>Sign in to post with your alias.</span>); })()}
                    </div>
                  </div>
                </div>

                <div className="notes-list">
                  {allNotes.length === 0 ? (
                    <div className="note-empty">No notes yet for this project.</div>
                  ) : (
                    allNotes.map((n) => (
                      <div key={n.id} className="note-item">
                        <div className="note-header">
                          <div className="note-meta">
                            <span className="note-uid" style={{ color: '#9fe9b8', fontWeight: 800 }}>{n.uid}</span>
                            <span className="note-dot">·</span>
                            <span className="note-alias">{n.authorAlias || 'guest'}</span>
                            {n.authorEmail && <span className="note-email">@{(n.authorEmail.split('@')[1] || '').split('.')[0]}</span>}
                            <span className="note-dot">·</span>
                            <span className="note-time">{new Date(n.ts).toLocaleString()}</span>
                          </div>
                        </div>
                        <div className="note-body">
                          <div className="note-text">{n.text}</div>
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>
            );
          })()}

          {/* Troubleshooting (below Project notes) */}
          {activeProjectId && (
            <TroubleshootingSection contextKey={`project:${activeProjectId}`} rows={Array.isArray(viewData?.OLSLinks) ? viewData.OLSLinks : []} />
          )}
        </>
      )}

      {/* Projects rail (visible on UID Assistant only) */}
      <div
        className={`projects-rail ${railCollapsed ? 'collapsed' : ''}`}
        style={railCollapsed ? undefined : { width: railWidth }}
        aria-label="Projects menu"
        role="complementary"
      >
        <div className="projects-rail-header">
          <span className="projects-rail-title">Projects</span>
          <button
            className="rail-toggle-btn"
            title={railCollapsed ? 'Expand projects' : 'Collapse projects'}
            aria-label={railCollapsed ? 'Expand projects' : 'Collapse projects'}
            onClick={() => setRailCollapsed(v => !v)}
          >
            {railCollapsed ? '›' : '‹'}
          </button>
        </div>
        {!railCollapsed && (
          <div
            className="projects-rail-resizer"
            title="Resize projects menu"
            aria-label="Resize projects menu"
            role="separator"
            aria-orientation="vertical"
            tabIndex={0}
            onMouseDown={onRailDragStart}
            onKeyDown={onRailKeyResize}
          />
        )}
        <div className="projects-rail-list">
          {(() => {
            if (projects.length === 0) return <div className="projects-rail-empty">No projects yet</div>;
            // Sort: pinned first, then by section (custom order: sections[] then unsectioned), then newest first
            const sectionOrder = new Map<string, number>();
            sections.forEach((s, i) => sectionOrder.set(s, i));
            const sorted = [...projects].sort((a, b) => {
              const ap = a.pinned ? 0 : 1;
              const bp = b.pinned ? 0 : 1;
              if (ap !== bp) return ap - bp;
              const as = a.section ? sectionOrder.get(a.section) ?? 9999 : 1e9;
              const bs = b.section ? sectionOrder.get(b.section) ?? 9999 : 1e9;
              if (as !== bs) return as - bs;
              return b.createdAt - a.createdAt;
            });

            // Optional text filter
            const visiblePre = sorted.filter(p => {
              const f = (projectFilter || '').trim().toLowerCase();
              if (!f) return true; // default shows all projects
              return String(p.name || '').toLowerCase().includes(f);
            });
            // Apply viewerSection filter (if chosen): match by section name
            const visible = viewerSection
              ? visiblePre.filter(p => (p.section || '') === viewerSection)
              : visiblePre;

            // Group by section
            const bySection = new Map<string, Project[]>();
            const keyFor = (s?: string) => (s && s.trim()) ? s.trim() : 'Archives';
            for (const p of visible) {
              const k = keyFor(p.section);
              const arr = bySection.get(k) || [];
              arr.push(p);
              bySection.set(k, arr);
            }

            const orderedSectionKeys: string[] = [];
            for (const s of sections) if (bySection.has(s)) orderedSectionKeys.push(s);
            if (bySection.has('Archives')) orderedSectionKeys.push('Archives');
            Array.from(bySection.keys()).forEach((k) => { if (!orderedSectionKeys.includes(k) && k !== 'Archives') orderedSectionKeys.push(k); });

            return (
              <>
                <div className="projects-rail-filter">
                  <input
                    className="projects-filter-input"
                    placeholder="Filter projects…"
                    value={projectFilter}
                    onChange={(e) => setProjectFilter(e.target.value)}
                  />
                  <div className="projects-filter-controls" style={{ justifyContent: 'space-between' }}>
                    <div style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                      <button
                        className="proj-navigate-btn tiny-icon-btn"
                        title="Previous person"
                        onClick={() => {
                          if (!sections.length) return;
                          if (!viewerSection) { setViewerSection(sections[0]); return; }
                          const idx = Math.max(0, sections.indexOf(viewerSection));
                          const next = (idx - 1 + sections.length) % sections.length;
                          setViewerSection(sections[next]);
                        }}
                        aria-label="Previous person"
                      >
                        ‹
                      </button>
                      <span style={{ color: '#cfe7ff', fontSize: 12, whiteSpace: 'nowrap' }}>
                        {viewerSection ? `Viewing: ${viewerSection}` : 'Viewing: All projects'}
                      </span>
                      <button
                        className="proj-navigate-btn tiny-icon-btn"
                        title="Next person"
                        onClick={() => {
                          if (!sections.length) return;
                          if (!viewerSection) { setViewerSection(sections[0]); return; }
                          const idx = Math.max(0, sections.indexOf(viewerSection));
                          const next = (idx + 1) % sections.length;
                          setViewerSection(sections[next]);
                        }}
                        aria-label="Next person"
                      >
                        ›
                      </button>
                    </div>
                    <button
                      className="sleek-btn repo"
                      style={{ padding: '4px 10px', fontSize: 12 }}
                      onClick={() => setViewerSection(null)}
                      title="Show all projects"
                    >
                      Show all
                    </button>
                  </div>
                </div>
                {(() => {
                  // When viewing a specific person (section), hide all other sections
                  const orderedSectionKeys: string[] = [];
                  if (viewerSection) {
                    orderedSectionKeys.push(viewerSection);
                  } else {
                    // Show all sections: ensure named sections first, then Archives, then any ad-hoc sections
                    sections.forEach(s => { if (!orderedSectionKeys.includes(s)) orderedSectionKeys.push(s); });
                    if (!orderedSectionKeys.includes('Archives')) orderedSectionKeys.push('Archives');
                    Array.from(bySection.keys()).forEach((k) => { if (!orderedSectionKeys.includes(k)) orderedSectionKeys.push(k); });
                  }

                  return orderedSectionKeys.map((sec, idx) => (
                    <div
                      key={sec}
                      className={`projects-rail-section ${hoveredSection===sec ? 'dropping' : ''}`}
                      onDragOver={(e) => { e.preventDefault(); }}
                      onDragEnter={() => setHoveredSection(sec)}
                      onDragLeave={(e) => {
                        // only clear if truly leaving the section
                        if ((e.currentTarget as any).contains(e.relatedTarget)) return;
                        setHoveredSection(prev => prev===sec ? null : prev);
                      }}
                      onDrop={(e) => {
                        e.preventDefault();
                        setHoveredSection(null);
                        const pid = e.dataTransfer.getData('text/plain') || dragProjectId;
                        if (!pid) return;
                        setDropProjectId(pid);
                        setDropTargetSection(sec === 'Archives' ? '' : sec);
                        setModalType('move-section');
                      }}
                    >
                      {idx > 0 && <div className="projects-rail-divider" aria-hidden />}
                      <div className="projects-rail-section-title">
                        <button
                          className="section-toggle"
                          title={collapsedSections.includes(sec) ? 'Expand section' : 'Collapse section'}
                          onClick={() => {
                            setCollapsedSections(prev => prev.includes(sec) ? prev.filter(x => x !== sec) : [...prev, sec]);
                          }}
                        >
                          {collapsedSections.includes(sec) ? '▸' : '▾'}
                        </button>
                        <span className="section-name" title={sec}>{sec}</span>
                        {sec !== 'Archives' && (
                          <span className="section-actions">
                            <button className="section-action" title="Rename section" onClick={() => requestRenameSection(sec)}>✎</button>
                            <button className="section-action del" title="Delete section" onClick={() => requestDeleteSection(sec)}>×</button>
                          </span>
                        )}
                      </div>
                      {!collapsedSections.includes(sec) && (() => {
                        const items = bySection.get(sec) || [];
                        if (!items.length) return <div className="projects-rail-empty-mini">No projects in this section</div>;
                        return items.map((p) => (
                          <div
                            key={p.id}
                            className={`projects-rail-item ${activeProjectId===p.id ? 'active' : ''} ${p.pinned ? 'pinned' : ''} ${p.urgent ? 'urgent' : ''}`}
                            draggable
                            onDragStart={(e) => { setDragProjectId(p.id); e.dataTransfer.setData('text/plain', p.id); }}
                            onDragEnd={() => setDragProjectId(null)}
                            onClick={() => { if (dragProjectId) return; setActiveProjectId(p.id); }}
                            role="button"
                            tabIndex={0}
                            onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setActiveProjectId(p.id); } }}
                          >
                            <div className="projects-rail-main">
                              <div className="projects-rail-name" title={p.name}>{p.name}</div>
                              {p.owners && p.owners.length > 0 && (
                                <div className="projects-rail-owners" title={p.owners.join(', ')}>{p.owners.join('\n')}</div>
                              )}
                              <div className="projects-rail-badges">
                                {(() => {
                                  const typeLabel = getProjectType(p);
                                  const typeCls = typeLabel.toLowerCase().includes('hybrid') ? 'hybrid' : typeLabel.toLowerCase().includes('owned') ? 'owned' : 'standard';
                                  return <span className={`proj-type-badge ${typeCls}`} title="Project Type">{typeLabel}</span>;
                                })()}
                                {p.urgent && (
                                  <span className="proj-urgent-badge" title="Marked urgent">Urgent</span>
                                )}
                              </div>
                              <div className="projects-rail-metrics">
                                {(() => {
                                  const dataNow = p.data || {};
                                  const cap = computeCapacity(dataNow.OLSLinks, dataNow?.KQLData?.Increment, dataNow?.KQLData?.DeviceA) as any;
                                  const dist = cap?.distribution || '';
                                  const count = cap?.count || 0;
                                  const total = cap?.main?.replace(/[^0-9G]/g,'') || '';
                                  return (
                                    <div className="proj-metric-line">
                                      <span className="proj-dist">{dist ? `${dist} Links` : `${count} link${count===1?'':'s'}`}</span>
                                      <span className="proj-sep">|</span>
                                      <span className="proj-cap">{total} CAP</span>
                                    </div>
                                  );
                                })()}
                                {(() => {
                                  const d = getProjectExpectedDelivery(p);
                                  return (
                                    <div className="proj-date-line">
                                      <span className="proj-date-label">Expected Delivery:</span>
                                      <span className="proj-date-value">{d || '—'}</span>
                                    </div>
                                  );
                                })()}
                              </div>
                              <div className="projects-rail-subrow">
                                <div className="projects-rail-sub">{p.data?.sourceUids?.length || 1} UID(s)</div>
                                <div className="projects-rail-actions-inline" onClick={(e) => e.stopPropagation()}>
                                  <button className="proj-action" title="Rename" onClick={() => renameProject(p.id)}>✎</button>
                                  <button className="proj-action" title="Owners" onClick={() => editOwners(p.id)}>👤</button>
                                  <button className={`proj-action pin ${p.pinned ? 'on' : ''}`} title={p.pinned ? 'Unpin' : 'Pin'} onClick={() => togglePin(p.id)}>★</button>
                                  <button className={`proj-action urgent ${p.urgent ? 'on' : ''}`} title={p.urgent ? 'Unmark urgent' : 'Mark urgent'} onClick={() => toggleUrgent(p.id)}>!</button>
                                  <button className="proj-action del" title="Delete project" onClick={() => requestDeleteProject(p.id)}>×</button>
                                </div>
                              </div>
                            </div>
                          </div>
                        ));
                      })()}
                    </div>
                  ));
                })()}
              </>
            );
          })()}
        </div>
      </div>

      {/* Generic input modal for project actions */}
      <Dialog
        hidden={!modalType}
        onDismiss={closeModal}
        dialogContentProps={{
          type: DialogType.normal,
          title:
            modalType === 'rename' ? 'Rename Project' :
            modalType === 'owners' ? 'Set Owners' :
            modalType === 'section' ? 'Assign Section' :
            modalType === 'new-section' ? 'New Section' :
            modalType === 'delete-section' ? 'Delete Section' :
            modalType === 'rename-section' ? 'Rename Section' :
            modalType === 'move-section' ? 'Move to Section' :
            modalType === 'delete-project' ? 'Delete Project' :
            modalType === 'create-project' ? 'Create Project' :
            modalType === 'confirm-merge' ? 'Merge projects with different SolutionID?' :
            'Action',
        }}
        modalProps={{ isBlocking: true }}
      >
        <div style={{ padding: '4px 2px' }}>
          {modalType === 'owners' ? (
            <TextField
              label="Owners"
              description="Enter one name per line (or separate with commas)."
              multiline
              rows={4}
              value={modalValue}
              onChange={(_e, v) => setModalValue(v || '')}
              autoFocus
            />
          ) : modalType === 'delete-section' ? (
            <div style={{ color: '#e6f6ff', lineHeight: 1.5 }}>
              Are you sure you want to delete the section <b style={{ color: '#ffd180' }}>{modalSection}</b>?<br/>
              Projects in this section will not be deleted; they will simply be unassigned from this section.
            </div>
          ) : modalType === 'move-section' ? (
            <div style={{ color: '#e6f6ff', lineHeight: 1.5 }}>
              Move the selected project to section <b style={{ color: '#c9ffd8' }}>{dropTargetSection || 'Archives'}</b>?
            </div>
          ) : modalType === 'delete-project' ? (
            <div style={{ color: '#e6f6ff', lineHeight: 1.5 }}>
              Are you sure you want to delete this project?
            </div>
          ) : modalType === 'create-project' ? (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
              <div style={{ color: '#cfe7ff' }}>
                Choose a section for this project, or create a new one.
              </div>
              <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  <label style={{ color: '#a6b7c6', fontSize: 12 }}>Existing section</label>
                  <select
                    className="sleek-select"
                    value={createSectionChoice}
                    onChange={(e) => { setCreateSectionChoice(e.target.value); setCreateError(null); }}
                    style={{ minWidth: 200 }}
                  >
                    <option value="">Choose…</option>
                    <option value="Archives">Archives</option>
                    {sections.map((s) => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                  </select>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  <label style={{ color: '#a6b7c6', fontSize: 12 }}>Or new section</label>
                  <input
                    className="projects-filter-input"
                    placeholder="Enter section name"
                    value={createNewSection}
                    onChange={(e) => { setCreateNewSection(e.target.value); setCreateError(null); }}
                    style={{ minWidth: 220 }}
                  />
                </div>
              </div>
              {createError && <div style={{ color: '#ff9e9e', fontSize: 12 }}>{createError}</div>}
            </div>
          ) : modalType === 'confirm-merge' ? (
            <div style={{ color: '#e6f6ff', lineHeight: 1.6 }}>
              {(() => {
                const p = projects.find(pp => pp.id === modalProjectId!);
                const curS = (getSolutionIds(data) || []).map(formatSolutionId).filter(Boolean);
                const projS = p ? (getSolutionIds(p.data) || []).map(formatSolutionId).filter(Boolean) : [];
                const curOne = curS.length ? curS[0] : '—';
                const projOne = projS.length ? projS[0] : '—';
                return (
                  <>
                    <div style={{ marginBottom: 8 }}>
                      The current UID&apos;s Solution ID (
                      <b style={{ color: '#a8f3c9' }}>{curOne}</b>
                      ) differs from the selected project&apos;s (
                      <b style={{ color: '#ff9e9e' }}>{projOne}</b>
                      ).
                    </div>
                    <div>
                      Are you sure you want to merge? You may receive mixed output across different solutions.
                    </div>
                  </>
                );
              })()}
            </div>
          ) : (
            <TextField
              label={modalType === 'rename' ? 'Title' : 'Section Name'}
              value={modalValue}
              onChange={(_e, v) => setModalValue(v || '')}
              autoFocus
            />
          )}
        </div>
        <DialogFooter>
          <PrimaryButton onClick={saveModal} text={(modalType === 'delete-section' || modalType === 'delete-project') ? 'Delete' : modalType === 'move-section' ? 'Move' : modalType === 'create-project' ? 'Create' : modalType === 'confirm-merge' ? 'Merge' : 'Save'} />
          <DefaultButton onClick={closeModal} text="Cancel" />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
}










// Use a consistent endpoint for saving/fetching notes to avoid mismatches
const NOTES_ENDPOINT = "https://optical360v2-ffa9ewbfafdvfyd8.westeurope-01.azurewebsites.net/api/HttpTrigger1";
