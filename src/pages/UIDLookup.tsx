import React, { useState, useEffect } from "react";
import { saveToStorage } from "../api/saveToStorage";
import { getNotesForUid, getStatusForUid, getProjectsForUid, getAllProjects, deleteNote as deleteNoteApi, NoteEntity } from "../api/items";
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
  Spinner,
  SpinnerSize,
} from '@fluentui/react';
import ThemedProgressBar from "../components/ThemedProgressBar";
import UIDSummaryPanel from "../components/UIDSummaryPanel";
import UIDStatusPanel from "../components/UIDStatusPanel";
import CapacityCircle from "../components/CapacityCircle";

import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import deriveLineForC0 from "../data/mappedlines";

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
    // 1) If an AssociatedUID row matches uidKey and contains WorkflowStatus, prefer that
    const u = (uidKey || '').toString();
    if (u && Array.isArray(src?.AssociatedUIDs)) {
      try {
        const match = (src.AssociatedUIDs as any[]).find((r: any) => String(r?.UID ?? r?.Uid ?? r?.uid ?? '') === u);
        if (match && (match?.WorkflowStatus ?? match?.Workflow ?? match?.WorkflowState)) {
          return niceWorkflowStatus(match.WorkflowStatus ?? match.Workflow ?? match.WorkflowState);
        }
      } catch {}
    }
    const map: Record<string, string> | undefined = (src as any)?.__WFStatusByUid;
    if (u && map && map[u]) {
      return niceWorkflowStatus(map[u]);
    }
    return niceWorkflowStatus(src?.KQLData?.WorkflowStatus ?? src?.WorkflowStatus);
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
    LinkWFs?: any[];
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
  // Projects are sourced from server AND a local cache. Projects created via the
  // UI are stored locally so users can save projects without any server-side
  // persistence. Server-side Projects (when present) are still fetched and
  // merged on load, but creation is local-only per current requirements.
  const LOCAL_PROJECTS_KEY = 'uidLocalProjects';
  const [projects, setProjects] = useState<Project[]>(() => {
    try {
      const raw = localStorage.getItem(LOCAL_PROJECTS_KEY);
      const arr = raw ? JSON.parse(raw) : [];
      if (!Array.isArray(arr)) return [];
      // Normalize any legacy or double-serialized entries so p.data is an object
      const norm = arr.map((p: any) => {
        try {
          const copy = { ...(p || {}) } as any;
          if (typeof copy.data === 'string') {
            try { copy.data = JSON.parse(copy.data); } catch { copy.data = {}; }
          }
          if (!copy.data || typeof copy.data !== 'object') copy.data = {};
          // ensure common arrays exist to avoid empty UI when opening projects
          copy.data.OLSLinks = Array.isArray(copy.data.OLSLinks) ? copy.data.OLSLinks : (Array.isArray(copy.data.LinkSummary) ? copy.data.LinkSummary : []);
          copy.data.AssociatedUIDs = Array.isArray(copy.data.AssociatedUIDs) ? copy.data.AssociatedUIDs : [];
          copy.data.MGFXA = Array.isArray(copy.data.MGFXA) ? copy.data.MGFXA : [];
          copy.data.MGFXZ = Array.isArray(copy.data.MGFXZ) ? copy.data.MGFXZ : [];
          copy.data.GDCOTickets = Array.isArray(copy.data.GDCOTickets) ? copy.data.GDCOTickets : (Array.isArray(copy.data.ReleatedTickets) ? copy.data.ReleatedTickets : []);
          if (!Array.isArray(copy.data.sourceUids)) copy.data.sourceUids = (Array.isArray(copy.data.sourceUids) ? copy.data.sourceUids : ([]));
          return copy as Project;
        } catch {
          return { id: String(p?.id || `${Date.now()}`), name: p?.name || 'Project', createdAt: Date.now(), data: {} } as Project;
        }
      });
      return norm as Project[];
    } catch {
      return [];
    }
  });

  // Persist only local-created projects (those without a __serverEntity marker)
  // into localStorage so they survive reloads. Server-origin projects are not
  // overwritten here and will be merged on startup when the app fetches them.
  useEffect(() => {
    try {
      const localOnly = (projects || []).filter(p => !(p as any).__serverEntity);
      localStorage.setItem(LOCAL_PROJECTS_KEY, JSON.stringify(localOnly || []));
    } catch {}
  }, [projects]);
  const [activeProjectId, setActiveProjectId] = useState<string | null>(null);
  const [projectFilter, setProjectFilter] = useState<string>("");
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
  // Pre-load Projects from server so project details are available before any UID search
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        // Use the default endpoint (proxy/API_BASE) to avoid cross-origin issues.
        const items = await getAllProjects();
        if (cancelled) return;
        if (!items || !items.length) return;
        // Map server entities into Project shape and merge with local projects
        const mapped: Project[] = items.map((e: NoteEntity) => {
          const id = e.rowKey || e.RowKey || `${Date.now()}-${Math.random().toString(36).slice(2,8)}`;
          const name = e.title || e.Title || e.projectName || e.ProjectName || `Project ${id}`;
          const createdAt = e.savedAt ? Date.parse(String(e.savedAt)) : Date.now();
          const parsed = (e.projectJson || e.ProjectJson || e.description) ? (() => {
            try { const raw = e.projectJson || e.ProjectJson || e.description || ''; return typeof raw === 'string' ? JSON.parse(raw) : raw; } catch { return null; }
          })() : null;
          // support two shapes:
          // 1) stored snapshot directly (contains sourceUids, AssociatedUIDs, etc.)
          // 2) stored full Project object (id, name, createdAt, data: { ...snapshot }, notes, section)
          let dataSnapshot: any = null;
          let finalId = id;
          let finalName = String(name || id);
          let finalCreatedAt = Number.isFinite(createdAt) ? createdAt : Date.now();
          let owners: string[] | undefined = undefined;
          let sectionVal: string | undefined = undefined;
          let notesVal: Record<string, any> | undefined = undefined;
          if (parsed) {
            if (parsed.data && (parsed.data.sourceUids || parsed.data.AssociatedUIDs || parsed.data.OLSLinks)) {
              // parsed is a full Project object
              dataSnapshot = parsed.data;
              finalId = parsed.id || finalId;
              finalName = String(parsed.name || finalName);
              finalCreatedAt = parsed.createdAt ? Number(parsed.createdAt) : finalCreatedAt;
              owners = parsed.owners;
              sectionVal = parsed.section;
              notesVal = parsed.notes;
            } else {
              // parsed is the snapshot itself
              dataSnapshot = parsed;
            }
          }
          return {
            id: finalId,
            name: finalName,
            createdAt: Number.isFinite(finalCreatedAt) ? finalCreatedAt : Date.now(),
            data: dataSnapshot || {},
            owners,
            section: sectionVal,
            notes: notesVal,
            __serverEntity: e,
          } as any;
        });
        setProjects(prev => {
          const existingIds = new Set(prev.map(p => p.id));
          const toAdd = mapped.filter(m => !existingIds.has(m.id));
          return toAdd.length ? [...toAdd, ...prev] : prev;
        });
      } catch (e) {
        // ignore
      }
    })();
    return () => { cancelled = true; };
  }, []);
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

  // Projects are no longer persisted to localStorage per requirements.
  // (Server sync occurs when creating/updating via saveToStorage elsewhere.)

  // ---- Project multi-UID loader state and helpers ----
  const [projectLoadingCount, setProjectLoadingCount] = useState<number>(0);
  const [projectTotalCount, setProjectTotalCount] = useState<number>(0);
  const [isProjectLoading, setIsProjectLoading] = useState<boolean>(false);
  const [projectLoadError, setProjectLoadError] = useState<string | null>(null);
  const [combinedData, setCombinedData] = useState<any | null>(null);

  // Combine results from multiple UID responses into the four target collections.
  const dedupeArray = (arr: any[]) => {
    const seen = new Set<string>();
    const out: any[] = [];
    for (const it of arr || []) {
      try {
        const k = JSON.stringify(it);
        if (!seen.has(k)) { seen.add(k); out.push(it); }
      } catch {
        if (!out.includes(it)) out.push(it);
      }
    }
    return out;
  };

  const combineResults = (resultsArray: any[]) => {
    // Accept two shapes for resultsArray items:
    // 1) legacy: each item is the raw payload
    // 2) new: each item is { uid, data }
    const flatLinkSummary: any[] = [];
    const mgfxAList: any[] = [];
  const mgfxZList: any[] = [];
  const utilizationList: any[] = [];
    const ticketsList: any[] = [];
    const byUid: Array<{ uid: string; links: any[] }> = [];

    for (const item of resultsArray || []) {
      const uid = (item && item.uid) ? String(item.uid) : null;
      const src = (item && item.data) ? item.data : item;
      // collect Utilization rows for combined view
      const utilRows = Array.isArray(src?.Utilization) ? src.Utilization : (Array.isArray(src?.utilization) ? src.utilization : []);
      if (utilRows && utilRows.length) utilizationList.push(...utilRows);
      const links = Array.isArray(src?.LinkSummary) ? src.LinkSummary : (Array.isArray(src?.OLSLinks) ? src.OLSLinks : (Array.isArray(src?.LinkSummaryArray) ? src.LinkSummaryArray : []));
      const mgfxA = Array.isArray(src?.MGFXA) ? src.MGFXA : [];
      const mgfxZ = Array.isArray(src?.MGFXZ) ? src.MGFXZ : [];
      const tickets = Array.isArray(src?.GDCOTickets) ? src.GDCOTickets : (Array.isArray(src?.ReleatedTickets) ? src.ReleatedTickets : []);
      if (uid) {
        byUid.push({ uid, links: dedupeArray(links) });
      }
      flatLinkSummary.push(...links);
      mgfxAList.push(...mgfxA);
      mgfxZList.push(...mgfxZ);
      ticketsList.push(...tickets);
    }

    const ls = dedupeArray(flatLinkSummary);
    const mgfxa = dedupeArray(mgfxAList);
    const mgfxz = dedupeArray(mgfxZList);
    const tickets = dedupeArray(ticketsList);
    const utilization = dedupeArray(utilizationList || []);
    return {
      LinkSummary: ls,
      OLSLinks: ls,
      MGFXA: mgfxa,
      MGFXZ: mgfxz,
      GDCOTickets: tickets,
      Utilization: utilization,
      // Optional grouped view for UI: array of { uid, links }
      OLSLinksByUid: byUid,
    };
  };

  // Load project data for a list of UIDs in parallel, update progressive state as each finishes.
  const loadProjectData = async (uids: string[]) => {
    if (!Array.isArray(uids) || !uids.length) {
      setProjectLoadError('No UIDs to load for this project.');
      return;
    }
    setProjectLoadError(null);
    // Prepare loading state: keep previous combinedData visible until at least
    // one UID returns meaningful content. This avoids clearing the tables
    // immediately on Refresh and provides a smoother UX.
    setProjectTotalCount(uids.length);
    setProjectLoadingCount(0);
    setIsProjectLoading(true);
    const partialResults: any[] = [];
    let firstMeaningfulSeen = false;

    const looksMeaningful = (json: any) => {
      if (!json || typeof json !== 'object') return false;
      if (Array.isArray(json.OLSLinks) && json.OLSLinks.length) return true;
      if (Array.isArray(json.AssociatedUIDs) && json.AssociatedUIDs.length) return true;
      if (Array.isArray(json.MGFXA) && json.MGFXA.length) return true;
      if (Array.isArray(json.MGFXZ) && json.MGFXZ.length) return true;
      if (Array.isArray(json.GDCOTickets) && json.GDCOTickets.length) return true;
      // If the payload contains any string/number keys beyond empty object, treat as meaningful
      if (Object.keys(json).length > 0) return true;
      return false;
    };

    const tasks = uids.map((u) => {
      const url = `${NOTES_ENDPOINT}?uid=${encodeURIComponent(String(u))}`;
      return fetch(url)
        .then(async (res) => {
          if (!res.ok) throw new Error(`HTTP ${res.status}`);
          const json = await res.json().catch(() => null);
          return json;
        })
        .then((json) => {
          partialResults.push({ uid: u, data: json || {} });
          if (!firstMeaningfulSeen && looksMeaningful(json)) firstMeaningfulSeen = true;
          // Only update combinedData once we've observed meaningful content
          if (firstMeaningfulSeen) {
            try { setCombinedData((_prev: any) => combineResults(partialResults)); } catch { setCombinedData(combineResults(partialResults)); }
          }
        })
        .catch((err) => {
          // on failure for this UID, record an empty object and continue
          partialResults.push({ uid: u, data: {} });
          if (firstMeaningfulSeen) {
            try { setCombinedData((_prev: any) => combineResults(partialResults)); } catch { setCombinedData(combineResults(partialResults)); }
          }
        })
        .finally(() => {
          setProjectLoadingCount((c) => c + 1);
        });
    });

    // Wait for all to settle. If none returned meaningful content, keep the
    // previous combinedData and show a lightweight error message instead of
    // blanking the UI.
    await Promise.allSettled(tasks);
    if (!firstMeaningfulSeen) {
      // leave previous data visible to avoid a jarring empty state; surface a message
      setProjectLoadError('Refresh completed: no new data returned for these UIDs.');
      // keep existing combinedData (do not overwrite)
    } else {
      // ensure final combined data reflects all partial results
      try { setCombinedData((_prev: any) => combineResults(partialResults)); } catch { setCombinedData(combineResults(partialResults)); }
      setProjectLoadError(null);
    }
    setIsProjectLoading(false);
    setIsProjectLoading(false);
  };

  // When the user clicks a project, load all UIDs for that project and switch to project view
  const handleProjectClick = (projectId: string) => {
    const p = projects.find(x => x.id === projectId);
    if (!p) {
      setActiveProjectId(projectId);
      return;
    }
    // Try to find sourceUids first, else extract from AssociatedUIDs
    const uids: string[] = Array.from(new Set([...(p.data?.sourceUids || []), ...(Array.isArray(p.data?.AssociatedUIDs) ? p.data.AssociatedUIDs.map((r: any) => String(r?.UID || r?.Uid || r?.uid || '')).filter(Boolean) : [])]));
    // Clear any previous combinedData so the stored project snapshot is used
    // immediately while we refresh the project's UIDs.
    setCombinedData(null);
    setProjectLoadError(null);
    setActiveProjectId(projectId);
    // kick off loading; progressive results will render as they come in
    void loadProjectData(uids);
  };

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

  // (Removed) top-level troubleshootingItems fetch: per-row persistence now handles Troubleshooting

  // Load Projects entries for current UID and merge into local projects state
  const [remoteProjectsForUid, setRemoteProjectsForUid] = useState<NoteEntity[] | null>(null);
  // reference remoteProjectsForUid to satisfy lint rules (used to drive debug/side-effects)
  useEffect(() => {
    if (!remoteProjectsForUid) return;
    // lightweight debug: count of remote projects loaded for the current UID
    try { console.debug('[UIDLookup] remoteProjectsForUid count=', remoteProjectsForUid.length); } catch {}
  }, [remoteProjectsForUid]);
  useEffect(() => {
    const keyUid = lastSearched || '';
    if (!keyUid) { setRemoteProjectsForUid(null); return; }
    let cancelled = false;
    (async () => {
      try {
  // Use proxy/default endpoint to avoid cross-origin issues
  const items = await getProjectsForUid(keyUid);
        if (cancelled) return;
        setRemoteProjectsForUid(items || []);
        // Merge server projects into local projects state (lightweight):
        try {
          if (items && items.length) {
            // Map server entities into Project shape where possible. We store the raw entity under data.__serverEntity
            const mapped = items.map((e) => {
              const id = e.rowKey || e.RowKey || `${Date.now()}-${Math.random().toString(36).slice(2,8)}`;
              const name = e.title || e.Title || e.projectName || e.ProjectName || `Project ${id}`;
              const createdAt = e.savedAt ? Date.parse(String(e.savedAt)) : Date.now();
              const parsed = (e.projectJson || e.ProjectJson || e.description) ? (() => {
                try { const raw = e.projectJson || e.ProjectJson || e.description || ''; return typeof raw === 'string' ? JSON.parse(raw) : raw; } catch { return null; }
              })() : null;
              let dataSnapshot: any = null;
              let finalId = id;
              let finalName = String(name || id);
              let finalCreatedAt = Number.isFinite(createdAt) ? createdAt : Date.now();
              let owners: string[] | undefined = undefined;
              let sectionVal: string | undefined = undefined;
              let notesVal: Record<string, any> | undefined = undefined;
              if (parsed) {
                if (parsed.data && (parsed.data.sourceUids || parsed.data.AssociatedUIDs || parsed.data.OLSLinks)) {
                  dataSnapshot = parsed.data;
                  finalId = parsed.id || finalId;
                  finalName = String(parsed.name || finalName);
                  finalCreatedAt = parsed.createdAt ? Number(parsed.createdAt) : finalCreatedAt;
                  owners = parsed.owners;
                  sectionVal = parsed.section;
                  notesVal = parsed.notes;
                } else {
                  dataSnapshot = parsed;
                }
              }
              return {
                id: finalId,
                name: finalName,
                createdAt: Number.isFinite(finalCreatedAt) ? finalCreatedAt : Date.now(),
                data: dataSnapshot || {},
                owners,
                section: sectionVal,
                notes: notesVal,
                __serverEntity: e,
              } as any;
            });
            // Prepend any remote projects that don't already exist locally by id
            setProjects(prev => {
              const existingIds = new Set(prev.map(p => p.id));
              const toAdd = mapped.filter(m => !existingIds.has(m.id));
              return toAdd.length ? [...toAdd, ...prev] : prev;
            });
          }
        } catch {}
      } catch (e) {
        setRemoteProjectsForUid(null);
      }
    })();
    return () => { cancelled = true; };
  }, [lastSearched]);

  // Load UID status (expected delivery date, etc.) when the current UID changes
  useEffect(() => {
    const keyUid = lastSearched || '';
    if (!keyUid) { return; }
    let cancelled = false;
    (async () => {
      try {
        const items = await getStatusForUid(keyUid, NOTES_ENDPOINT);
        if (cancelled) return;
        if (!items || !items.length) return;
        // Choose the most-recent entity (prefer savedAt/timestamp)
        const parseTime = (e: any) => {
          const cand = e?.savedAt || e?.timestamp || e?.Timestamp || e?.SavedAt || '';
          const t = Date.parse(String(cand || ''));
          return Number.isFinite(t) ? t : 0;
        };
        const sorted = items.slice().sort((a, b) => parseTime(b) - parseTime(a));
        const entity = sorted[0] || items[0];

        const normalizeDate = (v: any): string | null => {
          if (!v && v !== 0) return null;
          const s = String(v);
          const m = s.match(/\d{4}-\d{2}-\d{2}/);
          if (m) return m[0];
          const d = Date.parse(s);
          if (!isNaN(d)) return new Date(d).toISOString().slice(0, 10);
          return null;
        };

        const candidateDate = entity?.expectedDeliveryDate ?? entity?.expecteddeliverydate ?? entity?.etaForDelivery?.date ?? entity?.ETA ?? entity?.eta ?? entity?.Description ?? null;
        const normalized = normalizeDate(candidateDate);
        if (normalized) {
          try {
            const key = `uidStatus:${keyUid}`;
            const raw = localStorage.getItem(key);
            const base = raw ? JSON.parse(raw) : { configPush: "Not Started", circuitsQc: "Not Started", expectedDeliveryDate: null };
            const merged = { ...base, expectedDeliveryDate: normalized };
            localStorage.setItem(key, JSON.stringify(merged));
            // Also trigger a state update indirectly by touching `data` if present so panels re-read
            // (UIDStatusPanel reads from localStorage on uid change, so no further action needed)
          } catch {}
        }
      } catch {
        // ignore
      }
    })();
    return () => { cancelled = true; };
  }, [lastSearched]);

  // When the main view data changes (search result), pull Status rows for ALL associated UIDs
  // so the UIDStatusPanel and project summaries have up-to-date expectedDeliveryDate values.
  useEffect(() => {
    const p = projects.find(p => p.id === activeProjectId) || null;
    const view = p ? p.data : data;
    if (!view) return;
    const rows: any[] = Array.isArray(view.AssociatedUIDs) ? view.AssociatedUIDs : [];
    const uids = Array.from(new Set(rows.map(r => String(r?.UID ?? r?.Uid ?? r?.uid ?? '').trim()).filter(Boolean)));
    if (!uids.length) return;
    let cancelled = false;
    (async () => {
      // Helper to normalise date strings to YYYY-MM-DD (same logic as single-uid fetch)
      const normalizeDate = (v: any): string | null => {
        if (!v && v !== 0) return null;
        const s = String(v);
        const m = s.match(/\d{4}-\d{2}-\d{2}/);
        if (m) return m[0];
        const d = Date.parse(s);
        if (!isNaN(d)) return new Date(d).toISOString().slice(0, 10);
        return null;
      };

      // Limit concurrent requests to avoid flooding the Function/storage
      for (const u of uids) {
        if (cancelled) return;
        try {
          const items = await getStatusForUid(u, NOTES_ENDPOINT);
          if (!items || !items.length) continue;
          // prefer most recent
          const parseTime = (e: any) => { const cand = e?.savedAt || e?.timestamp || e?.Timestamp || e?.SavedAt || ''; const t = Date.parse(String(cand || '')); return Number.isFinite(t) ? t : 0; };
          const sorted = items.slice().sort((a,b)=> parseTime(b) - parseTime(a));
          const entity = sorted[0] || items[0];
          const candidateDate = entity?.expectedDeliveryDate ?? entity?.expecteddeliverydate ?? entity?.etaForDelivery?.date ?? entity?.ETA ?? entity?.eta ?? entity?.Description ?? null;
          const normalized = normalizeDate(candidateDate);
          if (normalized) {
            try {
              const key = `uidStatus:${u}`;
              const raw = localStorage.getItem(key);
              const base = raw ? JSON.parse(raw) : { configPush: "Not Started", circuitsQc: "Not Started", expectedDeliveryDate: null };
              const merged = { ...base, expectedDeliveryDate: normalized };
              localStorage.setItem(key, JSON.stringify(merged));
            } catch {}
          }
        } catch (e) {
          // ignore per-uid errors
        }
        // small pause to avoid hammering the Function
        await new Promise(resolve => setTimeout(resolve, 180));
      }
    })();
    return () => { /* signal cancellation */ cancelled = true; };
  }, [projects, activeProjectId, data]);
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
      // Prefer JobId from the AssociatedUID matching the UID the user searched; fallback to KQLData.JobId
      const assocRows: any[] = Array.isArray(data?.AssociatedUIDs) ? data.AssociatedUIDs : [];
      const assoc = lastSearched ? assocRows.find((r: any) => String(r?.UID ?? r?.Uid ?? r?.uid ?? '') === String(lastSearched)) : null;
      const jobId = assoc?.JobId ?? assoc?.JobID ?? data?.KQLData?.JobId;
      const link = jobId ? `https://azcis.trafficmanager.net/Public/NetworkingOptical/JobDetails/${jobId}` : null;
      setCancelDialogTitle(isCancelled ? 'WF Cancelled' : 'DECOM');
      setCancelDialogMsg(isCancelled ? 'This workflow has been cancelled. Please check the job in CIS below to confirm.' : 'This workflow appears to be decommissioned.');
      setCancelDialogLink(link);
      setShowCancelDialog(true);
      setLastPromptUid(lastSearched);
    }
  }, [data, lastSearched, lastPromptUid]);

  // Helper to get the dataset currently being viewed (live or project snapshot)
  const viewHasMeaningfulContent = (v: any) => {
    if (!v || typeof v !== 'object') return false;
    if (Array.isArray(v.OLSLinks) && v.OLSLinks.length) return true;
    if (Array.isArray(v.AssociatedUIDs) && v.AssociatedUIDs.length) return true;
    if (Array.isArray(v.MGFXA) && v.MGFXA.length) return true;
    if (Array.isArray(v.MGFXZ) && v.MGFXZ.length) return true;
    if (Array.isArray(v.GDCOTickets) && v.GDCOTickets.length) return true;
    if (Array.isArray(v.ReleatedTickets) && v.ReleatedTickets.length) return true;
    if (v.KQLData && Object.keys(v.KQLData).length) return true;
    return false;
  };

  const getViewData = React.useCallback(() => {
  // Prefer combinedData only when it contains meaningful rows; otherwise fall back
  // to the stored project snapshot so we don't blank the UI when a refresh
  // returns empty/partial results. When combinedData contains per-UID groups
  // prefer that grouped view so the UI can render separate Link Summary tables.
  if (activeProjectId && combinedData && viewHasMeaningfulContent(combinedData)) return combinedData;
    const p = projects.find(p => p.id === activeProjectId) || null;
    return p ? p.data : data;
  }, [projects, activeProjectId, data, combinedData]);
  // Helper to choose a primary UID for a given snapshot/view. When viewing a
  // saved project there may be no `lastSearched`, so prefer the first
  // `sourceUids` entry or the first AssociatedUID row's UID.
  const primaryUidFor = (src: any): string | null => {
    try {
      if (lastSearched) return String(lastSearched);
      if (!src) return null;
      if (Array.isArray(src.sourceUids) && src.sourceUids.length) return String(src.sourceUids[0]);
      if (Array.isArray(src.AssociatedUIDs) && src.AssociatedUIDs.length) {
        const a = src.AssociatedUIDs[0];
        return String(a?.UID ?? a?.Uid ?? a?.uid ?? '') || null;
      }
      return null;
    } catch { return null; }
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
  // removeProjectNote function removed (unused)

  const naturalSort = (a: string, b: string) =>
    a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });

  // Normalize WorkflowStatus strings: now using module-scope helpers above

  // Associated UIDs view filter: show only In Progress by default
  const [showAllAssociatedWF, setShowAllAssociatedWF] = useState<boolean>(false);
  // When creating a project from Associated UIDs we hold the UID list here
  const [createFromAssocPendingUids, setCreateFromAssocPendingUids] = useState<string[] | null>(null);
  const [createFromAssocMode, setCreateFromAssocMode] = useState<boolean>(false);
  // UI progress state for the associated-UIDs project creation
  const [createFromAssocRunning, setCreateFromAssocRunning] = useState<boolean>(false);
  const [createFromAssocCurrent, setCreateFromAssocCurrent] = useState<number>(0);
  const [createFromAssocTotal, setCreateFromAssocTotal] = useState<number>(0);
  const [createFromAssocMessage, setCreateFromAssocMessage] = useState<string | null>(null);
  const [createFromAssocFailedUids, setCreateFromAssocFailedUids] = useState<string[]>([]);
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
        const wfRaw = r?.WorkflowStatus ?? r?.Workflow ?? wfMap?.[String(uid)];
        const wf = niceWorkflowStatus(wfRaw) || '';
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
      // Prefer an explicit numeric OpticalSpeed on the row (may be provided from Utilization)
      const explicit = row?.OpticalSpeedGb ?? row?.OpticalSpeedGb ?? row?.OpticalSpeed;
      if (explicit != null && explicit !== '' && !isNaN(Number(explicit))) {
        // If OpticalSpeed is provided in Mbps (large numbers like 100000), convert to G
        const n = Number(explicit);
        if (n > 1000) return Math.round(n / 1000);
        return Math.round(n);
      }
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
  // Deep-clone the full normalized view object for project snapshots.
  // This preserves all top-level keys (including arrays like OLSLinks, AssociatedUIDs)
  // and also copies a few non-enumerable/internal properties that JSON.stringify
  // would otherwise drop (e.g., __AllWorkflowStatus, __WFStatusByUid).
  const deepCloneView = (src: any, srcUid: string) => {
    const copy: any = JSON.parse(JSON.stringify(src || {}));
    try {
      if ((src as any)?.__AllWorkflowStatus) copy.__AllWorkflowStatus = (src as any).__AllWorkflowStatus;
      if ((src as any)?.__WFStatusByUid) copy.__WFStatusByUid = (src as any).__WFStatusByUid;
      if ((src as any)?.ReleatedTickets && !copy.ReleatedTickets) copy.ReleatedTickets = (src as any).ReleatedTickets;
      // Ensure we have a sourceUids array so project listing can find UIDs quickly
      if (!Array.isArray(copy.sourceUids)) copy.sourceUids = [srcUid].filter(Boolean);
      // Ensure per-UID grouped links are preserved so projects show each UID's
      // Link Summary as a separate table. Prefer existing grouped data when
      // present; otherwise capture current OLSLinks/LinkSummary under the srcUid.
      try {
        const group = { uid: srcUid, links: Array.isArray(copy.OLSLinks) ? copy.OLSLinks : (Array.isArray(copy.LinkSummary) ? copy.LinkSummary : []) };
        if (!Array.isArray(copy.OLSLinksByUid)) copy.OLSLinksByUid = [];
        // avoid duplicating an existing uid entry
        const exists = (copy.OLSLinksByUid || []).some((g: any) => String(g?.uid) === String(srcUid));
        if (!exists && (group.links || []).length) copy.OLSLinksByUid.push(group);
      } catch {}
      // Ensure top-level OLSLinks is present (flatten per-UID groups when needed)
      try {
        if ((!Array.isArray(copy.OLSLinks) || !copy.OLSLinks.length) && Array.isArray(copy.OLSLinksByUid) && copy.OLSLinksByUid.length) {
          const flat: any[] = [];
          for (const g of copy.OLSLinksByUid) {
            if (Array.isArray(g?.links)) flat.push(...g.links);
          }
          try {
            const uniq = Array.from(new Set(flat.map((r: any) => JSON.stringify(r))));
            copy.OLSLinks = uniq.map((s: string) => JSON.parse(s));
          } catch { copy.OLSLinks = flat; }
        }
      } catch {}
    } catch {
      // best-effort copy; ignore errors
    }
    return copy;
  };

  // Keep refs to the latest getViewData and lastSearched so we can expose
  // a mount-only global helper without creating a hook dependency cycle.
  const getViewDataRef = React.useRef(getViewData);
  useEffect(() => { getViewDataRef.current = getViewData; }, [getViewData]);
  const lastSearchedRef = React.useRef(lastSearched);
  useEffect(() => { lastSearchedRef.current = lastSearched; }, [lastSearched]);

  // Expose a safe global helper so small helper components (outside this module)
  // can create a full snapshot of the current view for local-only saves.
  // This effect intentionally runs only on mount/unmount and reads the
  // up-to-date values via refs to avoid lint dependency warnings.
  useEffect(() => {
    try {
      (window as any).getCurrentViewSnapshot = () => {
        try {
          const cur = getViewDataRef.current ? getViewDataRef.current() : null;
          return deepCloneView(cur || {}, lastSearchedRef.current || '');
        } catch { return null; }
      };
    } catch {}
    return () => { try { delete (window as any).getCurrentViewSnapshot; } catch {} };
  }, []);
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
  // Merge two snapshots (or full view objects) without duplicating array entries.
  // Keeps base scalar/structured fields when present and appends missing array
  // entries from `add`. This is intentionally permissive to support both the
  // Snapshot shape and full normalized viewData objects returned by the UI.
  const mergeSnapshots = (base: any, add: any): any => {
    if (!base) return add ? JSON.parse(JSON.stringify(add)) : {};
    if (!add) return base ? JSON.parse(JSON.stringify(base)) : {};
    const out: any = { ...(base || {}) };
    // union sourceUids
    out.sourceUids = Array.from(new Set([...(base.sourceUids || []), ...(add.sourceUids || [])]));

    // Prefer base's detailed objects when they have meaningful content;
    // otherwise fall back to add's values.
    out.AExpansions = (base.AExpansions && Object.keys(base.AExpansions).length) ? base.AExpansions : (add.AExpansions || {});
    out.ZExpansions = (base.ZExpansions && Object.keys(base.ZExpansions).length) ? base.ZExpansions : (add.ZExpansions || {});
    out.KQLData = (base.KQLData && Object.keys(base.KQLData).length) ? base.KQLData : (add.KQLData || {});

  // Merge arrays with deduplication by JSON identity. Known arrays we merge:
  out.OLSLinks = dedupMerge(base.OLSLinks || [], add.OLSLinks || []);
  out.AssociatedUIDs = dedupMerge(base.AssociatedUIDs || [], add.AssociatedUIDs || []);
  out.GDCOTickets = dedupMerge(base.GDCOTickets || [], add.GDCOTickets || []);
  out.MGFXA = dedupMerge(base.MGFXA || [], add.MGFXA || []);
  out.MGFXZ = dedupMerge(base.MGFXZ || [], add.MGFXZ || []);
  out.LinkWFs = dedupMerge(base.LinkWFs || [], add.LinkWFs || []);
  out.ReleatedTickets = dedupMerge(base.ReleatedTickets || [], add.ReleatedTickets || []);
  // Merge Utilization (capitalized or lowercase) so per-link admin/oper/speed
  // values are preserved when merging snapshots from another UID.
  const baseUtil = Array.isArray(base.Utilization) ? base.Utilization : (Array.isArray(base.utilization) ? base.utilization : []);
  const addUtil = Array.isArray(add.Utilization) ? add.Utilization : (Array.isArray(add.utilization) ? add.utilization : []);
  out.Utilization = dedupMerge(baseUtil || [], addUtil || []);

  // Merge OLSLinksByUid: preserve per-UID groups and dedupe by uid; when same uid
  // exists in both, merge their links (deduped).
  try {
    const baseBy = Array.isArray(base.OLSLinksByUid) ? base.OLSLinksByUid : [];
    const addBy = Array.isArray(add.OLSLinksByUid) ? add.OLSLinksByUid : [];
    const byMap = new Map<string, any[]>();
    for (const g of baseBy) {
      try { byMap.set(String(g.uid), Array.isArray(g.links) ? g.links.slice() : []); } catch { }
    }
    for (const g of addBy) {
      try {
        const k = String(g.uid);
        const existing = byMap.get(k) || [];
        const mergedLinks = dedupMerge(existing, Array.isArray(g.links) ? g.links : []);
        byMap.set(k, mergedLinks);
      } catch { }
    }
    // also ensure any single srcUid present in `add` as top-level sourceUids are represented
    try {
      const allSrc = Array.isArray(out.sourceUids) ? out.sourceUids : [];
      for (const u of allSrc) {
        const k = String(u);
        if (!byMap.has(k)) {
          // attempt to find a matching flat OLSLinks subset for this uid in add (best-effort)
          const candidate = (Array.isArray(add.OLSLinks) && add.OLSLinks.length) ? add.OLSLinks : (Array.isArray(add.LinkSummary) ? add.LinkSummary : []);
          if (candidate && candidate.length) byMap.set(k, candidate.slice());
        }
      }
    } catch {}
    out.OLSLinksByUid = Array.from(byMap.entries()).map(([uid, links]) => ({ uid, links: dedupeArray(links || []) }));
  } catch {}

  // Merge top-level structured fields shallowly so missing keys from base
  // are filled in from the added snapshot (preserve base values when present).
  out.AExpansions = { ...(add.AExpansions || {}), ...(base.AExpansions || {}) };
  out.ZExpansions = { ...(add.ZExpansions || {}), ...(base.ZExpansions || {}) };
  out.KQLData = { ...(add.KQLData || {}), ...(base.KQLData || {}) };

    // For any other keys present in `add` that aren't in out, copy them over.
    for (const k of Object.keys(add || {})) {
      if (out[k] === undefined) out[k] = add[k];
    }

    return out;
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

  // Get SRLG with fallback: prefer the AssociatedUID matching uidKey (if present),
  // then AExpansions.SRLG, then KQLData.SRLG.
  const getSrlgFrom = (src: any, uidKey?: string): string | null => {
    try {
      // 1) Prefer AssociatedUIDs row matching the entered UID
      const rows: any[] = Array.isArray(src?.AssociatedUIDs) ? src.AssociatedUIDs : [];
      if (rows.length) {
        const match = uidKey ? rows.find(r => String(r?.UID || r?.Uid || r?.uid || '') === String(uidKey)) : null;
        const r = match || rows[0] || {};
        const val = r['Srlg'] ?? r['SRLG'] ?? r['SRLGName'] ?? r['SrlgName'] ?? r['SRLG'] ?? '';
        if (val != null && String(val).trim()) return String(val).trim();
      }

      // 2) Fallback to AExpansions then KQLData
      const a = src?.AExpansions?.SRLG ?? src?.AExpansions?.Srlg;
      if (a != null && String(a).trim()) return String(a).trim();
      const k = src?.KQLData?.SRLG ?? src?.KQLData?.Srlg;
      if (k != null && String(k).trim()) return String(k).trim();

      return null;
    } catch { return null; }
  };
  const computeProjectTitle = (src: any, uidKey: string): string => {
    try {
      // Prefer SolutionId from the AssociatedUID matching uidKey when available
      let sol = '';
      try {
        const assocRows: any[] = Array.isArray(src?.AssociatedUIDs) ? src.AssociatedUIDs : [];
        const assoc = uidKey ? assocRows.find((r: any) => String(r?.UID ?? r?.Uid ?? r?.uid ?? '') === String(uidKey)) : (assocRows[0] || null);
        const assocSol = assoc?.SolutionId ?? assoc?.SolutionID ?? assoc?.Solution ?? null;
        if (assocSol) {
          if (Array.isArray(assocSol)) sol = (assocSol || []).map((v: any) => formatSolutionId(String(v))).filter(Boolean)[0] || '';
          else sol = formatSolutionId(String(assocSol));
        }
      } catch {}
      if (!sol) {
        const sols = getSolutionIds(src).map(formatSolutionId).filter(Boolean);
        sol = sols[0] || '';
      }
      const sites = getFirstSites(src, uidKey);
      const a = (sites.a || '').toString().trim();
      const z = (sites.z || '').toString().trim();
  if (sol && a && z) return `${sol} - ${a} ↔ ${z}`;
      if (sol && (a || z)) return `${sol} - ${a || z}`;
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

  // Helper: wait until a handleSearch-driven load for `targetUid` completes
  // (Removed unused wait helper to satisfy lint rules)
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
    // No warning needed, merge immediately. Mark the merged project as locally
    // modified (clear any __serverEntity marker) so it will be persisted to
    // localStorage and survive reloads.
    setProjects(prev => prev.map(pp => {
      if (pp.id !== targetId) return pp;
      const merged = mergeSnapshots(pp.data, deepCloneView(data, lastSearched));
      return { ...pp, data: merged, notes: { ...(pp.notes || {}) }, __serverEntity: undefined, __localModified: true } as any;
    }));
    setActiveProjectId(targetId);
  };

  // Initiate a Create Project flow which will sequentially open each Associated UID
  // in the same tab (via handleSearch) and merge their snapshots into a new project.
  // We trigger the modal first for section selection, then perform the sequence
  // when the user confirms the modal.
  const createProjectFromAssociatedUIDs = (uids: string[]) => {
    if (!Array.isArray(uids) || !uids.length) return;
    setCreateFromAssocPendingUids(uids.slice());
    setCreateFromAssocMode(true);
    setCreateSectionChoice("");
    setCreateNewSection("");
    setCreateError(null);
    setModalType('create-project');
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
  // addSection removed (unused) to satisfy lint rules
  const closeModal = () => { setModalType(null); setModalProjectId(null); setModalValue(''); };
  const saveModal = async () => {
    const value = (modalValue || '').trim();
    if (!modalType) return;
    if (modalType === 'confirm-merge' && modalProjectId) {
      // Proceed with merge now that user confirmed
      const targetId = modalProjectId;
      if (!data || !lastSearched) { closeModal(); return; }
      setProjects(prev => prev.map(pp => {
        if (pp.id !== targetId) return pp;
        const merged = mergeSnapshots(pp.data, deepCloneView(data, lastSearched));
        return { ...pp, data: merged, notes: { ...(pp.notes || {}) }, __serverEntity: undefined, __localModified: true } as any;
      }));
      setActiveProjectId(targetId);
      closeModal();
      return;
    }
    if (modalType === 'create-project') {
      // If we're in create-from-associated mode, sequentially load each UID in
      // the same tab (via handleSearch) and merge their snapshots into one project.
      if (createFromAssocMode && Array.isArray(createFromAssocPendingUids) && createFromAssocPendingUids.length) {
        const uids = createFromAssocPendingUids.slice();
        const chosen = (createSectionChoice || '').trim();
        const newName = (createNewSection || '').trim();
        let finalSection: string | undefined = undefined;
        if (newName) {
          if (!sections.includes(newName)) setSections([...sections, newName]);
          finalSection = newName;
        } else if (chosen) {
          finalSection = (chosen === 'Archives') ? undefined : chosen;
        }
        if (!newName && !chosen) {
          setCreateError('Please choose a section or enter a new section name.');
          return;
        }

        // Sequentially open each UID and merge
        let mergedSnapshot: any = null;
        // Helper to ensure top-level OLSLinks exist (flatten OLSLinksByUid if needed)
        const ensureOlsLinks = (snap: any) => {
          try {
            if (!snap) return snap;
            if (!Array.isArray(snap.OLSLinks) || !snap.OLSLinks.length) {
              if (Array.isArray(snap.LinkSummary) && snap.LinkSummary.length) snap.OLSLinks = snap.LinkSummary.slice();
              else if (Array.isArray(snap.OLSLinksByUid) && snap.OLSLinksByUid.length) {
                const flat: any[] = [];
                for (const g of snap.OLSLinksByUid) {
                  if (Array.isArray(g?.links)) flat.push(...g.links);
                }
                try {
                  const uniq = Array.from(new Set(flat.map((r: any) => JSON.stringify(r))));
                  snap.OLSLinks = uniq.map((s: string) => JSON.parse(s));
                } catch { snap.OLSLinks = flat; }
              }
            }
          } catch {}
          return snap;
        };

        // Close the modal immediately so the overlay can show and block input
        closeModal();
        setCreateFromAssocRunning(true);
        setCreateFromAssocTotal(uids.length);
        setCreateFromAssocCurrent(0);
        setCreateFromAssocMessage(null);
        setCreateFromAssocFailedUids([]);
        // Create a provisional project that will be updated as each UID is merged
        const provisionalId = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
        const provisionalProject: Project = {
          id: provisionalId,
          name: `Building project...`,
          createdAt: Date.now(),
          data: { sourceUids: [], OLSLinks: [], AssociatedUIDs: [], GDCOTickets: [], MGFXA: [], MGFXZ: [] },
          section: finalSection,
        };
        setProjects(prev => [provisionalProject, ...prev]);
        setActiveProjectId(provisionalId);

        const failedList: string[] = [];
        for (let i = 0; i < uids.length; i++) {
          const u = uids[i];
          try {
            // update progress index (1-based)
            setCreateFromAssocCurrent(i + 1);
            // trigger a UI search in the same tab
            // eslint-disable-next-line no-await-in-loop
            await handleSearch(u);
            // small pause to allow React state to settle (faster than full wait helper)
            // eslint-disable-next-line no-await-in-loop
            await new Promise(resolve => setTimeout(resolve, 100));
            try {
              const snapRaw = deepCloneView(getViewData() || data || {}, String(u));
              const snap = ensureOlsLinks(snapRaw);
              mergedSnapshot = mergedSnapshot ? mergeSnapshots(mergedSnapshot, snap) : snap;
              // update provisional project incrementally so UI reflects additions immediately
              setProjects(prev => prev.map(p => {
                if (p.id !== provisionalId) return p;
                try {
                  const mergedNow = mergeSnapshots(p.data || {}, snap);
                  const normalizedNow = ensureOlsLinks(mergedNow);
                  return { ...p, data: normalizedNow } as Project;
                } catch { return p; }
              }));
            } catch (err) {
              failedList.push(String(u));
            }
          } catch (err) {
            failedList.push(String(u));
          }
        }

        // Update the provisional project with final merged snapshot and name
        const id = provisionalId;
        const primaryForTitle = (uids && uids.length) ? String(uids[0]) : (lastSearched || '');
        const finalData = ensureOlsLinks(mergedSnapshot || {});
        const updatedProj: Project = {
          id,
          name: computeProjectTitle(finalData, primaryForTitle),
          createdAt: Date.now(),
          data: finalData,
          section: finalSection,
        };
        setProjects(prev => prev.map(p => p.id === provisionalId ? updatedProj : p));
        setActiveProjectId(id);
        // set failed uids state and show final success message then clear overlay after a short delay
        setCreateFromAssocFailedUids(failedList);
  setCreateFromAssocMessage(`Project created (${uids.length - failedList.length} succeeded${failedList.length ? `, ${failedList.length} failed` : ''})`);
  setCreateFromAssocRunning(false);
  // Ensure any running top-level progress bar is signalled complete and hidden
  try { setProgressComplete(true); setProgressVisible(false); } catch (e) { /* ignore if states unavailable */ }
        setCreateSectionChoice('');
        setCreateNewSection('');
        setCreateError(null);
        setCreateFromAssocPendingUids(null);
        setCreateFromAssocMode(false);
        // hide the overlay message after a few seconds
        setTimeout(() => {
          setCreateFromAssocMessage(null);
          setCreateFromAssocFailedUids([]);
        }, 3500);
        return;
      }

      // default single-UID create flow (existing behaviour)
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
        // Save a full deep clone of the current view so the project contains the
        // whole page state (not just a reduced snapshot). This prevents fields
        // from disappearing when viewing a saved project.
        data: deepCloneView(data, lastSearched),
        section: finalSection,
        notes: Object.keys(notesMap).length ? notesMap : undefined,
      };
      setProjects((prev) => [proj, ...prev]);
      // Persist locally only. We intentionally do NOT save created projects to the
      // server anymore; local projects are kept in localStorage by the effect
      // that watches `projects` (LOCAL_PROJECTS_KEY). This keeps the Projects
      // side menu fully functional while avoiding server writes.
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
    // Call backend proxy (LogicAppProxy) instead of direct Logic App
  try {
    const isJson = (r: Response) =>
      /application\/json/i.test(r.headers.get("content-type") || "");

    const res = await fetch("/api/LogicAppProxy", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ type: "UID", uid: query }),
    });

    if (!(res.ok && isJson(res))) {
      const text = await res.text().catch(() => "");
      const statusPart = `HTTP ${res.status}`;
      const bodyPart = text ? `: ${text.slice(0, 220)}` : "";
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

  // The Logic App now returns the UID payload under `OtherData`. Prefer that as the
  // view model so the UI remains unchanged. Also provide a fallback for MGFX
  // information: if the A/Z arrays are missing or empty in OtherData, try to
  // use `MGFXbySLS` (several naming variants supported) and normalise its rows
  // to the same shape the UI expects.
  const normalizeMgfxRows = (rows: any[] | undefined) => {
    if (!Array.isArray(rows)) return [];
    return rows.map((r: any) => ({
      XOMT: r.XOMT ?? r.xomt ?? r.XomT ?? r['XOMT'] ?? '',
      // preserve original keys where present; downstream mapping functions
      // also accept many variants so we keep a permissive mapping here
      'C0 Device': r['C0 Device'] ?? r.C0Device ?? r['C0_Device'] ?? r.c0Device ?? '',
      'C0 Port': r['C0 Port'] ?? r.C0Port ?? r['C0_Port'] ?? r.c0Port ?? '',
      StartHardwareSku: r.StartHardwareSku ?? r.HardwareSku ?? r.SKU ?? r.sku ?? '',
      'M0 Device': r['M0 Device'] ?? r.M0Device ?? r['M0_Device'] ?? r.m0Device ?? '',
      'M0 Port': r['M0 Port'] ?? r.M0Port ?? r['M0_Port'] ?? r.m0Port ?? '',
      'C0 DIFF': r['C0 DIFF'] ?? r.C0_DIFF ?? r.C0Diff ?? '',
      'M0 DIFF': r['M0 DIFF'] ?? r.M0_DIFF ?? r.M0Diff ?? '',
    }));
  };

  // Build the primary view object: prefer OtherData when present.
  // Some Logic App responses wrap the payload in a `body` property, so
  // prefer that when present.
  const topPayload = (result && typeof result === 'object' && result.body && typeof result.body === 'object') ? result.body : result;
  let normalized: any = topPayload;
  if (topPayload?.OtherData && typeof topPayload.OtherData === 'object') {
    normalized = { ...topPayload.OtherData };
    // carry useful top-level pieces across if they aren't present inside OtherData
    if (!normalized.KQLData && topPayload.KQLData) normalized.KQLData = topPayload.KQLData;
    if (!normalized.OLSLinks && topPayload.OLSLinks) normalized.OLSLinks = topPayload.OLSLinks;
    if (!normalized.AssociatedUIDs && topPayload.AssociatedUIDs) normalized.AssociatedUIDs = topPayload.AssociatedUIDs;
  if (!normalized.GDCOTickets && topPayload.GDCOTickets) normalized.GDCOTickets = topPayload.GDCOTickets;
  // Logic App may return ReleatedTickets (note spelling). Ensure we copy it into normalized so downstream
  // consumers (getGdcoRows, status panel) can find tickets regardless of wrapper placement.
  if (!normalized.ReleatedTickets && topPayload.ReleatedTickets) normalized.ReleatedTickets = topPayload.ReleatedTickets;
    if (!normalized.MGFXA && topPayload.MGFXA) normalized.MGFXA = topPayload.MGFXA;
    if (!normalized.MGFXZ && topPayload.MGFXZ) normalized.MGFXZ = topPayload.MGFXZ;
    if (!normalized.LinkWFs && topPayload.LinkWFs) normalized.LinkWFs = topPayload.LinkWFs;
    // Carry across WorkflowsString (and common name variants) so the UI can find ordered workflow URLs
    if (!normalized.WorkflowsString) {
      normalized.WorkflowsString = topPayload.WorkflowsString ?? topPayload.WorkflowsStringRaw ?? topPayload.Workflows ?? topPayload.WorkflowUrls ?? null;
    }
    // copy the attached workflow maps if present
    if ((topPayload as any).__AllWorkflowStatus) {
      try { Object.defineProperty(normalized, '__AllWorkflowStatus', { value: (topPayload as any).__AllWorkflowStatus, enumerable: false }); } catch { (normalized as any).__AllWorkflowStatus = (topPayload as any).__AllWorkflowStatus; }
    }
    if ((topPayload as any).__WFStatusByUid) {
      try { Object.defineProperty(normalized, '__WFStatusByUid', { value: (topPayload as any).__WFStatusByUid, enumerable: false }); } catch { (normalized as any).__WFStatusByUid = (topPayload as any).__WFStatusByUid; }
    }
  }

  // Some Logic App responses return collections wrapped in an object with a
  // `value` property (e.g. { value: [...] }). Unwrap these into plain arrays so
  // downstream code can treat them uniformly.
  const unwrapValueArray = (obj: any, key: string) => {
    try {
      const v = obj?.[key];
      if (!v) return;
      if (Array.isArray(v)) return;
      if (v && typeof v === 'object' && Array.isArray(v.value)) obj[key] = v.value;
    } catch {}
  };
  ['AssociatedUIDs', 'OLSLinks', 'MGFXA', 'MGFXZ', 'GDCOTickets', 'ReleatedTickets', 'AssociatedTickets'].forEach(k => unwrapValueArray(normalized, k));
  // Also unwrap LinkWFs if wrapped under { value: [...] }
  unwrapValueArray(normalized, 'LinkWFs');

  // Some Logic App payloads use alternative names for the link-summary array.
  // Map common alternatives into `OLSLinks` so the existing UI logic can render them.
  try {
    if ((!normalized.OLSLinks || !Array.isArray(normalized.OLSLinks) || !normalized.OLSLinks.length) && Array.isArray(normalized.Base) && normalized.Base.length) {
      normalized.OLSLinks = normalized.Base;
    }
    // lowercase variants
    if ((!normalized.OLSLinks || !Array.isArray(normalized.OLSLinks) || !normalized.OLSLinks.length) && Array.isArray(normalized.base) && normalized.base.length) {
      normalized.OLSLinks = normalized.base;
    }
    // other possible names
    if ((!normalized.OLSLinks || !Array.isArray(normalized.OLSLinks) || !normalized.OLSLinks.length) && Array.isArray(normalized.LinkSummary) && normalized.LinkSummary.length) {
      normalized.OLSLinks = normalized.LinkSummary;
    }
  } catch {}

  // Prefer MGFXbySLS as the primary MGFX source when present (topPayload variants supported).
  // If MGFXbySLS is present, it will populate/overwrite normalized.MGFXA and normalized.MGFXZ.
  const candidate = (topPayload?.MGFXbySLS || topPayload?.MGFXBySLS || topPayload?.MGFX_by_SLS || topPayload?.MgfxBySLS || topPayload?.MGFXBY_SLS) || null;
  if (candidate) {
    // If candidate is an object with explicit A/Z arrays, use them directly
    if (candidate && typeof candidate === 'object' && !Array.isArray(candidate)) {
      const aSrc = candidate.A || candidate.Aside || candidate.MGFXA || candidate.mgfxA || candidate['MGFXA'];
      const zSrc = candidate.Z || candidate.Zside || candidate.MGFXZ || candidate.mgfxZ || candidate['MGFXZ'];
      if (Array.isArray(aSrc)) normalized.MGFXA = normalizeMgfxRows(aSrc);
      if (Array.isArray(zSrc)) normalized.MGFXZ = normalizeMgfxRows(zSrc);
    }

    // If candidate is an array, attempt to parse common MGFX-by-SLS shapes.
    if (Array.isArray(candidate)) {
      const arr = candidate as any[];
      // If rows look like StartDevice/EndDevice pairs, group by StartDevice (XOMT)
      const looksLikePairs = arr.some(it => it && (it.StartDevice || it.StartPort || it.EndDevice || it.EndPort));
      if (looksLikePairs) {
        const groups = new Map<string, any[]>();
        for (const it of arr) {
          if (!it) continue;
          const start = String(it.StartDevice ?? it.StartDeviceName ?? it['StartDevice'] ?? '').trim();
          const end = String(it.EndDevice ?? it.EndDeviceName ?? it['EndDevice'] ?? '').trim();
          const startPort = String(it.StartPort ?? it.StartPortName ?? it['StartPort'] ?? it.StartPort ?? '').trim();
          const endPort = String(it.EndPort ?? it.EndPortName ?? it['EndPort'] ?? it.EndPort ?? '').trim();
          if (!start || !end) continue;
          // ignore entries where either device is an OLT (per instruction)
          const endsWithOlt = (s: string) => /olt$/i.test(s);
          if (endsWithOlt(start) || endsWithOlt(end)) continue;
          const endSkuVal = String(it.EndSku ?? it.endSku ?? it.EndHardwareSku ?? it.EndHardware ?? it.endHardware ?? '').trim();
          const list = groups.get(start) || [];
          list.push({ start, startPort, end, endPort, endSku: endSkuVal });
          groups.set(start, list);
        }

        const aRows: any[] = [];
        const zRows: any[] = [];
  

        const makeDiffLink = (hostname: string) => `https://phynet.trafficmanager.net/ConfigMon/ConfigDiff?Hostname=${encodeURIComponent(hostname)}&DiffGroups=&Timestamp`;

        // Use ordering from Logic Apps: rows are sent in A→Z order.
        // Determine a simple group prefix (first two dash-separated segments)
        // and treat the first encountered prefix as the A-side; any subsequent
        // different prefix is treated as Z-side. This avoids relying on Site A/Z
        // fields which are not always accurate.
        const xomtPrefixKey = (s: string) => {
          const parts = String(s || '').split('-').filter(Boolean);
          if (parts.length >= 2) return (parts[0] + '-' + parts[1]).toLowerCase();
          return parts[0]?.toLowerCase() || String(s || '').toLowerCase();
        };
        const xomtBaseKey = (s: string) => {
          const parts = String(s || '').split('-').filter(Boolean);
          return parts[0]?.toLowerCase() || String(s || '').toLowerCase();
        };
        let firstPrefix: string | null = null;

        // If there are many different prefixes present, prefer using the
        // optical devices from the Link Summary to identify A vs Z prefixes
        // (use only the base prefix like 'gvx11' or 'osl20'). Fall back to
        // order-based assignment when optical hints aren't available.
        let optAPrefix: string | null = null;
        let optZPrefix: string | null = null;
        try {
          const links = Array.isArray(normalized?.OLSLinks) ? normalized.OLSLinks : [];
          // Helper to safely read a device name from a link row
          const pickDevice = (r: any, keys: string[]) => {
            for (const k of keys) {
              const v = r?.[k];
              if (v) return String(v).trim();
            }
            return null;
          };
          if (links && links.length) {
            const aDev = pickDevice(links.find(Boolean), ['A Optical Device', 'AOpticalDevice', 'AOpticalDevice', 'A Optical Device', 'ADevice', 'A Device', 'ADevice', 'ADevice']);
            const zDev = pickDevice(links.find(Boolean), ['Z Optical Device', 'ZOpticalDevice', 'ZOpticalDevice', 'Z Optical Device', 'ZDevice', 'Z Device', 'ZDevice', 'ZDevice']);
            if (aDev) optAPrefix = xomtBaseKey(aDev);
            if (zDev) optZPrefix = xomtBaseKey(zDev);
          }
        } catch {}

        const makeTargetFor = (x: string) => {
          const cur = xomtPrefixKey(x);
          const base = xomtBaseKey(x);
          if (!firstPrefix) firstPrefix = cur;
          // If we have optical-derived prefixes and the payload contains
          // many different prefixes, prefer optical matching for assignment.
          try {
            const allPrefixes = Array.from(groups.keys()).map(k => xomtBaseKey(String(k))).filter(Boolean);
            const distinct = Array.from(new Set(allPrefixes));
            if (distinct.length > 2 && (optAPrefix || optZPrefix)) {
              if (optAPrefix && base === optAPrefix) return aRows;
              if (optZPrefix && base === optZPrefix) return zRows;
              // if not matched, fall through to order-based
            }
          } catch {}
          return cur === firstPrefix ? aRows : zRows;
        };
        for (const [xomt, items] of Array.from(groups.entries())) {

          // find c0 and m0 entries in the group, and capture any EndSku available
          let c0Dev = '';
          let c0Port = '';
          let c0Sku = '';
          let m0Dev = '';
          let m0Port = '';
          for (const it of items) {
            const e = (it.end || '').toLowerCase();
            const endSku = String(it.endSku ?? it.EndSku ?? it.EndHardwareSku ?? it.EndHardware ?? it.endHardware ?? '').trim();
            if (/\bc0\b|c0$/i.test(e) || /-c0/i.test(e)) {
              c0Dev = it.end;
              c0Port = it.endPort || c0Port;
              if (endSku) c0Sku = endSku;
            } else if (/\bm0\b|m0$/i.test(e) || /-m0/i.test(e)) {
              m0Dev = it.end;
              m0Port = it.endPort || m0Port;
            } else {
              // If we can't detect, attempt heuristics: devices containing 'c0' -> c0, 'm0' -> m0
              if (it.end.toLowerCase().includes('c0') && !c0Dev) { c0Dev = it.end; c0Port = it.endPort || c0Port; if (endSku) c0Sku = endSku; }
              if (it.end.toLowerCase().includes('m0') && !m0Dev) { m0Dev = it.end; m0Port = it.endPort || m0Port; }
            }
          }

          const row: any = {
            XOMT: xomt,
            'C0 Device': c0Dev || '',
            'C0 Port': c0Port || '',
            'Line': '', // per request: no line calculation for fallback (we set StartHardwareSku so downstream mapping can compute Line)
            'M0 Device': m0Dev || '',
            'M0 Port': m0Port || '',
            'C0 DIFF': c0Dev ? makeDiffLink(c0Dev) : '',
            'M0 DIFF': m0Dev ? makeDiffLink(m0Dev) : '',
            StartHardwareSku: c0Sku || '',
          };
          // assign to A or Z based on observed ordering/prefix
          const target = makeTargetFor(xomt);
          target.push(row);
        }

        if (aRows.length) normalized.MGFXA = aRows;
        if (zRows.length) normalized.MGFXZ = zRows;
      } else {
        // previous behavior: try Side marker split or heuristic half/half
        const aSide = arr.filter(it => String(it?.Side ?? it?.side ?? '').toLowerCase().includes('a'));
        const zSide = arr.filter(it => String(it?.Side ?? it?.side ?? '').toLowerCase().includes('z'));
        if (aSide.length) normalized.MGFXA = normalizeMgfxRows(aSide);
        if (zSide.length) normalized.MGFXZ = normalizeMgfxRows(zSide);
        if (!aSide.length && !zSide.length) {
          const mid = Math.ceil(arr.length / 2);
          normalized.MGFXA = normalizeMgfxRows(arr.slice(0, mid));
          normalized.MGFXZ = normalizeMgfxRows(arr.slice(mid));
        }
      }
    }
  }

  // Ensure stable sorting for MGFX lists
  normalized.MGFXA?.sort && normalized.MGFXA.sort((a: any, b: any) => naturalSort(a.XOMT, b.XOMT));
  normalized.MGFXZ?.sort && normalized.MGFXZ.sort((a: any, b: any) => naturalSort(a.XOMT, b.XOMT));

  // Ensure MGFX A/Z each contain placeholders for XOMT 01..06 per base prefix.
  // For any prefix (e.g. "gvx01-335-") that has some XOMT rows, if any of 01-06
  // are missing, insert placeholder rows (only XOMT populated) before that group's rows.
  const insertMissingXomtsForSide = (rows: any[] | undefined) => {
    if (!Array.isArray(rows) || rows.length === 0) return rows || [];
    const prefixRegex = /^(.*?)(\d{2})xomt$/i;
    // Group rows by prefix while preserving encounter order
    const groups: Array<{ prefix: string | null; items: any[] }> = [];
    const seenPrefixes = new Set<string>();
    for (const r of rows) {
      const xomt = String(r?.XOMT || '');
      const m = xomt.match(prefixRegex);
      const prefix = m ? m[1] : null;
      const key = prefix || '____NO_PREFIX____';
      if (!seenPrefixes.has(key)) {
        seenPrefixes.add(key);
        groups.push({ prefix: prefix, items: [] });
      }
      const grp = groups[groups.length - 1];
      grp.items.push(r);
    }

    const pad2 = (n: number) => (n < 10 ? '0' + n : String(n));
    const makePlaceholder = (prefix: string | null, n: number) => {
      const x = prefix ? `${prefix}${pad2(n)}xomt` : `${pad2(n)}xomt`;
      return {
        XOMT: x,
        'C0 Device': '',
        'C0 Port': '',
        StartHardwareSku: '',
        'M0 Device': '',
        'M0 Port': '',
        'C0 DIFF': '',
        'M0 DIFF': '',
      };
    };

    const out: any[] = [];
    for (const g of groups) {
      const prefix = g.prefix; // may be null
      // Determine existing numeric suffixes (only consider 01-06 range)
      const existing = new Set<number>();
      for (const it of g.items) {
        const x = String(it?.XOMT || '');
        const m = x.match(prefixRegex);
        if (m) {
          const num = Number(m[2]);
          if (!isNaN(num) && num >= 1 && num <= 99) existing.add(num);
        }
      }
      // compute missing in 1..6
      const missing: number[] = [];
      for (let n = 1; n <= 6; n++) if (!existing.has(n)) missing.push(n);
      // If there are missing entries and we have a prefix (i.e., rows look like gvx..NNxomt),
      // insert placeholders before the group's existing rows. If there's no prefix, don't insert.
      if (missing.length && prefix) {
        // insert placeholders in ascending order
        for (const n of missing) out.push(makePlaceholder(prefix, n));
      }
      // then append the original group rows
      out.push(...g.items);
    }
    return out;
  };

  normalized.MGFXA = insertMissingXomtsForSide(normalized.MGFXA);
  normalized.MGFXZ = insertMissingXomtsForSide(normalized.MGFXZ);

  // Refined MGFX filtering: previously required a hard-coded "-352-" segment which
  // caused valid XOMTs (e.g. ams22-53313-01xomt) to be excluded. Now we:
  // 1) Match rows whose XOMT starts with "<site>-" (case-insensitive)
  // 2) If that yields zero but site exists, fallback to rows whose XOMT contains the site code
  // 3) Only apply filtering when we have a site value; otherwise keep original lists
  try {
    const sitesForMgfx = getFirstSites(normalized, query /* current UID */);
    const siteAL = String(sitesForMgfx.a || '').trim().toLowerCase();
    const siteZL = String(sitesForMgfx.z || '').trim().toLowerCase();
    const filterBySite = (rows: any[] | undefined, site: string) => {
      if (!site) return rows || [];
      const src = Array.isArray(rows) ? rows : [];
      const primary = src.filter(r => {
        const x = String(r?.XOMT || r?.xomt || '').toLowerCase();
        return x.startsWith(site + '-');
      });
      if (primary.length) return primary;
      // fallback: contains site anywhere
      const alt = src.filter(r => String(r?.XOMT || r?.xomt || '').toLowerCase().includes(site));
      return alt.length ? alt : src; // if still none, return original to avoid emptying list
    };
    if (siteAL) normalized.MGFXA = insertMissingXomtsForSide(filterBySite(normalized.MGFXA, siteAL));
    if (siteZL) normalized.MGFXZ = insertMissingXomtsForSide(filterBySite(normalized.MGFXZ, siteZL));
  } catch {
    // Non-fatal: keep existing MGFX lists if any error arises
  }

  // Ensure final MGFX lists are sorted ascending by XOMT (numeric-aware),
  // after any filtering and placeholder insertion.
  try {
    if (Array.isArray(normalized.MGFXA)) {
      normalized.MGFXA.sort((a: any, b: any) => naturalSort(String(a?.XOMT || a?.xomt || ''), String(b?.XOMT || b?.xomt || '')));
    }
    if (Array.isArray(normalized.MGFXZ)) {
      normalized.MGFXZ.sort((a: any, b: any) => naturalSort(String(a?.XOMT || a?.xomt || ''), String(b?.XOMT || b?.xomt || '')));
    }
  } catch {}

  // Parse AllWorkflowStatus from whichever place it is present (top-level or inside OtherData)
  try {
    let wfList: any[] = [];
    if (Array.isArray(normalized.AllWorkflowStatus)) wfList = normalized.AllWorkflowStatus;
    else if (Array.isArray(result.AllWorkflowStatus)) wfList = result.AllWorkflowStatus;
    else if (typeof normalized.AllWorkflowStatus === 'string') {
      const s = String(normalized.AllWorkflowStatus || '').trim();
      if (s.startsWith('[')) wfList = JSON.parse(s);
    } else if (typeof result.AllWorkflowStatus === 'string') {
      const s = String(result.AllWorkflowStatus || '').trim();
      if (s.startsWith('[')) wfList = JSON.parse(s);
    }
    if (Array.isArray(wfList) && wfList.length) {
      const wfMap: Record<string, string> = {};
      for (const it of wfList) {
        const uid = String(it?.Uid ?? it?.UID ?? it?.uid ?? '').trim();
        if (!uid) continue;
        wfMap[uid] = niceWorkflowStatus(it?.WorkflowStatus);
      }
      try { Object.defineProperty(normalized, '__AllWorkflowStatus', { value: wfList, enumerable: false }); } catch { (normalized as any).__AllWorkflowStatus = wfList; }
      try { Object.defineProperty(normalized, '__WFStatusByUid', { value: wfMap, enumerable: false }); } catch { (normalized as any).__WFStatusByUid = wfMap; }
    }
  } catch {
    // ignore parsing errors
  }

  // Lightweight normalization for common collection key names so the UI
  // reliably finds UID rows and link fields even when upstream uses
  // slightly different casing (e.g., "Uid" vs "UID").
  try {
    if (Array.isArray(normalized.AssociatedUIDs)) {
      normalized.AssociatedUIDs = normalized.AssociatedUIDs.map((r: any) => {
        const out = { ...(r || {}) };
        // canonical UID key used across the UI
        if (!out.UID) out.UID = out.Uid ?? out.uid ?? out.Uid ?? '';
        // normalise device/port keys for downstream convenience
        if (!out['Device A']) out['Device A'] = out['DeviceA'] ?? out.DeviceA ?? out.DeviceA ?? out['ADevice'] ?? out.ADevice ?? out.DeviceA ?? '';
        if (!out['Device Z']) out['Device Z'] = out['DeviceZ'] ?? out.DeviceZ ?? out['ZDevice'] ?? out.ZDevice ?? '';
        if (!out['Site A']) out['Site A'] = out.SiteA ?? out['ASite'] ?? out.ASite ?? '';
        if (!out['Site Z']) out['Site Z'] = out.SiteZ ?? out['ZSite'] ?? out.ZSite ?? '';
        return out;
      });
    }

    if (Array.isArray(normalized.OLSLinks)) {
      normalized.OLSLinks = normalized.OLSLinks.map((r: any) => {
        const out = { ...(r || {}) };
        if (!out['A Device']) out['A Device'] = out.ADevice ?? out['ADevice'] ?? out['DeviceA'] ?? out.DeviceA ?? '';
        if (!out['A Port']) out['A Port'] = out.APort ?? out['APort'] ?? out['PortA'] ?? out.PortA ?? '';
        if (!out['Z Device']) out['Z Device'] = out.ZDevice ?? out['ZDevice'] ?? out['DeviceZ'] ?? out.DeviceZ ?? '';
        if (!out['Z Port']) out['Z Port'] = out.ZPort ?? out['ZPort'] ?? out['PortZ'] ?? out.PortZ ?? '';
        if (!out['A Optical Device']) out['A Optical Device'] = out.AOpticalDevice ?? out['AOpticalDevice'] ?? out['A Optical Device'] ?? '';
        if (!out['A Optical Port']) out['A Optical Port'] = out.AOpticalPort ?? out['AOpticalPort'] ?? out['A Optical Port'] ?? '';
        if (!out['Z Optical Device']) out['Z Optical Device'] = out.ZOpticalDevice ?? out['ZOpticalDevice'] ?? out['Z Optical Device'] ?? '';
        if (!out['Z Optical Port']) out['Z Optical Port'] = out.ZOpticalPort ?? out['ZOpticalPort'] ?? out['Z Optical Port'] ?? '';
        return out;
      });
    }
  } catch {
    // non-fatal; if normalization fails, we'll still render raw data as before
  }
      // Sort Associated UIDs in the normalized view (numeric-aware)
      if (Array.isArray(normalized.AssociatedUIDs)) {
        normalized.AssociatedUIDs.sort((a: any, b: any) => {
          const uidA = String(a?.UID || a?.Uid || a?.uid || "");
          const uidB = String(b?.UID || b?.Uid || b?.uid || "");
          return uidA.localeCompare(uidB, undefined, { numeric: true });
        });
      }

      // Ensure OLSLinks are stable-sorted on the normalized object as well
      if (Array.isArray(normalized.OLSLinks)) {
        normalized.OLSLinks.sort((a: any, b: any) => naturalSort(a.APort || a?.APort || '', b.APort || b?.APort || ''));
      }

      setData(normalized);
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

  // Build WAN Checker and Deployment Validator links on-the-fly
  const getWanLinkForSide = (src: any, side: 'A' | 'Z'): string | null => {
    try {
      if (!src) return null;
      const links: any[] = Array.isArray(src.OLSLinks) ? src.OLSLinks : [];
      const device = side === 'A'
        ? (src?.KQLData?.DeviceA || (links[0]?.['A Device'] ?? links[0]?.ADevice) || '')
        : (src?.KQLData?.DeviceZ || (links[0]?.['Z Device'] ?? links[0]?.ZDevice) || '');
      const dev = String(device || '').trim();
      if (!dev) return null;
      // Collect unique interfaces for the selected side
      const set = new Set<string>();
      for (const r of links) {
        const p = side === 'A'
          ? (r['A Port'] ?? r.APort ?? r.PortA ?? r['Port A'])
          : (r['Z Port'] ?? r.ZPort ?? r.PortZ ?? r['Port Z']);
        const v = String(p ?? '').trim();
        if (v) set.add(v);
      }
      if (!set.size) return null;
      const interfaces = Array.from(set.values()).join(';');
      return `https://phynet.trafficmanager.net/WAN?deviceName=${encodeURIComponent(dev)}&interfaces=${interfaces}`;
    } catch {
      return null;
    }
  };

  const getDeploymentValidatorLinkForSide = (src: any, side: 'A' | 'Z'): string | null => {
    try {
      if (!src) return null;
      const rows: any[] = side === 'A' ? (src.MGFXA || []) : (src.MGFXZ || []);
      if (!Array.isArray(rows) || rows.length === 0) return null;
      const set = new Set<string>();
      for (const r of rows) {
        const x = String(r?.XOMT ?? r?.xomt ?? '').trim();
        if (x) set.add(x);
      }
      if (!set.size) return null;
      const devices = Array.from(set.values()).join(',');
      return `https://phynet.trafficmanager.net/Optical/DeploymentValidator?devices=${devices}`;
    } catch {
      return null;
    }
  };

  // Normalise GDCO ticket rows from either GDCOTickets or AssociatedTickets.
  // Return objects with only the visible columns used by the UI and export,
  // and attach a non-enumerable __hiddenLink when a ticket link is present so
  // the table can render the TicketId as a clickable link without adding
  // extra visible columns.
  // Accept explicit UID override for GDCO rows
  const getGdcoRows = (src: any, forceUid?: string): any[] => {
    if (!src) return [];
  // Prefer new Logic App shape: ReleatedTickets (note: source may have this exact spelled key).
  // Fallbacks: GDCOTickets, AssociatedTickets. Some Logic Apps put tickets inside AssociatedUIDs
  // (legacy/combined payload) — detect those too by looking for TicketId/CleanTitle.
  const primary = Array.isArray(src.ReleatedTickets) && src.ReleatedTickets.length ? src.ReleatedTickets : (Array.isArray(src.GDCOTickets) && src.GDCOTickets.length ? src.GDCOTickets : null);
  const alternate = Array.isArray(src.AssociatedTickets) && src.AssociatedTickets.length ? src.AssociatedTickets : null;
    let rows: any[] = primary || alternate || [];
    let sourcePicked = primary ? 'GDCOTickets' : (alternate ? 'AssociatedTickets' : null);
    if ((!rows || !rows.length) && Array.isArray(src.AssociatedUIDs) && src.AssociatedUIDs.length) {
      // Filter AssociatedUIDs for entries that look like tickets
      const candidates = src.AssociatedUIDs.filter((r: any) => r && (r.TicketId || r.CleanTitle || r.TicketLink || r.TicketID || r.DatacenterCode));
      if (candidates.length) {
        rows = candidates;
        sourcePicked = 'AssociatedUIDs';
      }
    }
    // Debug info to help diagnose missing rows (can be removed later)
    try {
      // eslint-disable-next-line no-console
      console.debug('[getGdcoRows] sourcePicked=', sourcePicked, 'rowsCount=', (rows || []).length, 'sample=', (rows || []).slice(0,3));
    } catch {}
    // Get the searched UID from global state if available
    // Always use the searched UID (from lastSearched or uid) for every row
    let searchedUid = '';
    try {
      searchedUid = (window && (window as any).lastSearched) || (window && (window as any).uid) || '';
    } catch {}
    if (!searchedUid && src && typeof src === 'object') {
      searchedUid = src.lastSearched || src.uid || (Array.isArray(src.sourceUids) && src.sourceUids[0]) || '';
    }
    const mapped = (rows || []).map((r: any) => {
      const ticketId = String(r?.TicketId ?? r?.TicketID ?? r?.['Ticket Id'] ?? r?.['Ticket Id'] ?? r?.Ticket ?? '').trim();
      const dc = String(r?.DatacenterCode ?? r?.DCCode ?? r?.['DC Code'] ?? r?.Datacenter ?? r?.DC ?? '').trim();
      const title = String(r?.CleanTitle ?? r?.Title ?? r?.cleanTitle ?? '').trim();
      const state = String(r?.State ?? r?.Status ?? '').trim();
      const assigned = String(r?.CleanAssignedTo ?? r?.AssignedTo ?? r?.Owner ?? r?.Assigned ?? '').trim();
      const link = String(r?.TicketLink ?? r?.TicketLinkUrl ?? r?.TicketURL ?? r?.TicketUrl ?? r?.Ticket_Link ?? r?.TicketLink ?? r?.Link ?? r?.link ?? r?.URL ?? r?.Url ?? '').trim() || null;
      // Always use forceUid if provided, else blank
      const obj: any = {
        UID: forceUid || '',
        "Ticket Id": ticketId,
        "DC Code": dc,
        "Title": title,
        "State": state,
        "Assigned To": assigned,
      };
      if (link) {
        try { Object.defineProperty(obj, '__hiddenLink', { value: link, enumerable: false }); } catch { (obj as any).__hiddenLink = link; }
      }
      return obj;
    });
    // Filter out rows that have no visible content except UID (all other fields empty)
    return mapped.filter((m: any) => {
      return (
        String(m['Ticket Id'] || '').trim() ||
        String(m['Title'] || '').trim() ||
        String(m['DC Code'] || '').trim() ||
        String(m['State'] || '').trim() ||
        String(m['Assigned To'] || '').trim()
      );
    });
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
  const fallbackUid = primaryUidFor(dataNow);
  const rawStatus = String(getWFStatusFor(dataNow, fallbackUid) || '').trim();
      const isCancelled = /cancel|cancelled|canceled/i.test(rawStatus);
      const isDecom = /decom/i.test(rawStatus);
      const statusDisplay = isCancelled ? 'WF Cancelled' : isDecom ? 'DECOM' : (rawStatus || '—');
  

      const detailsHeaders = ["SRLGID", "SRLG", "SolutionID", "Status", "CIS Workflow"];
      // Prefer SRLG/SRLGID and JobId/SolutionId from the AssociatedUID matching the current UID when present
      const assocRows: any[] = Array.isArray((dataNow as any).AssociatedUIDs) ? (dataNow as any).AssociatedUIDs : [];
      // Prefer assoc matching the last searched UID; when viewing a project (no lastSearched)
      // fall back to the first AssociatedUID so SolutionID / JobId (CIS Workflow) remain present.
      let assoc: any = null;
      try {
        if (lastSearched) assoc = assocRows.find((r: any) => String(r?.UID ?? r?.Uid ?? r?.uid ?? '') === String(lastSearched ?? ''));
        if (!assoc && assocRows.length) assoc = assocRows[0];
      } catch { assoc = assocRows.length ? assocRows[0] : null; }
      // Prefer JobId from assoc if present (for CIS Workflow link)
      const jobIdPrefer = assoc?.JobId ?? dataNow?.KQLData?.JobId;
      const cisLinkPrefer = jobIdPrefer ? `https://azcis.trafficmanager.net/Public/NetworkingOptical/JobDetails/${jobIdPrefer}` : '';
      const solutionPrefer = (() => {
        try {
          const assocSol = assoc?.SolutionId ?? assoc?.SolutionID ?? assoc?.Solution ?? null;
          if (assocSol) {
            if (Array.isArray(assocSol)) return assocSol.map((v: any) => formatSolutionId(String(v))).filter(Boolean).join(', ');
            return formatSolutionId(String(assocSol));
          }
          const base = dataNow?.Base ?? dataNow?.base ?? null;
          if (base) {
            if (Array.isArray(base) && base.length) {
              const b0 = base[0];
              const bSol = b0?.SolutionId ?? b0?.SolutionID ?? b0?.Solution ?? null;
              if (bSol) {
                if (Array.isArray(bSol)) return bSol.map((v: any) => formatSolutionId(String(v))).filter(Boolean).join(', ');
                return formatSolutionId(String(bSol));
              }
            } else if (typeof base === 'object') {
              const bSol = base?.SolutionId ?? base?.SolutionID ?? base?.Solution ?? null;
              if (bSol) {
                if (Array.isArray(bSol)) return bSol.map((v: any) => formatSolutionId(String(v))).filter(Boolean).join(', ');
                return formatSolutionId(String(bSol));
              }
            }
          }
          return (getSolutionIds(dataNow) || []).map(formatSolutionId).filter(Boolean).join(', ');
        } catch { return ''; }
      })();
      const detailsRows = [
        {
          SRLGID: String(assoc?.SrlgId ?? assoc?.SRLGID ?? getSrlgIdFrom(dataNow, lastSearched) ?? ""),
          SRLG: String(assoc?.Srlg ?? assoc?.SRLG ?? getSrlgFrom(dataNow, lastSearched) ?? ""),
          SolutionID: solutionPrefer,
          Status: statusDisplay,
          "CIS Workflow": cisLinkPrefer,
        },
      ].map((r) => Object.values(r).reduce((acc: any, v: any, i: number) => ({ ...acc, [detailsHeaders[i]]: v }), {}));
      text += formatTableText("Details", detailsRows as any, detailsHeaders);
      // Tools / quick links (A/Z WAN checker + Deployment Validator)
      try {
        const aWan = getWanLinkForSide(dataNow, 'A');
        const aDeploy = getDeploymentValidatorLinkForSide(dataNow, 'A');
        const zWan = getWanLinkForSide(dataNow, 'Z');
        const zDeploy = getDeploymentValidatorLinkForSide(dataNow, 'Z');
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
  const wf = niceWorkflowStatus(r?.WorkflowStatus ?? r?.Workflow ?? wfMap?.[String(uid)]) || '';
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

    // GDCO Tickets - prefer normalized rows (this will pick up AssociatedTickets when present)
    try {
      const gdco = getGdcoRows(dataNow || {});
      const exportRows = (gdco || []).map((r: any) => ({ ...r, Link: (r as any).__hiddenLink || '' }));
      text += formatTableText("GDCO Tickets", exportRows, ["Ticket Id", "DC Code", "Title", "State", "Assigned To", "Link"]);
    } catch {
      text += formatTableText("GDCO Tickets", [], ["Ticket Id", "DC Code", "Title", "State", "Assigned To", "Link"]);
    }





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
    // Prefer values from the AssociatedUID that matches lastSearched when present
    const assocRowsForExport: any[] = Array.isArray((dataNow as any).AssociatedUIDs) ? (dataNow as any).AssociatedUIDs : [];
    // Prefer assoc matching the primary UID for this view; fallback to first assoc row
    let assocForExport: any = null;
    try {
      const pick = primaryUidFor(dataNow);
      if (pick) assocForExport = assocRowsForExport.find((r: any) => String(r?.UID ?? r?.Uid ?? r?.uid ?? '') === String(pick));
      if (!assocForExport && assocRowsForExport.length) assocForExport = assocRowsForExport[0];
    } catch { assocForExport = assocRowsForExport.length ? assocRowsForExport[0] : null; }
    const jobIdForExport = assocForExport?.JobId ?? assocForExport?.JobID ?? dataNow?.KQLData?.JobId ?? null;
    const solutionForExport = (() => {
      try {
        const assocSol = assocForExport?.SolutionId ?? assocForExport?.SolutionID ?? assocForExport?.Solution ?? null;
        if (assocSol) {
          if (Array.isArray(assocSol)) return assocSol.map((v: any) => formatSolutionId(String(v))).filter(Boolean).join(', ');
          return formatSolutionId(String(assocSol));
        }
        const base = dataNow?.Base ?? dataNow?.base ?? null;
        if (base) {
          if (Array.isArray(base) && base.length) {
            const b0 = base[0];
            const bSol = b0?.SolutionId ?? b0?.SolutionID ?? b0?.Solution ?? null;
            if (bSol) {
              if (Array.isArray(bSol)) return bSol.map((v: any) => formatSolutionId(String(v))).filter(Boolean).join(', ');
              return formatSolutionId(String(bSol));
            }
          } else if (typeof base === 'object') {
            const bSol = base?.SolutionId ?? base?.SolutionID ?? base?.Solution ?? null;
            if (bSol) {
              if (Array.isArray(bSol)) return bSol.map((v: any) => formatSolutionId(String(v))).filter(Boolean).join(', ');
              return formatSolutionId(String(bSol));
            }
          }
        }
        return (getSolutionIds(dataNow) || []).map(formatSolutionId).filter(Boolean).join(', ');
      } catch { return ''; }
    })();
    const detailsRow = [
      {
        SRLGID: String(assocForExport?.SrlgId ?? assocForExport?.SRLGID ?? getSrlgIdFrom(dataNow, lastSearched) ?? ""),
        SRLG: String(assocForExport?.Srlg ?? assocForExport?.SRLG ?? getSrlgFrom(dataNow, lastSearched) ?? ""),
        SolutionID: solutionForExport,
                Status: String(getWFStatusFor(dataNow, primaryUidFor(dataNow)) || ""),
        CIS_Workflow: jobIdForExport ? `https://azcis.trafficmanager.net/Public/NetworkingOptical/JobDetails/${jobIdForExport}` : "",
      },
    ];

    const toolsRows = [
      {
        Tool: 'A WAN Checker',
        URL: String(getWanLinkForSide(dataNow, 'A') || ""),
      },
      {
        Tool: 'A Deployment Validator',
        URL: String(getDeploymentValidatorLinkForSide(dataNow, 'A') || ""),
      },
      {
        Tool: 'Z WAN Checker',
        URL: String(getWanLinkForSide(dataNow, 'Z') || ""),
      },
      {
        Tool: 'Z Deployment Validator',
        URL: String(getDeploymentValidatorLinkForSide(dataNow, 'Z') || ""),
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
          const wf = niceWorkflowStatus(r?.WorkflowStatus ?? r?.Workflow ?? wfMap?.[String(uid)]) || '';
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
  // When combinedData provides per-UID grouped links, render a separate Link Summary
  // table for each UID so the UI shows the UID above each table.
  ...(Array.isArray((dataNow as any)?.OLSLinksByUid) && (dataNow as any).OLSLinksByUid.length ? (() => {
    const map: Record<string, any[]> = {};
    for (const g of (dataNow as any).OLSLinksByUid) {
      try { const key = `Link Summary — ${String(g.uid || '')}`; map[key] = Array.isArray(g.links) ? g.links : []; } catch { }
    }
    return map;
  })() : { "Link Summary": dataNow.OLSLinks }),
  "Associated UIDs": associatedRows,
      "GDCO Tickets": ((): any[] => {
        try {
          const gd = getGdcoRows(dataNow || {}, lastSearched || uid);
          return (gd || []).map((r: any) => ({ ...r, Link: (r as any).__hiddenLink || '' }));
        } catch { return []; }
      })(),
      "MGFX A-Side": mapMgfx(dataNow.MGFXA),
      "MGFX Z-Side": mapMgfx(dataNow.MGFXZ),
    } as Record<string, any[]>;

    // Preferred header ordering for known sections (used for per-sheet and consolidated All Details)
    const preferredHeaders: Record<string, string[]> = {
      'Details': ['SRLGID', 'SRLG', 'SolutionID', 'Status', 'CIS Workflow'],
      'Link Summary': [
        'A Device','A Port','Z Device','Z Port','A Optical Device','A Optical Port','Z Optical Device','Z Optical Port','Wirecheck'
      ],
      'Associated UIDs': ['UID','SrlgId','Action','Type','Device A','Device Z','Site A','Site Z','WF Status'],
      'GDCO Tickets': ['UID','Ticket Id','DC Code','Title','State','Assigned To','Link'],
      'MGFX A-Side': ['XOMT','C0 Device','C0 Port','Line','M0 Device','M0 Port','C0 DIFF','M0 DIFF'],
      'MGFX Z-Side': ['XOMT','C0 Device','C0 Port','Line','M0 Device','M0 Port','C0 DIFF','M0 DIFF'],
      'Tools': ['Tool','URL'],
      'Project UIDs': ['UID'],
    };

    // Build consolidated "All Details" sheet by stacking each section as its own table
    try {
      const aoa: any[] = [];
      for (const [title, rows] of Object.entries(sections)) {
        if (!Array.isArray(rows) || !rows.length) continue;
        // section title row
        aoa.push([title]);

        // compute headers for this section (preferred order + extras)
        const keysSet = new Set<string>();
        rows.forEach((r: any) => { if (r && typeof r === 'object') Object.keys(r).forEach(k => keysSet.add(k)); else keysSet.add('Value'); });
        let headersForSection: string[];
        if (preferredHeaders[title]) {
          const pref = preferredHeaders[title];
          const extras = Array.from(keysSet).filter(k => !pref.includes(k));
          headersForSection = [...pref, ...extras];
        } else {
          headersForSection = Array.from(keysSet);
        }

        // header row for the section
        aoa.push(headersForSection);

        // data rows
        for (const r of rows) {
          if (r && typeof r === 'object') {
            aoa.push(headersForSection.map(h => r[h] ?? ''));
          } else {
            const rowArr = headersForSection.map((_, idx) => idx === 0 ? String(r ?? '') : '');
            aoa.push(rowArr);
          }
        }

        // blank separator row
        aoa.push([]);
      }

      if (aoa.length) {
        const wsAll = XLSX.utils.aoa_to_sheet(aoa);
        XLSX.utils.book_append_sheet(wb, wsAll, 'All Details'.slice(0, 31));
      }
    } catch (e) {
      // non-fatal
    }

    // Export each section as its own sheet (match on-page tables)
    for (const [title, rows] of Object.entries(sections)) {
      if (!Array.isArray(rows) || !rows.length) continue;
      // Determine a stable set of headers (union of keys) so columns are consistent.
      const keysSet = new Set<string>();
      rows.forEach((r: any) => { if (r && typeof r === 'object') Object.keys(r).forEach(k => keysSet.add(k)); else keysSet.add('Value'); });
      let headers: string[];
      if (preferredHeaders[title]) {
        const pref = preferredHeaders[title];
        const extras = Array.from(keysSet).filter(k => !pref.includes(k));
        headers = [...pref, ...extras];
      } else {
        headers = Array.from(keysSet);
      }

      const normalized = (rows as any[]).map(r => {
        if (r && typeof r === 'object') {
          const out: any = {};
          headers.forEach(h => out[h] = r[h] ?? '');
          return out;
        }
        const out: any = {};
        headers.forEach(h => out[h] = '');
        out[headers[0]] = String(r ?? '');
        return out;
      });

      const ws = XLSX.utils.json_to_sheet(normalized, { header: headers });
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

  // Global toggle to expand/collapse troubleshooting rows in the Link Summary table
  const [troubleshootExpandedAll] = useState<boolean>(false);

  // Track selected Associated UIDs for the Associated UIDs table (UID -> checked)
  const [assocSelected, setAssocSelected] = useState<Record<string, boolean>>({});

  const Table = ({ title, headers, rows, highlightUid, headerRight, contextUid }: any) => {
    // Determine keys from first row to ensure consistent ordering and sorting (safe fallback)
    const keys = rows && rows[0] ? Object.keys(rows[0]) : [];

    const [sortKey, setSortKey] = useState<string | null>(null);
    const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");

    const effectiveHeaders = headers && headers.length === keys.length ? headers : keys;

  const isLinkSummary = title === 'Link Summary';

    // Build the list of displayed header/key indices for this table. For Link Summary
    // we previously hid the Wirecheck column; restore it so Wirecheck/Open buttons
    // are visible in the Link Summary table.
    const displayIndices = effectiveHeaders.map((h: string, i: number) => i);
  const displayHeaders = displayIndices.map((i: number) => effectiveHeaders[i]);
  const displayKeys = displayIndices.map((i: number) => keys[i]);

  // Column widths (only used for Link Summary). Persisted per-UID in localStorage.
    const WIDTHS_KEY = `linkSummaryWidths:${(contextUid || lastSearched) || 'global'}`;
    const [columnWidths, setColumnWidths] = useState<number[]>(() => {
      try {
        if (!isLinkSummary) return [];
        const raw = localStorage.getItem(WIDTHS_KEY);
        if (!raw) return [];
        const parsed = JSON.parse(raw || '[]');
        return Array.isArray(parsed) ? parsed : [];
      } catch { return []; }
    });
    // Keep drag state in a ref so mousemove handler can access it without re-registering
    const dragRef = React.useRef<{ index: number | null; startX: number; startWidth: number } | null>(null);

    useEffect(() => {
      if (!isLinkSummary) return;
      try { localStorage.setItem(WIDTHS_KEY, JSON.stringify(columnWidths || [])); } catch {}
    }, [columnWidths, WIDTHS_KEY, isLinkSummary]);

    const onResizeStart = (e: React.MouseEvent, colIndex: number) => {
      e.preventDefault();
      const el = e.currentTarget as HTMLElement;
      const th = el.closest('th') as HTMLElement | null;
      const startWidth = th ? th.offsetWidth : 120;
      dragRef.current = { index: colIndex, startX: e.clientX, startWidth };

      const onMove = (ev: MouseEvent) => {
        if (!dragRef.current) return;
        const delta = ev.clientX - dragRef.current.startX;
        const next = Math.max(40, Math.round(dragRef.current.startWidth + delta));
        setColumnWidths(prev => {
          const copy = prev ? prev.slice() : [];
          copy[colIndex] = next;
          return copy;
        });
      };
      const onUp = () => {
        dragRef.current = null;
        window.removeEventListener('mousemove', onMove);
        window.removeEventListener('mouseup', onUp);
      };
      window.addEventListener('mousemove', onMove);
      window.addEventListener('mouseup', onUp);
    };

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

    // Per-table persisted troubleshooting comments (keyed by current UID + row content)
    const storageKey = `troubleshootComments:${(contextUid || lastSearched) || 'global'}`;
    const [comments, setComments] = useState<Record<string, any>>(() => {
      try {
        if (!isLinkSummary) return {};
        return JSON.parse(localStorage.getItem(storageKey) || '{}');
      } catch { return {}; }
    });
    // Fetch troubleshooting notes from backend when Link Summary table loads for a UID
    useEffect(() => {
      let cancelled = false;
      async function loadTroubleshooting() {
        if (!isLinkSummary) return;
        const uidToFetch = contextUid || lastSearched;
        if (!uidToFetch) return;
        try {
          const notes = await import('../api/items').then(m => m.getTroubleshootingForUid(uidToFetch));
          if (cancelled) return;
          // notes: array of NoteEntity, each with description (JSON string of {aDevice, aOpt, zDevice, zOpt}), rowKey, etc.
          const loaded: Record<string, any> = {};
          for (const n of notes) {
            let desc = {};
            try { desc = n.description ? JSON.parse(n.description) : {}; } catch {}
            // Use LinkKey if present, else fallback to rowKey
            const linkKey = n.LinkKey || n.linkKey || n.rowKey || n.RowKey;
            if (linkKey) {
              loaded[linkKey] = { ...desc, rowKey: n.rowKey || n.RowKey, savedAt: n.savedAt || n.Timestamp };
            }
          }
          // Merge loaded notes with any local comments (local wins if edited after fetch)
          setComments(prev => {
            const merged = { ...loaded, ...prev };
            try { localStorage.setItem(storageKey, JSON.stringify(merged)); } catch {}
            return merged;
          });
        } catch (e) {
          // eslint-disable-next-line no-console
          console.warn('[Troubleshooting] Failed to load notes', e);
        }
      }
      loadTroubleshooting();
      return () => { cancelled = true; };
    }, [isLinkSummary, contextUid, storageKey]);
    useEffect(() => {
      try {
        if (!isLinkSummary) { setComments({}); return; }
        setComments(JSON.parse(localStorage.getItem(storageKey) || '{}'));
      } catch { setComments({}); }
    }, [storageKey, isLinkSummary]);
    const saveComments = (c: Record<string, any>) => {
      try { localStorage.setItem(storageKey, JSON.stringify(c || {})); } catch {}
      setComments(c || {});
    };
    // Persist a single row's comments to the Troubleshooting table when in UID or Project context
    const persistSingle = async (rowKey: string, rowObj: any) => {
      try {
        // Determine server-side uid param and partition
        // Prefer explicit contextUid (when Table is rendered for a project or provided UID),
        // otherwise fall back to the last searched UID so saves from the main UID lookup work.
        const serverUid = contextUid ? String(contextUid) : (activeProjectId ? `PROJECT_${activeProjectId}` : (lastSearched ? String(lastSearched) : null));
        if (!serverUid) return; // nothing sensible to save against
        const partition = contextUid ? `UID_${String(contextUid)}` : (activeProjectId ? `PROJECT_${activeProjectId}` : `UID_${String(lastSearched)}`);
        const alias = getAlias(getEmail());
        const payload = {
          aDevice: rowObj.aDevice || '',
          aOpt: rowObj.aOpt || '',
          zDevice: rowObj.zDevice || '',
          zOpt: rowObj.zOpt || '',
        } as any;

        // If payload is empty and we have an existing rowKey, attempt delete
        const isEmpty = !payload.aDevice && !payload.aOpt && !payload.zDevice && !payload.zOpt;
        if (isEmpty && rowObj.rowKey) {
          try {
            await deleteNoteApi(partition, rowObj.rowKey, NOTES_ENDPOINT, 'Troubleshooting');
          } catch (e) { /* best-effort */ }
          // remove local copy
          const next = { ...(comments || {}) };
          delete next[rowKey];
          saveComments(next);
          return;
        }

        const resText = await saveToStorage({
          endpoint: NOTES_ENDPOINT,
          category: 'Troubleshooting',
          uid: serverUid,
          title: 'Troubleshooting',
          description: JSON.stringify(payload),
          owner: alias || '',
          rowKey: rowObj.rowKey,
          extras: {
            LinkKey: rowKey,
            TableName: 'Troubleshooting',
            tableName: 'Troubleshooting',
            targetTable: 'Troubleshooting',
            ContextType: contextUid ? 'uid' : (activeProjectId ? 'project' : 'uid'),
            ContextId: contextUid ? String(contextUid) : (activeProjectId ? String(activeProjectId) : String(lastSearched)),
          },
        });
        try {
          const parsed = JSON.parse(resText);
          const entity = parsed?.entity || parsed?.Entity;
          const rk = entity ? (entity.RowKey || entity.rowKey) : undefined;
          const ts = entity ? (entity.Timestamp || entity.timestamp) : undefined;
          const next = { ...(comments || {}) };
          next[rowKey] = { ...(next[rowKey] || {}), ...(rowObj || {}), rowKey: rk, savedAt: ts };
          saveComments(next);
        } catch (e) {
          // ignore parse errors but keep local
        }
      } catch (e) {
        // eslint-disable-next-line no-console
        console.warn('[Table] persistSingle failed', e);
      }
    };
    // Per-table Troubleshoot expanded state (persisted per-context so each UID/table can expand independently)
    const EXPANDED_KEY = `troubleshootExpanded:${(contextUid || lastSearched) || 'global'}`;
    const [perTableExpanded, setPerTableExpanded] = useState<boolean>(() => {
      try { const raw = localStorage.getItem(EXPANDED_KEY); return raw == null ? false : raw === '1'; } catch { return false; }
    });
    useEffect(() => { try { localStorage.setItem(EXPANDED_KEY, perTableExpanded ? '1' : '0'); } catch {} }, [perTableExpanded, EXPANDED_KEY]);
    const isScrollCandidate =
      title === 'GDCO Tickets' ||
      title === 'Associated UIDs' ||
      title === 'MGFX A-Side' ||
      title === 'MGFX Z-Side';
    const shouldScroll = isScrollCandidate && Array.isArray(rows) && rows.length > 10;

    return (
      <div className="table-container">
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text className="section-title">{title}</Text>
          <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
            {/* Per-table troubleshoot toggle (operates per-UID when contextUid present). */}
            {isLinkSummary && (
              <button
                className="sleek-btn repo"
                onClick={() => setPerTableExpanded(v => !v)}
                title={perTableExpanded ? 'Hide troubleshoot rows (this table)' : 'Show troubleshoot rows (this table)'}
                style={{ marginRight: 8 }}
              >
                {perTableExpanded ? 'Hide Troubleshoot' : 'Troubleshoot'}
              </button>
            )}
            {headerRight}
            <CopyIconInline onCopy={() => copyTableText(title, rows, effectiveHeaders)} message="Table copied" />
          </div>
        </Stack>
        {noRows ? (
          <div style={{ padding: '8px 0', color: '#a6b7c6' }}>No rows to display.</div>
        ) : (
          <div style={shouldScroll ? { maxHeight: 360, overflowY: 'auto', marginTop: 4 } : undefined}>
  <table className={`data-table ${isLinkSummary ? 'compact-link-summary' : ''}`}>
          <thead>
            <tr>
              {/* Selection column for Associated UIDs (narrow, no title text) */}
              {title === 'Associated UIDs' ? (
                <th key="select-col" style={{ width: 44, minWidth: 44, textAlign: 'center' }}>
                  {/* Select-all checkbox: checked when all visible rows are selected */}
                  <input
                    type="checkbox"
                    aria-label="Select all"
                    checked={Array.isArray(rows) && rows.length > 0 && rows.every((r: any) => !!assocSelected[String(r?.UID ?? r?.Uid ?? r?.uid ?? '')])}
                    onChange={(e) => {
                      try {
                        const newMap = { ...assocSelected };
                        const visible: string[] = (rows || []).map((r: any) => String(r?.UID ?? r?.Uid ?? r?.uid ?? '').trim()).filter(Boolean);
                        if (e.target.checked) visible.forEach(u => { if (u) newMap[u] = true; });
                        else visible.forEach(u => { if (u) delete newMap[u]; });
                        setAssocSelected(newMap);
                      } catch {}
                    }}
                    style={{ accentColor: '#2fb85b' }}
                  />
                </th>
              ) : null}
              {displayHeaders.map((h: string, idx: number) => {
                const i = displayIndices[idx];
                const k = keys[i] ?? h;
                const active = sortKey === k;
                const headerLower = String(h || '').toLowerCase();
                const isStatusMini = isLinkSummary && (/admin|oper|state/i.test(String(k)) || /admin|state/i.test(String(h)));
                const isWirecheckHeader = isLinkSummary && headerLower.includes('wirecheck');
                return (
                  <th
                      key={i}
                      onClick={() => toggleSort(k)}
                      style={{
                        cursor: 'pointer',
                        userSelect: 'none',
                        textAlign: isStatusMini ? 'center' : undefined,
                        // Give the Wirecheck column more room to avoid ellipsis; still allow user resizing
                        width: columnWidths && columnWidths[i] ? columnWidths[i] : (isStatusMini ? 24 : (isWirecheckHeader ? 120 : undefined)),
                        minWidth: isStatusMini ? 24 : (isWirecheckHeader ? 96 : undefined),
                        position: 'relative',
                      }}
                    >
                      <span>{h}</span>
                      <span style={{ marginLeft: 6 }}>{active ? (sortDir === 'asc' ? '▲' : '▼') : '↕'}</span>
                      {/* Resizer handle (only show for Link Summary) */}
                      {isLinkSummary && (
                        <div
                          onMouseDown={(ev) => onResizeStart(ev as any, i)}
                          className="col-resizer"
                          title="Resize column"
                          style={{ position: 'absolute', right: 0, top: 0, bottom: 0, width: 8, cursor: 'col-resize', zIndex: 4 }}
                        />
                      )}
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
              const rowKey = JSON.stringify(keys.map(k => row[k] ?? ''));
              const rowComments = (comments && comments[rowKey]) || {};

              return (
                <React.Fragment key={i}>
                  <tr className={highlight ? 'highlight-row' : ''}>
                    {title === 'Associated UIDs' ? (
                      <td style={{ width: 44, minWidth: 44, textAlign: 'center' }}>
                        <input
                          type="checkbox"
                          checked={!!assocSelected[String(uidVal ?? '')]}
                          onChange={(e) => {
                            try {
                              const id = String(uidVal ?? '');
                              if (e.target.checked) setAssocSelected(prev => ({ ...(prev || {}), [id]: true }));
                              else setAssocSelected(prev => { const copy = { ...(prev || {}) }; delete copy[id]; return copy; });
                            } catch {}
                          }}
                          title={`Select UID ${String(uidVal ?? '')}`}
                          style={{ accentColor: '#2fb85b' }}
                        />
                      </td>
                    ) : null}
                    {displayKeys.map((key: string, j: number) => {
                      const val = row[key];

                      // Admin/Oper compact indicator
                      if (isLinkSummary && /admin|oper|state/i.test(String(key))) {
                        const v = String(val ?? '').trim();
                        const isUp = v === '1' || v.toLowerCase() === 'up' || v === 'true';
                        const isDown = v === '0' || v.toLowerCase() === 'down' || v === 'false';
                        return (
                          <td key={j} style={{ textAlign: 'center', width: 24, minWidth: 24 }} title={isUp ? 'Up' : isDown ? 'Down' : String(val ?? '')}>
                            <span className={`status-arrow ${isUp ? 'up' : isDown ? 'down' : ''}`} style={{ color: isUp ? '#107c10' : isDown ? '#d13438' : '#a6b7c6', fontWeight: 800, fontSize: 12, lineHeight: '14px' }}>
                              {isUp ? '▲' : isDown ? '▼' : ''}
                            </span>
                          </td>
                        );
                      }

                      // Link-like columns (Open / Copy)
                      const keyLower = String(key).toLowerCase();
                      const headerLower = String(displayHeaders[j] || '').toLowerCase();
                      const looksLikeLink = ['workflow', 'diff', 'ticketlink', 'url', 'link', 'wirecheck'].some(s => keyLower.includes(s) || headerLower.includes(s));
                      if (looksLikeLink) {
                        const link = val;
                        const isWirecheckCol = keyLower.includes('wirecheck') || headerLower.includes('wirecheck');
                        const fromLinkWF = isWirecheckCol && (row as any).__wirecheckFrom === 'linkwfs';
                        return (
                          <td key={j} title={String(val ?? '')}>
                            {link ? (
                              <>
                                <button
                                  className="open-btn"
                                  onClick={() => window.open(link, '_blank')}
                                  style={fromLinkWF ? { background: '#107c10', borderColor: '#0b5a0b', color: '#ffffff' } : undefined}
                                  title={fromLinkWF ? 'Matched from LinkWFs' : undefined}
                                >
                                  Open
                                </button>
                                {isWirecheckCol ? (
                                  <CopyIconInline onCopy={() => { try { navigator.clipboard.writeText(String(link)); } catch {} }} message="Link copied" />
                                ) : (
                                  <CopyIconInline onCopy={() => { try { navigator.clipboard.writeText(String(link)); } catch {} }} message="Link copied" />
                                )}
                              </>
                            ) : null}
                          </td>
                        );
                      }

                      // UID clickable behavior
                      if ((title === 'Associated UIDs' || title === 'Project UIDs') && String(key).toLowerCase() === 'uid') {
                        const v = val;
                        return (
                          <td key={j} title={`Search UID ${v}`}>
                            <span
                              className="uid-click"
                              onClick={() => {
                                const url = `${window.location.pathname}?uid=${encodeURIComponent(String(v))}`;
                                window.open(url, '_blank');
                              }}
                            >
                              {v}
                            </span>
                          </td>
                        );
                      }

                      // WF status badge for Associated UIDs
                      if (title === 'Associated UIDs' && (String(key).toLowerCase() === 'wf status' || String(displayHeaders[j]).toLowerCase() === 'wf status')) {
                        const s = String(val ?? '').trim();
                        const isCancelled = /cancel|cancelled|canceled/i.test(s);
                        const isDecom = /decom/i.test(s);
                        const isFinished = /wf\s*finished|finished/i.test(s);
                        const isInProgress = /in\s*-?\s*progress|running/i.test(s);
                        const display = s || '—';
                        if (isFinished) return (<td key={j} style={{ textAlign: 'center' }}><span className="wf-finished-badge wf-finished-pulse" style={{ color: '#00c853', fontWeight: 900, fontSize: 12, padding: '2px 8px', borderRadius: 10, border: '1px solid rgba(0,200,83,0.45)' }}>{display}</span></td>);
                        if (isInProgress) return (<td key={j} style={{ textAlign: 'center' }}><span className="wf-inprogress-badge wf-inprogress-pulse" style={{ color: '#50b3ff', fontWeight: 800, fontSize: 11, padding: '1px 6px', borderRadius: 10, border: '1px solid rgba(80,179,255,0.28)' }}>{display}</span></td>);
                        const color = (isCancelled || isDecom) ? '#d13438' : '#a6b7c6';
                        const border = (isCancelled || isDecom) ? '1px solid rgba(209,52,56,0.45)' : '1px solid rgba(166,183,198,0.35)';
                        return (<td key={j} style={{ textAlign: 'center' }}><span style={{ color, fontWeight: 700, fontSize: 12, padding: '2px 8px', borderRadius: 10, border }}>{display}</span></td>);
                      }

                      // Ticket link behavior
                      if (String(key).toLowerCase().includes('ticket') || String(displayHeaders[j]).toLowerCase().includes('ticket')) {
                        const link = findLinkForRow(row);
                        if (link) return (<td key={j}><a className="uid-click" href={String(link)} target="_blank" rel="noopener noreferrer">{val}</a>{title !== 'GDCO Tickets' && (<button className="open-btn" onClick={() => window.open(String(link), '_blank')}>Open</button>)}</td>);
                      }

                      // Default cell — show notes icon for monitored columns
                      const monitored = isLinkSummary && (key === 'A Device' || key === 'A Optical Device' || key === 'Z Device' || key === 'Z Optical Device');
                      const hasNote = Boolean(rowComments && (rowComments.aDevice || rowComments.aOpt || rowComments.zDevice || rowComments.zOpt));
                      return (
                        <td key={j} title={String(val ?? '')}>
                          {val}
                          {monitored && hasNote ? (<span title="Troubleshoot notes present" style={{ marginLeft: 6, color: '#ffd166', fontSize: 12 }}>📝</span>) : null}
                        </td>
                      );
                    })}
                  </tr>

                  {isLinkSummary && (
                    <tr className="troubleshoot-row" style={{ display: (troubleshootExpandedAll || perTableExpanded) ? 'table-row' : 'none', background: 'rgba(13,20,28,0.65)' }}>
                      {(() => {
                        const cells: any[] = [];
                        for (let idx = 0; idx < displayHeaders.length; idx++) {
                          const header = String(displayHeaders[idx] || '').toLowerCase();
                          // A Device + A Port
                          if (header === 'a device') {
                            cells.push(
                              <td key={`a-dev-${idx}`} colSpan={2}>
                                <input
                                  className="troubleshoot-input"
                                  value={rowComments.aDevice || ''}
                                  onChange={(ev) => { const next = { ...(comments || {}) }; next[rowKey] = { ...(next[rowKey] || {}), aDevice: ev.target.value }; saveComments(next); void persistSingle(rowKey, next[rowKey]); }}
                                  onKeyDown={(e) => { if ((e as any).key === 'Enter') try { (e.target as HTMLInputElement).blur(); } catch {} }}
                                  style={{ width: '100%', padding: '4px 6px', borderRadius: 2, border: '1px solid rgba(166,183,198,0.10)', background: 'transparent', color: '#d0e7ff', fontSize: 13, lineHeight: '16px' }}
                                />
                              </td>
                            );
                            idx++; // skip next (A Port)
                            continue;
                          }
                          if (header === 'a optical device') {
                            cells.push(
                              <td key={`a-opt-${idx}`} colSpan={2}>
                                <input
                                  className="troubleshoot-input"
                                  value={rowComments.aOpt || ''}
                                  onChange={(ev) => { const next = { ...(comments || {}) }; next[rowKey] = { ...(next[rowKey] || {}), aOpt: ev.target.value }; saveComments(next); void persistSingle(rowKey, next[rowKey]); }}
                                  onKeyDown={(e) => { if ((e as any).key === 'Enter') try { (e.target as HTMLInputElement).blur(); } catch {} }}
                                  style={{ width: '100%', padding: '4px 6px', borderRadius: 2, border: '1px solid rgba(166,183,198,0.10)', background: 'transparent', color: '#d0e7ff', fontSize: 13, lineHeight: '16px' }}
                                />
                              </td>
                            );
                            idx++; // skip A Optical Port
                            continue;
                          }
                          if (header === 'z device') {
                            cells.push(
                              <td key={`z-dev-${idx}`} colSpan={2}>
                                <input
                                  className="troubleshoot-input"
                                  value={rowComments.zDevice || ''}
                                  onChange={(ev) => { const next = { ...(comments || {}) }; next[rowKey] = { ...(next[rowKey] || {}), zDevice: ev.target.value }; saveComments(next); void persistSingle(rowKey, next[rowKey]); }}
                                  onKeyDown={(e) => { if ((e as any).key === 'Enter') try { (e.target as HTMLInputElement).blur(); } catch {} }}
                                  style={{ width: '100%', padding: '4px 6px', borderRadius: 2, border: '1px solid rgba(166,183,198,0.10)', background: 'transparent', color: '#d0e7ff', fontSize: 13, lineHeight: '16px' }}
                                />
                              </td>
                            );
                            idx++; // skip Z Port
                            continue;
                          }
                          if (header === 'z optical device') {
                            cells.push(
                              <td key={`z-opt-${idx}`} colSpan={2}>
                                <input
                                  className="troubleshoot-input"
                                  value={rowComments.zOpt || ''}
                                  onChange={(ev) => { const next = { ...(comments || {}) }; next[rowKey] = { ...(next[rowKey] || {}), zOpt: ev.target.value }; saveComments(next); void persistSingle(rowKey, next[rowKey]); }}
                                  onKeyDown={(e) => { if ((e as any).key === 'Enter') try { (e.target as HTMLInputElement).blur(); } catch {} }}
                                  style={{ width: '100%', padding: '4px 6px', borderRadius: 2, border: '1px solid rgba(166,183,198,0.10)', background: 'transparent', color: '#d0e7ff', fontSize: 13, lineHeight: '16px' }}
                                />
                              </td>
                            );
                            idx++; // skip Z Optical Port
                            continue;
                          }
                          // keep table alignment
                          cells.push(<td key={`empty-${idx}`} />);
                        }
                        return cells;
                      })()}
                    </tr>
                  )}
                </React.Fragment>
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

    // Prefer Increment from the AssociatedUID that matches the current UID (this is where Optic/Increment are placed by Logic Apps)
    try {
      const assocRows: any[] = Array.isArray(viewData?.AssociatedUIDs) ? viewData.AssociatedUIDs : [];
      // Prefer an AssociatedUID that matches the currently searched UID when available.
      // When viewing a saved project there may be no `lastSearched`, so fall back
      // to the first AssociatedUID row to recover Increment/OpticalSpeed values.
      const assocMatch = assocRows.find((r: any) => String(r?.UID ?? r?.Uid ?? r?.uid ?? '') === String(lastSearched ?? '')) || null;
      const assoc = assocMatch || (assocRows.length ? assocRows[0] : null);
      const incCandidate = assoc?.Increment ?? assoc?.increment ?? assoc?.OpticalSpeed ?? assoc?.IncrementGb ?? assoc?.OpticalSpeedGb ?? assoc?.Speed ?? null;
      const incNum = incCandidate != null && incCandidate !== '' && !isNaN(Number(incCandidate)) ? Number(incCandidate) : null;
      if (incNum != null && effectiveLinks.length) {
        // Use increment x number of links as requested
        const total = Math.round(incNum * effectiveLinks.length);
        const linkCount = effectiveLinks.length;
        return { main: `${total}G`, sub: `${linkCount} link${linkCount === 1 ? '' : 's'}`, count: linkCount };
      }
    } catch {}

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
  const solutionIdDisplay = (() => {
    try {
      if (!viewData) return '';
      // 1) Prefer SolutionId from the AssociatedUID that matches the current UID
      const assocRows: any[] = Array.isArray(viewData?.AssociatedUIDs) ? viewData.AssociatedUIDs : [];
      // Prefer the assoc row matching the current UID. When viewing a saved project
      // there may be no `lastSearched`; in that case prefer the first AssociatedUID
      // as the representative source so SolutionID / JobId (CIS Workflow) are preserved.
      let assoc: any = null;
      try {
        if (lastSearched) {
          assoc = assocRows.find((r: any) => String(r?.UID ?? r?.Uid ?? r?.uid ?? '') === String(lastSearched));
        }
        if (!assoc && assocRows.length) assoc = assocRows[0];
      } catch { assoc = assocRows.length ? assocRows[0] : null; }
      const assocSol = assoc?.SolutionId ?? assoc?.SolutionID ?? assoc?.Solution ?? null;
      if (assocSol) {
        if (Array.isArray(assocSol)) return assocSol.map((v: any) => formatSolutionId(String(v))).filter(Boolean).join(', ');
        return formatSolutionId(String(assocSol));
      }

      // 2) Fallback: prefer SolutionId from Base if present (some payloads use Base for link-summary rows)
      const base = viewData?.Base ?? viewData?.base ?? null;
      if (base) {
        if (Array.isArray(base) && base.length) {
          const b0 = base[0];
          const bSol = b0?.SolutionId ?? b0?.SolutionID ?? b0?.Solution ?? null;
          if (bSol) {
            if (Array.isArray(bSol)) return bSol.map((v: any) => formatSolutionId(String(v))).filter(Boolean).join(', ');
            return formatSolutionId(String(bSol));
          }
        } else if (typeof base === 'object') {
          const bSol = base?.SolutionId ?? base?.SolutionID ?? base?.Solution ?? null;
          if (bSol) {
            if (Array.isArray(bSol)) return bSol.map((v: any) => formatSolutionId(String(v))).filter(Boolean).join(', ');
            return formatSolutionId(String(bSol));
          }
        }
      }

      // 3) Last fallback: derive from KQLData / OLSLinks as before
      return (getSolutionIds(viewData) || []).map(formatSolutionId).filter(Boolean).join(', ');
    } catch {
      return '';
    }
  })();

  // Troubleshooting per-row UI was moved into the per-table inputs and persisted directly.
  // The previous bottom-page TroubleshootingSection implementation was removed to avoid duplication.

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
  {projectLoadError && <MessageBar messageBarType={MessageBarType.error}>{projectLoadError}</MessageBar>}

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
                  <button className="sleek-btn repo accent-cta" onClick={createProjectFromCurrent} title="Create a new project from the current UID">Create Project</button>
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
                  <button
                    className="sleek-btn repo"
                    onClick={() => {
                      const ap = projects.find(pp => pp.id === activeProjectId);
                      const uids: string[] = Array.from(new Set([...(ap?.data?.sourceUids || []), ...(Array.isArray(ap?.data?.AssociatedUIDs) ? (ap?.data?.AssociatedUIDs || []).map((r: any)=>String(r?.UID||r?.Uid||r?.uid||'')).filter(Boolean) : [])]));
                      void loadProjectData(uids);
                    }}
                    title="Reload all UIDs for this project"
                    style={{ marginLeft: 8, color: '#fff' }}
                    disabled={isProjectLoading}
                  >
                    {projectLoadingCount < projectTotalCount && projectTotalCount > 0
                      ? `Loading ${projectLoadingCount}/${projectTotalCount} UIDs...`
                      : 'Refresh Project Data'}
                  </button>
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
            // Ensure a stable, numeric-aware ascending order for project UIDs
            uids.sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
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
            <table className="data-table details-table" style={{ borderTop: 'none' }}>
              <thead>
                <tr>
                  <th style={{ borderBottom: 'none' }}>SRLGID</th>
                  <th style={{ borderBottom: 'none' }}>SRLG</th>
                  <th style={{ borderBottom: 'none' }}>SolutionID</th>
                  <th style={{ textAlign: 'center', borderBottom: 'none' }}>Status</th>
                  <th style={{ borderBottom: 'none' }}>CIS Workflow</th>
                  <th style={{ borderBottom: 'none' }}>Repository</th>
                  <th style={{ borderBottom: 'none' }}>Fiber Planner</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td>{getSrlgIdFrom(viewData, lastSearched) || ''}</td>
                  <td>{getSrlgFrom(viewData, lastSearched) || ''}</td>
                  <td>{solutionIdDisplay || '—'}</td>
                  <td style={{ textAlign: 'center' }}>
                    {(() => {
                      const raw = String(getWFStatusFor(viewData, primaryUidFor(viewData)) || '').trim();
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
                      // Prefer JobId from the AssociatedUID that matches the primary UID for this view;
                      // when viewing a saved project there may be no lastSearched so primaryUidFor
                      // will choose a sensible fallback (sourceUids[0] or first AssociatedUID).
                      const assocRows: any[] = Array.isArray(viewData?.AssociatedUIDs) ? viewData.AssociatedUIDs : [];
                      const pickUid = primaryUidFor(viewData);
                      let assoc: any = null;
                      try {
                        if (pickUid) assoc = assocRows.find((r: any) => String(r?.UID ?? r?.Uid ?? r?.uid ?? '') === String(pickUid));
                        if (!assoc && assocRows.length) assoc = assocRows[0];
                      } catch { assoc = assocRows.length ? assocRows[0] : null; }
                      const jobId = assoc?.JobId ?? assoc?.JobID ?? viewData?.KQLData?.JobId;
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

          {/* WAN Buttons (formulated links) - only show for single-UID live searches, not when viewing a saved project */}
          {!activeProjectId && (
            <div className="button-header-align-left">
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <Text className="side-label">A Side:</Text>
                    {(() => { const url = getWanLinkForSide(viewData, 'A'); return url; })() && (
                      <>
                        <button
                          className="sleek-btn wan"
                          onClick={() => { const u = getWanLinkForSide(viewData, 'A'); if (u) window.open(u, "_blank"); }}
                        >
                          WAN Checker
                        </button>
                        <CopyIconInline onCopy={() => { const u = getWanLinkForSide(viewData, 'A'); if (u) navigator.clipboard.writeText(String(u)); }} message="Link copied" />
                      </>
                    )}
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    {(() => { const url = getDeploymentValidatorLinkForSide(viewData, 'A'); return url; })() && (
                      <>
                        <button
                          className="sleek-btn optical"
                          onClick={() => { const u = getDeploymentValidatorLinkForSide(viewData, 'A'); if (u) window.open(u, "_blank"); }}
                        >
                          Deployment Validator
                        </button>
                        <CopyIconInline onCopy={() => { const u = getDeploymentValidatorLinkForSide(viewData, 'A'); if (u) navigator.clipboard.writeText(String(u)); }} message="Link copied" />
                      </>
                    )}
                  </div>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                  <Text className="side-label">Z Side:</Text>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    {(() => { const url = getWanLinkForSide(viewData, 'Z'); return url; })() && (
                      <>
                        <button
                          className="sleek-btn wan"
                          onClick={() => { const u = getWanLinkForSide(viewData, 'Z'); if (u) window.open(u, "_blank"); }}
                        >
                          WAN Checker
                        </button>
                        <CopyIconInline onCopy={() => { const u = getWanLinkForSide(viewData, 'Z'); if (u) navigator.clipboard.writeText(String(u)); }} message="Link copied" />
                      </>
                    )}
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    {(() => { const url = getDeploymentValidatorLinkForSide(viewData, 'Z'); return url; })() && (
                      <>
                        <button
                          className="sleek-btn optical"
                          onClick={() => { const u = getDeploymentValidatorLinkForSide(viewData, 'Z'); if (u) window.open(u, "_blank"); }}
                        >
                          Deployment Validator
                        </button>
                        <CopyIconInline onCopy={() => { const u = getDeploymentValidatorLinkForSide(viewData, 'Z'); if (u) navigator.clipboard.writeText(String(u)); }} message="Link copied" />
                      </>
                    )}
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Tables */}
                  {!(Array.isArray((viewData as any)?.OLSLinksByUid) && (viewData as any).OLSLinksByUid.length > 0) && (
                  <Table
                    title="Link Summary"
            headers={[
              "A Device",
              "A Port",
              "A Admin",
              "A Oper",
              "A Optical Device",
              "A Optical Port",
              "Z Device",
              "Z Port",
              "Z Admin",
              "Z Oper",
              "Z Optical Device",
              "Z Optical Port",
              "Speed",
              "Wirecheck",
            ]}
            rows={(() => {
              const links: any[] = Array.isArray(viewData.OLSLinks) ? viewData.OLSLinks : [];
              // If there are no link rows, synthesize a single fallback row from AssociatedUIDs (preferred) or KQLData
              if (!links.length) {
                const assocRows: any[] = Array.isArray((viewData as any)?.AssociatedUIDs) ? (viewData as any).AssociatedUIDs : [];
                const assocMatch = lastSearched
                  ? assocRows.find((r: any) => String(r?.UID ?? r?.Uid ?? r?.uid ?? '') === String(lastSearched))
                  : null;
                const assoc = assocMatch || assocRows[0] || null;
                const aDevAssoc = assoc ? String(assoc['A Device'] ?? assoc['Device A'] ?? assoc.ADevice ?? assoc.DeviceA ?? '').trim() : '';
                const zDevAssoc = assoc ? String(assoc['Z Device'] ?? assoc['Device Z'] ?? assoc.ZDevice ?? assoc.DeviceZ ?? '').trim() : '';
                const kd = (viewData as any)?.KQLData || {};
                const aDev = aDevAssoc || String(kd?.DeviceA ?? '').trim();
                const zDev = zDevAssoc || String(kd?.DeviceZ ?? '').trim();
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

              const utilRows: any[] = Array.isArray(viewData?.Utilization) ? viewData.Utilization : (Array.isArray(viewData?.utilization) ? viewData.utilization : []);
              // Parse WorkflowsString (Logic App may provide a multiline string with one URL per link row)
              const rawWorkflows: any = viewData?.WorkflowsString ?? viewData?.WorkflowsStringRaw ?? viewData?.Workflows ?? viewData?.WorkflowUrls ?? null;
              let workflowsArr: string[] = [];
              try {
                if (typeof rawWorkflows === 'string') {
                  workflowsArr = rawWorkflows.split(/\r?\n/).map((s: string) => String(s || '').trim()).filter((s: string) => !!s);
                } else if (Array.isArray(rawWorkflows)) {
                  workflowsArr = rawWorkflows.map((s: any) => String(s || '').trim()).filter((s: string) => !!s);
                }
              } catch {
                workflowsArr = [];
              }
              // Prepare AssociatedUID device fallback for per-row mapping as well
              const assocRows: any[] = Array.isArray((viewData as any)?.AssociatedUIDs) ? (viewData as any).AssociatedUIDs : [];
              const assocMatch = lastSearched
                ? assocRows.find((r: any) => String(r?.UID ?? r?.Uid ?? r?.uid ?? '') === String(lastSearched))
                : null;
              const assoc = assocMatch || assocRows[0] || null;
              const aDevAssoc = assoc ? String(assoc['A Device'] ?? assoc['Device A'] ?? assoc.ADevice ?? assoc.DeviceA ?? '').trim() : '';
              const zDevAssoc = assoc ? String(assoc['Z Device'] ?? assoc['Device Z'] ?? assoc.ZDevice ?? assoc.DeviceZ ?? '').trim() : '';

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
                // Prefer a per-payload ordered WorkflowsString if present (first URL -> first row, etc.)
                const idx = Array.isArray(links) ? links.indexOf(r) : 0;
                const wfFromArr = (workflowsArr && workflowsArr.length && typeof idx === 'number') ? (workflowsArr[idx] ?? null) : null;
                // Prefer LinkWFs match (case-insensitive A/Z device+port) when available
                const linkWFs: any[] = Array.isArray((viewData as any)?.LinkWFs) ? (viewData as any).LinkWFs : [];
                const makeKey = (ad: string, ap: string, zd: string, zp: string) => `${(ad||'').toLowerCase().replace(/\s+/g,'').trim()}|${(ap||'').toLowerCase().replace(/\s+/g,'').trim()}|${(zd||'').toLowerCase().replace(/\s+/g,'').trim()}|${(zp||'').toLowerCase().replace(/\s+/g,'').trim()}`;
                const mapWF = (() => {
                  const m = new Map<string, string>();
                  for (const it of linkWFs) {
                    const ad = String(it?.ADevice ?? it?.Adevice ?? it?.adevice ?? '').trim();
                    const ap = String(it?.APort ?? it?.Aport ?? it?.aport ?? '').trim();
                    const zd = String(it?.ZDevice ?? it?.Zdevice ?? it?.zdevice ?? '').trim();
                    const zp = String(it?.ZPort ?? it?.Zport ?? it?.zport ?? '').trim();
                    const url = String(it?.Workflow ?? it?.workflow ?? it?.Link ?? it?.URL ?? '').trim();
                    if (!ad || !ap || !zd || !zp || !url) continue;
                    const k1 = makeKey(ad, ap, zd, zp);
                    const k2 = makeKey(zd, zp, ad, ap); // reverse direction
                    m.set(k1, url);
                    if (!m.has(k2)) m.set(k2, url);
                  }
                  return m;
                })();

                const aDevNorm = String(aDevRaw ?? '').trim();
                const zDevNorm = String(zDevRaw ?? '').trim();
                const k = makeKey(aDevNorm, String(aPort||''), zDevNorm, String(zPort||''));
                const wfMatched = mapWF.get(k) || null;

                const workflow = (wfMatched ? String(wfMatched).trim() : '') || (wfFromArr ? String(wfFromArr).trim() : '') || (r["Workflow"] ?? r["workflow"] ?? r["Link"] ?? r["link"] ?? r["URL"] ?? r["Url"] ?? "");

                // Fallback to AssociatedUIDs DeviceA/DeviceZ, then KQLData DeviceA/DeviceZ if per-row device fields are blank
                const aDev = (String(aDevRaw ?? '').trim() || aDevAssoc || String(viewData?.KQLData?.DeviceA ?? '').trim());
                const zDev = (String(zDevRaw ?? '').trim() || zDevAssoc || String(viewData?.KQLData?.DeviceZ ?? '').trim());

                // Admin/Oper status for A/Z sides (support multiple possible key names; fallback to global AdminStatus/OperStatus)
                let aAdmin = r["AAdminStatus"] ?? r["AdminStatusA"] ?? r["AdminStatus_A"] ?? r["A_AdminStatus"] ?? r["A AdminStatus"] ?? r["AdminStatus"] ?? '';
                let aOper = r["AOperStatus"] ?? r["OperStatusA"] ?? r["OperStatus_A"] ?? r["A_OperStatus"] ?? r["A OperStatus"] ?? r["OperStatus"] ?? '';
                let zAdmin = r["ZAdminStatus"] ?? r["AdminStatusZ"] ?? r["AdminStatus_Z"] ?? r["Z_AdminStatus"] ?? r["Z AdminStatus"] ?? r["AdminStatus"] ?? '';
                let zOper = r["ZOperStatus"] ?? r["OperStatusZ"] ?? r["OperStatus_Z"] ?? r["Z_OperStatus"] ?? r["Z OperStatus"] ?? r["OperStatus"] ?? '';

                // Try to find a matching utilization row (match both directions and also partial matches)
                const aDevL = String(aDev || '').toLowerCase();
                const aPortL = String(aPort || '').toLowerCase();
                const zDevL = String(zDev || '').toLowerCase();
                const zPortL = String(zPort || '').toLowerCase();
                const utilMatch = utilRows.find((u: any) => {
                  const sd = String(u.StartDevice ?? u.startDevice ?? '').toLowerCase();
                  const sp = String(u.StartPort ?? u.startPort ?? '').toLowerCase();
                  const ed = String(u.EndDevice ?? u.endDevice ?? '').toLowerCase();
                  const ep = String(u.EndPort ?? u.endPort ?? '').toLowerCase();
                  if (sd === aDevL && sp === aPortL && ed === zDevL && ep === zPortL) return true;
                  if (sd === zDevL && sp === zPortL && ed === aDevL && ep === aPortL) return true;
                  if (sd === aDevL && sp === aPortL) return true;
                  if (ed === zDevL && ep === zPortL) return true;
                  return false;
                }) || null;

                // If we found utilization data, merge statuses and add speed (store as opticalGb)
                let opticalGb: number | null = null;
                if (utilMatch) {
                  const opticalSpeedRaw = utilMatch.OpticalSpeed ?? utilMatch.opticalSpeed ?? utilMatch.Optical_Speed ?? null;
                  if (opticalSpeedRaw != null && opticalSpeedRaw !== '' && !isNaN(Number(opticalSpeedRaw))) {
                    const n = Number(opticalSpeedRaw);
                    opticalGb = n > 1000 ? Math.round(n / 1000) : Math.round(n);
                  }
                  // prefer utilization-provided admin/oper statuses if present
                  aAdmin = utilMatch.AdminStatus ?? utilMatch.Admin ?? aAdmin;
                  aOper = utilMatch.OperStatus ?? utilMatch.Oper ?? aOper;
                  zAdmin = utilMatch.AdminStatus ?? utilMatch.Admin ?? zAdmin;
                  zOper = utilMatch.OperStatus ?? utilMatch.Oper ?? zOper;
                }

                const defaultInc = viewData?.KQLData?.Increment ?? null;
                const speedDisplay = opticalGb != null ? `${opticalGb}G` : (defaultInc ? `${defaultInc}G` : '');

                // Return only the visible Link Summary columns (keep Speed).
                // Per-row SRLG/router-optic details remain available on the original
                // row objects (e.g. AOpticalDevice/AOpticalPort/etc) so they can be
                // used in the Details section and AI summary panel.
                const outRow: any = {
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
                  "Speed": speedDisplay,
                  "Wirecheck": workflow,
                };
                if (wfMatched) {
                  try { Object.defineProperty(outRow, '__wirecheckFrom', { value: 'linkwfs', enumerable: false }); } catch { (outRow as any).__wirecheckFrom = 'linkwfs'; }
                }
                return outRow;
              });
            })()}
            headerRight={(() => {
              // Display a fixed "Latest Refresh" ISO timestamp converted to the user's local format,
              // and also keep any per-payload TIMESTAMP (if present) to the right of it.
              const latestIso = '2025-11-11T10:00:00Z';
              const latestLocal = formatTimestamp(latestIso);
              const ts = formatTimestamp(getTimestamp(viewData));
                return (
                <>
                  {latestLocal ? (
                    <span style={{ color: '#a6b7c6', fontSize: 12, marginRight: 8 }} title={latestIso}>
                      Latest Refresh: <b style={{ color: '#d0e7ff' }}>{latestLocal}</b>
                    </span>
                  ) : null}
                  {ts ? (
                    // per-payload timestamp (kept smaller / secondary)
                    <span style={{ color: '#a6b7c6', fontSize: 12 }} title={ts}><b style={{ color: '#d0e7ff' }}>{ts}</b></span>
                  ) : null}
                </>
                );
            })()}
          />
          )}

          {/* Per-UID Link Summary tables (appear below the main Link Summary) */}
          {Array.isArray((viewData as any)?.OLSLinksByUid) && (viewData as any).OLSLinksByUid.length > 0 && (
            <div style={{ marginTop: 8 }}>
              {(() => {
                // make a numeric-ascending copy of the per-UID groups
                const groups = Array.isArray((viewData as any).OLSLinksByUid) ? Array.from((viewData as any).OLSLinksByUid) : [];
                groups.sort((a: any, b: any) => {
                  const na = Number(a?.uid);
                  const nb = Number(b?.uid);
                  if (!isNaN(na) && !isNaN(nb)) return na - nb;
                  return String(a?.uid ?? '').localeCompare(String(b?.uid ?? ''));
                });
                return groups.map((g: any, idx: number) => {
                  const uidLabel = `UID: ${String(g?.uid || `UID-${idx}`)}`;
                  const links = Array.isArray(g?.links) ? g.links : [];
                // Reuse main Link Summary mapping logic: produce rows suitable for the Table component
                const mapLinksToRows = (linksArr: any[]) => {
                  try {
                    const utilRows: any[] = Array.isArray(viewData?.Utilization) ? viewData.Utilization : (Array.isArray(viewData?.utilization) ? viewData.utilization : []);
                    const rawWorkflows: any = viewData?.WorkflowsString ?? viewData?.WorkflowsStringRaw ?? viewData?.Workflows ?? viewData?.WorkflowUrls ?? null;
                    let workflowsArr: string[] = [];
                    try {
                      if (typeof rawWorkflows === 'string') workflowsArr = rawWorkflows.split(/\r?\n/).map((s: string) => String(s || '').trim()).filter(Boolean);
                      else if (Array.isArray(rawWorkflows)) workflowsArr = rawWorkflows.map((s: any) => String(s || '').trim()).filter(Boolean);
                    } catch { workflowsArr = []; }
                    const linkWFs: any[] = Array.isArray((viewData as any)?.LinkWFs) ? (viewData as any).LinkWFs : [];
                    const makeKey = (ad: string, ap: string, zd: string, zp: string) => `${(ad||'').toLowerCase().replace(/\s+/g,'').trim()}|${(ap||'').toLowerCase().replace(/\s+/g,'').trim()}|${(zd||'').toLowerCase().replace(/\s+/g,'').trim()}|${(zp||'').toLowerCase().replace(/\s+/g,'').trim()}`;
                    const mapWF = (() => {
                      const m = new Map<string,string>();
                      for (const it of linkWFs) {
                        const ad = String(it?.ADevice ?? it?.Adevice ?? it?.adevice ?? '').trim();
                        const ap = String(it?.APort ?? it?.Aport ?? it?.aport ?? '').trim();
                        const zd = String(it?.ZDevice ?? it?.Zdevice ?? it?.zdevice ?? '').trim();
                        const zp = String(it?.ZPort ?? it?.Zport ?? it?.zport ?? '').trim();
                        const url = String(it?.Workflow ?? it?.workflow ?? it?.Link ?? it?.URL ?? it?.Url ?? '').trim();
                        if (!ad || !ap || !zd || !zp || !url) continue;
                        const k1 = makeKey(ad, ap, zd, zp);
                        const k2 = makeKey(zd, zp, ad, ap);
                        m.set(k1, url);
                        if (!m.has(k2)) m.set(k2, url);
                      }
                      return m;
                    })();

                    return linksArr.map((r: any) => {
                      const aDevRaw = r["ADevice"] ?? r["A Device"] ?? r["A_Device"] ?? r["DeviceA"] ?? r["Device A"] ?? "";
                      const aPort = r["APort"] ?? r["A Port"] ?? r["A_Port"] ?? r["PortA"] ?? r["Port A"] ?? "";
                      const zDevRaw = r["ZDevice"] ?? r["Z Device"] ?? r["Z_Device"] ?? r["DeviceZ"] ?? r["Device Z"] ?? "";
                      const zPort = r["ZPort"] ?? r["Z Port"] ?? r["Z_Port"] ?? r["PortZ"] ?? r["Port Z"] ?? "";
                      const aOptDev = r["AOpticalDevice"] ?? r["A Optical Device"] ?? r["A_Optical_Device"] ?? r["A OpticalDevice"] ?? r["A Optical"] ?? "";
                      const aOptPort = r["AOpticalPort"] ?? r["A Optical Port"] ?? r["A_Optical_Port"] ?? r["A OpticalPort"] ?? "";
                      const zOptDev = r["ZOpticalDevice"] ?? r["Z Optical Device"] ?? r["Z_Optical_Device"] ?? r["Z OpticalDevice"] ?? r["Z Optical"] ?? "";
                      const zOptPort = r["ZOpticalPort"] ?? r["Z Optical Port"] ?? r["Z_Optical_Port"] ?? r["Z OpticalPort"] ?? "";
                      const idx = Array.isArray(linksArr) ? linksArr.indexOf(r) : 0;
                      const wfFromArr = (workflowsArr && workflowsArr.length && typeof idx === 'number') ? (workflowsArr[idx] ?? null) : null;
                      const aDevNorm = String(aDevRaw ?? '').trim();
                      const zDevNorm = String(zDevRaw ?? '').trim();
                      const k = makeKey(aDevNorm, String(aPort||''), zDevNorm, String(zPort||''));
                      const wfMatched = mapWF.get(k) || null;
                      const workflow = (wfMatched ? String(wfMatched).trim() : '') || (wfFromArr ? String(wfFromArr).trim() : '') || (r["Workflow"] ?? r["workflow"] ?? r["Link"] ?? r["link"] ?? r["URL"] ?? r["Url"] ?? "");

                      // Fallbacks
                      const assocRows: any[] = Array.isArray((viewData as any)?.AssociatedUIDs) ? (viewData as any).AssociatedUIDs : [];
                      // Prefer the AssociatedUID that corresponds to this per-UID group (g.uid).
                      // Fall back to lastSearched or the first AssociatedUID when necessary.
                      const pickUid = String(g?.uid ?? lastSearched ?? '');
                      const assocMatch = pickUid ? (assocRows.find((ar: any) => String(ar?.UID ?? ar?.Uid ?? ar?.uid ?? '') === String(pickUid)) || null) : (assocRows.length ? assocRows[0] : null);
                      const aDevAssoc = assocMatch ? String(assocMatch['A Device'] ?? assocMatch['Device A'] ?? assocMatch.ADevice ?? assocMatch.DeviceA ?? '').trim() : '';
                      const zDevAssoc = assocMatch ? String(assocMatch['Z Device'] ?? assocMatch['Device Z'] ?? assocMatch.ZDevice ?? assocMatch.DeviceZ ?? '').trim() : '';
                      const aDev = (String(aDevRaw ?? '').trim() || aDevAssoc || String(viewData?.KQLData?.DeviceA ?? '').trim());
                      const zDev = (String(zDevRaw ?? '').trim() || zDevAssoc || String(viewData?.KQLData?.DeviceZ ?? '').trim());

                      // Admin/Oper and Utilization matching
                      const aDevL = String(aDev || '').toLowerCase();
                      const aPortL = String(aPort || '').toLowerCase();
                      const zDevL = String(zDev || '').toLowerCase();
                      const zPortL = String(zPort || '').toLowerCase();
                      const utilMatch = utilRows.find((u: any) => {
                        const sd = String(u.StartDevice ?? u.startDevice ?? '').toLowerCase();
                        const sp = String(u.StartPort ?? u.startPort ?? '').toLowerCase();
                        const ed = String(u.EndDevice ?? u.endDevice ?? '').toLowerCase();
                        const ep = String(u.EndPort ?? u.endPort ?? '').toLowerCase();
                        if (sd === aDevL && sp === aPortL && ed === zDevL && ep === zPortL) return true;
                        if (sd === zDevL && sp === zPortL && ed === aDevL && ep === aPortL) return true;
                        if (sd === aDevL && sp === aPortL) return true;
                        if (ed === zDevL && ep === zPortL) return true;
                        return false;
                      }) || null;
                      let opticalGb: number | null = null;
                      let aAdmin = r["AAdminStatus"] ?? r["AdminStatusA"] ?? r["AdminStatus_A"] ?? r["A_AdminStatus"] ?? r["A AdminStatus"] ?? r["AdminStatus"] ?? '';
                      let aOper = r["AOperStatus"] ?? r["OperStatusA"] ?? r["OperStatus_A"] ?? r["A_OperStatus"] ?? r["A OperStatus"] ?? r["OperStatus"] ?? '';
                      let zAdmin = r["ZAdminStatus"] ?? r["AdminStatusZ"] ?? r["AdminStatus_Z"] ?? r["Z_AdminStatus"] ?? r["Z AdminStatus"] ?? r["AdminStatus"] ?? '';
                      let zOper = r["ZOperStatus"] ?? r["OperStatusZ"] ?? r["OperStatus_Z"] ?? r["Z_OperStatus"] ?? r["Z OperStatus"] ?? r["OperStatus"] ?? '';
                      if (utilMatch) {
                        const opticalSpeedRaw = utilMatch.OpticalSpeed ?? utilMatch.opticalSpeed ?? utilMatch.Optical_Speed ?? null;
                        if (opticalSpeedRaw != null && opticalSpeedRaw !== '' && !isNaN(Number(opticalSpeedRaw))) {
                          const n = Number(opticalSpeedRaw);
                          opticalGb = n > 1000 ? Math.round(n / 1000) : Math.round(n);
                        }
                        aAdmin = utilMatch.AdminStatus ?? utilMatch.Admin ?? aAdmin;
                        aOper = utilMatch.OperStatus ?? utilMatch.Oper ?? aOper;
                        zAdmin = utilMatch.AdminStatus ?? utilMatch.Admin ?? zAdmin;
                        zOper = utilMatch.OperStatus ?? utilMatch.Oper ?? zOper;
                      }
                      const defaultInc = viewData?.KQLData?.Increment ?? null;
                      const speedDisplay = opticalGb != null ? `${opticalGb}G` : (defaultInc ? `${defaultInc}G` : '');
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
                        "Speed": speedDisplay,
                        "Wirecheck": workflow,
                      };
                    });
                  } catch (e) { return []; }
                };

                const rows = mapLinksToRows(links || []);
                // Precompute per-UID headerRight buttons (null when not applicable)
                let headerRightButtonsVar: any = null;
                try {
                  // Always compute per-UID header controls so each UID table can render them
                  // inside its own header (including the primary/first UID). Previously we
                  // only showed them for additional UIDs which caused the first UID's
                  // controls to appear elsewhere on the page.
                  const showPerUidControls = true; // always show in-table for clarity
                  if (showPerUidControls) {
                    const assocRows: any[] = Array.isArray((viewData as any)?.AssociatedUIDs) ? (viewData as any).AssociatedUIDs : [];
                    const assocMatch = assocRows.find((r: any) => String(r?.UID ?? r?.Uid ?? r?.uid ?? '') === String(g?.uid));
                    const uidView = {
                      ...(viewData || {}),
                      OLSLinks: Array.isArray(links) ? links : [],
                      KQLData: {
                        ...(viewData?.KQLData || {}),
                        DeviceA: assocMatch ? (assocMatch['A Device'] ?? assocMatch.ADevice ?? assocMatch.DeviceA) : (viewData?.KQLData?.DeviceA),
                        DeviceZ: assocMatch ? (assocMatch['Z Device'] ?? assocMatch.ZDevice ?? assocMatch.DeviceZ) : (viewData?.KQLData?.DeviceZ),
                      },
                    };
                    const uidWanA = getWanLinkForSide(uidView, 'A');
                    const uidWanZ = getWanLinkForSide(uidView, 'Z');
                    const uidDeployA = getDeploymentValidatorLinkForSide(viewData, 'A');
                    const uidDeployZ = getDeploymentValidatorLinkForSide(viewData, 'Z');
                    const sites = getFirstSites(viewData, String(g?.uid));
                    const showSitesLabel = !activeProjectId;
                    headerRightButtonsVar = (
                      <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                          <Text className="side-label">A Side:</Text>
                          {uidWanA && (
                            <>
                              <button className="sleek-btn wan" onClick={() => { if (uidWanA) window.open(uidWanA, '_blank'); }}>
                                WAN Checker
                              </button>
                              <CopyIconInline onCopy={() => { try { navigator.clipboard.writeText(String(uidWanA)); } catch {} }} message="Link copied" />
                            </>
                          )}
                          {uidDeployA && (
                            <>
                              <button className="sleek-btn optical" onClick={() => { if (uidDeployA) window.open(uidDeployA, '_blank'); }}>
                                Deployment Validator
                              </button>
                              <CopyIconInline onCopy={() => { try { navigator.clipboard.writeText(String(uidDeployA)); } catch {} }} message="Link copied" />
                            </>
                          )}
                        </div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                          {showSitesLabel ? (<div style={{ marginRight: 8, color: '#98b6d4' }}>{sites?.a || ''}{sites?.a && sites?.z ? ' ↔ ' : ''}{sites?.z || ''}</div>) : null}
                          <Text className="side-label">Z Side:</Text>
                          {uidWanZ && (
                            <>
                              <button className="sleek-btn wan" onClick={() => { if (uidWanZ) window.open(uidWanZ, '_blank'); }}>
                                WAN Checker
                              </button>
                              <CopyIconInline onCopy={() => { try { navigator.clipboard.writeText(String(uidWanZ)); } catch {} }} message="Link copied" />
                            </>
                          )}
                          {uidDeployZ && (
                            <>
                              <button className="sleek-btn optical" onClick={() => { if (uidDeployZ) window.open(uidDeployZ, '_blank'); }}>
                                Deployment Validator
                              </button>
                              <CopyIconInline onCopy={() => { try { navigator.clipboard.writeText(String(uidDeployZ)); } catch {} }} message="Link copied" />
                            </>
                          )}
                        </div>
                      </div>
                    );
                  }
                } catch {}
                return (
                  <div key={`uid-links-${idx}`} style={{ marginTop: 6 }}>
                    <div style={{ color: '#cfe7ff', fontWeight: 700, margin: '6px 0' }}>{uidLabel}</div>
                    {/* per-UID headerRight computed above (headerRightButtonsVar) */}
                    <Table
                      contextUid={String(g?.uid ?? '')}
                      title={`Link Summary`}
                      headerRight={headerRightButtonsVar}
                      headers={[
                        "A Device",
                        "A Port",
                        "A Admin",
                        "A Oper",
                        "A Optical Device",
                        "A Optical Port",
                        "Z Device",
                        "Z Port",
                        "Z Admin",
                        "Z Oper",
                        "Z Optical Device",
                        "Z Optical Port",
                        "Speed",
                        "Wirecheck",
                      ]}
                      rows={rows}
                      highlightUid={lastSearched || String(g?.uid ?? '')}
                    />
                  </div>
                );
                });
              })()}
            </div>
          )}

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
                  const wf = niceWorkflowStatus(r?.WorkflowStatus ?? r?.Workflow ?? wfMap?.[String(uid)]) || '';
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
                <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                  <button
                    className="sleek-btn repo"
                    onClick={() => setShowAllAssociatedWF(v => !v)}
                    title={showAllAssociatedWF ? 'Show only In Progress' : 'Show all UIDs'}
                  >
                    {showAllAssociatedWF ? 'Show In Progress only' : 'Show All'}
                  </button>
                  <button
                    className="sleek-btn repo accent-cta"
                    onClick={() => {
                      try {
                        const selected = Object.keys(assocSelected || {}).filter(k => !!assocSelected[k]);
                        if (!selected.length) {
                          setProjectLoadError('Select at least one Associated UID to create a project.');
                          return;
                        }
                        createProjectFromAssociatedUIDs(selected);
                      } catch (e) {
                        setProjectLoadError('Failed to start create-from-associated flow.');
                      }
                    }}
                    title={Object.keys(assocSelected || {}).filter(k => !!assocSelected[k]).length ? 'Create project from selected Associated UIDs' : 'Select at least one Associated UID to create a project'}
                    disabled={Object.keys(assocSelected || {}).filter(k => !!assocSelected[k]).length === 0}
                    style={Object.keys(assocSelected || {}).filter(k => !!assocSelected[k]).length === 0 ? { opacity: 0.45, cursor: 'not-allowed' } : undefined}
                  >
                    Create Project
                  </button>
                </div>
              )}
              highlightUid={lastSearched || uid}
            />
            <Table
              title="GDCO Tickets"
              headers={["UID", "Ticket Id", "DC Code", "Title", "State", "Assigned To"]}
              rows={(() => {
                const searchedUid = lastSearched || uid;
                const rows = getGdcoRows(viewData || {}, lastSearched || uid) || [];
                // Always add UID column for each row
                const withUid = rows.map((r: any) => ({ UID: searchedUid, ...r }));
                if (activeProjectId) {
                  // sort ascending by UID when viewing a project
                  return withUid.slice().sort((a: any, b: any) => String(a?.UID || '').localeCompare(String(b?.UID || ''), undefined, { numeric: true }));
                }
                return withUid;
              })()}
            />
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 20 }} styles={{ root: { width: '100%' } }} className="equal-tables-row">
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
                  <div className="notice-banner warning" style={{ marginBottom: 8 }}>
                    <div className="banner-icon">!</div>
                    <div
                      className="banner-text"
                      style={{
                        fontSize: 12,
                        lineHeight: '1.15',
                        maxHeight: '2.3em',
                        overflow: 'hidden',
                        display: '-webkit-box',
                        WebkitLineClamp: 2,
                        WebkitBoxOrient: 'vertical',
                        textOverflow: 'ellipsis',
                      }}
                    >
                      Projects are only saved locally - ensure you export them.
                    </div>
                  </div>
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
                            onClick={() => { if (dragProjectId) return; handleProjectClick(p.id); }}
                            role="button"
                            tabIndex={0}
                            onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handleProjectClick(p.id); } }}
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
                                      <span className="proj-cap-total" title="Total capacity" style={{ marginLeft: 6, fontSize: 11, color: '#9adfbf' }}>{cap?.main || ''}</span>
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
        className={modalType ? `dialog-${modalType}` : undefined}
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
            <div className="modal-text">
              Are you sure you want to delete this project?
            </div>
          ) : modalType === 'create-project' ? (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
              <div className="create-project-instruction">
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
          <PrimaryButton
            className={modalType === 'create-project' ? 'accent-cta' : undefined}
            onClick={saveModal}
            text={(modalType === 'delete-section' || modalType === 'delete-project') ? 'Delete' : modalType === 'move-section' ? 'Move' : modalType === 'create-project' ? 'Create' : modalType === 'confirm-merge' ? 'Merge' : 'Save'}
          />
          <DefaultButton onClick={closeModal} text="Cancel" />
        </DialogFooter>
      </Dialog>
      {/* Overlay shown while creating project from Associated UIDs */}
      {(createFromAssocRunning || createFromAssocMessage) && (
        <div className="create-from-assoc-overlay" style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', zIndex: 9999, display: 'flex', alignItems: 'center', justifyContent: 'center', pointerEvents: 'auto' }}>
          <div className="create-from-assoc-panel" style={{
            background: createFromAssocRunning ? '#01121a' : '#072b12',
            border: `1px solid ${createFromAssocRunning ? '#234' : '#1f7a3f'}`,
            padding: 24,
            borderRadius: 8,
            minWidth: 320,
            maxWidth: '80%',
            color: createFromAssocRunning ? '#e6f6ff' : '#dff6e6',
            textAlign: 'center'
          }}>
            {createFromAssocRunning ? (
              <>
                <Spinner size={SpinnerSize.large} label={`Adding ${createFromAssocCurrent || 0} of ${createFromAssocTotal || 0} UID${(createFromAssocTotal || 0) === 1 ? '' : 's'}`} />
                  <div className="create-from-assoc-message" style={{ marginTop: 12 }}>Please wait — the project is being assembled. This will process each UID one-by-one.</div>
              </>
            ) : (
              <>
                  <div className="create-from-assoc-message create-from-assoc-success">✓ {createFromAssocMessage}</div>
                {createFromAssocFailedUids && createFromAssocFailedUids.length > 0 && (
                  <div style={{ fontSize: 12, color: '#ffb3b3', marginTop: 6 }}>
                    Failed UIDs: {createFromAssocFailedUids.join(', ')}
                  </div>
                )}
              </>
            )}
          </div>
        </div>
      )}
    </Stack>
  );
}










// Use a consistent endpoint for saving/fetching notes to avoid mismatches
// NOTE: this should be the function name (no '/api' prefix). `saveToStorage`
// will join this with `API_BASE` so having 'api/HttpTrigger1' caused
// duplicate `/api/api/...` routes when `API_BASE` already contains `/api`.
const NOTES_ENDPOINT = "HttpTrigger1";