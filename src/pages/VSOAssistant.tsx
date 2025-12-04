import React, { useState, useMemo, useEffect, useRef } from "react";
import { useNavigate } from "react-router-dom";
import {
  ComboBox,
  IComboBox,
  IComboBoxOption,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  IconButton,
  Checkbox,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Dialog,
  DialogType,
  DialogFooter,
  DefaultButton,
  TextField,
  Text,
  DatePicker,
  TooltipHost,
} from "@fluentui/react";
import "../Theme.css";
import datacenterOptions from "../data/datacenterOptions";
import COUNTRIES from "../data/Countries";
import { getRackElevationUrl } from "../data/MappedREs";
import VSOCalendar, { VsoCalendarEvent } from "../components/VSOCalendar";
import { getCalendarEntries } from "../api/items";
import { computeScopeStage } from "./utils/scope";
import { saveToStorage } from "../api/saveToStorage";
import useTelemetry from "../hooks/useTelemetry";
import { apiFetch } from "../api/http";
import { logAction } from "../api/log";

interface SpanData {
  SpanID: string;
  Status: string;
  Color: string;
  Diversity?: string;
  IDF_A?: string;
  SpliceRackA?: string;
  WiringScope?: string;
  OpticalLink?: string;
  FacilityCodeA?: string;
  FacilityCodeZ?: string;
  SpliceRackA_Unit?: string;
  SpliceRackZ_Unit?: string;
  OpticalDeviceA?: string;
  OpticalRackA_Unit?: string;
  OpticalDeviceZ?: string;
  OpticalRackZ_Unit?: string;
}

interface LogicAppResponse {
  Spans: SpanData[];
  RackElevationUrl?: string;
  DataCenter?: string;
}
interface MaintenanceWindow {
  startDate: Date | null;
  startTime: string;
  endDate: Date | null;
  endTime: string;
}

const VSOAssistant: React.FC = () => {
  const navigate = useNavigate();
  // Fluent UI icons are initialized at app startup in `src/index.tsx`

  useTelemetry('VSOAssistant');
  useEffect(() => {
    try {
      const email = localStorage.getItem("loggedInEmail") || "";
      logAction(email, "View VSO Assistant");
    } catch {
      logAction("", "View VSO Assistant");
    }
  }, []);
  const isLightTheme = typeof document !== 'undefined' && (document.documentElement.classList.contains('light-theme') || document.body.classList.contains('light-theme'));
  const labelStyles = (size: number, weight: number, mb?: number) => ({ root: { color: isLightTheme ? 'var(--accent)' : '#ccc', fontSize: size, fontWeight: `${weight}`, marginBottom: mb ?? 0 } });
  const [facilityCodeA, setFacilityCodeA] = useState<string>("");
  const [facilityCodeZ, setFacilityCodeZ] = useState<string>("");
  const [diversity, setDiversity] = useState<string[]>([]);
  const [spliceRackA, setSpliceRackA] = useState<string>();
  const [spliceRackZ, setSpliceRackZ] = useState<string>();
  const [loading, setLoading] = useState<boolean>(false);
  const [result, setResult] = useState<SpanData[]>([]);
  const [selectedSpans, setSelectedSpans] = useState<string[]>([]);
  // For drag selection
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [dragStartIndex, setDragStartIndex] = useState<number | null>(null);
  const [dragSelecting, setDragSelecting] = useState<boolean | null>(null); // true: selecting, false: deselecting
  const [error, setError] = useState<string | null>(null);
  const [showAll, setShowAll] = useState<boolean>(false);
  const [, setRackUrl] = useState<string>();
  const [rackDC, setRackDC] = useState<string>();
  const [dcSearch, setDcSearch] = useState<string>("");
  const [dcSearchZ, setDcSearchZ] = useState<string>("");
  const [country, setCountry] = useState<string>("");
  const [countrySearch, setCountrySearch] = useState<string>("");
  const countryComboRef = React.useRef<IComboBox | null>(null);
  const [isDecomMode, setIsDecomMode] = useState<boolean>(false);
  const dcComboRef = React.useRef<IComboBox | null>(null);
  const dcComboRefZ = React.useRef<IComboBox | null>(null);

  // Track whether a search was completed to show no-results banner
  const [searchDone, setSearchDone] = useState<boolean>(false);
  const [oppositePrompt, setOppositePrompt] = useState<{ show: boolean; from: 'A' | 'Z' | null }>({ show: false, from: null });
  const [oppositePromptUsed, setOppositePromptUsed] = useState<boolean>(false);
  const [, setTriedSides] = useState<{ A: boolean; Z: boolean }>({ A: false, Z: false });
  const [triedBothNoResults, setTriedBothNoResults] = useState<boolean>(false);
  // When true, ignore the A/Z exclusivity filtering so all options are visible again
  const [showAllOptions, setShowAllOptions] = useState<boolean>(false);
  // Simplified vs Details view for results table. Default to simplified (simplified=true).
  // The UI exposes a "Detailed view" toggle which is OFF by default (showing the simplified view).
  const [simplifiedView, setSimplifiedView] = useState<boolean>(true);
  // Key to force remount of the search form controls so internal component state (e.g. ComboBox freeform text)
  // is fully reset when the user hits Reset.
  const [formKey, setFormKey] = useState<number>(0);
  // UI tab state for the new tabbed search layout: A-Z, Facility (both), Z-A, Decommissioned
  const [currentTab, setCurrentTab] = useState<'A-Z' | 'Facility' | 'Z-A' | 'Decom'>('Facility');

  // Sorting state for results table
  const [sortBy, setSortBy] = useState<string>("");
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");
  // Removed per-column filters in favor of simple clickable sort

  // === Stage 2: Compose Email state ===
  const [composeOpen, setComposeOpen] = useState<boolean>(false);
  const EMAIL_TO = "opticaldri@microsoft.com"; // fixed
  const [subject, setSubject] = useState<string>("");
  const [notificationType, setNotificationType] = useState<string>("New Maintenance Scheduled");
  const [location, setLocation] = useState<string>("");
  const [lat, setLat] = useState<string>("");
  const [lng, setLng] = useState<string>("");
  // Tags: selected tag keys (strings). We'll use a themed multi-select Dropdown (no freeform entries).
  const [tags, setTags] = useState<string[]>([]);
  const [maintenanceReason, setMaintenanceReason] = useState<string>("");
  const [impactExpected, setImpactExpected] = useState<boolean>(true);
  const [startDate, setStartDate] = useState<Date | null>(null);
  const [startWarning, setStartWarning] = useState<string | null>(null);
  const [pendingEmergency, setPendingEmergency] = useState<boolean>(false);
  const [showEmergencyDialog, setShowEmergencyDialog] = useState<boolean>(false);
  // Generic field error dialog (used for past dates / invalid time ranges)
  const [showFieldErrorDialog, setShowFieldErrorDialog] = useState<boolean>(false);
  const [fieldErrorMessageDialog, setFieldErrorMessageDialog] = useState<string | null>(null);
  const [endDate, setEndDate] = useState<Date | null>(null);
  const [startTime, setStartTime] = useState<string>("00:00");
  const [endTime, setEndTime] = useState<string>("00:00");
  const [additionalWindows, setAdditionalWindows] = useState<MaintenanceWindow[]>([]);
  const [userEmail, setUserEmail] = useState<string>("");
  const [cc, setCc] = useState<string>("");

  // Refs for compose inputs so we can focus/scroll to the first invalid field on submit
  const subjectRef = React.useRef<any>(null);
  const startDateRef = React.useRef<any>(null);
  const startTimeRef = React.useRef<any>(null);
  const endDateRef = React.useRef<any>(null);
  const endTimeRef = React.useRef<any>(null);
  const locationRef = React.useRef<any>(null);
  const maintenanceReasonRef = React.useRef<any>(null);
  const ccRef = React.useRef<any>(null);

  // === Calendar state (persisted locally) ===
  const [vsoEvents, setVsoEvents] = useState<VsoCalendarEvent[]>([]);
  const [calendarDate, setCalendarDate] = useState<Date | null>(null);

  // Load persisted calendar entries from server so all users see the same events.
  useEffect(() => {
    let mounted = true;
    let timer: any = null;
    const mapItems = (items: any[]): VsoCalendarEvent[] => (items || []).map((it: any) => {
          const title = it.title || it.Title || '';
          const desc = it.description || it.Description || '';
          const savedAt = it.savedAt || it.SavedAt || it.rowKey || it.RowKey || null;
          // try parse Start:/End: ISO timestamps from description
          const startMatch = /Start:\s*([\dTZ:+.\u002D]+)\b/i.exec(desc);
          const endMatch = /End:\s*([\dTZ:+.\u002D]+)\b/i.exec(desc);
          const parseDate = (s: string | null) => {
            try { return s ? new Date(s) : null; } catch { return null; }
          };
          const start = startMatch ? parseDate(startMatch[1]) : (savedAt ? parseDate(savedAt) : null) || new Date();
          const _start = start || new Date();
          const end = endMatch ? parseDate(endMatch[1]) : new Date(_start.getFullYear(), _start.getMonth(), _start.getDate() + 1);
          const spansMatch = /Spans:\s*([^\n\r]+)/i.exec(desc);
          const spans = spansMatch ? spansMatch[1].split(',').map((s: string) => s.trim()).filter(Boolean) : [];
          return {
            id: it.rowKey || it.RowKey || `${title}-${savedAt || Math.random().toString(36).slice(2,6)}`,
            title: title,
            start: start as Date,
            end: end as Date,
            // Prefer explicit Status field if present
            status: (it.Status || it.status || 'Draft') as any,
            summary: desc ? String(desc).slice(0, 160) : undefined,
            dcCode: undefined,
            spans,
            subject: undefined,
            notificationType: undefined,
            location: undefined,
            maintenanceReason: desc || undefined,
          } as VsoCalendarEvent;
    });

    const loadOnce = async () => {
      try {
        const items = await getCalendarEntries('VsoCalendar');
        const mapped = mapItems(items);
        if (!mounted) return;
        setVsoEvents((prev) => {
          const byId = new Map(prev.map((p) => [p.id, p]));
          for (const s of mapped) byId.set(s.id, s);
          return Array.from(byId.values());
        });
      } catch (e) {
        // eslint-disable-next-line no-console
        console.warn('Failed to load calendar entries', e);
      }
    };

    // initial load
    loadOnce();
    // poll every 30s for updates so status changes propagate to other users
    timer = setInterval(loadOnce, 30_000);
    return () => { mounted = false; if (timer) clearInterval(timer); };
  }, []);

  // Ensure unique IDs across sessions
  const ensureUnique = (arr: VsoCalendarEvent[]) => {
    const seen = new Set<string>();
    const out: VsoCalendarEvent[] = [];
    for (const e of arr) {
      const id = String(e.id || "");
      if (!id) continue;
      if (seen.has(id)) continue;
      seen.add(id);
      out.push(e);
    }
    return out;
  };

  // simple validation state
  const [fieldErrors, setFieldErrors] = useState<Record<string, string>>({});
  const [showValidation, setShowValidation] = useState<boolean>(false);

  // Try to detect signed-in user's email from App Service/Static Web Apps auth
  useEffect(() => {
    try {
      // Read persisted login email (if any) so CC preview is available immediately
      const stored = localStorage.getItem("loggedInEmail");
      if (stored && stored.length > 3) setUserEmail(stored);
    } catch (e) {}

    // Load persisted calendar events (with backup fallback)
    try {
      const raw = localStorage.getItem("vsoEvents");
      const rawBackup = localStorage.getItem("vsoEventsBackup");
      const loadList = (txt?: string | null) => {
        if (!txt) return [] as any[];
        try { const a = JSON.parse(txt); return Array.isArray(a) ? a : []; } catch { return []; }
      };
  const arr = loadList(raw);
      const arrBackup = loadList(rawBackup);
      const source = (arr && arr.length ? arr : arrBackup);
      if (source && source.length) {
        const parsed: VsoCalendarEvent[] = (source || []).map((e) => {
          // Prefer date-only reconstruction when available to avoid timezone drift
          const parseYmd = (ymd?: string, fallback?: string) => {
            try {
              if (ymd && /^\d{4}-\d{2}-\d{2}$/.test(ymd)) {
                const [yy, mm, dd] = ymd.split('-').map((x: string) => parseInt(x, 10));
                return new Date(yy, (mm || 1) - 1, dd || 1);
              }
            } catch {}
            if (fallback) {
              const d = new Date(fallback);
              // Normalize to local midnight to keep it on the intended day
              return new Date(d.getFullYear(), d.getMonth(), d.getDate());
            }
            return new Date();
          };
          const start = parseYmd((e as any).startYMD, (e as any).start);
          const end = parseYmd((e as any).endYMD, (e as any).end);
          const status = (e as any).status || 'Draft';
          const id = (e as any).id || `restored-${start?.getTime() || Date.now()}-${Math.random().toString(36).slice(2,6)}`;
          const title = (e as any).title || 'Fiber Maintenance';
          return { ...(e as any), id, title, status, start, end } as VsoCalendarEvent;
        });
        setVsoEvents(ensureUnique(parsed));
      }
    } catch {}

    // Restore last viewed calendar month if available
    try {
      const saved = localStorage.getItem("vsoCalendarDate");
      if (saved) {
        const d = new Date(saved);
        setCalendarDate(new Date(d.getFullYear(), d.getMonth(), 1));
      }
    } catch {}

    const fetchUserEmail = async () => {
      try {
        const res = await apiFetch("/.auth/me", { credentials: "include" });
        if (!res.ok) return;
        const data = await res.json();
        // Handle both App Service ([identities]) and Static Web Apps ({clientPrincipal}) shapes
        const identities = Array.isArray(data)
          ? data
          : data?.clientPrincipal
          ? [{ user_claims: data.clientPrincipal?.claims || [] }]
          : [];
        for (const id of identities) {
          const claims = id?.user_claims || [];
          const getClaim = (t: string) => claims.find((c: any) => c?.typ === t)?.val || "";
          const mail =
            getClaim("emails") ||
            getClaim("email") ||
            getClaim("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress");
          if (mail && mail.length > 3) {
            setUserEmail(mail);
            try { localStorage.setItem("loggedInEmail", mail); } catch (e) {}
            break;
          }
        }
      } catch (e) {
        // ignore auth lookup failures
      }
    };

    // Invoke once to populate userEmail if available
    fetchUserEmail();
  }, []); // end login/email/calendar initialization effect

  // Persist calendar events whenever they change (with backup + timestamp)
  useEffect(() => {
    try {
      const serializable = (vsoEvents || []).map(e => {
        const start = e.start instanceof Date ? e.start : new Date(e.start as any);
        const end = e.end instanceof Date ? e.end : new Date(e.end as any);
        const ymd = (d: Date) => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
        return {
          id: e.id,
          title: e.title,
          status: e.status,
          start: start.toISOString(),
          end: end.toISOString(),
          startYMD: ymd(start),
          endYMD: ymd(end),
          summary: (e as any).summary || undefined,
          dcCode: (e as any).dcCode || undefined,
          spans: (e as any).spans || [],
          subject: (e as any).subject || undefined,
          notificationType: (e as any).notificationType || undefined,
          location: (e as any).location || undefined,
          maintenanceReason: (e as any).maintenanceReason || undefined,
        };
      });
      localStorage.setItem("vsoEvents", JSON.stringify(serializable));
      localStorage.setItem("vsoEventsBackup", JSON.stringify(serializable));
      localStorage.setItem("vsoEventsLastSaved", String(Date.now()));
    } catch {}
  }, [vsoEvents]);

  // Watch for external/local changes to storage and auto-restore if needed
  useEffect(() => {
    const onStorage = (ev: StorageEvent) => {
      if (ev.storageArea !== localStorage) return;
      if (ev.key !== 'vsoEvents') return;
      try {
        const primary = localStorage.getItem('vsoEvents');
        const backup = localStorage.getItem('vsoEventsBackup');
        const parseList = (txt?: string | null) => {
          if (!txt) return [] as any[];
          try { const a = JSON.parse(txt); return Array.isArray(a) ? a : []; } catch { return []; }
        };
        const p = parseList(primary);
        if (p && p.length) return; // still has events; no action
        const b = parseList(backup);
        if (b && b.length) {
          // Rehydrate from backup
          const restored: VsoCalendarEvent[] = b.map((e: any) => ({
            ...e,
            start: new Date(e.start),
            end: new Date(e.end),
          }));
          setVsoEvents(ensureUnique(restored));
        }
      } catch {}
    };
    window.addEventListener('storage', onStorage);
    return () => window.removeEventListener('storage', onStorage);
  }, []);

  const addWindow = () =>
    setAdditionalWindows((w) => [
      ...w,
      { startDate: null, startTime: "00:00", endDate: null, endTime: "00:00" },
    ]);

  const removeWindow = (index: number) =>
    setAdditionalWindows((w) => w.filter((_, i) => i !== index));
  const [sendLoading, setSendLoading] = useState<boolean>(false);
  const [sendSuccess, setSendSuccess] = useState<string | null>(null);
  const [sendError, setSendError] = useState<string | null>(null);
  const [showSendSuccessDialog, setShowSendSuccessDialog] = useState<boolean>(false);
  // Calendar event dialog
  const [showEventDialog, setShowEventDialog] = useState<boolean>(false);
  const [activeEventId, setActiveEventId] = useState<string | null>(null);

  const resetForm = () => {
    // Close compose and clear all compose-related state so form is empty next time
    setComposeOpen(false);
    setSelectedSpans([]);
    setSubject("");
    setNotificationType("New Maintenance Scheduled");
    setLocation("");
    setLat("");
    setLng("");
  // clear tags
  setTags([]);
    setMaintenanceReason("");
    setImpactExpected(true);
    setStartDate(null);
    setStartTime("00:00");
    setEndDate(null);
    setEndTime("00:00");
    setAdditionalWindows([]);
    setCc("");
    setFieldErrors({});
    setShowValidation(false);
    setSendError(null);
    setSendSuccess(null);
    setShowSendSuccessDialog(false);
    setStartWarning(null);
    setPendingEmergency(false);
  };

  // Fully reset page (search + compose) and return to base route
  const resetAll = () => {
    resetForm();
    // Clear search and results
    setFacilityCodeA("");
    setFacilityCodeZ("");
    setDiversity([]);
    setSpliceRackA(undefined);
    setSpliceRackZ(undefined);
    setLoading(false);
    setResult([]);
    setSelectedSpans([]);
    setError(null);
    setShowAll(false);
    setRackUrl(undefined);
    setRackDC(undefined);
    setDcSearch("");
    setDcSearchZ("");
    setSearchDone(false);
    setSortBy("");
    setSortDir("asc");
  // Ensure simplified view is the default when resetting (Detailed view toggle will be off)
  setSimplifiedView(true);
  // no-op (filters removed)
    setComposeOpen(false);
    // Force remount of the search form so any uncontrolled/internal component state is cleared
    setFormKey((f) => f + 1);
    // reset to default tab
    setCurrentTab('A-Z');
  };

  // === Diversity options ===
  const diversityOptions: IDropdownOption[] = [
    // Blank option to allow clearing selection
    { key: "", text: "" },
    { key: "West,West1,West 1,West2,West 2", text: "West, West 1, West 2" },
    { key: "East,East1,East 1,East2,East 2", text: "East, East 1, East 2" },
    { key: "North", text: "North" },
    { key: "South", text: "South" },
    { key: "Y", text: "Y" },
    { key: "Z", text: "Z" },
    // Combined options (display differs from value)
    { key: "West,North,Z", text: "West / North / Z" },
    { key: "East,South,Y", text: "East / South / Y" },
  ];

  // === Filter DCs based on input ===
  const filteredDcOptions: IComboBoxOption[] = useMemo(() => {
    // include a blank option at the top so users can clear selection
    let base = [{ key: "", text: "" }, ...datacenterOptions.map((d) => ({ key: d.key, text: d.text }))];
    const search = dcSearch.toLowerCase().trim();

    // Respect showAllOptions: when true, skip exclusivity filtering so user can see all choices again
    if (!showAllOptions) {
      // If the user has already chosen Facility Code Z, keep only the blank option for A (A choices hidden)
      if (facilityCodeZ) {
        base = base.filter((o) => o.key === "");
      }

      // If the user has chosen Facility Code A, remove any options that include 'z' in key/text so Z-related DCs disappear
      if (facilityCodeA) {
        base = base.filter((opt) => {
          if (opt.key === "") return true;
          const k = (opt.key || "").toString().toLowerCase();
          const t = (opt.text || "").toString().toLowerCase();
          return !(k.includes('z') || t.includes('z')) || opt.key === facilityCodeA;
        });
      }
    }

    const items = !search
      ? base
      : base.filter(
          (opt) =>
            opt.key.toString().toLowerCase().includes(search) ||
            opt.text.toString().toLowerCase().includes(search)
        );
    // Remove explicit (None); clicking selected option will now deselect
    return items;
  }, [dcSearch, facilityCodeA, facilityCodeZ, showAllOptions]);

  const filteredDcOptionsZ: IComboBoxOption[] = useMemo(() => {
    // include a blank option at the top so users can clear selection
    let base = [{ key: "", text: "" }, ...datacenterOptions.map((d) => ({ key: d.key, text: d.text }))];
    const search = dcSearchZ.toLowerCase().trim();

    // Respect showAllOptions: when true, skip exclusivity filtering so user can see all choices again
    if (!showAllOptions) {
      // If the user has already chosen Facility Code A, keep only the blank option for Z (Z choices hidden)
      if (facilityCodeA) {
        base = base.filter((o) => o.key === "");
      }

      // If the user has chosen Facility Code Z, remove any options that include 'a' in key/text so A-related DCs disappear
      if (facilityCodeZ) {
        base = base.filter((opt) => {
          if (opt.key === "") return true;
          const k = (opt.key || "").toString().toLowerCase();
          const t = (opt.text || "").toString().toLowerCase();
          return !(k.includes('a') || t.includes('a')) || opt.key === facilityCodeZ;
        });
      }
    }

    const items = !search
      ? base
      : base.filter(
          (opt) =>
            opt.key.toString().toLowerCase().includes(search) ||
            opt.text.toString().toLowerCase().includes(search)
        );
    return items;
  }, [dcSearchZ, facilityCodeA, facilityCodeZ, showAllOptions]);

  // === Submit ===
  const handleSubmit = async (alreadyAttemptedOpposite: boolean = false) => {
    // Require at least one facility code (A or Z). When both are present (Facility tab)
    // that's allowed and maps to the facility-pair stages (12-15).
    const hasA = !!facilityCodeA;
    const hasZ = !!facilityCodeZ;
    // If the user is on the Facility tab, Facility Code is mandatory
    if (currentTab === 'Facility' && !hasA) {
      alert("Please select a Facility Code for the Facility tab before submitting.");
      return;
    }
    if (!hasA && !hasZ) {
      alert("Please select at least one Facility Code (A or Z) before submitting.");
      return;
    }

  // Track which side(s) we're searching this request.
  // Use functional update to avoid races with previous state updates.
  const searchingA = !!facilityCodeA || !!spliceRackA;
  const searchingZ = !!facilityCodeZ || !!spliceRackZ;
  // Compute and update triedSides in a functional manner so concurrent updates don't lose data
  let computedNext: { A: boolean; Z: boolean } = { A: false, Z: false };
  setTriedSides((prev) => {
    const next = { A: prev.A || searchingA, Z: prev.Z || searchingZ };
    computedNext = next;
    return next;
  });
  // reset both-no-results flag for a fresh search unless we've already exhausted both
  if (!(computedNext.A && computedNext.Z)) setTriedBothNoResults(false);

  const email = (() => {
      try {
        return localStorage.getItem("loggedInEmail") || "";
      } catch {
        return "";
      }
    })();
    logAction(email, "Submit VSO Search", {
      facilityCodeA,
      facilityCodeZ,
      diversity,
      spliceRackA,
      spliceRackZ,
      currentTab,
    });

    setLoading(true);
    setError(null);
    setResult([]);
    setSearchDone(false);

    try {
      // Format diversity as comma-separated string for payload
      const diversityValue = (diversity && diversity.length > 0)
        ? diversity.join(', ')
        : "N";

      // If the user is on the Facility tab, use the Facility-specific payload and stage mapping
      let payload: any = { Diversity: diversityValue };
      if (currentTab === 'Facility') {
  const hasDiv = !!(diversityValue && diversityValue !== 'N');
        const hasSp = !!(spliceRackA && String(spliceRackA).trim());
        // Compute stage for Facility tab according to the specified rules
        let stage = "12";
        if (hasDiv && hasSp) stage = "14"; // Facility + Diversity + SpliceRack
        else if (hasSp) stage = "15"; // Facility + SpliceRack
        else if (hasDiv) stage = "13"; // Facility + Diversity
        else stage = "12"; // Facility only

        payload = {
          Facility: facilityCodeA || "",
          SpliceRack: spliceRackA || "",
          Diversity: diversityValue,
          Stage: stage,
        };
      } else {
        // Default behaviour: keep existing A/Z style payload
        const stage = computeScopeStage({
          facilityA: facilityCodeA,
          facilityZ: facilityCodeZ,
          diversity: diversityValue === "N" ? "" : diversityValue,
          spliceA: spliceRackA,
          spliceZ: spliceRackZ,
        });
        payload = {
          Diversity: diversityValue,
          Stage: stage,
        };
        if (facilityCodeA) payload.FacilityCodeA = facilityCodeA;
        if (facilityCodeZ) payload.FacilityCodeZ = facilityCodeZ;
        if (spliceRackA) payload.SpliceRackA = spliceRackA;
        if (spliceRackZ) payload.SpliceRackZ = spliceRackZ;
      }

      const response = await fetch("/api/LogicAppProxy", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ type: "VSO", ...payload }),
      });

      if (!response.ok) throw new Error(`HTTP error! Status: ${response.status}`);
      const data: LogicAppResponse = await response.json();

      if (data?.Spans) {
        // Ensure we always store an array in `result` even if the Logic App
        // returns a single object for Spans. This prevents runtime errors
        // like "result.filter is not a function" when callers assume an array.
        const spans = Array.isArray(data.Spans) ? data.Spans : [data.Spans];
        if (!Array.isArray(data.Spans)) console.warn('Logic App returned a non-array Spans payload, coercing to array', data.Spans);
        setResult(spans as any);
      }
      if (data?.RackElevationUrl) setRackUrl(data.RackElevationUrl);
      if (data?.DataCenter) setRackDC(data.DataCenter);
      // If no results and user searched by a splice rack, offer to try from the opposite splice side
      const noSpans = !data?.Spans || (Array.isArray(data.Spans) && data.Spans.length === 0);
      if (noSpans) {
        // If we've now tried both A and Z, show the 'adjust search' message instead of prompting
        if (computedNext.A && computedNext.Z) {
          setOppositePrompt({ show: false, from: null });
          setTriedBothNoResults(true);
          // Make all search options visible again so user can adjust selections
          setShowAllOptions(true);
        } else if (!oppositePromptUsed && !alreadyAttemptedOpposite) {
          // Only offer the opposite-side prompt if we haven't already used it for this action.
          if (spliceRackA) {
            setOppositePrompt({ show: true, from: 'A' });
          } else if (spliceRackZ) {
            setOppositePrompt({ show: true, from: 'Z' });
          }
        } else {
          // If we didn't prompt because this was already an opposite-side retry, show all options so user can edit
          if (alreadyAttemptedOpposite) setShowAllOptions(true);
        }
      }
    } catch (err: any) {
      setError(err.message || "Unknown error occurred.");
    } finally {
      setLoading(false);
      setSearchDone(true);
    }
  };

  const handleOppositeSearch = async (from: 'A' | 'Z') => {
    // Prevent repeated prompting
    setOppositePromptUsed(true);
    setOppositePrompt({ show: false, from: null });

    // Copy facility/diversity and splice rack into the opposite side and clear the original side
    if (from === 'A') {
      // Search from Z using the same inputs entered on A
      const origFacility = facilityCodeA;
      const origDiversity = diversity;
      const origSplice = spliceRackA;
      // Clear A side and set Z side values
      setFacilityCodeA("");
      setSpliceRackA(undefined);
      setFacilityCodeZ(origFacility || "");
      setSpliceRackZ(origSplice || undefined);
      setDiversity(origDiversity);
    } else {
      // Search from A using the same inputs entered on Z
      const origFacility = facilityCodeZ;
      const origDiversity = diversity;
      const origSplice = spliceRackZ;
      setFacilityCodeZ("");
      setSpliceRackZ(undefined);
      setFacilityCodeA(origFacility || "");
      setSpliceRackA(origSplice || undefined);
      setDiversity(origDiversity);
    }

    // Trigger a search after the state updates. Small timeout to ensure state is applied.
    // Pass `true` to indicate this submit is the opposite-side retry so we don't re-prompt.
    setTimeout(() => {
      handleSubmit(true);
    }, 50);
  };

  const toggleSelectSpan = (spanId: string) => {
    setSelectedSpans((prev) =>
      prev.includes(spanId) ? prev.filter((id) => id !== spanId) : [...prev, spanId]
    );
  };

  // Drag selection handlers
  const handleRowMouseDown = (rowIdx: number, spanId: string) => (e: React.MouseEvent) => {
    e.preventDefault();
    setIsDragging(true);
    setDragStartIndex(rowIdx);
    // Determine if we're selecting or deselecting based on initial state
    setDragSelecting(!selectedSpans.includes(spanId));
    // Immediately update selection for the first row
    setSelectedSpans((prev) => {
      if (!selectedSpans.includes(spanId)) return [...prev, spanId];
      else return prev.filter((id) => id !== spanId);
    });
  };

  const handleRowMouseEnter = (rowIdx: number, spanId: string) => (e: React.MouseEvent) => {
    if (!isDragging || dragStartIndex === null || dragSelecting === null) return;
    // Select all between dragStartIndex and rowIdx
    const min = Math.min(dragStartIndex, rowIdx);
    const max = Math.max(dragStartIndex, rowIdx);
    const spanIdsInRange = sortedResults.slice(min, max + 1).map(r => r.SpanID);
    setSelectedSpans((prev) => {
      if (dragSelecting) {
        // Add all in range
        return Array.from(new Set([...prev, ...spanIdsInRange]));
      } else {
        // Remove all in range
        return prev.filter(id => !spanIdsInRange.includes(id));
      }
    });
  };

  const handleRowMouseUp = () => {
    setIsDragging(false);
    setDragStartIndex(null);
    setDragSelecting(null);
  };

  // End drag if mouse leaves table
  const handleTableMouseLeave = () => {
    setIsDragging(false);
    setDragStartIndex(null);
    setDragSelecting(null);
  };

  const filteredResultsBase = showAll
    ? result
    : result.filter((r) => {
        // Exclude spans with State === 'New' from the main (production) results
        const state = ((r as any)?.State || "").toString().toLowerCase();
        if (state === "new") return false;
        return ((r && (r.Status || "")).toString().toLowerCase() === "inproduction");
      });

  // Accessor for sorting
  const getSortValue = (row: SpanData, key: string): string | number => {
    const v = (row as any)[key];
    if (v === undefined || v === null) return "";
    return v as any;
  };

  const sortedResults = useMemo(() => {
    const rows = [...filteredResultsBase];
    if (!sortBy) return rows;
    rows.sort((a, b) => {
      const av = getSortValue(a, sortBy);
      const bv = getSortValue(b, sortBy);
      // numeric-aware compare where possible
      const an = typeof av === "string" ? Number(av) : (av as number);
      const bn = typeof bv === "string" ? Number(bv) : (bv as number);
      const aIsNum = !isNaN(an as number) && an !== (undefined as any) && an !== null && av !== "";
      const bIsNum = !isNaN(bn as number) && bn !== (undefined as any) && bn !== null && bv !== "";
      if (aIsNum && bIsNum) {
        return (an as number) - (bn as number);
      }
      const as = (av ?? "").toString().toLowerCase();
      const bs = (bv ?? "").toString().toLowerCase();
      return as.localeCompare(bs);
    });
    if (sortDir === "desc") rows.reverse();
    return rows;
  }, [filteredResultsBase, sortBy, sortDir]);

  // Export current filtered results (full detailed view) as CSV for Excel
  const handleExportSpansToCsv = () => {
    const rows = [...filteredResultsBase];
    if (!rows.length) return;

    // Export all detailed span fields except Color and OpticalLink;
    // append Status as the last column for readability in Excel.
    const columns: string[] = [
      'SpanID',
      'Diversity',
      'IDF_A',
      'SpliceRackA',
      'WiringScope',
      'FacilityCodeA',
      'FacilityCodeZ',
      'SpliceRackA_Unit',
      'SpliceRackZ_Unit',
      'OpticalDeviceA',
      'OpticalRackA_Unit',
      'OpticalDeviceZ',
      'OpticalRackZ_Unit',
      'SpanType',
      'Status',
    ];

    const esc = (val: any): string => {
      if (val === undefined || val === null) return '';
      const s = String(val);
      const mustQuote = /[",\n\r]/.test(s);
      const inner = s.replace(/"/g, '""');
      return mustQuote ? `"${inner}"` : inner;
    };

    const header = columns.join(',');
    const body = rows
      .map((r) => columns.map((c) => esc((r as any)[c])).join(','))
      .join('\r\n');

    const csv = `${header}\r\n${body}`;
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;

    // Build a friendly Excel filename based on FacilityCode and optional Diversity
    const baseFacility = (facilityCodeA || facilityCodeZ || 'Spans').toString().trim() || 'Spans';
    const diversityLabel = (diversity && diversity.length ? `_${diversity.join('-')}` : '');
    const safeBase = `${baseFacility}${diversityLabel}_Spans`
      .replace(/[^A-Za-z0-9_-]+/g, '_')
      .replace(/_+/g, '_')
      .replace(/^_+|_+$/g, '');
    a.download = `${safeBase || 'Spans'}.csv`;

    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleSort = (key: string) => {
    if (sortBy === key) {
      setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    } else {
      setSortBy(key);
      setSortDir("asc");
    }
  };

  const comboBoxStyles = {
    root: { width: "100%" },
    input: {
      color: "var(--vso-input-text)",
      backgroundColor: "var(--vso-input-bg)",
      height: 42,
      border: "1px solid var(--vso-input-border)",
      borderRadius: 8,
      paddingLeft: 10,
    },
    callout: { background: "var(--vso-dropdown-bg)", maxHeight: 240, overflowY: "auto" },
    optionsContainer: { background: "var(--vso-dropdown-bg)" },
    caretDownWrapper: { color: "var(--vso-dropdown-text)" },
  } as const;

  // Table ref for column resizing
  const tableRef = useRef<HTMLTableElement | null>(null);

  // Column resize handler: when the user drags the resizer, adjust the
  // corresponding th width and all body cells in that column to match.
  const startColumnResize = (ev: React.MouseEvent, resizerEl: HTMLElement) => {
    ev.preventDefault();
    ev.stopPropagation();
    if (!tableRef.current) return;
    const th = resizerEl.parentElement as HTMLTableCellElement | null;
    if (!th) return;
    const table = tableRef.current;
    const startX = (ev.nativeEvent as MouseEvent).clientX;
    const startWidth = th.getBoundingClientRect().width;
    // ensure we operate on visible rows
    const rows = Array.from(table.tBodies[0]?.rows || [] as any[]);

    // Prevent text selection while dragging
    const prevUserSelect = document.body.style.userSelect;
    document.body.style.userSelect = 'none';
    document.body.style.cursor = 'col-resize';

    const onMove = (e: MouseEvent) => {
      const dx = e.clientX - startX;
      const next = Math.max(36, Math.round(startWidth + dx));
      try {
        th.style.width = `${next}px`;
        // Apply to each row cell at the same index
        const cellIndex = (th as any).cellIndex;
        for (const r of rows) {
          const cell = r.cells[cellIndex];
          if (cell) cell.style.width = `${next}px`;
        }
      } catch {}
    };

    const onUp = () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      document.body.style.userSelect = prevUserSelect || '';
      document.body.style.cursor = '';
    };

    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };

  // Build country options from Countries list (allow typing to filter but require selection)
  const countryOptions: IComboBoxOption[] = useMemo(() => {
    const base = [...COUNTRIES].sort((a, b) => a.localeCompare(b)).map((c) => ({ key: c, text: c }));
    const search = (countrySearch || "").toString().trim().toLowerCase();
    if (!search) return base;
    // Prefer startsWith matches then includes
    const starts = base.filter(o => (o.key as string).toLowerCase().startsWith(search));
    if (starts.length) return starts.concat(base.filter(o => !starts.includes(o) && (o.key as string).toLowerCase().includes(search)));
    return base.filter(o => (o.key as string).toLowerCase().includes(search) || (o.text || '').toString().toLowerCase().includes(search));
  }, [countrySearch]);

  const dropdownStyles = {
    root: { width: "100%" },
    dropdown: { backgroundColor: "var(--vso-input-bg)", color: "var(--vso-input-text)", borderRadius: 8 },
    title: {
      background: "var(--vso-input-bg)",
      color: "var(--vso-input-text)",
      border: "1px solid var(--vso-input-border)",
      borderRadius: 8,
      height: 42,
      display: "flex",
      alignItems: "center",
      paddingLeft: 10,
    },
    caretDownWrapper: { color: "var(--vso-dropdown-text)" },
    dropdownItemsWrapper: { background: "var(--vso-dropdown-bg)" },
    callout: { background: "var(--vso-dropdown-bg)" },
    dropdownItem: {
      background: "transparent",
      color: "var(--vso-dropdown-text)",
      selectors: {
        ':hover': {
          background: 'var(--vso-dropdown-item-hover-bg)',
          color: 'var(--vso-dropdown-text)',
        },
        ':active': {
          background: 'var(--vso-dropdown-item-selected-bg)',
          color: 'var(--vso-dropdown-text)',
        },
      },
    },
    dropdownItemSelected: {
      background: 'var(--vso-dropdown-item-selected-bg)',
      color: 'var(--vso-dropdown-text)',
      selectors: {
        ':hover': {
          background: 'var(--vso-dropdown-item-hover-bg)',
          color: 'var(--vso-dropdown-text)',
        },
      },
    },
  } as const;

  // Make Diversity placeholder text exactly match TextField placeholder (e.g., AM111)
  const diversityDropdownStyles = {
    ...dropdownStyles,
    title: {
      ...dropdownStyles.title,
    },
    titleIsPlaceHolder: {
      ...dropdownStyles.title,
      color: "#a6b7c6",
      opacity: 0.8,
      fontStyle: "normal",
      fontSize: 14,
      fontWeight: 400,
    },
  } as const;

  const textFieldStyles = {
    fieldGroup: {
      background: "var(--vso-input-bg)",
      border: "1px solid var(--vso-input-border)",
      borderRadius: 8,
      height: 42,
      selectors: {
        ':hover': {
          border: '1px solid var(--vso-input-border-active)',
        },
        ':focus': {
          border: '1px solid var(--vso-input-border-active)',
        },
      },
    },
    field: {
      background: "var(--vso-input-bg)",
      color: "var(--vso-input-text)",
      selectors: {
        '::placeholder': {
          color: 'var(--vso-input-placeholder)',
        },
      },
    },
  } as const;

  // Time dropdown styles (reuse dropdownStyles but allow container width to control it)
  const timeDropdownStyles = {
    ...dropdownStyles,
    root: { width: '100%' },
  } as const;

  // Light-theme friendly styles for the Decommissioned Spans country ComboBox
  const decomCountryComboStyles = {
    ...comboBoxStyles,
    root: { width: '100%' },
    input: {
      color: 'var(--vso-input-text)',
      background: 'var(--vso-input-bg)',
    },
  } as any;

  // Option renderer to force dark readable text when light theme is active (overrides global white !important)
  const renderLightOption = (option?: IComboBoxOption): JSX.Element => {
    const label = (option?.text || option?.key || '') as string;
    return <span style={{ color: 'var(--vso-input-text)', fontSize: 13 }}>{label}</span>;
  };

  // Dark DatePicker styles to avoid white-on-white
  const datePickerStyles: any = {
    root: { width: 220 },
    textField: {
      fieldGroup: {
        background: "var(--vso-input-bg)",
        border: "1px solid var(--vso-input-border)",
        borderRadius: 8,
        height: 42,
        selectors: {
          ':hover': {
            border: '1px solid var(--vso-input-border-active)',
          },
          ':focus': {
            border: '1px solid var(--vso-input-border-active)',
          },
        },
      },
      field: {
        color: "var(--vso-input-text)",
        background: "var(--vso-input-bg)",
        selectors: { '::placeholder': { color: 'var(--vso-input-placeholder)', opacity: 0.9 } },
      },
    },
    callout: { background: "var(--vso-dropdown-bg)" },
    // dayPicker (calendar) styles to keep popover theme-aware
    dayPicker: {
      root: { background: 'var(--vso-dropdown-bg)', color: 'var(--vso-dropdown-text)' },
      monthPickerVisible: {},
      showWeekNumbers: {},
    },
  };

  // Helpers to return error-aware styles (highlight red when validation marks a field)
  const getTextFieldStyles = (key: string) => {
    const base: any = textFieldStyles as any;
    const has = showValidation && !!(fieldErrors as any)[key];
    if (!has) return base;
    const fg = { ...(base.fieldGroup || {}) };
    const selectors = { ...(fg.selectors || {}) };
    fg.border = '1px solid #a80000';
    selectors[':hover'] = { border: '1px solid #a80000' };
    selectors[':focus'] = { border: '1px solid #a80000' };
    fg.selectors = selectors;
    return { ...base, fieldGroup: fg } as any;
  };

  const getDatePickerStyles = (key: string) => {
    const base: any = datePickerStyles || {};
    const has = showValidation && !!(fieldErrors as any)[key];
    if (!has) return base;
    const tg = { ...(base.textField || {}) };
    const fg = { ...(tg.fieldGroup || {}) };
    const selectors = { ...(fg.selectors || {}) };
    fg.border = '1px solid #a80000';
    selectors[':hover'] = { border: '1px solid #a80000' };
    selectors[':focus'] = { border: '1px solid #a80000' };
    fg.selectors = selectors;
    tg.fieldGroup = fg;
    return { ...base, textField: tg } as any;
  };

  const getDropdownStyles = (key: string, baseStyles: any = dropdownStyles) => {
    const base: any = baseStyles || dropdownStyles;
    const has = showValidation && !!(fieldErrors as any)[key];
    if (!has) return base;
    const t = { ...(base.title || {}) };
    t.border = '1px solid #a80000';
    return { ...base, title: t } as any;
  };

  // Unique DC codes present in the current result set (Facility A/Z columns)
  const availableDcOptions: IDropdownOption[] = useMemo(() => {
    const set = new Set<string>();
    for (const r of result || []) {
      if (!r || typeof r !== 'object') continue;
      const dc = (r as any).Datacenter || (r as any).DataCenter;
      if (dc && typeof dc === 'string' && dc.trim()) set.add(dc.trim());
    }
    // Convert to sorted options (only Datacenter codes present in the DC column)
    return Array.from(set)
      .sort()
      .map((d) => ({ key: d, text: d }));
  }, [result]);

  // Build 30-min interval time options in 24h format
  const timeOptions: IDropdownOption[] = useMemo(() => {
    const opts: IDropdownOption[] = [];
    for (let h = 0; h < 24; h++) {
      for (let m = 0; m < 60; m += 30) {
        const hh = h.toString().padStart(2, "0");
        const mm = m.toString().padStart(2, "0");
        const text = `${hh}:${mm}`;
        opts.push({ key: text, text });
      }
    }
    return opts;
  }, []);

  // Tag options for the themed multi-select Dropdown (no freeform creation)
  const tagOptions: IDropdownOption[] = [
    { key: "1st Party Maintenance", text: "1st Party Maintenance" },
    { key: "3rd Party Maintenance", text: "3rd Party Maintenance" },
    { key: "50%Impact", text: "50%Impact" },
    { key: "AXENT", text: "AXENT" },
    { key: "AznetIDC: WAN", text: "AznetIDC: WAN" },
    { key: "Beanfield", text: "Beanfield" },
    { key: "Colt Technology Services", text: "Colt Technology Services" },
    { key: "CORE", text: "CORE" },
    { key: "drain", text: "drain" },
    { key: "East-West Spans", text: "East-West Spans" },
    { key: "euNetworks", text: "euNetworks" },
    { key: "FIber_Activity_Core", text: "FIber_Activity_Core" },
    { key: "LTIM_Optical", text: "LTIM_Optical" },
    { key: "Microsoft", text: "Microsoft" },
    { key: "no WAN links", text: "no WAN links" },
    { key: "nodrain", text: "nodrain" },
    { key: "npa-im", text: "npa-im" },
    { key: "Open Fiber", text: "Open Fiber" },
    { key: "Open Fiber S.p.A.", text: "Open Fiber S.p.A." },
    { key: "Optical", text: "Optical" },
    { key: "pass-fm", text: "pass-fm" },
    { key: "peering", text: "peering" },
    { key: "Sipartech", text: "Sipartech" },
    { key: "WAN", text: "WAN" },
    { key: "WANOKR-Ga", text: "WANOKR-Ga" },
    { key: "WanOptical", text: "WanOptical" },
    { key: "Zayo", text: "Zayo" },
  ];

  const spansComma = useMemo(() => selectedSpans.join(","), [selectedSpans]);

  const formatUtcString = (date: Date | null, time: string) => {
    if (!date) return "";
    // Format selected date + time as MM/DD/YYYY HH:MM (local time)
    const [hh, mm] = time.split(":").map((s) => parseInt(s, 10));
    const y = date.getFullYear();
    const m = (date.getMonth() + 1).toString().padStart(2, "0");
    const d = date.getDate().toString().padStart(2, "0");
    const H = (isNaN(hh) ? 0 : hh).toString().padStart(2, "0");
    const M = (isNaN(mm) ? 0 : mm).toString().padStart(2, "0");
    return `${m}/${d}/${y} ${H}:${M}`;
  };

  const parseTimeToDate = (date: Date | null, time: string | null) => {
    if (!date || !time) return null;
    const [hh, mm] = (time || "00:00").split(":").map((s) => parseInt(s, 10));
    return new Date(date.getFullYear(), date.getMonth(), date.getDate(), isNaN(hh) ? 0 : hh, isNaN(mm) ? 0 : mm);
  };

  const isPastDay = (date: Date | null) => {
    if (!date) return false;
    const today = new Date();
    const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const t = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    return d.getTime() < t.getTime();
  };

  const isWithinDays = (date: Date | null, days: number) => {
    if (!date) return false;
    const now = new Date();
    // normalize to start of day for comparison convenience
    const diff = date.getTime() - now.getTime();
    return diff < days * 24 * 60 * 60 * 1000;
  };

  const startUtc = useMemo(() => formatUtcString(startDate, startTime), [startDate, startTime]);
  const endUtc = useMemo(() => formatUtcString(endDate, endTime), [endDate, endTime]);

  // Prefill subject when entering compose
  useEffect(() => {
    if (!composeOpen) return;
    if (!subject || subject.trim().length === 0) {
      // Prefer FacilityCodeA (user requested) but fall back to rackDC or FacilityCodeZ if needed
      const dc = facilityCodeA || rackDC || facilityCodeZ || "Region";
      const spansPart = (selectedSpans && selectedSpans.length > 0) ? selectedSpans.join(", ") : "<Enter Spans here>";
      setSubject(`Fiber Maintenance scheduled in ${dc} for Spans ${spansPart}`);
    }
  }, [composeOpen, subject, rackDC, facilityCodeA, facilityCodeZ, selectedSpans]);

  const latLongCombined = useMemo(() => (lat && lng ? `${lat},${lng}` : ""), [lat, lng]);

  // Prefill CC from detected user email when compose opens or userEmail changes
  useEffect(() => {
    if (composeOpen && !cc && userEmail) {
      setCc(userEmail);
    }
  }, [composeOpen, userEmail, cc]);

  // Helpers to determine emergency tag state across windows and update subject
  const anyWindowWithin7Days = (primary: Date | null, windows: MaintenanceWindow[]) => {
    if (isWithinDays(primary, 7)) return true;
    for (const w of windows) {
      if (isWithinDays(w.startDate, 7)) return true;
    }
    return false;
  };

  const addEmergencyTag = () => {
    const s = (subject || "").trim();
    if (!/\[EMERGENCY\]/i.test(s)) {
      setSubject(`[EMERGENCY] ${s}`.trim());
    } else if (!/^\s*\[EMERGENCY\]/i.test(s)) {
      // If tag exists but not at the start, move it to the start
      const without = s.replace(/\s*\[EMERGENCY\]\s*/i, "").trim();
      setSubject(`[EMERGENCY] ${without}`.trim());
    }
  };

  const removeEmergencyTag = () => {
    if (/\[EMERGENCY\]/i.test(subject || "")) {
      setSubject(((subject || "") as string).replace(/\s*\[EMERGENCY\]\s*/i, "").trim());
    }
  };

  const emailBody = useMemo(() => {
    const impactStr = impactExpected ? "Yes/True" : "No/False";

    // Build comma-separated start/end lists: primary first, then any additional windows
    const startList: string[] = [];
    const endList: string[] = [];
    if (startUtc) startList.push(startUtc);
    if (endUtc) endList.push(endUtc);
    additionalWindows.forEach((w) => {
      const s = formatUtcString(w.startDate, w.startTime);
      const e = formatUtcString(w.endDate, w.endTime);
      if (s) startList.push(s);
      if (e) endList.push(e);
    });

    // Diversity can be set from the dropdown, but may be undefined or empty
    // Collect unique Diversity values from selected spans in the result table
    let diversityStr = '';
    if (selectedSpans && selectedSpans.length > 0 && result && result.length > 0) {
      const selectedRows = result.filter((r) => selectedSpans.includes(r.SpanID));
      const diversitySet = new Set<string>();
      selectedRows.forEach((row) => {
        if (row.Diversity && typeof row.Diversity === 'string' && row.Diversity.trim()) {
          row.Diversity.split(',').forEach((d) => {
            const val = d.trim();
            if (val) diversitySet.add(val);
          });
        }
      });
      diversityStr = Array.from(diversitySet).join(', ');
    }

    const parts: string[] = [
      `To: ${EMAIL_TO}`,
      `From: Fibervsoassistant@microsoft.com`,
      `CC: ${cc || ""}`,
      `Subject: ${subject}`,
      ``,
      `----------------------------------------`,
      `SentBy: VSO Assistant`,
      `Timezone: Local`,
      `CircuitIds: ${spansComma}`,
      `Diversity: ${diversityStr}`,
      `StartDatetime: ${startList.join(', ')}`,
      `EndDatetime: ${endList.join(', ')}`,
      `NotificationType: ${notificationType}`,
      `MaintenanceReason: ${maintenanceReason}`,
      `Location: ${location}`,
      `Tags: ${tags && tags.length ? tags.join('; ') : ''}`,
      `ImpactExpected: ${impactStr}`,
    ];

    return parts.map((p) => p || "").join("\n");
  }, [EMAIL_TO, subject, spansComma, startUtc, endUtc, notificationType, location, maintenanceReason, tags, impactExpected, additionalWindows, cc, selectedSpans, result]);

  const canSend = useMemo(() => {
    return (
      selectedSpans.length > 0 &&
      !!subject &&
      !!startDate && !!startTime &&
      !!endDate && !!endTime &&
      !!location &&
  // Tags are optional now; do not require ISP / ISP Ticket
      (impactExpected === true || impactExpected === false) &&
      !!maintenanceReason &&
      !!(cc && cc.trim())
    );
  }, [selectedSpans.length, subject, startDate, startTime, endDate, endTime, location, impactExpected, maintenanceReason, cc]);

  // Validate compose fields and return the first invalid field key (or null if valid)
  const validateCompose = (): string | null => {
    const errors: Record<string, string> = {};
    if (!subject || !subject.trim()) errors.subject = "Required";
    if (!startDate) errors.startDate = "Required";
    if (!startTime) errors.startTime = "Required";
    if (!endDate) errors.endDate = "Required";
    if (!endTime) errors.endTime = "Required";
    if (!location?.trim()) errors.location = "Required";
    // Tags are optional; no ISP / ISP Ticket validation
    if (!(impactExpected === true || impactExpected === false)) errors.impactExpected = "Required";
    if (!maintenanceReason?.trim()) errors.maintenanceReason = "Required";
    if (!cc?.trim()) errors.cc = "Required";
    // Additional checks: dates not in past and end after start
    if (startDate && isPastDay(startDate)) {
      errors.startDate = 'Start date cannot be in the past';
    }
    if (endDate && isPastDay(endDate)) {
      errors.endDate = 'End date cannot be in the past';
    }
    // Primary window ordering
    const primaryStart = parseTimeToDate(startDate, startTime);
    const primaryEnd = parseTimeToDate(endDate, endTime);
    if (primaryStart && primaryEnd && primaryEnd.getTime() <= primaryStart.getTime()) {
      errors.endTime = 'End must be after start';
    }
    // Additional windows ordering
    for (let i = 0; i < (additionalWindows || []).length; i++) {
      const w = additionalWindows[i];
      const s = parseTimeToDate(w.startDate, w.startTime);
      const e = parseTimeToDate(w.endDate || w.startDate, w.endTime);
      if (s && e && e.getTime() <= s.getTime()) {
        errors.endTime = 'End must be after start';
        break;
      }
    }

    setFieldErrors(errors);

    // Decide first invalid field in desired order
    const order = [
      'subject',
      'startDate',
      'startTime',
      'endDate',
      'endTime',
      'location',
      'maintenanceReason',
      'cc',
    ];
    for (const k of order) {
      if (errors[k]) return k;
    }
    return null;
  };

  const friendlyFieldNames: Record<string, string> = {
    startDate: "Start Date",
    startTime: "Start Time",
    endDate: "End Date",
    endTime: "End Time",
    location: "Location",
    impactExpected: "Impact Expected",
    maintenanceReason: "Maintenance Reason",
    cc: "CC",
    subject: "Subject",
  };

  const handleSend = async () => {
    // Enable showing validation UI once the user attempts to send
    setShowValidation(true);
    const firstInvalid = validateCompose();
    if (firstInvalid) {
      // Scroll/focus the first invalid field so user can correct it
      const refMap: Record<string, React.RefObject<any>> = {
        subject: subjectRef,
        startDate: startDateRef,
        startTime: startTimeRef,
        endDate: endDateRef,
        endTime: endTimeRef,
        location: locationRef,
        maintenanceReason: maintenanceReasonRef,
        cc: ccRef,
      };
      const r = refMap[firstInvalid];
      try {
        if (r && r.current) {
          // Try to focus the control; many Fluent controls expose focus() via componentRef
          if (typeof r.current.focus === 'function') {
            try { r.current.focus(); } catch (e) { /* ignore */ }
          }
          // Try to select text where possible (TextField exposes inputElement)
          try {
            if (r.current.inputElement && typeof r.current.inputElement.select === 'function') {
              r.current.inputElement.select();
            } else if (r.current.refs && r.current.refs.input && typeof r.current.refs.input.select === 'function') {
              r.current.refs.input.select();
            }
          } catch (e) { /* ignore */ }

          // If the componentRef wraps the native input, try to scroll into view
          if (typeof r.current.scrollIntoView === 'function') {
            r.current.scrollIntoView({ behavior: 'smooth', block: 'center' });
          } else if (r.current.rootElement && typeof r.current.rootElement.scrollIntoView === 'function') {
            r.current.rootElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
          } else if (r.current.domElement && typeof r.current.domElement.scrollIntoView === 'function') {
            r.current.domElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
          } else {
            // fallback: scroll to top
            window.scrollTo({ top: 0, behavior: 'smooth' });
          }
        } else {
          window.scrollTo({ top: 0, behavior: 'smooth' });
        }
      } catch (e) {
        window.scrollTo({ top: 0, behavior: 'smooth' });
      }
      // Show a concise dialog explaining the required field
      try {
        const friendly = friendlyFieldNames[firstInvalid] || firstInvalid;
        setFieldErrorMessageDialog(`${friendly} cannot be empty. Please complete this field.`);
        setShowFieldErrorDialog(true);
      } catch (e) {
        // ignore
      }
      return;
    }
    // Clear validation UI if passed
    setShowValidation(false);
    setSendError(null);
    setSendSuccess(null);
    setSendLoading(true);
    try {
      // Build comma-separated StartDatetime and EndDatetime (primary + any additional)
      const startList: string[] = [];
      const endList: string[] = [];
      if (startUtc) startList.push(startUtc);
      if (endUtc) endList.push(endUtc);
      additionalWindows.forEach((w) => {
        const s = formatUtcString(w.startDate, w.startTime);
        const e = formatUtcString(w.endDate, w.endTime);
        if (s) startList.push(s);
        if (e) endList.push(e);
      });

      const payload = {
        FacilityCodeA: facilityCodeA || "",
        Diversity: (diversity && diversity.length > 0) ? diversity.join(', ') : "",
        SpliceRackA: spliceRackA || "",
        Stage: "9",
        CC: cc || "",
        Subject: subject || "",
        CircuitIds: spansComma || "",
        StartDatetime: startList.join(', '),
        EndDatetime: endList.join(', '),
        LatLong: latLongCombined || "",
        NotificationType: notificationType || "",
        MaintenanceReason: maintenanceReason || "",
        Location: location || "",
        Tags: tags && tags.length ? tags.join('; ') : "",
        ImpactExpected: impactExpected ? "True" : "False",
      };

      const resp = await fetch("/api/LogicAppProxy", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ type: "VSO", ...payload }),
      });

      if (!resp.ok) {
        const text = await resp.text().catch(() => ``);
        throw new Error(`HTTP ${resp.status} ${text}`);
      }

      // Add events to calendar (primary + additional windows) as Draft
      const newEvents: VsoCalendarEvent[] = [];
      const dcCode = rackDC || facilityCodeA;
      const spansShort = (() => {
        if (selectedSpans.length <= 3) return selectedSpans.join(", ");
        return `${selectedSpans.slice(0, 3).join(", ")} (+${selectedSpans.length - 3} more)`;
      })();
      const title = `Fiber Maintenance ${dcCode || ""} - ${spansShort || "Spans"}`.trim();
      const fullReason = (maintenanceReason || "").trim();
      const summary = fullReason.slice(0, 160);

      const makeAllDayRange = (d: Date | null) => {
        if (!d) return null;
        const start = new Date(d.getFullYear(), d.getMonth(), d.getDate());
        const end = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1); // exclusive
        return { start, end };
      };

      // Primary window
      const primary = makeAllDayRange(startDate);
      if (primary) {
        newEvents.push({
          id: `vso-${Date.now()}-0-${Math.random().toString(36).slice(2,6)}`,
          title,
          start: primary.start,
          end: primary.end,
          status: "Draft",
          summary,
          maintenanceReason: fullReason,
          dcCode: dcCode || undefined,
          spans: [...selectedSpans],
          startTimeUtc: startTime,
          endTimeUtc: endTime,
          subject,
          notificationType,
          location,
          tags: tags && tags.length ? tags : [],
          impactExpected,
        } as any);
      }
      // Additional windows
      additionalWindows.forEach((w, i) => {
        const r = makeAllDayRange(w.startDate);
        if (!r) return;
  newEvents.push({
          id: `vso-${Date.now()}-${i + 1}-${Math.random().toString(36).slice(2,6)}`,
          title,
          start: r.start,
          end: r.end,
          status: "Draft",
          summary,
          maintenanceReason: fullReason,
          dcCode: dcCode || undefined,
          spans: [...selectedSpans],
          startTimeUtc: w.startTime,
          endTimeUtc: w.endTime,
          subject,
          notificationType,
          location,
          tags: tags && tags.length ? tags : [],
          impactExpected,
  } as any);
      });
      setVsoEvents((prev) => ensureUnique([...prev, ...newEvents]));
      // Focus calendar on the first window's month so users see it after reload
      if (!calendarDate && (startDate || additionalWindows[0]?.startDate)) {
        const d = startDate || additionalWindows[0]?.startDate || null;
        if (d) setCalendarDate(new Date(d.getFullYear(), d.getMonth(), 1));
      }

      setSendSuccess("Request submitted to Logic App successfully.");
      setShowSendSuccessDialog(true);
    } catch (e: any) {
      setSendError(e?.message || "Failed to send email.");
    } finally {
      setSendLoading(false);
    }
  };

  // Map status text to visual pill styles defined in Theme.css (.status-label.*)
  const getStatusClass = (status?: string) => {
    const t = (status || "").toLowerCase();
    if (t.includes("inproduction") || t === "in production" || t === "production") return "good";
    if (t.includes('in progress') || t.includes('inprogress') || t.includes('wf in progress') || t.includes('wf inprogress')) return 'warning';
    if (
      t.includes("decom") ||
      t.includes("retired") ||
      t.includes("outofservice") ||
      t.includes("out of service") ||
      t.includes("warning")
    )
      return "warning";
    return "accent";
  };

  // Map diversity text to color-coded pill
  const getDiversityClass = (div?: string) => {
    const t = (div || "").toLowerCase().trim();
    if (t.includes("east 1")) return "accent"; // blue
    if (t.includes("east 2")) return "good"; // green
    if (t === "south") return "accent"; // blue
    if (t === "y") return "accent"; // blue
    if (t.includes("west 1")) return "danger"; // red
    if (t.includes("west 2")) return "warning"; // yellow
    if (t === "north") return "danger"; // red
    if (t === "z") return "danger"; // red
    if (t.startsWith("east")) return "accent";
    if (t.startsWith("west")) return "danger";
    return "accent";
  };

  // Map SpanType text to visual pill styles: WAN=red, Fabric=blue, others=yellow
  const getSpanTypeClass = (t?: string) => {
    const v = (t || "").toString().toLowerCase().trim();
    if (!v) return 'warning';
    if (v.includes('wan')) return 'danger';
    if (v.includes('fabric')) return 'accent';
    return 'warning';
  };

  // Helper to compute display text for datacenter combo boxes so typed/selected
  // values persist across tabs. Prefer the pending typed text, then the
  // combo's own selected key, then fall back to the opposite side's text.
  const dcTextFor = (key?: string) => {
    if (!key) return '';
    const opt = datacenterOptions.find(d => String(d.key) === String(key));
    return opt?.text?.toString() || String(key) || '';
  };

  const azText = dcSearch || (facilityCodeA ? dcTextFor(facilityCodeA) : (facilityCodeZ ? dcTextFor(facilityCodeZ) : ''));
  const facilityText = dcSearch || (facilityCodeA ? dcTextFor(facilityCodeA) : (facilityCodeZ ? dcTextFor(facilityCodeZ) : ''));
  const zaText = dcSearchZ || (facilityCodeZ ? dcTextFor(facilityCodeZ) : (facilityCodeA ? dcTextFor(facilityCodeA) : ''));

  return (
    <div className="main-content fade-in">
  {/* Increased container width (~30%) to allow more horizontal room for table/content */}
  <div className="vso-form-container glow" style={{ width: "92%", maxWidth: 1300 }}>
        <div className="banner-title">
          <span className="title-text">Fiber VSO Assistant</span>
          <span className="title-sub">Simplifying Span Lookup and VSO Creation.</span>
        </div>

        {!composeOpen && (
          <div key={formKey}>
        {/* Tabbed layout: A-Z, Facility, Z-A. Each tab shows the corresponding inputs. */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', padding: 10, borderRadius: 12, background: 'linear-gradient(180deg, rgba(0,0,0,0.28), rgba(3,12,18,0.18))', boxShadow: '0 8px 30px rgba(0,96,140,0.06), inset 0 1px 0 rgba(255,255,255,0.02)', border: '1px solid rgba(255,255,255,0.03)' }}>
            {/* Tab buttons */}
            {(['A-Z', 'Facility', 'Z-A', 'Decom'] as const).map((t) => (
              <button
                key={t}
                onClick={() => setCurrentTab(t)}
                aria-pressed={currentTab === t}
                className={`tab-btn ${currentTab === t ? 'active' : ''}`}
                style={{
                  padding: '10px 20px',
                  borderRadius: 10,
                  border: 'none',
                  background: currentTab === t ? 'linear-gradient(180deg,#06435a,#003b6f)' : 'linear-gradient(180deg, rgba(255,255,255,0.01), rgba(255,255,255,0.005))',
                  color: currentTab === t ? '#e6fbff' : '#9fb3c6',
                  cursor: 'pointer',
                  fontWeight: 700,
                  boxShadow: currentTab === t ? '0 6px 18px rgba(0,120,200,0.18), 0 1px 0 rgba(255,255,255,0.03) inset' : 'none',
                }}
              >
                <span style={{ display: 'inline-block', minWidth: 130, textAlign: 'center' }}>
                  {t === 'A-Z' ? 'A → Z' : t === 'Z-A' ? 'Z → A' : t === 'Decom' ? 'Decommissioned' : 'Facility'}
                </span>
              </button>
            ))}
          </div>

          {/* Tab contents */}
          <div style={{ display: 'flex', gap: 12, alignItems: 'flex-start' }}>
            {/* A-Z tab: show Facility A input */}
            {currentTab === 'A-Z' && (
              <div style={{ flex: 1 }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <Text styles={labelStyles(13, 700)}>
                    Facility Code A <span style={{ color: "#ff4d4d" }}>*</span>
                  </Text>
                  <TooltipHost content={"This search will provide results from the FacilityCodeA side using the Datacenter code entered."}>
                    <IconButton iconProps={{ iconName: 'Info' }} title="Facility selection info" styles={{ root: { color: '#a6b7c6', height: 20, width: 20 } }} />
                  </TooltipHost>
                </div>
                <ComboBox
                  placeholder="Type or select Facility Code A"
                  options={filteredDcOptions}
                  selectedKey={facilityCodeA || undefined}
                  text={azText}
                  allowFreeform={true}
                  autoComplete="on"
                  useComboBoxAsMenuWidth
                  calloutProps={{ className: 'combo-dark-callout' }}
                  componentRef={dcComboRef}
                  styles={comboBoxStyles}
                  onRenderOption={isLightTheme ? renderLightOption : undefined}
                  onChange={(_, option, index, value) => {
                    const typed = (value || "").toString().toLowerCase();
                    const found = datacenterOptions.find((d) => {
                      const keyStr = d.key?.toString().toLowerCase();
                      const textStr = d.text?.toString().toLowerCase();
                      return textStr === typed || keyStr === typed;
                    });
                    if (option) {
                      const selectedKey = option.key?.toString() ?? "";
                      if (selectedKey === "") { setFacilityCodeA(""); return; }
                      if (selectedKey === facilityCodeA) { setFacilityCodeA(""); setDcSearch(""); }
                      else { setFacilityCodeA(selectedKey); setDcSearch(option.text?.toString() || selectedKey); }
                    } else if (found) { setFacilityCodeA(found.key.toString()); }
                    else setFacilityCodeA("");
                  }}
                  onPendingValueChanged={(option, index, value) => setDcSearch(value || "")}
                  onMenuDismiss={() => setDcSearch("")}
                />

                {/* Diversity for A tab */}
                <div style={{ marginTop: 8 }}>
                  <Text styles={labelStyles(13, 700, 6)}>Diversity <span className="optional-text">(Optional)</span></Text>
                  <Dropdown
                    placeholder=""
                    options={diversityOptions}
                    calloutProps={{ className: 'combo-dark-callout' }}
                    styles={diversityDropdownStyles}
                    selectedKeys={diversity}
                    multiSelect
                    onChange={(_, option) => {
                      if (!option) return;
                      const key = option.key?.toString() ?? "";
                      if (key === "") { setDiversity([]); return; }
                      setDiversity(prev => {
                        if (option.selected) return [...prev, key];
                        return prev.filter(k => k !== key);
                      });
                    }}
                  />
                </div>

                <div style={{ marginTop: 8 }}>
                  <Text styles={labelStyles(13, 700)}>Splice Rack A <span className="optional-text">(Optional)</span></Text>
                  <TextField
                    placeholder="e.g. AM111"
                    onChange={(_, value) => { const v = value || undefined; setSpliceRackA(v); }}
                    styles={textFieldStyles}
                    disabled={!!spliceRackZ}
                  />
                </div>
              </div>
            )}

            {/* Facility tab: single Facility search (Facility Code, Diversity, Splice Rack) */}
            {currentTab === 'Facility' && (
              <div style={{ flex: 1 }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <Text styles={labelStyles(13, 700)}>
                    Facility Code <span style={{ color: "#ff4d4d" }}>*</span>
                  </Text>
                  <TooltipHost content={"This search will provide results from both the FacilityCodeA and FacilityCodeZ side using the Datacenter code entered."}>
                    <IconButton iconProps={{ iconName: 'Info' }} title="Facility selection info" styles={{ root: { color: '#a6b7c6', height: 20, width: 20 } }} />
                  </TooltipHost>
                </div>
                <ComboBox
                  placeholder="Type or select Facility Code"
                  options={filteredDcOptions}
                  selectedKey={facilityCodeA || undefined}
                  text={facilityText}
                  allowFreeform={true}
                  autoComplete="on"
                  useComboBoxAsMenuWidth
                  calloutProps={{ className: 'combo-dark-callout' }}
                  componentRef={dcComboRef}
                  styles={comboBoxStyles}
                  onRenderOption={isLightTheme ? renderLightOption : undefined}
                  onChange={(_, option, index, value) => {
                    const typed = (value || "").toString().toLowerCase();
                    const found = datacenterOptions.find((d) => {
                      const keyStr = d.key?.toString().toLowerCase();
                      const textStr = d.text?.toString().toLowerCase();
                      return textStr === typed || keyStr === typed;
                    });
                    if (option) {
                      const selectedKey = option.key?.toString() ?? "";
                      if (selectedKey === "") { setFacilityCodeA(""); return; }
                      if (selectedKey === facilityCodeA) { setFacilityCodeA(""); setDcSearch(""); }
                      else { setFacilityCodeA(selectedKey); setDcSearch(option.text?.toString() || selectedKey); }
                    } else if (found) { setFacilityCodeA(found.key.toString()); }
                    else setFacilityCodeA("");
                  }}
                  onPendingValueChanged={(option, index, value) => setDcSearch(value || "")}
                  onMenuDismiss={() => setDcSearch("")}
                />

                <div style={{ display: 'flex', gap: 12, marginTop: 12 }}>
                  <div style={{ flex: 1 }}>
                      <Text styles={labelStyles(13, 700, 6)}>Diversity <span className="optional-text">(Optional)</span></Text>
                    <Dropdown
                      placeholder=""
                      options={diversityOptions}
                      calloutProps={{ className: 'combo-dark-callout' }}
                      styles={diversityDropdownStyles}
                      selectedKeys={diversity}
                      multiSelect
                      onChange={(_, option) => {
                        if (!option) return;
                        const key = option.key?.toString() ?? "";
                        if (key === "") { setDiversity([]); return; }
                        setDiversity(prev => {
                          if (option.selected) return [...prev, key];
                          return prev.filter(k => k !== key);
                        });
                      }}
                    />
                  </div>
                </div>
                <div style={{ display: 'flex', gap: 12, marginTop: 12 }}>
                  <div style={{ flex: 1 }}>
                    <Text styles={labelStyles(13, 700, 6)}>Splice Rack <span className="optional-text">(Optional)</span></Text>
                    <TextField placeholder="e.g. AM111" value={spliceRackA || ''} onChange={(_, v) => { const val = v || undefined; setSpliceRackA(val); }} styles={textFieldStyles} />
                  </div>
                </div>
              </div>
            )}

            {/* Z-A tab: show Facility Z input */}
            {currentTab === 'Z-A' && (
              <div style={{ flex: 1 }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <Text styles={labelStyles(13, 700)}>Facility Code Z <span style={{ color: "#ff4d4d" }}>*</span></Text>
                  <TooltipHost content={"This search will provide results from the FacilityCodeZ side using the Datacenter code entered."}>
                    <IconButton iconProps={{ iconName: 'Info' }} title="Facility selection info" styles={{ root: { color: '#a6b7c6', height: 20, width: 20 } }} />
                  </TooltipHost>
                </div>
                <ComboBox
                  placeholder="Type or select Facility Code Z"
                  options={filteredDcOptionsZ}
                  selectedKey={facilityCodeZ || undefined}
                  text={zaText}
                  allowFreeform={true}
                  autoComplete="on"
                  useComboBoxAsMenuWidth
                  calloutProps={{ className: 'combo-dark-callout' }}
                  componentRef={dcComboRefZ}
                  styles={comboBoxStyles}
                  onRenderOption={isLightTheme ? renderLightOption : undefined}
                  onChange={(_, option, index, value) => {
                    const typed = (value || "").toString().toLowerCase();
                    const found = datacenterOptions.find((d) => {
                      const keyStr = d.key?.toString().toLowerCase();
                      const textStr = d.text?.toString().toLowerCase();
                      return textStr === typed || keyStr === typed;
                    });
                    if (option) {
                      const selectedKey = option.key?.toString() ?? "";
                      if (selectedKey === "") { setFacilityCodeZ(""); return; }
                      if (selectedKey === facilityCodeZ) { setFacilityCodeZ(""); setDcSearchZ(""); }
                      else { setFacilityCodeZ(selectedKey); setDcSearchZ(option.text?.toString() || selectedKey); }
                    } else if (found) { setFacilityCodeZ(found.key.toString()); }
                    else setFacilityCodeZ("");
                  }}
                  onPendingValueChanged={(option, index, value) => setDcSearchZ(value || "")}
                  onMenuDismiss={() => setDcSearchZ("")}
                />

                {/* Diversity for Z tab */}
                <div style={{ marginTop: 8 }}>
                  <Text styles={labelStyles(13, 700, 6)}>Diversity <span className="optional-text">(Optional)</span></Text>
                  <Dropdown
                    placeholder=""
                    options={diversityOptions}
                    calloutProps={{ className: 'combo-dark-callout' }}
                    styles={diversityDropdownStyles}
                    selectedKeys={diversity}
                    multiSelect
                    onChange={(_, option) => {
                      if (!option) return;
                      const key = option.key?.toString() ?? "";
                      if (key === "") { setDiversity([]); return; }
                      setDiversity(prev => {
                        if (option.selected) return [...prev, key];
                        return prev.filter(k => k !== key);
                      });
                    }}
                  />
                </div>

                <div style={{ marginTop: 8 }}>
                  <Text styles={labelStyles(13, 700)}>Splice Rack Z <span className="optional-text">(Optional)</span></Text>
                  <TextField placeholder="e.g. AJ1508" onChange={(_, value) => { const v = value || undefined; setSpliceRackZ(v); }} styles={textFieldStyles} disabled={!!spliceRackA} />
                </div>
              </div>
            )}

            {/* Decommissioned tab: country lookup moved here (was a collapsed panel previously) */}
            {currentTab === 'Decom' && (
              <div style={{ flex: 1 }}>
                <div className="decom-card" style={{ background: '#081518', border: '1px solid #123238', padding: 16, borderRadius: 10, boxShadow: '0 6px 18px rgba(0,80,100,0.12)', maxWidth: 680 }}>
                  <div className="decom-card-header" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 10 }}>
                    <div>
                        <div style={{ fontSize: 16, fontWeight: 700 }}>Decommissioned Spans</div>
                      </div>
                    <IconButton iconProps={{ iconName: 'WarningSolid' }} title="Decommissioned search info" styles={{ root: { color: '#f1c232', height: 28, width: 28 } }} />
                  </div>

                  <div className="decom-card-inner" style={{ background: '#071821', border: '1px solid #20343f', padding: 12, borderRadius: 8, marginBottom: 12 }}>
                    <div style={{ fontSize: 13, marginBottom: 6 }}>This search will pull all spans that are in a Decommission ready state by the country entered.</div>
                    <div style={{ fontSize: 12 }}>Select a country, then click Lookup to fetch spans marked ready for decommission in that country.</div>
                  </div>

                  <div style={{ display: 'flex', gap: 12, alignItems: 'center' }}>
                    <div style={{ flex: 1 }}>
                      <ComboBox
                        placeholder="Type or select a country"
                        options={countryOptions}
                        selectedKey={country || undefined}
                        allowFreeform={true}
                        autoComplete="on"
                        useComboBoxAsMenuWidth
                        calloutProps={{ className: 'combo-dark-callout' }}
                        componentRef={countryComboRef}
                        onClick={() => countryComboRef.current?.focus(true)}
                        onFocus={() => countryComboRef.current?.focus(true)}
                        styles={decomCountryComboStyles}
                        onRenderOption={isLightTheme ? renderLightOption : undefined}
                        onChange={(_, option, index, value) => {
                          const typed = (value || "").toString().toLowerCase();
                          const found = countryOptions.find((c) => {
                            const keyStr = (c.key || "").toString().toLowerCase();
                            const textStr = (c.text || "").toString().toLowerCase();
                            return textStr === typed || keyStr === typed;
                          });

                          if (option) {
                            const selectedKey = option.key?.toString() ?? "";
                            if (selectedKey === "") { setCountry(""); return; }
                            if (selectedKey === country) {
                              setCountry("");
                            } else {
                              setCountry(selectedKey);
                            }
                          } else if (found) {
                            setCountry(found.key.toString());
                          } else {
                            // reset if invalid typed text
                            setCountry("");
                          }
                        }}
                        onPendingValueChanged={(option, index, value) => {
                          setCountrySearch(value || "");
                        }}
                        onMenuDismiss={() => setCountrySearch("")}
                      />
                    </div>
                    <div style={{ width: 120, display: 'flex', alignItems: 'center' }}>
                      <PrimaryButton
                        className="lookup-btn"
                        text="Lookup"
                        disabled={!country}
                        styles={{ root: { height: 40 } }}
                        onClick={async () => {
                          setLoading(true);
                          setError(null);
                          setResult([]);
                          setIsDecomMode(true);
                          setShowAll(true); // show all decom'd spans
                          try {
                            const payload = { Stage: "10", Country: country };
                            const resp = await fetch("/api/LogicAppProxy", { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ type: "VSO", ...payload }) });
                            if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
                            const data: LogicAppResponse = await resp.json();
                            const spans = Array.isArray(data?.Spans) ? data.Spans : (data?.Spans ? [data.Spans] : []);
                            setResult(spans as any);
                            if (data?.RackElevationUrl) setRackUrl(data.RackElevationUrl);
                            if (data?.DataCenter) setRackDC(data.DataCenter);
                          } catch (e: any) {
                            setError(String(e?.message || e));
                          } finally {
                            setLoading(false);
                            setSearchDone(true);
                          }
                        }}
                      />
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Facility-specific inputs moved into the Facility tab above. */}

        {currentTab !== 'Decom' && (
          <div className="form-buttons" style={{ marginTop: 16 }}>
            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <div style={{ display: 'flex', gap: 8 }}>
                <button className="submit-btn" onClick={() => handleSubmit()}>
                  Submit
                </button>
                {searchDone && (
                  <button
                    className="sleek-btn danger"
                    onClick={() => { resetAll(); setIsDecomMode(false); }}
                    title="Reset search and form"
                    aria-label="Reset"
                    style={{ minWidth: 96 }}
                  >
                    Reset
                  </button>
                )}
              </div>
              {/* Decommissioned tab omits Submit/Reset (only Lookup button inside tab). */}
            </div>
          </div>
        )}

        {loading && <Spinner label="Loading results..." size={SpinnerSize.medium} styles={{ root: { marginTop: 15 } }} />}

        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
            {error}
          </MessageBar>
        )}

        {!loading && !error && searchDone && result.length === 0 && (
          <div>
            {oppositePrompt.show ? (
              <MessageBar
                messageBarType={MessageBarType.info}
                isMultiline={false}
                styles={{
                  root: {
                    marginTop: 8,
                    background: '#141414',
                    color: '#dfefff',
                    border: '1px solid #333',
                    borderRadius: 8,
                    padding: 12,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                  },
                  content: { display: 'flex', alignItems: 'center', gap: 12 },
                }}
              >
                <div style={{ display: 'flex', flexDirection: 'column', flex: 1, marginRight: 16 }}>
                  <div style={{ marginBottom: 14 }}>
                    <span style={{ fontWeight: 600, whiteSpace: 'normal' }}>{`No results were found for that search. Would you like to try searching from Splice Rack ${oppositePrompt.from === 'A' ? 'Z' : 'A'} instead?`}</span>
                  </div>
                  <div style={{ display: 'flex', gap: 8 }}>
                    <PrimaryButton text={`Yes`} onClick={() => handleOppositeSearch(oppositePrompt.from || 'A')} />
                    <DefaultButton text="No" onClick={() => setOppositePrompt({ show: false, from: null })} />
                  </div>
                </div>
              </MessageBar>
            ) : (
              <MessageBar
                messageBarType={MessageBarType.info}
                isMultiline={false}
                styles={{ root: { marginTop: 8, background: '#141414', color: '#dfefff', border: '1px solid #333', borderRadius: 8, padding: 12 } }}
              >
                {triedBothNoResults
                  ? 'No results were found searching both Splice Rack A and Z. Please adjust your search criteria.'
                  : 'There were no results for the selections you made. Try adjusting your search.'}
              </MessageBar>
            )}
          </div>
        )}

        {result.length > 0 && (
          // Prevent horizontal scrollbar by hiding overflow and forcing tighter column widths
          <div className="table-container" style={{ marginTop: 14, overflowX: 'hidden' }}>
            <div style={{ display: 'flex', alignItems: 'center', marginBottom: 8 }}>
              {/* Left: spans summary and info icon */}
              <div style={{ flex: 1, display: 'flex', justifyContent: 'flex-start', alignItems: 'center', gap: 10 }}>
                {(() => {
                  const totalSpans = result.length;
                  const productionCount = result.filter(r => {
                    const state = (((r as any).State || '') as string).toString().toLowerCase();
                    if (state === 'new') return false;
                    return ((r.Status || '') as string).toString().toLowerCase() === 'inproduction';
                  }).length;
                  return (
                    <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                      <div style={{ background: '#071821', border: '1px solid #20343f', padding: '6px 10px', borderRadius: 8, display: 'flex', gap: 8, alignItems: 'center', boxShadow: '0 4px 14px rgba(0,0,0,0.4)' }}>
                        <div style={{ textAlign: 'right' }}>
                          <div style={{ fontSize: 18, fontWeight: 700, color: '#dfefff' }}>{totalSpans}</div>
                          <div style={{ fontSize: 11, color: '#9fb3c6' }}>Total Spans</div>
                        </div>
                        <div style={{ textAlign: 'right' }}>
                          <div style={{ fontSize: 18, fontWeight: 700, color: '#8fe3a3' }}>{productionCount}</div>
                          <div style={{ fontSize: 11, color: '#9fb3c6' }}>In Production</div>
                        </div>
                      </div>
                      {/* Info icon for drag selection */}
                      <TooltipHost content="Tip: You can select multiple spans at once by clicking and dragging your mouse over the rows.">
                        <IconButton iconProps={{ iconName: 'Info' }} title="Multi-select info" styles={{ root: { color: '#a6b7c6', height: 22, width: 22, marginLeft: 2 } }} />
                      </TooltipHost>
                    </div>
                  );
                })()}
              </div>

              {/* Center: slim rack elevation dropdown */}
              <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', width: 220 }}>
                  {(() => {
                  // Build leading options so the rack elevations dropdown prioritizes
                  // any FacilityCodeA/FacilityCodeZ values that appear in the results table
                  // (i.e., values from the row columns), then fall back to rackDC and explicit inputs.
                  const leadingCodes: string[] = [];
                  // Collect facility codes from result rows (FacilityCodeA / FacilityCodeZ)
                  for (const r of result || []) {
                    try {
                      const a = (r as any).FacilityCodeA;
                      const z = (r as any).FacilityCodeZ;
                      if (a && String(a).trim()) leadingCodes.push(String(a).trim());
                      if (z && String(z).trim()) leadingCodes.push(String(z).trim());
                    } catch {}
                  }
                  // Also include rackDC and any explicit inputs as fallback (preserve ordering)
                  if (rackDC && String(rackDC).trim()) leadingCodes.push(String(rackDC).trim());
                  if (facilityCodeA && String(facilityCodeA).trim()) leadingCodes.push(String(facilityCodeA).trim());
                  if (facilityCodeZ && String(facilityCodeZ).trim()) leadingCodes.push(String(facilityCodeZ).trim());

                  // Deduplicate while preserving order
                  const seen = new Set<string>();
                  const leadingUnique = leadingCodes.filter((c) => {
                    if (!c) return false;
                    if (seen.has(c)) return false;
                    seen.add(c);
                    return true;
                  });

                  // Create options: put leadingUnique first (sorted ascending), then any availableDcOptions not already included (also sorted)
                  // Filter out any 'unknown' keys or labels
                  const isKnown = (v: any) => {
                    if (!v) return false;
                    const s = String(v).toLowerCase();
                    return s !== 'unknown';
                  };
                  const leadingUniqueSorted = [...leadingUnique].filter(isKnown).sort((a, b) => a.localeCompare(b));
                  const leadingOptions = leadingUniqueSorted.map((c) => ({ key: c, text: c }));
                  const remaining = availableDcOptions
                    .filter((o) => !seen.has(String(o.key)) && isKnown(o.key) && isKnown(o.text))
                    .sort((a, b) => String(a.key).localeCompare(String(b.key)));
                  const options = leadingOptions.length ? [...leadingOptions, ...remaining] : remaining;

                  const headerDropdownStyles = {
                    ...dropdownStyles,
                    root: { width: 180 },
                    title: { ...dropdownStyles.title, background: '#003b6f', color: '#fff', height: 32, borderRadius: 6, fontSize: 13 },
                    dropdownItem: { background: 'transparent', color: '#fff' },
                    dropdownItemSelected: { background: '#004b8a', color: '#fff' },
                    callout: { background: '#181818' },
                  } as const;

                  return (
                    <div style={{ minWidth: 160 }}>
                      <Dropdown
                        placeholder="GNS Rack Elevations"
                        options={options}
                        styles={headerDropdownStyles}
                        onChange={(_, opt) => {
                          if (!opt) return;
                          const key = opt.key?.toString();
                          if (!key) return;
                          const url = getRackElevationUrl(key);
                          if (url) window.open(url, '_blank');
                          else alert(`No rack elevation URL available for ${opt.text}`);
                        }}
                        disabled={options.length === 0}
                      />
                    </div>
                  );
                })()}
              </div>

              {/* Right: controls */}
              <div
                className="span-header-actions"
                style={{ flex: 1, display: 'flex', justifyContent: 'flex-end', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}
              >
                <button
                  type="button"
                  className="rack-btn slim"
                  onClick={handleExportSpansToCsv}
                  title="Export to Excel"
                >
                  Export Spans
                </button>

                {!isDecomMode && (
                  <button className="rack-btn slim" onClick={() => setShowAll(!showAll)}>
                    {showAll ? "Show Only Production" : "Show All Spans"}
                  </button>
                )}

                <button
                  type="button"
                  className={`rack-btn slim pill-toggle ${!simplifiedView ? 'pill-toggle-on' : 'pill-toggle-off'}`}
                  onClick={() => setSimplifiedView(!simplifiedView)}
                >
                  <span className="pill-toggle-label">Detailed view</span>
                  <span className="pill-toggle-knob" aria-hidden="true" />
                </button>
              </div>
            </div>

            {(() => {
              const hasValue = (key: string) => filteredResultsBase.some((r: SpanData) => {
                const v = (r as any)[key];
                if (v === undefined || v === null) return false;
                const s = String(v).trim();
                return s.length > 0;
              });
              const candidate = [
                { key: 'Diversity', label: 'Diversity', render: (row: SpanData) => (
                  <span
                    className={`status-label ${getDiversityClass(row.Diversity)}`}
                    style={{ display: 'inline-block', padding: '1px 6px', whiteSpace: 'nowrap', marginRight: 6 }}
                    title={row.Diversity}
                  >
                    {row.Diversity}
                  </span>
                ) },
                { key: 'SpanID', label: 'SpanID', render: (row: SpanData) => (
                  <span
                    role="button"
                    tabIndex={0}
                    className="uid-click"
                    onClick={(e) => {
                      e.stopPropagation();
                      try {
                        const q = encodeURIComponent(String(row.SpanID || ''));
                        const url = `${window.location.origin}/fiber-span-utilization?spans=${q}`;
                        window.open(url, '_blank');
                      } catch (err) {
                        // fallback: open route without params in new tab
                        try { window.open(`${window.location.origin}/fiber-span-utilization`, '_blank'); } catch { /* ignore */ }
                      }
                    }}
                    onKeyDown={(e: React.KeyboardEvent) => {
                      if (e.key === 'Enter' || e.key === ' ') {
                        e.preventDefault();
                        e.stopPropagation();
                        try {
                          const q = encodeURIComponent(String(row.SpanID || ''));
                          const url = `${window.location.origin}/fiber-span-utilization?spans=${q}`;
                          window.open(url, '_blank');
                        } catch (err) {
                          try { window.open(`${window.location.origin}/fiber-span-utilization`, '_blank'); } catch { /* ignore */ }
                        }
                      }
                    }}
                    style={{ fontSize: 14, fontWeight: 600, cursor: 'pointer', outline: 'none' }}
                  >
                    {row.SpanID}
                  </span>
                ) },
                { key: 'Datacenter', label: 'DC', render: (row: SpanData) => {
                  // Prefer explicit Datacenter/ DataCenter on the row, fallback to rackDC or facility codes
                  const v = (row as any).Datacenter || (row as any).DataCenter || rackDC || (row as any).FacilityCodeA || (row as any).FacilityCodeZ || '';
                  return String(v || '');
                } },
                { key: 'FacilityCodeA', label: 'Facility A' },
                { key: 'FacilityCodeZ', label: 'Facility Z' },
                { key: 'IDF_A', label: 'IDF A' },
                { key: 'SpliceRackA', label: 'Splice A' },
                { key: 'SpliceRackA_Unit', label: 'Splice Rack A' },
                { key: 'SpliceRackZ_Unit', label: 'Splice Rack Z' },
                { key: 'OpticalDeviceA', label: 'Optical A' },
                { key: 'OpticalRackA_Unit', label: 'Rack A' },
                { key: 'OpticalDeviceZ', label: 'Optical Z' },
                { key: 'OpticalRackZ_Unit', label: 'Rack Z' },
                { key: 'WiringScope', label: 'Scope' },
                { key: 'SpanType', label: 'Type', render: (row: SpanData) => (
                  <span
                    className={`status-label ${getSpanTypeClass((row as any).SpanType)}`}
                    style={{ display: 'inline-block', padding: '1px 6px', whiteSpace: 'nowrap', marginRight: 6 }}
                    title={(row as any).SpanType}
                  >
                    {(row as any).SpanType}
                  </span>
                ) },
                { key: 'Status', label: 'Status', render: (row: SpanData) => {
                  const stateVal = (((row as any).State || '') as string).toString().toLowerCase();
                  const isNewState = stateVal === 'new';
                  const displayStatus = isNewState ? 'New' : (row.Status || '');
                  return (
                    <span
                      className={`status-label ${getStatusClass(displayStatus)}`}
                      style={{ display: 'inline-block', padding: '1px 6px', whiteSpace: 'nowrap' }}
                      title={displayStatus}
                    >
                      {displayStatus}
                    </span>
                  );
                } },
              ];
              // Determine which columns to show depending on simplifiedView
              // For Facility tab, always show both Splice Rack (A) and Splice Rack Z columns
              const getCandidate = (k: string) => candidate.find(c => c.key === k);

              let dynamicCols: any[];
                  if (currentTab === 'Facility') {
                if (simplifiedView) {
                  // Facility tab simplified view: fixed subset
                  const facilityKeys = ['Diversity', 'SpanID', 'FacilityCodeA', 'FacilityCodeZ', 'SpliceRackA_Unit', 'SpliceRackZ_Unit', 'WiringScope', 'SpanType', 'Status'];
                  dynamicCols = facilityKeys.map((k) => {
                    const found = getCandidate(k);
                    if (found) return found;
                    return { key: k, label: k, render: (row: SpanData) => ((row as any)[k] ?? '') };
                  });
                } else {
                  // Facility tab detailed view: show all columns with data, like other tabs
                  dynamicCols = candidate.filter(c => c.key === 'SpanID' || hasValue(c.key));
                }
              } else if (simplifiedView) {
                // For Z-A tab, show SpliceRackZ_Unit as the default splice column
                let splicePreferredKey;
                if (currentTab === 'Z-A') {
                  splicePreferredKey = 'SpliceRackZ_Unit';
                } else if (facilityCodeA) {
                  splicePreferredKey = 'SpliceRackA_Unit';
                } else if (spliceRackA) {
                  splicePreferredKey = 'SpliceRackA';
                } else {
                  splicePreferredKey = 'SpliceRackZ';
                }
                const simplifiedKeys = ['Diversity', 'SpanID', 'FacilityCodeA', 'FacilityCodeZ', splicePreferredKey, 'WiringScope', 'SpanType', 'Status'];
                dynamicCols = simplifiedKeys.map((k) => {
                  const found = getCandidate(k);
                  if (found) return found;
                  // Provide a fallback for SpliceRackZ which might be returned as SpliceRackZ or SpliceRackZ_Unit
                  if (k === 'SpliceRackZ') {
                    return {
                      key: 'SpliceRackZ',
                      label: 'Splice Z',
                      render: (row: SpanData) => ((row as any).SpliceRackZ || (row as any).SpliceRackZ_Unit || ''),
                    };
                  }
                  return { key: k, label: k, render: (row: SpanData) => ((row as any)[k] ?? '') };
                });
              } else {
                dynamicCols = candidate.filter(c => c.key === 'SpanID' || hasValue(c.key));
              }
              // Compute a fair percent width per column for the detailed view so columns fit evenly
              // Use more aggressive compression: allow as small as 3% per column when many columns present
              const colWidthPercent = dynamicCols && dynamicCols.length ? Math.max(3, Math.floor(100 / dynamicCols.length)) : 12;

              return (
                <table
                  ref={tableRef}
                  className="data-table compact"
                  style={{ fontSize: 11, tableLayout: 'fixed', width: '100%', borderCollapse: 'collapse', marginLeft: -6 }}
                  onMouseLeave={handleTableMouseLeave}
                >
                  <thead>
                    <tr>
                      <th style={{ width: 30, padding: '2px 4px' }}></th>
                      {dynamicCols.map(c => (
                        <th
                                key={c.key}
                                onClick={() => handleSort(c.key)}
                                style={{
                                // reduce horizontal padding so content shifts left and leaves more room to the right
                                padding: '1px 6px',
                            cursor: 'pointer',
                            whiteSpace: 'nowrap',
                            fontWeight: 500,
                            // Use even percentage widths for detailed view so columns fit; keep small fixed widths for certain keys
                            // Give Diversity a bit more space so the colored pill doesn't get clipped.
                                width: (c.key === 'SpanID') ? '8ch'
                                  : (c.key === 'Datacenter') ? '4ch'
                                  : (c.key === 'Diversity') ? '7%'
                                  : (c.key === 'FacilityCodeA') ? '5ch'
                                  : (c.key === 'FacilityCodeZ') ? '5ch'
                                  : (c.key === 'IDF_A') ? '5ch'
                                  : (c.key === 'SpliceRackZ' || c.key === 'SpliceRackZ_Unit') ? '5ch'
                                  : (c.key === 'WiringScope') ? '6ch'
                                  : (c.key === 'SpanType') ? '5ch'
                                  : (c.key === 'Status') ? '12%'
                                  : `${colWidthPercent}%`,
                            ...(c.key === 'Diversity' ? { paddingLeft: 10 } : {}),
                            ...(c.key === 'SpanID' ? { textAlign: 'center' as const } : {})
                          }}
                        >
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                            {c.label} {sortBy === c.key && (sortDir === 'asc' ? '▲' : '▼')}
                          </span>
                          {/* resize handle: narrow draggable area at the right edge of the header */}
                          <div
                            className="col-resizer"
                            onMouseDown={(e) => startColumnResize(e, e.currentTarget as HTMLElement)}
                            title="Drag to resize"
                          />
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {sortedResults.map((row, i) => {
                      const isSelected = selectedSpans.includes(row.SpanID);
                      const bg = isSelected ? '#10324a' : (i % 2 === 0 ? '#0f0f0f' : '#121212');
                      return (
                        <tr
                          key={i}
                          className={isSelected ? 'highlight-row' : ''}
                          style={{ cursor: 'pointer', background: bg }}
                          onMouseDown={handleRowMouseDown(i, row.SpanID)}
                          onMouseEnter={handleRowMouseEnter(i, row.SpanID)}
                          onMouseUp={handleRowMouseUp}
                        >
                          <td style={{ padding: '2px 4px', width: 38, display: 'flex', alignItems: 'flex-end', justifyContent: 'center' }}>
                            <div
                              onMouseDown={(e: React.MouseEvent<HTMLDivElement>) => { e.stopPropagation(); e.preventDefault(); }}
                              onClick={e => e.stopPropagation()}
                              style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', width: '100%' }}
                            >
                              <Checkbox
                                checked={isSelected}
                                onChange={(ev, checked) => {
                                  if (ev) ev.stopPropagation();
                                  if (checked) {
                                    if (!selectedSpans.includes(row.SpanID)) toggleSelectSpan(row.SpanID);
                                  } else {
                                    if (selectedSpans.includes(row.SpanID)) toggleSelectSpan(row.SpanID);
                                  }
                                }}
                                styles={{ root: { margin: 0, padding: 0, display: 'block' } } as any}
                              />
                            </div>
                          </td>
                          {dynamicCols.map(c => {
                            const rawValue = c.render ? c.render(row) : ((row as any)[c.key] ?? '');
                            // Prepare displayValue: for detailed view, truncate long OpticalDevice columns to keep table compact
                            let displayValue: any = rawValue;
                            // Aggressively truncate long device/splice identifiers in detailed view to avoid horizontal scrolling
                            if (!simplifiedView) {
                              if (c.key === 'OpticalDeviceA' || c.key === 'OpticalDeviceZ') {
                                if (typeof rawValue === 'string' && rawValue.length > 8) displayValue = rawValue.slice(0, 8) + '...';
                              }
                              if (c.key === 'SpliceRackA_Unit' || c.key === 'SpliceRackA' || c.key === 'SpliceRackZ') {
                                if (typeof rawValue === 'string' && rawValue.length > 8) displayValue = rawValue.slice(0, 8) + '...';
                              }
                            }

                            // Width tweaks: shrink Span/Facility A/Z; otherwise use the computed percent so columns fit evenly
                            const truncateSplice = (val: string) => {
                              if (!val) return val;
                              const s = String(val);
                              if (s.length <= 10) return s;
                              if (s.includes('-')) {
                                const parts = s.split('-');
                                if (parts.length >= 2) {
                                  const after = parts[1] ? parts[1].slice(0, 1) : '';
                                  return `${parts[0]}-${after}...`;
                                }
                              }
                              return s.slice(0, 8) + '...';
                            };

                            const baseStyle: React.CSSProperties = {
                              // reduce horizontal padding so content shifts left and gives more room for status pills
                              padding: '1px 6px',
                              whiteSpace: 'nowrap',
                              // Always clip overflow and use ellipsis to avoid overlapping adjacent columns
                              overflow: 'hidden',
                              textOverflow: 'ellipsis',
                              // keep small fixed widths for id columns; otherwise use percent. Use small maxWidth to compress.
                              width: (c.key === 'SpanID') ? '8ch'
                                : (c.key === 'Datacenter') ? '4ch'
                                : (c.key === 'Diversity') ? '7%'
                                : (c.key === 'FacilityCodeA') ? '5ch'
                                : (c.key === 'FacilityCodeZ') ? '5ch'
                                : (c.key === 'IDF_A') ? '5ch'
                                : (c.key === 'SpliceRackZ' || c.key === 'SpliceRackZ_Unit') ? '5ch'
                                : (c.key === 'WiringScope') ? '6ch'
                                : (c.key === 'SpanType') ? '5ch'
                                : (c.key === 'Status') ? '12%'
                                : `${colWidthPercent}%`,
                              maxWidth: (c.key === 'SpanID') ? '8ch' : (c.key === 'Datacenter' ? '4ch' : (c.key === 'FacilityCodeA' ? '5ch' : (c.key === 'FacilityCodeZ' ? '5ch' : (c.key === 'IDF_A' ? '5ch' : '8ch')))),
                              ...(c.key === 'Diversity' ? { paddingLeft: 10 } : {}),
                              ...(c.key === 'SpanID' ? { textAlign: 'center' as const } : {})
                            };
                            return (
                              <td
                                key={`${i}-${c.key}`}
                                style={baseStyle}
                                title={typeof rawValue === 'string' && (c.key !== 'Status' && c.key !== 'Diversity') ? rawValue : undefined}
                              >
                                {/* For splice/rack columns, apply a short readable truncation so values like "ZRH21-1..." appear */}
                                {((c.key === 'SpliceRackA_Unit' || c.key === 'SpliceRackA' || c.key === 'SpliceRackZ' || c.key === 'SpliceRackZ_Unit') && typeof displayValue === 'string') ? truncateSplice(displayValue) : displayValue}
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              );
            })()}

            <div style={{ textAlign: "center", marginTop: 12 }}>
              <div style={{ display: 'inline-flex', gap: 8, alignItems: 'center' }}>
                {/* Disable when none selected or when selection exceeds 20 */}
                <TooltipHost content={selectedSpans.length > 20 ? 'Max 20 spans at a time. Reduce selection to enable this action.' : undefined}>
                  <div>
                    <PrimaryButton
                      text="Span Traffic"
                      disabled={selectedSpans.length === 0 || selectedSpans.length > 20}
                      onClick={() => {
                        try {
                          const q = encodeURIComponent(selectedSpans.join(','));
                          const url = `${window.location.origin}/fiber-span-utilization?spans=${q}`;
                          window.open(url, '_blank');
                        } catch (e) {
                          // fallback: open without params in new tab
                          try { window.open(`${window.location.origin}/fiber-span-utilization`, '_blank'); } catch { /* ignore */ }
                        }
                      }}
                      styles={{ root: { backgroundColor: selectedSpans.length > 20 ? undefined : '#6a00ff', borderColor: '#5a00e6', height: 36, borderRadius: 6, color: '#fff' } }}
                    />
                  </div>
                </TooltipHost>
                <PrimaryButton
                  text={`Continue (${selectedSpans.length} selected)`}
                  disabled={selectedSpans.length === 0}
                  onClick={() => setComposeOpen(true)}
                />
              </div>
            </div>
          </div>
        )}
          </div>
        )}

        {/* === Compose Section === */}
        {composeOpen && (
          <div className="table-container compose-container" style={{ marginTop: 16 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
              <IconButton
                className="back-button"
                iconProps={{ iconName: 'ChevronLeft' }}
                title="Back"
                ariaLabel="Back"
                onClick={() => setComposeOpen(false)}
              />
              <div className="section-title" style={{ margin: 0 }}>Compose Maintenance Email</div>
            </div>

            {/* policy-info removed per user request */}

            {/* success is shown as a dialog to match Emergency UX */}
            {sendError && (
              <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                {sendError}
              </MessageBar>
            )}
            {showValidation && Object.keys(fieldErrors).length > 0 && (
              <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                Please complete the required fields (these cannot be empty): {Object.keys(fieldErrors).map(k => friendlyFieldNames[k] || k).join(', ')}
              </MessageBar>
            )}

            <Dialog
              hidden={!showEmergencyDialog}
              className="dialog-emergency"
              onDismiss={() => { setShowEmergencyDialog(false); setStartWarning(null); setPendingEmergency(false); }}
              dialogContentProps={{
                type: DialogType.normal,
                title: 'Confirm Emergency Maintenance',
                subText: startWarning || '',
              }}
              modalProps={{ isBlocking: true }}
            >
              <DialogFooter>
                <PrimaryButton
                  text="Confirm & Mark Emergency"
                  onClick={() => {
                    if (pendingEmergency) {
                      addEmergencyTag();
                      setPendingEmergency(false);
                    }
                    setShowEmergencyDialog(false);
                    setStartWarning(null);
                  }}
                />
                <DefaultButton
                  text="Cancel"
                  onClick={() => { setShowEmergencyDialog(false); setStartWarning(null); setPendingEmergency(false); }}
                />
              </DialogFooter>
            </Dialog>

            <Dialog
              hidden={!showFieldErrorDialog}
              className="dialog-field-error"
              onDismiss={() => { setShowFieldErrorDialog(false); setFieldErrorMessageDialog(null); }}
              dialogContentProps={{
                type: DialogType.normal,
                title: 'Invalid Selection',
                subText: fieldErrorMessageDialog || 'Selected value is invalid. Please correct and try again.'
              }}
              modalProps={{ isBlocking: false }}
            >
              <DialogFooter>
                <PrimaryButton text="OK" onClick={() => { setShowFieldErrorDialog(false); setFieldErrorMessageDialog(null); }} />
              </DialogFooter>
            </Dialog>

            <Dialog
              hidden={!showSendSuccessDialog}
              className="dialog-send-success"
              onDismiss={() => setShowSendSuccessDialog(false)}
              dialogContentProps={{
                type: DialogType.normal,
                title: 'Email Sent',
                subText: sendSuccess || 'Email has been sent successfully.'
              }}
              modalProps={{ isBlocking: false }}
            >
              <DialogFooter>
                <PrimaryButton
                  text="Start Over"
                  onClick={() => {
                    // Reset everything and navigate forcing a refresh key
                    resetAll();
                    navigate(`/vso?reset=${Date.now()}`, { replace: true });
                  }}
                />
                <DefaultButton text="Close" onClick={() => setShowSendSuccessDialog(false)} />
              </DialogFooter>
            </Dialog>


            <div style={{ display: 'flex', gap: 12, alignItems: 'flex-start' }}>
              <div style={{ flex: 1 }}>
                <div style={{ display: 'flex', gap: 12, alignItems: 'center' }}>
                  <div style={{ flex: 2 }}>
                    <TextField
                      componentRef={subjectRef}
                      label="Subject"
                      placeholder="Fiber Maintenance scheduled in <FacilityCode> for Spans <Enter Spans here>"
                      value={subject}
                      onChange={(_, v) => setSubject(v || "")}
                      styles={getTextFieldStyles('subject')}
                      required
                      errorMessage={showValidation ? fieldErrors.subject : undefined}
                    />
                  </div>
                  <div style={{ flex: 1 }}>
                    <TextField
                      componentRef={ccRef}
                      label="CC"
                      placeholder="name@contoso.com"
                      value={cc}
                      onChange={(_, v) => setCc(v || "")}
                      styles={getTextFieldStyles('cc')}
                      required
                      errorMessage={showValidation ? fieldErrors.cc : undefined}
                    />
                  </div>
                  <div style={{ width: 320, flexShrink: 0 }}>
                    <Dropdown
                      label="Notification Type"
                      options={[
                        { key: "New Maintenance Scheduled", text: "New Maintenance Scheduled" },
                        { key: "Rescheduled", text: "Rescheduled" },
                        { key: "Maintenance Cancelled", text: "Maintenance Cancelled" },
                        { key: "Maintenance Reminder", text: "Maintenance Reminder" },
                      ]}
                      selectedKey={notificationType}
                      onChange={(_, opt) => opt && setNotificationType(opt.key.toString())}
                      styles={dropdownStyles}
                      required
                    />
                  </div>
                </div>

                <div className="compose-datetime-row">
                    <div className="dt-field">
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <Text styles={{ root: { color: "var(--vso-label-color)", fontSize: 12, fontWeight: 600 } }}>Start Date</Text>
                        <TooltipHost content={"Times shown are in your local time and will be converted automatically when the VSO is created."}>
                          <IconButton iconProps={{ iconName: 'Info' }} title="Time conversion info" styles={{ root: { color: '#a6b7c6', height: 18, width: 18 } }} />
                        </TooltipHost>
                      </div>
                      <DatePicker
                        componentRef={startDateRef}
                        placeholder="Select start date"
                        value={startDate || undefined}
                        minDate={new Date()}
                        onSelectDate={(d) => {
                          const selected = d || null;
                          // Prevent past-day selection (and show popup)
                          if (isPastDay(selected)) {
                            setFieldErrorMessageDialog("Selected Start Date is in the past. Please choose today or a future date.");
                            setShowFieldErrorDialog(true);
                            return;
                          }
                          setStartDate(selected);
                          if (isWithinDays(selected, 7)) {
                            setStartWarning(
                              "You have selected a date less than 7 days in advance. If this is required please press confirm to continue and the email will be updated with the Emergency Tag."
                            );
                            setPendingEmergency(true);
                            setShowEmergencyDialog(true);
                          } else {
                            setStartWarning(null);
                            setPendingEmergency(false);
                            // If there are no windows (primary + additional) within 7 days, remove the emergency tag
                            if (!anyWindowWithin7Days(selected, additionalWindows)) removeEmergencyTag();
                          }
                        }}
                        styles={getDatePickerStyles('startDate')}
                        isRequired
                        aria-label="Start Date"
                      />
                      {showValidation && fieldErrors.startDate ? (
                        <Text styles={{ root: { color: '#a80000', fontSize: 12 } }}>Required</Text>
                      ) : null}
                    </div>

                    <div className="dt-time">
                      <Text styles={{ root: { color: "var(--vso-label-color)", fontSize: 12, fontWeight: 600 } }}>Start Time</Text>
                      <Dropdown
                        componentRef={startTimeRef}
                        options={timeOptions}
                        selectedKey={startTime}
                        onChange={(_, opt) => {
                          if (!opt) return;
                          const next = opt.key.toString();
                          setStartTime(next);
                          // If end is already selected, ensure end > new start
                          const sDT = parseTimeToDate(startDate, next);
                          const eDT = parseTimeToDate(endDate, endTime);
                          if (sDT && eDT && eDT.getTime() <= sDT.getTime()) {
                            setFieldErrors((prev) => ({ ...prev, endTime: 'End must be after start' }));
                            setFieldErrorMessageDialog('End date/time must be after start date/time.');
                            setShowFieldErrorDialog(true);
                          } else {
                            setFieldErrors((prev) => { const copy = { ...prev }; delete copy.endTime; return copy; });
                          }
                        }}
                        styles={getDropdownStyles('startTime', timeDropdownStyles)}
                        required
                        errorMessage={showValidation ? fieldErrors.startTime : undefined}
                      />
                    </div>

                    <div className="dt-field">
                      <Text styles={{ root: { color: "var(--vso-label-color)", fontSize: 12, fontWeight: 600 } }}>End Date</Text>
                      <DatePicker
                        componentRef={endDateRef}
                        placeholder="Select end date"
                        value={endDate || undefined}
                        minDate={new Date()}
                        onSelectDate={(d) => {
                          const selected = d || null;
                          if (isPastDay(selected)) {
                            setFieldErrorMessageDialog("Selected End Date is in the past. Please choose today or a future date.");
                            setShowFieldErrorDialog(true);
                            return;
                          }
                          setEndDate(selected || null);
                          // If both dates present, ensure end (date+time) > start (date+time)
                          const sDT = parseTimeToDate(startDate, startTime);
                          const eDT = parseTimeToDate(selected, endTime);
                          if (sDT && eDT && eDT.getTime() <= sDT.getTime()) {
                            setFieldErrors((prev) => ({ ...prev, endTime: 'End must be after start' }));
                            setFieldErrorMessageDialog('End date/time must be after start date/time.');
                            setShowFieldErrorDialog(true);
                          } else {
                            setFieldErrors((prev) => { const copy = { ...prev }; delete copy.endTime; return copy; });
                          }
                        }}
                        styles={getDatePickerStyles('endDate')}
                        isRequired
                        aria-label="End Date"
                      />
                      {showValidation && fieldErrors.endDate ? (
                        <Text styles={{ root: { color: '#a80000', fontSize: 12 } }}>Required</Text>
                      ) : null}
                    </div>

                    <div className="dt-time">
                      <Text styles={{ root: { color: "var(--vso-label-color)", fontSize: 12, fontWeight: 600 } }}>End Time</Text>
                      <Dropdown
                        componentRef={endTimeRef}
                        options={timeOptions}
                        selectedKey={endTime}
                        onChange={(_, opt) => {
                          if (!opt) return;
                          const next = opt.key.toString();
                          // If start date/time known, validate ordering
                          const sDT = parseTimeToDate(startDate, startTime);
                          const eDT = parseTimeToDate(endDate || startDate, next);
                          if (sDT && eDT && eDT.getTime() <= sDT.getTime()) {
                            setFieldErrors((prev) => ({ ...prev, endTime: 'End must be after start' }));
                            setFieldErrorMessageDialog('End date/time must be after start date/time.');
                            setShowFieldErrorDialog(true);
                            setEndTime(next);
                            return;
                          }
                          setFieldErrors((prev) => { const copy = { ...prev }; delete copy.endTime; return copy; });
                          setEndTime(next);
                        }}
                        styles={timeDropdownStyles}
                        required
                        errorMessage={showValidation ? fieldErrors.endTime : undefined}
                      />
                    </div>

                    <div className="dt-actions">
                      <button type="button" className="tiny-icon-btn add-window-btn" aria-label="Add Window" onClick={addWindow} title="Add Window">
                        <span className="glyph">+</span>
                      </button>
                    </div>
                </div>

              {/* Additional maintenance windows */}
              <div style={{ marginTop: 8 }}>
                {additionalWindows.map((w, i) => (
                  <div key={i} className="compose-datetime-row additional-window">
                    <div className="dt-field">
                      <Text styles={{ root: { color: "var(--vso-label-color)", fontSize: 12, fontWeight: 600 } }}>Start Date</Text>
                      <DatePicker
                        placeholder="Select start date"
                        value={w.startDate || undefined}
                        minDate={new Date()}
                        onSelectDate={(d) => {
                          const selected = d || null;
                          if (isPastDay(selected)) {
                            setFieldErrorMessageDialog("Selected Start Date is in the past. Please choose today or a future date.");
                            setShowFieldErrorDialog(true);
                            return;
                          }
                          setAdditionalWindows((arr) => {
                            const next = [...arr];
                            next[i] = { ...next[i], startDate: selected };
                            return next;
                          });

                          if (isWithinDays(selected, 7)) {
                            setStartWarning(
                              "You have selected a date less than 7 days in advance. If this is required please press confirm to continue and the email will be updated with the Emergency Tag."
                            );
                            setPendingEmergency(true);
                            setShowEmergencyDialog(true);
                          } else {
                            // Re-evaluate across primary + updated additional windows
                            const updated = additionalWindows.map((w, idx) => (idx === i ? { ...w, startDate: selected } : w));
                            if (!anyWindowWithin7Days(startDate, updated)) {
                              setStartWarning(null);
                              setPendingEmergency(false);
                              removeEmergencyTag();
                            }
                          }
                        }}
                        styles={datePickerStyles}
                      />
                    </div>
                    <div className="dt-time">
                      <Text styles={{ root: { color: "var(--vso-label-color)", fontSize: 12, fontWeight: 600 } }}>Start Time</Text>
                      <Dropdown
                        options={timeOptions}
                        selectedKey={w.startTime}
                        onChange={(_, opt) => {
                          if (!opt) return;
                          const nextTime = opt.key.toString();
                          setAdditionalWindows((arr) => {
                            const next = [...arr];
                            next[i] = { ...next[i], startTime: nextTime };
                            return next;
                          });
                          // Validate with end time for this window
                          const win = additionalWindows[i];
                          const sDT = parseTimeToDate(win?.startDate || null, nextTime);
                          const eDT = parseTimeToDate(win?.endDate || win?.startDate || null, win?.endTime || null);
                          if (sDT && eDT && eDT.getTime() <= sDT.getTime()) {
                            setFieldErrors((prev) => ({ ...prev, [`additional-${i}-endTime`]: 'End must be after start' }));
                            setFieldErrorMessageDialog('End date/time must be after start date/time for the additional window.');
                            setShowFieldErrorDialog(true);
                          } else {
                            setFieldErrors((prev) => { const copy = { ...prev }; delete copy[`additional-${i}-endTime`]; return copy; });
                          }
                        }}
                        styles={getDropdownStyles('endTime', timeDropdownStyles)}
                      />
                    </div>
                    <div className="dt-field">
                      <Text styles={{ root: { color: "var(--vso-label-color)", fontSize: 12, fontWeight: 600 } }}>End Date</Text>
                      <DatePicker
                        placeholder="Select end date"
                        value={w.endDate || undefined}
                        minDate={new Date()}
                        onSelectDate={(d) => {
                          const selected = d || null;
                          if (isPastDay(selected)) {
                            setFieldErrorMessageDialog("Selected End Date is in the past. Please choose today or a future date.");
                            setShowFieldErrorDialog(true);
                            return;
                          }
                          setAdditionalWindows((arr) => {
                            const next = [...arr];
                            next[i] = { ...next[i], endDate: selected || null };
                            return next;
                          });
                          const win = additionalWindows[i];
                          const sDT = parseTimeToDate(win?.startDate || null, win?.startTime || null);
                          const eDT = parseTimeToDate(selected, win?.endTime || null);
                          if (sDT && eDT && eDT.getTime() <= sDT.getTime()) {
                            setFieldErrors((prev) => ({ ...prev, [`additional-${i}-endTime`]: 'End must be after start' }));
                            setFieldErrorMessageDialog('End date/time must be after start date/time for the additional window.');
                            setShowFieldErrorDialog(true);
                          } else {
                            setFieldErrors((prev) => { const copy = { ...prev }; delete copy[`additional-${i}-endTime`]; return copy; });
                          }
                        }}
                        styles={datePickerStyles}
                      />
                    </div>
                    <div className="dt-time">
                      <Text styles={{ root: { color: "var(--vso-label-color)", fontSize: 12, fontWeight: 600 } }}>End Time</Text>
                      <Dropdown
                        options={timeOptions}
                        selectedKey={w.endTime}
                        onChange={(_, opt) => {
                          if (!opt) return;
                          const nextTime = opt.key.toString();
                          setAdditionalWindows((arr) => {
                            const next = [...arr];
                            next[i] = { ...next[i], endTime: nextTime };
                            return next;
                          });
                          const win = additionalWindows[i];
                          const sDT = parseTimeToDate(win?.startDate || null, win?.startTime || null);
                          const eDT = parseTimeToDate(win?.endDate || win?.startDate || null, nextTime);
                          if (sDT && eDT && eDT.getTime() <= sDT.getTime()) {
                            setFieldErrors((prev) => ({ ...prev, [`additional-${i}-endTime`]: 'End must be after start' }));
                            setFieldErrorMessageDialog('End date/time must be after start date/time for the additional window.');
                            setShowFieldErrorDialog(true);
                          } else {
                            setFieldErrors((prev) => { const copy = { ...prev }; delete copy[`additional-${i}-endTime`]; return copy; });
                          }
                        }}
                        styles={timeDropdownStyles}
                      />
                    </div>
                    <div className="dt-actions">
                      {/* only show a compact remove button for additional windows */}
                      <button type="button" className="tiny-icon-btn remove-window-btn" aria-label={`Remove window ${i + 2}`} title="Remove Window" onClick={() => removeWindow(i)}>
                        <span className="glyph">−</span>
                      </button>
                    </div>
                  </div>
                ))}
              </div>

                <div style={{ display: 'flex', gap: 12, marginTop: 8 }}>
                <div style={{ flex: 1 }}>
                  <TextField componentRef={locationRef} label="Location" value={location} onChange={(_, v) => setLocation(v || "")} styles={textFieldStyles} required errorMessage={showValidation ? fieldErrors.location : undefined} />
                </div>
                <div style={{ flex: 1 }}>
                  <TextField label="Latitude" value={lat} onChange={(_, v) => setLat((v || '').trim())} styles={textFieldStyles} />
                </div>
                <div style={{ flex: 1 }}>
                  <TextField label="Longitude" value={lng} onChange={(_, v) => setLng((v || '').trim())} styles={textFieldStyles} />
                </div>
              </div>

                <div style={{ display: 'flex', gap: 12, marginTop: 8, alignItems: 'flex-end' }}>
                    <div style={{ width: 200 }}>
                      <Text styles={{ root: { color: 'var(--vso-label-color)', fontSize: 12, fontWeight: 600, marginBottom: 6 } }}>Tags</Text>
                        <Dropdown
                        placeholder="Select tags"
                        multiSelect
                        options={tagOptions}
                        selectedKeys={tags}
                        onChange={(_, option) => {
                          if (!option) return;
                          const key = option.key?.toString() || '';
                          if ((option as any).selected) {
                            setTags((prev) => (prev.includes(key) ? prev : [...prev, key]));
                          } else {
                            setTags((prev) => prev.filter((k) => k !== key));
                          }
                        }}
                        styles={{
                          // use theme variables so light/dark adapt automatically
                          ...dropdownStyles,
                          root: { width: 200 },
                          dropdownItem: { background: 'transparent', color: 'var(--vso-dropdown-text)', selectors: { ':hover': { background: 'var(--vso-dropdown-item-hover-bg)', color: 'var(--vso-dropdown-text)' } } },
                          dropdownItemSelected: { background: 'var(--vso-dropdown-item-selected-bg)', color: 'var(--vso-dropdown-text)', selectors: { ':hover': { background: 'var(--vso-dropdown-item-hover-bg)', color: 'var(--vso-dropdown-text)' } } },
                          callout: { background: 'var(--vso-dropdown-bg)' },
                          title: { ...dropdownStyles.title, height: 42 },
                        }}
                      />
                    </div>
                  <div style={{ width: 200 }}>
                    <Dropdown
                      label="Impact Expected"
                      options={[{ key: "true", text: "Yes/True" }, { key: "false", text: "No/False" }]}
                      selectedKey={impactExpected ? "true" : "false"}
                      onChange={(_, opt) => opt && setImpactExpected(opt.key === "true")}
                      styles={getDropdownStyles('impactExpected', dropdownStyles)}
                      required
                      errorMessage={showValidation ? fieldErrors.impactExpected : undefined}
                    />
                  </div>
                </div>

                <div style={{ display: 'flex', gap: 8, marginTop: 8, alignItems: 'center' }}>
                  {lat && lng ? (
                    <a className="uid-click" href={`https://www.bing.com/maps?q=${encodeURIComponent(lat+','+lng)}`} target="_blank" rel="noopener noreferrer">Open Map</a>
                  ) : null}
                </div>
              </div>
            </div>

            <div style={{ marginTop: 10 }}>
              {/* Reason wrapper with counter */}
              <div className="reason-wrapper" style={{ position: 'relative' }}>
                <TextField
                  componentRef={maintenanceReasonRef}
                  className="reason-field"
                  label="Reason for Maintenance"
                  multiline
                  autoAdjustHeight
                  value={maintenanceReason}
                  onChange={(_, v) => setMaintenanceReason(v || "")}
                  styles={{
                    ...getTextFieldStyles('maintenanceReason'),
                    field: { ...(getTextFieldStyles('maintenanceReason') as any).field, minHeight: 220, paddingBottom: 28 },
                    fieldGroup: { ...((getTextFieldStyles('maintenanceReason') as any).fieldGroup || {}), height: 'auto' },
                  }}
                  // limit to a reasonable amount
                  maxLength={2000}
                  aria-label="Reason for Maintenance"
                  required
                  errorMessage={showValidation ? fieldErrors.maintenanceReason : undefined}
                />

                <div className="reason-counter" aria-hidden style={{ position: 'absolute', right: 10, bottom: 8, fontSize: 12 }}>
                  {`${maintenanceReason.length}/2000`}
                </div>
              </div>
            </div>

            <div className="section-title email-preview-header" style={{ marginTop: 6 }}>Email Body Preview</div>
            <div
              className="email-preview"
              style={{
                borderRadius: 8,
                padding: 12,
                whiteSpace: "pre-wrap",
                border: "1px solid #333",
                background: "#0f0f0f",
                color: "#dfefff",
              }}
            >
              {emailBody}
            </div>

            <div style={{ display: "flex", justifyContent: "space-between", alignItems: 'center', marginTop: 12 }}>
              <button className="sleek-btn danger" onClick={() => setComposeOpen(false)}>
                Back
              </button>
              <button
                className={`sleek-btn wan ${!canSend && !sendLoading ? 'muted-disabled' : ''}`}
                // Keep clickable so handleSend can show validation; only truly disable while sending
                disabled={sendLoading}
                onClick={handleSend}
                title={!canSend && !sendLoading ? 'Some required fields are missing. Click to validate.' : undefined}
              >
                {sendLoading ? "Sending..." : "Confirm & Send"}
              </button>
            </div>
          </div>
        )}

        <hr />
        <div className="disclaimer">
        This tool is intended for internal use within Microsoft’s Data Center Operations and Network Delivery environments. Always verify critical data before taking operational action. The information provided is automatically retrieved from validated sources but may not reflect the most recent updates, configurations, or status changes in live systems. Users are responsible for ensuring all details are accurate before proceeding with submitting a VSO. This application is developed and maintained by <b>Josh Maclean</b>, supported by the <b>CIA | Network Delivery</b> team. For any issues or requests please <a href="https://teams.microsoft.com/l/chat/0/0?users=joshmaclean@microsoft.com" className="uid-click">send a message</a>. 
        </div>
      </div>
      {/* Big calendar below the card */}
      <VSOCalendar
        events={vsoEvents}
        date={calendarDate || undefined}
        onNavigate={(d) => {
          setCalendarDate(d);
          try { localStorage.setItem("vsoCalendarDate", d.toISOString()); } catch {}
        }}
        onEventClick={(ev) => {
          setActiveEventId(ev.id);
          setShowEventDialog(true);
        }}
      />

      {/* Event Details Dialog - rendered globally so it works from anywhere */}
      <Dialog
        hidden={!showEventDialog}
        className="dialog-event-details"
        onDismiss={() => setShowEventDialog(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Maintenance Details',
        }}
        modalProps={{ isBlocking: false }}
      >
        {(() => {
          const ev = vsoEvents.find((e) => e.id === activeEventId);
          if (!ev) return null;
          return (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              <div style={{ display: 'flex', gap: 12 }}>
                <div style={{ flex: 1 }}>
                  <TextField componentRef={locationRef} label="Location" value={location} onChange={(_, v) => setLocation(v || "")} styles={getTextFieldStyles('location')} required errorMessage={showValidation ? fieldErrors.location : undefined} />
                </div>
                <div style={{ width: 240 }}>
                  <Dropdown
                    label="Status"
                    options={[
                      { key: 'Draft', text: 'Draft' },
                      { key: 'Approved', text: 'Approved' },
                      { key: 'Rejected', text: 'Rejected' },
                    ]}
                    selectedKey={ev.status}
                    onChange={(_, opt) => {
                      if (!opt) return;
                      const next = opt.key.toString() as any;
                      setVsoEvents((arr) => arr.map((x) => (x.id === ev.id ? { ...x, status: next } : x)));
                      // Persist status change to server so all users see it
                      (async () => {
                        try {
                          const owner = (() => { try { return localStorage.getItem('loggedInEmail') || 'VSO Calendar'; } catch { return 'VSO Calendar'; } })();
                          await saveToStorage({
                            category: 'Calendar',
                            uid: 'VsoCalendar',
                            title: ev.title || 'VSO Event',
                            description: ev.maintenanceReason || ev.summary || '',
                            owner,
                            timestamp: ev.start || new Date(),
                            rowKey: ev.id,
                            extras: { Status: next },
                          });
                        } catch (e) {
                          // eslint-disable-next-line no-console
                          console.warn('Failed to persist calendar status change', e);
                        }
                      })();
                    }}
                    styles={{
                      ...dropdownStyles,
                    }}
                  />
                </div>
              </div>

              <div className="equal-tables-row" style={{ gap: 12 }}>
                <div className="table-container details-fit" style={{ flex: 1 }}>
                  <div className="section-title">Schedule</div>
                  <table className="data-table compact details-table">
                    <tbody>
                      <tr><td>Day</td><td>{ev.start.toLocaleDateString()}</td></tr>
                      <tr><td>Time</td><td>{ev.startTimeUtc || '--'} - {ev.endTimeUtc || '--'}</td></tr>
                    </tbody>
                  </table>
                </div>
                <div className="table-container details-fit" style={{ flex: 1 }}>
                  <div className="section-title">Context</div>
                  <table className="data-table compact details-table">
                    <tbody>
                      <tr><td>DC</td><td>{ev.dcCode || '--'}</td></tr>
                      <tr><td>Notification</td><td>{ev.notificationType || '--'}</td></tr>
                      <tr><td>Subject</td><td>{ev.subject || '--'}</td></tr>
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="equal-tables-row" style={{ gap: 12 }}>
                <div className="table-container details-fit" style={{ flex: 1 }}>
                  <div className="section-title">Location</div>
                  <table className="data-table compact details-table">
                    <tbody>
                      <tr><td>Location</td><td>{ev.location || '--'}</td></tr>
                      <tr><td>ISP</td><td>{ev.isp || '--'}</td></tr>
                      <tr><td>ISP Ticket</td><td>{ev.ispTicket || '--'}</td></tr>
                      <tr><td>Impact Expected</td><td>{ev.impactExpected ? 'Yes/True' : 'No/False'}</td></tr>
                    </tbody>
                  </table>
                </div>
                <div className="table-container details-fit" style={{ flex: 1 }}>
                  <div className="section-title">Spans</div>
                  <div style={{ padding: 8, color: 'var(--text-1)' }}>{(ev.spans || []).join(', ') || '--'}</div>
                </div>
              </div>

              <div className="table-container details-fit">
                <div className="section-title">Reason</div>
                <div style={{ padding: 8, color: 'var(--text-1)', whiteSpace: 'pre-wrap' }}>{ev.maintenanceReason || ev.summary || '--'}</div>
              </div>

              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
                <DefaultButton text="Close" onClick={() => setShowEventDialog(false)} />
              </div>
            </div>
          );
        })()}
      </Dialog>
    </div>
  );
};

export default VSOAssistant;
