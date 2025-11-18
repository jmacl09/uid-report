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
  Toggle,
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
  initializeIcons,
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
  // Ensure Fluent UI icon font is available for this page
  useEffect(() => {
    try { initializeIcons(); } catch {}
  }, []);
  const [facilityCodeA, setFacilityCodeA] = useState<string>("");
  const [facilityCodeZ, setFacilityCodeZ] = useState<string>("");
  const [diversity, setDiversity] = useState<string>();
  const [spliceRackA, setSpliceRackA] = useState<string>();
  const [spliceRackZ, setSpliceRackZ] = useState<string>();
  const [loading, setLoading] = useState<boolean>(false);
  const [result, setResult] = useState<SpanData[]>([]);
  const [selectedSpans, setSelectedSpans] = useState<string[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [showAll, setShowAll] = useState<boolean>(false);
  const [, setRackUrl] = useState<string>();
  const [rackDC, setRackDC] = useState<string>();
  const [dcSearch, setDcSearch] = useState<string>("");
  const [dcSearchZ, setDcSearchZ] = useState<string>("");
  const [country, setCountry] = useState<string>("");
  const [countrySearch, setCountrySearch] = useState<string>("");
  const countryComboRef = React.useRef<IComboBox | null>(null);
  const [countryPanelOpen, setCountryPanelOpen] = useState<boolean>(false);
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

  // Sorting state for results table
  const [sortBy, setSortBy] = useState<string>("");
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");
  // Removed per-column filters in favor of simple clickable sort

  // === Stage 2: Compose Email state ===
  const [composeOpen, setComposeOpen] = useState<boolean>(false);
  const EMAIL_TO = "opticaldri@microsoft.com"; // fixed
  const EMAIL_LOGIC_APP_URL = "https://fibertools-dsavavdcfdgnh2cm.westeurope-01.azurewebsites.net:443/api/VSO/triggers/When_an_HTTP_request_is_received/invoke?api-version=2022-05-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=6ViXNM-TmW5F7Qd9_e4fz3IhRNqmNzKwovWvcmuNJto";
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
    // Read persisted login email (if any) so CC preview is available immediately
    try {
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
        const res = await fetch("/.auth/me", { credentials: "include" });
        if (!res.ok) return;
        const data = await res.json();
        // Debugging: print the full /.auth/me response so we can see available claims
        try {
          // eslint-disable-next-line no-console
          console.debug("/.auth/me response:", data);
        } catch (e) {}
        // Handle both App Service ([identities]) and Static Web Apps ({clientPrincipal}) shapes
        const identities = Array.isArray(data)
          ? data
          : data?.clientPrincipal
          ? [{ user_claims: data.clientPrincipal?.claims || [] }]
          : [];
        for (const id of identities) {
          const claims = id?.user_claims || [];
          const getClaim = (t: string) => claims.find((c: any) => c?.typ === t)?.val || "";
          const email =
            getClaim("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress") ||
            getClaim("emails") || // SWA sometimes
            getClaim("preferred_username") ||
            getClaim("upn") ||
            "";
          if (email) {
            setUserEmail(email);
            try {
              localStorage.setItem("loggedInEmail", email);
            } catch (e) {}
            return;
          }
        }
      } catch {}
    };
    fetchUserEmail();
  }, []);

  // When the user edits any search input, clear the temporary "show all options" state
  // and reset prompt/exhaustion flags so they can get fresh prompts on a new search.
  useEffect(() => {
    // Clear any temporary 'show all' or prompt/exhaustion flags when the user edits search inputs.
    // We explicitly set values rather than reading them to avoid stale-read dependency issues.
    setShowAllOptions(false);
    setOppositePromptUsed(false);
    setTriedBothNoResults(false);
    setOppositePrompt({ show: false, from: null });
    setTriedSides({ A: false, Z: false });
    // Intentionally do not clear search results or the no-results banner here; keep that visible until the user re-submits.
  }, [facilityCodeA, facilityCodeZ, spliceRackA, spliceRackZ, diversity]);

  // Persist calendar events whenever they change
  useEffect(() => {
    try {
      const serializable = vsoEvents.map((e) => ({
        ...e,
        // Persist both ISO and date-only for robust reload across timezones
        start: e.start.toISOString(),
        end: e.end.toISOString(),
        startYMD: `${e.start.getFullYear()}-${String(e.start.getMonth() + 1).padStart(2, '0')}-${String(e.start.getDate()).padStart(2, '0')}`,
        endYMD: `${e.end.getFullYear()}-${String(e.end.getMonth() + 1).padStart(2, '0')}-${String(e.end.getDate()).padStart(2, '0')}`,
      }));
      localStorage.setItem("vsoEvents", JSON.stringify(serializable));
      // Also write a backup copy to guard against accidental clears/overwrites
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
    setDiversity(undefined);
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
  };

  // === Diversity options ===
  const diversityOptions: IDropdownOption[] = [
    // Blank option to allow clearing selection
    { key: "", text: "" },
    { key: "West", text: "West, West 1, West 2" },
    { key: "East", text: "East, East 1, East 2" },
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
    // Require exactly one facility code (A or Z). Selecting one disables the other in the UI.
    const hasA = !!facilityCodeA;
    const hasZ = !!facilityCodeZ;
    if ((hasA && hasZ) || (!hasA && !hasZ)) {
      alert("Please select either Facility Code A or Facility Code Z (choose one).");
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

  setLoading(true);
    setError(null);
    setResult([]);
    setSearchDone(false);

    try {
      const logicAppUrl =
        "https://fibertools-dsavavdcfdgnh2cm.westeurope-01.azurewebsites.net:443/api/VSO/triggers/When_an_HTTP_request_is_received/invoke?api-version=2022-05-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=6ViXNM-TmW5F7Qd9_e4fz3IhRNqmNzKwovWvcmuNJto";

      const diversityValue = (() => {
        const raw = (diversity || "").toString();
        if (!raw) return "N";
        // Preserve comma-separated groups and normalize spacing: "A, B, C"
        const normalized = raw.split(',').map((s) => s.trim()).filter(Boolean).join(', ');
        return normalized || "N";
      })();
      const stage = computeScopeStage({
        facilityA: facilityCodeA,
        facilityZ: facilityCodeZ,
        diversity: diversityValue === "N" ? "" : diversityValue,
        spliceA: spliceRackA,
        spliceZ: spliceRackZ,
      });
      const payload: any = {
        Diversity: diversityValue,
        Stage: stage,
      };
      if (facilityCodeA) payload.FacilityCodeA = facilityCodeA;
      if (facilityCodeZ) payload.FacilityCodeZ = facilityCodeZ;
      if (spliceRackA) payload.SpliceRackA = spliceRackA;
      if (spliceRackZ) payload.SpliceRackZ = spliceRackZ;

      const response = await fetch(logicAppUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
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
      color: "#fff",
      backgroundColor: "#141414",
      height: 42,
      border: "1px solid #333",
      borderRadius: 8,
      paddingLeft: 10,
    },
    callout: { background: "#181818", maxHeight: 240, overflowY: "auto" },
    optionsContainer: { background: "#181818" },
    caretDownWrapper: { color: "#fff" },
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
    dropdown: { backgroundColor: "#141414", color: "#fff", borderRadius: 8 },
    title: {
      background: "#141414",
      color: "#fff",
      border: "1px solid #333",
      borderRadius: 8,
      height: 42,
      display: "flex",
      alignItems: "center",
      paddingLeft: 10,
    },
    caretDownWrapper: { color: "#fff" },
    dropdownItemsWrapper: { background: "#181818" },
    dropdownItem: { background: "transparent", color: "#fff" },
    dropdownItemSelected: { background: "#003b6f", color: "#fff" },
    callout: { background: "#181818" },
  };

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
    fieldGroup: { backgroundColor: "#141414", border: "1px solid #333", borderRadius: 8, height: 42 },
    field: { color: "#fff" },
  };

  // Time dropdown styles (reuse dropdownStyles but allow container width to control it)
  const timeDropdownStyles = {
    ...dropdownStyles,
    root: { width: '100%' },
  } as const;

  // Dark DatePicker styles to avoid white-on-white
  const datePickerStyles: any = {
    root: { width: 220 },
    textField: {
      fieldGroup: { backgroundColor: "#141414", border: "1px solid #333", borderRadius: 8, height: 42 },
      field: { color: "#fff", selectors: { '::placeholder': { color: '#a6b7c6', opacity: 0.8 } } },
    },
    callout: { background: "#181818" },
    // dayPicker (calendar) styles to avoid white popover
    dayPicker: { root: { background: '#181818', color: '#fff' }, monthPickerVisible: {}, showWeekNumbers: {} },
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

    const parts: string[] = [
      `To: ${EMAIL_TO}`,
  `From: Fibervsoassistant@microsoft.com`,
      `CC: ${cc || ""}`,
      `Subject: ${subject}`,
      ``,
      `----------------------------------------`,
      `CircuitIds: ${spansComma}`,
      `StartDatetime: ${startList.join(', ')}`,
      `EndDatetime: ${endList.join(', ')}`,
    ];

    parts.push(
      `NotificationType: ${notificationType}`,
      `MaintenanceReason: ${maintenanceReason}`,
      `Location: ${location}`,
    `Tags: ${tags && tags.length ? tags.join('; ') : ''}`,
      `ImpactExpected: ${impactStr}`,
    );

    return parts.map((p) => p || "").join("\n");
  }, [EMAIL_TO, subject, spansComma, startUtc, endUtc, notificationType, location, maintenanceReason, tags, impactExpected, additionalWindows, cc]);

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
          if (typeof r.current.focus === 'function') r.current.focus();
          // If the componentRef wraps the native input, try to scroll into view
          if (r.current && typeof r.current.scrollIntoView === 'function') {
            r.current.scrollIntoView({ behavior: 'smooth', block: 'center' });
          } else if (r.current && r.current.rootElement && typeof r.current.rootElement.scrollIntoView === 'function') {
            r.current.rootElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
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
        Diversity:
          (() => {
            const raw = (diversity || "").toString();
            if (!raw) return "";
            // Preserve comma-separated groups and normalize spacing
            return raw.split(',').map((s) => s.trim()).filter(Boolean).join(', ');
          })(),
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

      const resp = await fetch(EMAIL_LOGIC_APP_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
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
        {/* Facility Code A / Z - always visible. Selecting one disables the other. */}
        <div style={{ display: 'flex', gap: 12, alignItems: 'flex-start' }}>
          {!facilityCodeZ && (
            <div style={{ flex: 1 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <Text styles={{ root: { color: "#ccc", fontSize: 15, fontWeight: 500 } }}>
                  Facility Code A <span style={{ color: "#ff4d4d" }}>*</span>
                </Text>
                <TooltipHost content="You can only select one Facility Code at a time. Choosing A hides Z options.">
                  <IconButton iconProps={{ iconName: 'Info' }} title="Facility selection info" styles={{ root: { color: '#a6b7c6', height: 20, width: 20 } }} />
                </TooltipHost>
              </div>
            <ComboBox
              placeholder="Type or select Facility Code A"
              options={filteredDcOptions}
              selectedKey={facilityCodeA || undefined}
              allowFreeform={true}
              autoComplete="on"
              useComboBoxAsMenuWidth
              calloutProps={{ className: 'combo-dark-callout' }}
              componentRef={dcComboRef}
              onClick={() => dcComboRef.current?.focus(true)}
              onFocus={() => dcComboRef.current?.focus(true)}
              styles={comboBoxStyles}
              disabled={!!facilityCodeZ}
              onChange={(_, option, index, value) => {
                const typed = (value || "").toString().toLowerCase();
                const found = datacenterOptions.find((d) => {
                  const keyStr = d.key?.toString().toLowerCase();
                  const textStr = d.text?.toString().toLowerCase();
                  return textStr === typed || keyStr === typed;
                });

                if (option) {
                  const selectedKey = option.key?.toString() ?? "";
                  // selecting the blank option ("") clears the selection
                  if (selectedKey === "") {
                    setFacilityCodeA("");
                    return;
                  }
                  // Toggle off if the selected option is clicked again
                  if (selectedKey === facilityCodeA) {
                    setFacilityCodeA("");
                  } else {
                    setFacilityCodeA(selectedKey);
                    // Enforce mutual exclusivity: clear Z selections when A chosen
                    setFacilityCodeZ("");
                    setSpliceRackZ(undefined);
                  }
                } else if (found) {
                  setFacilityCodeA(found.key.toString());
                  setFacilityCodeZ("");
                  setSpliceRackZ(undefined);
                } else {
                  // reset if invalid typed text
                  setFacilityCodeA("");
                }
              }}
              onPendingValueChanged={(option, index, value) => {
                setDcSearch(value || "");
              }}
              onMenuDismiss={() => setDcSearch("")}
            />
          </div>
          )}

          {!facilityCodeA && (
          <div style={{ flex: 1 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <Text styles={{ root: { color: "#ccc", fontSize: 15, fontWeight: 500, marginTop: 0 } }}>
                Facility Code Z
              </Text>
              <TooltipHost content="You can only select one Facility Code at a time. Choosing Z hides A options.">
                <IconButton iconProps={{ iconName: 'Info' }} title="Facility selection info" styles={{ root: { color: '#a6b7c6', height: 20, width: 20 } }} />
              </TooltipHost>
            </div>
            <ComboBox
              placeholder="Type or select Facility Code Z"
              options={filteredDcOptionsZ}
              selectedKey={facilityCodeZ || undefined}
              allowFreeform={true}
              autoComplete="on"
              useComboBoxAsMenuWidth
              calloutProps={{ className: 'combo-dark-callout' }}
              componentRef={dcComboRefZ}
              onClick={() => dcComboRefZ.current?.focus(true)}
              onFocus={() => dcComboRefZ.current?.focus(true)}
              styles={comboBoxStyles}
              disabled={!!facilityCodeA}
              onChange={(_, option, index, value) => {
                const typed = (value || "").toString().toLowerCase();
                const found = datacenterOptions.find((d) => {
                  const keyStr = d.key?.toString().toLowerCase();
                  const textStr = d.text?.toString().toLowerCase();
                  return textStr === typed || keyStr === typed;
                });

                if (option) {
                  const selectedKey = option.key?.toString() ?? "";
                  if (selectedKey === "") {
                    setFacilityCodeZ("");
                    return;
                  }
                  if (selectedKey === facilityCodeZ) {
                    setFacilityCodeZ("");
                  } else {
                    setFacilityCodeZ(selectedKey);
                    // Enforce mutual exclusivity: clear A selections when Z chosen
                    setFacilityCodeA("");
                    setSpliceRackA(undefined);
                  }
                } else if (found) {
                  setFacilityCodeZ(found.key.toString());
                  setFacilityCodeA("");
                  setSpliceRackA(undefined);
                } else {
                  setFacilityCodeZ("");
                }
              }}
              onPendingValueChanged={(option, index, value) => {
                setDcSearchZ(value || "");
              }}
              onMenuDismiss={() => setDcSearchZ("")}
            />
            {(facilityCodeA && facilityCodeZ) && (
              <MessageBar messageBarType={MessageBarType.warning} isMultiline={false}>
                Only one Facility Code may be selected at a time.
              </MessageBar>
            )}
          </div>
          )}
        </div>

        {/* === Diversity Dropdown === */}
        <Text styles={{ root: { color: "#ccc", fontSize: 15, fontWeight: 500, marginTop: 10 } }}>
          Diversity Path (Optional)
        </Text>
        <Dropdown
          placeholder=""
          options={diversityOptions}
          calloutProps={{ className: 'combo-dark-callout' }}
          styles={diversityDropdownStyles}
          selectedKey={diversity === undefined || diversity === "" ? undefined : diversity}
          onChange={(_, option) => {
            if (!option) return;
            const nextKey = option.key?.toString() ?? "";
            // Selecting the blank option always clears the selection
            if (nextKey === "") {
              setDiversity(undefined);
              return;
            }
            // Toggle off if the same diversity option is clicked
            if ((diversity || "") === nextKey) {
              setDiversity(undefined);
            } else {
              setDiversity(nextKey);
            }
          }}
        />

        {/* Removed duplicate top-level Decommissioned-by-Country control; use the Decommissioned collapsed panel button instead. */}

        {/* Splice Rack A/Z - conditionally shown; selecting one hides the other */}
        <div style={{ display: 'flex', gap: 12, marginTop: 8 }}>
          {!facilityCodeZ && (
            <div style={{ flex: 1 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <Text styles={{ root: { color: "#ccc", fontSize: 13, fontWeight: 500 } }}>Splice Rack A (Optional)</Text>
                <TooltipHost content="Only one Splice Rack may be selected at a time. Choosing A hides Z.">
                  <IconButton iconProps={{ iconName: 'Info' }} title="Splice Rack selection info" styles={{ root: { color: '#a6b7c6', height: 18, width: 18 } }} />
                </TooltipHost>
              </div>
              <TextField
                placeholder="e.g. AM111"
                onChange={(_, value) => {
                  const v = value || undefined;
                  setSpliceRackA(v);
                  if (v) {
                    // enforce exclusivity: clear Z-side selections when A splice entered
                    setSpliceRackZ(undefined);
                    setFacilityCodeZ("");
                  }
                }}
                styles={textFieldStyles}
                disabled={!!spliceRackZ}
              />
            </div>
          )}

          {!facilityCodeA && (
            <div style={{ flex: 1 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <Text styles={{ root: { color: "#ccc", fontSize: 13, fontWeight: 500 } }}>Splice Rack Z (Optional)</Text>
                <TooltipHost content="Only one Splice Rack may be selected at a time. Choosing Z hides A.">
                  <IconButton iconProps={{ iconName: 'Info' }} title="Splice Rack selection info" styles={{ root: { color: '#a6b7c6', height: 18, width: 18 } }} />
                </TooltipHost>
              </div>
              <TextField
                placeholder="e.g. AJ1508"
                onChange={(_, value) => {
                  const v = value || undefined;
                  setSpliceRackZ(v);
                  if (v) {
                    setSpliceRackA(undefined);
                    setFacilityCodeA("");
                  }
                }}
                styles={textFieldStyles}
                disabled={!!spliceRackA}
              />
            </div>
          )}
        </div>
        {(spliceRackA && spliceRackZ) && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline={false}>
            Only one Splice Rack may be selected at a time.
          </MessageBar>
        )}

        <div className="form-buttons" style={{ marginTop: 16 }}>
            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <div style={{ display: 'flex', gap: 8 }}>
                <button className="submit-btn" onClick={() => handleSubmit()}>
                  Submit
                </button>
                {searchDone && (
                  <button
                    className="sleek-btn danger"
                    onClick={() => { resetAll(); setCountryPanelOpen(false); setIsDecomMode(false); }}
                    title="Reset search and form"
                    aria-label="Reset"
                    style={{ minWidth: 96 }}
                  >
                    Reset
                  </button>
                )}
              </div>

              {/* Right-aligned collapsed country panel toggle and content */}
              <div style={{ marginLeft: 12, display: 'flex', alignItems: 'center' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <DefaultButton
                    text={countryPanelOpen ? 'Hide' : 'Decommissioned Spans'}
                    onClick={() => setCountryPanelOpen(o => !o)}
                    styles={{ root: { height: 36, backgroundColor: '#f1c232', color: '#000', borderRadius: 6, border: '1px solid #b38f16' } }}
                  />
                </div>
                {countryPanelOpen && (
                  <div style={{ display: 'flex', gap: 8, marginLeft: 8, alignItems: 'center', width: 520 }}>
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
                        styles={{ ...comboBoxStyles, root: { width: 340 } } as any}
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
                    <div style={{ width: 100, display: 'flex', alignItems: 'center' }}>
                      <PrimaryButton
                        text="Lookup"
                        disabled={!country}
                        styles={{ root: { height: 36 } }}
                        onClick={async () => {
                          setLoading(true);
                          setError(null);
                          setResult([]);
                          setIsDecomMode(true);
                          setShowAll(true); // show all decom'd spans
                          try {
                            const logicAppUrl = "https://fibertools-dsavavdcfdgnh2cm.westeurope-01.azurewebsites.net:443/api/VSO/triggers/When_an_HTTP_request_is_received/invoke?api-version=2022-05-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=6ViXNM-TmW5F7Qd9_e4fz3IhRNqmNzKwovWvcmuNJto";
                            const payload = { Stage: "10", Country: country };
                            const resp = await fetch(logicAppUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
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
                )}
              </div>
            </div>
        </div>

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
              {/* Left: spans summary */}
              <div style={{ flex: 1, display: 'flex', justifyContent: 'flex-start' }}>
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
                              {/* divider removed to tighten layout */}
                              <div style={{ textAlign: 'right' }}>
                                <div style={{ fontSize: 18, fontWeight: 700, color: '#8fe3a3' }}>{productionCount}</div>
                                <div style={{ fontSize: 11, color: '#9fb3c6' }}>In Production</div>
                              </div>
                      </div>
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
                  const leadingUniqueSorted = [...leadingUnique].sort((a, b) => a.localeCompare(b));
                  const leadingOptions = leadingUniqueSorted.map((c) => ({ key: c, text: c }));
                  const remaining = availableDcOptions
                    .filter((o) => !seen.has(String(o.key)))
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
              <div style={{ flex: 1, display: 'flex', justifyContent: 'flex-end', gap: 12, alignItems: 'center' }}>
                {!isDecomMode && (
                  <button className="rack-btn slim" onClick={() => setShowAll(!showAll)}>
                    {showAll ? "Show Only Production" : "Show All Spans"}
                  </button>
                )}
                <div style={{ display: 'flex', alignItems: 'center' }}>
                  <Toggle
                    label="Detailed view"
                    inlineLabel
                    onText="On"
                    offText="Off"
                    // Toggle reflects Detailed view state; invert simplifiedView for checked
                    checked={!simplifiedView}
                    onChange={(_, v) => setSimplifiedView(!(!!v))}
                    // Reduce space between the label text and the toggle control
                    styles={{ root: { display: 'flex', alignItems: 'center' }, label: { marginRight: 6 } }}
                  />
                </div>
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
                  <a
                    href={row.OpticalLink}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="uid-click"
                    onClick={(e) => e.stopPropagation()}
                    style={{ fontSize: 14, fontWeight: 600 }}
                  >
                    {row.SpanID}
                  </a>
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
              // Prefer the unit-level splice column when the user searched by Facility Code A
              // Logic app returns a SpliceRackA_Unit field; show that when Facility Code A is used.
              const splicePreferredKey = facilityCodeA
                ? 'SpliceRackA_Unit'
                : spliceRackA
                ? 'SpliceRackA'
                : 'SpliceRackZ';

              const getCandidate = (k: string) => candidate.find(c => c.key === k);

              let dynamicCols: any[];
              if (simplifiedView) {
                // Include FacilityCodeA and FacilityCodeZ in simplified view per request
                const simplifiedKeys = ['Diversity', 'SpanID', 'FacilityCodeA', 'FacilityCodeZ', splicePreferredKey, 'WiringScope', 'Status'];
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
                <table ref={tableRef} className="data-table compact" style={{ fontSize: 11, tableLayout: 'fixed', width: '100%', borderCollapse: 'collapse', marginLeft: -6 }}>
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
                          onClick={() => toggleSelectSpan(row.SpanID)}
                          style={{ cursor: 'pointer', background: bg }}
                        >
                          <td style={{ padding: '2px 4px', width: 38, display: 'flex', alignItems: 'flex-end', justifyContent: 'center' }}>
                            <Checkbox
                              checked={isSelected}
                              onChange={() => toggleSelectSpan(row.SpanID)}
                              styles={{ root: { margin: 0, padding: 0, display: 'block' } } as any}
                            />
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
              <PrimaryButton
                text={`Continue (${selectedSpans.length} selected)`}
                disabled={selectedSpans.length === 0}
                onClick={() => setComposeOpen(true)}
              />
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
                Please complete the required fields: {Object.keys(fieldErrors).map(k => friendlyFieldNames[k] || k).join(', ')}
              </MessageBar>
            )}

            <Dialog
              hidden={!showEmergencyDialog}
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
              hidden={!showSendSuccessDialog}
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
                      styles={textFieldStyles}
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
                      styles={textFieldStyles}
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
                        <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>Start Date</Text>
                        <TooltipHost content={"Times shown are in your local time and will be converted automatically when the VSO is created."}>
                          <IconButton iconProps={{ iconName: 'Info' }} title="Time conversion info" styles={{ root: { color: '#a6b7c6', height: 18, width: 18 } }} />
                        </TooltipHost>
                      </div>
                      <DatePicker
                        componentRef={startDateRef}
                        placeholder="Select start date"
                        value={startDate || undefined}
                        onSelectDate={(d) => {
                          const selected = d || null;
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
                        styles={datePickerStyles}
                        isRequired
                        aria-label="Start Date"
                      />
                      {showValidation && fieldErrors.startDate ? (
                        <Text styles={{ root: { color: '#a80000', fontSize: 12 } }}>Required</Text>
                      ) : null}
                    </div>

                    <div className="dt-time">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>Start Time</Text>
                      <Dropdown
                        componentRef={startTimeRef}
                        options={timeOptions}
                        selectedKey={startTime}
                        onChange={(_, opt) => opt && setStartTime(opt.key.toString())}
                        styles={timeDropdownStyles}
                        required
                        errorMessage={showValidation ? fieldErrors.startTime : undefined}
                      />
                    </div>

                    <div className="dt-field">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>End Date</Text>
                      <DatePicker
                        componentRef={endDateRef}
                        placeholder="Select end date"
                        value={endDate || undefined}
                        onSelectDate={(d) => setEndDate(d || null)}
                        styles={datePickerStyles}
                        isRequired
                        aria-label="End Date"
                      />
                      {showValidation && fieldErrors.endDate ? (
                        <Text styles={{ root: { color: '#a80000', fontSize: 12 } }}>Required</Text>
                      ) : null}
                    </div>

                    <div className="dt-time">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>End Time</Text>
                      <Dropdown
                        componentRef={endTimeRef}
                        options={timeOptions}
                        selectedKey={endTime}
                        onChange={(_, opt) => opt && setEndTime(opt.key.toString())}
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
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>Start Date</Text>
                      <DatePicker
                        placeholder="Select start date"
                        value={w.startDate || undefined}
                        onSelectDate={(d) => {
                          const selected = d || null;
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
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>Start Time</Text>
                      <Dropdown
                        options={timeOptions}
                        selectedKey={w.startTime}
                        onChange={(_, opt) => opt && setAdditionalWindows((arr) => {
                          const next = [...arr];
                          next[i] = { ...next[i], startTime: opt.key.toString() };
                          return next;
                        })}
                        styles={timeDropdownStyles}
                      />
                    </div>
                    <div className="dt-field">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>End Date</Text>
                      <DatePicker
                        placeholder="Select end date"
                        value={w.endDate || undefined}
                        onSelectDate={(d) => {
                          setAdditionalWindows((arr) => {
                            const next = [...arr];
                            next[i] = { ...next[i], endDate: d || null };
                            return next;
                          });
                        }}
                        styles={datePickerStyles}
                      />
                    </div>
                    <div className="dt-time">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>End Time</Text>
                      <Dropdown
                        options={timeOptions}
                        selectedKey={w.endTime}
                        onChange={(_, opt) => opt && setAdditionalWindows((arr) => {
                          const next = [...arr];
                          next[i] = { ...next[i], endTime: opt.key.toString() };
                          return next;
                        })}
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
                      <Text styles={{ root: { color: '#ccc', fontSize: 14, fontWeight: 600, marginBottom: 6 } }}>Tags</Text>
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
                          // match the theme and ensure selected/hovered items keep readable text
                          ...dropdownStyles,
                          root: { width: 200 },
                          dropdownItem: { background: 'transparent', color: '#fff', selectors: { ':hover': { background: '#004b8a', color: '#fff' } } },
                          dropdownItemSelected: { background: '#004b8a', color: '#fff', selectors: { ':hover': { background: '#003b6f', color: '#fff' } } },
                          callout: { background: '#181818' },
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
                      styles={dropdownStyles}
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
                    ...textFieldStyles,
                    field: { ...(textFieldStyles as any).field, minHeight: 220, paddingBottom: 28 },
                    fieldGroup: { ...(textFieldStyles as any).fieldGroup, height: 'auto' },
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

            <div className="section-title" style={{ marginTop: 6 }}>Email Body Preview</div>
            <div style={{ background: "#0f0f0f", border: "1px solid #333", borderRadius: 8, padding: 12, whiteSpace: "pre-wrap", color: "#dfefff" }}>
              {emailBody}
            </div>

            <div style={{ display: "flex", justifyContent: "space-between", alignItems: 'center', marginTop: 12 }}>
              <button className="sleek-btn danger" onClick={() => setComposeOpen(false)}>
                Back
              </button>
              <button className="sleek-btn wan" disabled={!canSend || sendLoading} onClick={handleSend}>
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
                <div style={{ flex: 2 }}>
                  <Text styles={{ root: { color: '#a6b7c6', fontSize: 12 } }}>Title</Text>
                  <div style={{ color: '#e6f6ff' }}>{ev.title}</div>
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
                  <div style={{ padding: 8, color: '#e6f6ff' }}>{(ev.spans || []).join(', ') || '--'}</div>
                </div>
              </div>

              <div className="table-container details-fit">
                <div className="section-title">Reason</div>
                <div style={{ padding: 8, color: '#dfefff', whiteSpace: 'pre-wrap' }}>{ev.maintenanceReason || ev.summary || '--'}</div>
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
