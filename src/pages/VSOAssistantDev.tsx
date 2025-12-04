import React, { useEffect, useMemo, useRef, useState } from "react";
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
} from "@fluentui/react";
import "../Theme.css";
import datacenterOptions from "../data/datacenterOptions";
import { getRackElevationUrl } from "../data/MappedREs";
import VSOCalendar, { VsoCalendarEvent } from "../components/VSOCalendar";
import { getCalendarEntries } from "../api/items";
import { API_BASE } from "../api/config";
import { logAction } from "../api/log";

interface SpanData {
  SpanID: string;
  Diversity: string;
  IDF_A: string;
  SpliceRackA: string;
  WiringScope: string;
  Status: string;
  Color: string;
  OpticalLink?: string;
}

interface BackendResponse {
  Spans: SpanData[];
  DataCenter?: string;
}
interface MaintenanceWindow {
  startDate: Date | null;
  startTime: string;
  endDate: Date | null;
  endTime: string;
}

const VSOAssistantDev: React.FC = () => {
  // Fluent UI icons are initialized once in `src/index.tsx`

  // Search state
  const [facilityCodeA, setFacilityCodeA] = useState<string>("");
  const [diversity, setDiversity] = useState<string | undefined>(undefined);
  const [spliceRackA, setSpliceRackA] = useState<string | undefined>(undefined);
  const [loading, setLoading] = useState<boolean>(false);
  const [result, setResult] = useState<SpanData[]>([]);
  const [selectedSpans, setSelectedSpans] = useState<string[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [showAll, setShowAll] = useState<boolean>(false);
  const [rackUrl, setRackUrl] = useState<string | undefined>(undefined);
  const [rackDC, setRackDC] = useState<string | undefined>(undefined);
  const [dcSearch, setDcSearch] = useState<string>("");
  const dcComboRef = useRef<IComboBox | null>(null);
  const [searchDone, setSearchDone] = useState<boolean>(false);

  // Compose state
  const [composeOpen, setComposeOpen] = useState<boolean>(false);
  const [subject, setSubject] = useState<string>("");
  const [notificationType, setNotificationType] = useState<string>("New Maintenance Scheduled");
  const [location, setLocation] = useState<string>("");
  const [lat, setLat] = useState<string>("");
  const [lng, setLng] = useState<string>("");
  const [isp, setIsp] = useState<string>("");
  const [ispTicket, setIspTicket] = useState<string>("");
  const [maintenanceReason, setMaintenanceReason] = useState<string>("");
  const [impactExpected, setImpactExpected] = useState<boolean>(true);
  const [startDate, setStartDate] = useState<Date | null>(null);
  const [endDate, setEndDate] = useState<Date | null>(null);
  const [startTime, setStartTime] = useState<string>("00:00");
  const [endTime, setEndTime] = useState<string>("00:00");
  const [additionalWindows, setAdditionalWindows] = useState<MaintenanceWindow[]>([]);
  const [userEmail, setUserEmail] = useState<string>("");
  const [cc, setCc] = useState<string>("");
  const [sendLoading, setSendLoading] = useState<boolean>(false);
  const [sendSuccess, setSendSuccess] = useState<string | null>(null);
  const [sendError, setSendError] = useState<string | null>(null);
  const [showSendSuccessDialog, setShowSendSuccessDialog] = useState<boolean>(false);

  // Calendar state
  const [vsoEvents, setVsoEvents] = useState<VsoCalendarEvent[]>([]);
  const [calendarDate, setCalendarDate] = useState<Date | null>(null);

  useEffect(() => {
    const email = (() => {
      try {
        return localStorage.getItem("loggedInEmail") || "";
      } catch {
        return "";
      }
    })();
    logAction(email, "View VSO Assistant Dev");
  }, []);

  // Load persisted calendar entries from server so everyone sees the same calendar
  useEffect(() => {
    let mounted = true;
    let timer: any = null;

    const mapItems = (items: any[]): VsoCalendarEvent[] => (items || []).map((it: any) => {
      const title = it.title || it.Title || '';
      const desc = it.description || it.Description || '';
      const savedAt = it.savedAt || it.SavedAt || it.rowKey || it.RowKey || null;
      const startMatch = /Start:\s*([\dTZ:+.\u002D]+)\b/i.exec(desc);
      const endMatch = /End:\s*([\dTZ:+.\u002D]+)\b/i.exec(desc);
      const parseDate = (s: string | null) => { try { return s ? new Date(s) : null; } catch { return null; } };
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

    loadOnce();
    timer = setInterval(loadOnce, 30_000);
    return () => { mounted = false; if (timer) clearInterval(timer); };
  }, []);

  // Types/validation
  const [fieldErrors, setFieldErrors] = useState<Record<string, string>>({});
  const [showValidation, setShowValidation] = useState<boolean>(false);

  // Options & styles (copy from existing page)
  const diversityOptions: IDropdownOption[] = [
    { key: "", text: "" },
    { key: "West", text: "West, West 1, West 2" },
    { key: "East", text: "East, East 1, East 2" },
    { key: "North", text: "North" },
    { key: "South", text: "South" },
    { key: "Y", text: "Y" },
    { key: "Z", text: "Z" },
  ];
  const comboBoxStyles = {
    root: { width: "100%" },
    input: { color: "#fff", backgroundColor: "#141414", height: 42, border: "1px solid #333", borderRadius: 8, paddingLeft: 10 },
    callout: { background: "#181818", maxHeight: 240, overflowY: "auto" },
    optionsContainer: { background: "#181818" },
    caretDownWrapper: { color: "#fff" },
  } as const;
  const dropdownStyles = {
    root: { width: "100%" },
    dropdown: { backgroundColor: "#141414", color: "#fff", borderRadius: 8 },
    title: { background: "#141414", color: "#fff", border: "1px solid #333", borderRadius: 8, height: 42, display: "flex", alignItems: "center", paddingLeft: 10 },
    caretDownWrapper: { color: "#fff" },
    dropdownItemsWrapper: { background: "#181818" },
    dropdownItem: { background: "transparent", color: "#fff" },
    dropdownItemSelected: { background: "#003b6f", color: "#fff" },
    callout: { background: "#181818" },
  };
  const diversityDropdownStyles = { ...dropdownStyles, titleIsPlaceHolder: { ...dropdownStyles.title, color: "#a6b7c6", opacity: 0.8, fontStyle: "normal", fontSize: 14, fontWeight: 400 } } as const;
  const textFieldStyles = { fieldGroup: { backgroundColor: "#141414", border: "1px solid #333", borderRadius: 8, height: 42 }, field: { color: "#fff" } } as const;
  const timeDropdownStyles = { ...dropdownStyles, root: { width: "100%" } } as const;
  const datePickerStyles: any = {
    root: { width: 220 },
    textField: { fieldGroup: { backgroundColor: "#141414", border: "1px solid #333", borderRadius: 8, height: 42 }, field: { color: "#fff", selectors: { "::placeholder": { color: "#a6b7c6", opacity: 0.8 } } } },
    callout: { background: "#181818" },
    dayPicker: { root: { background: "#181818", color: "#fff" } },
  };

  // Derived state/helpers
  const filteredDcOptions: IComboBoxOption[] = useMemo(() => {
    const base = datacenterOptions.map((d) => ({ key: d.key, text: d.text }));
    const search = dcSearch.toLowerCase().trim();
    return !search
      ? base
      : base.filter((opt) => opt.key.toString().toLowerCase().includes(search) || opt.text.toString().toLowerCase().includes(search));
  }, [dcSearch]);

  const filteredResults = showAll ? result : result.filter((r) => {
    const state = (((r as any).State || '') as string).toLowerCase();
    if (state === 'new') return false;
    return ((r.Status || '') as string).toLowerCase() === 'inproduction';
  });
  const [sortBy, setSortBy] = useState<string>("");
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");
  const getSortValue = (row: SpanData, key: string): string | number => {
    switch (key) {
      case "diversity": return row.Diversity || "";
      case "span": return row.SpanID || "";
      case "idf": return row.IDF_A || "";
      case "splice": return row.SpliceRackA || "";
      case "scope": return row.WiringScope || "";
      case "status": return row.Status || "";
      default: return "";
    }
  };
  const sortedResults = useMemo(() => {
    const rows = [...filteredResults];
    if (!sortBy) return rows;
    rows.sort((a, b) => {
      const av = getSortValue(a, sortBy);
      const bv = getSortValue(b, sortBy);
      const an = typeof av === "string" ? Number(av) : (av as number);
      const bn = typeof bv === "string" ? Number(bv) : (bv as number);
      const aIsNum = !isNaN(an as number) && an !== (undefined as any) && an !== null && av !== "";
      const bIsNum = !isNaN(bn as number) && bn !== (undefined as any) && bn !== null && bv !== "";
      if (aIsNum && bIsNum) return (an as number) - (bn as number);
      const as = (av ?? "").toString().toLowerCase();
      const bs = (bv ?? "").toString().toLowerCase();
      return as.localeCompare(bs);
    });
    if (sortDir === "desc") rows.reverse();
    return rows;
  }, [filteredResults, sortBy, sortDir]);
  const handleSort = (key: string) => { if (sortBy === key) setSortDir((d) => (d === "asc" ? "desc" : "asc")); else { setSortBy(key); setSortDir("asc"); } };

  // Helpers
  const toggleSelectSpan = (spanId: string) => setSelectedSpans((prev) => (prev.includes(spanId) ? prev.filter((id) => id !== spanId) : [...prev, spanId]));
  const spansComma = useMemo(() => selectedSpans.join(","), [selectedSpans]);
  const formatUtcString = (date: Date | null, time: string) => {
    if (!date) return "";
    const [hh, mm] = time.split(":").map((s) => parseInt(s, 10));
    const y = date.getFullYear();
    const m = (date.getMonth() + 1).toString().padStart(2, "0");
    const d = date.getDate().toString().padStart(2, "0");
    const H = (isNaN(hh) ? 0 : hh).toString().padStart(2, "0");
    const M = (isNaN(mm) ? 0 : mm).toString().padStart(2, "0");
    return `${m}/${d}/${y} ${H}:${M}`;
  };

  // GET /.auth/me to prefill user email for CC
  useEffect(() => {
    try { const stored = localStorage.getItem("loggedInEmail"); if (stored) setUserEmail(stored); } catch {}
    const fetchUserEmail = async () => {
      try {
        const res = await fetch("/.auth/me", { credentials: "include" });
        if (!res.ok) return;
        const data = await res.json();
        const identities = Array.isArray(data) ? data : data?.clientPrincipal ? [{ user_claims: data.clientPrincipal?.claims || [] }] : [];
        for (const id of identities) {
          const claims = (id as any)?.user_claims || [];
          const getClaim = (t: string) => claims.find((c: any) => c?.typ === t)?.val || "";
          const email = getClaim("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress") || getClaim("emails") || getClaim("preferred_username") || getClaim("upn") || "";
          if (email) { setUserEmail(email); try { localStorage.setItem("loggedInEmail", email); } catch {} return; }
        }
      } catch {}
    };
    fetchUserEmail();
  }, []);

  // Stage 1: search spans via backend
  const handleSubmit = async () => {
    const email = (() => {
      try {
        return localStorage.getItem("loggedInEmail") || "";
      } catch {
        return "";
      }
    })();
    logAction(email, "Submit VSO Dev Search", {
      facilityCodeA,
      diversity,
      spliceRackA,
    });

    if (!facilityCodeA) { alert("Please select a valid Data Center first."); return; }
    setLoading(true); setError(null); setResult([]); setSearchDone(false);
    try {
      const payload = {
        FacilityCodeA: facilityCodeA,
        Diversity: (() => { const raw = (diversity || "").toString(); if (!raw) return "N"; const parsed = raw.split(",")[0].trim().replace(/,$/, ""); return parsed.replace(/[^A-Za-z0-9 ]/g, "").trim() || "N"; })(),
        SpliceRackA: (spliceRackA && spliceRackA.trim()) ? spliceRackA.trim() : "N",
        Stage: "VSO_Details",
      } as const;
      const response = await fetch(`${API_BASE}/vso2`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
      const raw = await response.text();
      let data: BackendResponse;
      try {
        data = JSON.parse(raw);
      } catch (e) {
        throw new Error(`Backend did not return JSON (HTTP ${response.status}). Body: ${raw.slice(0, 240)}`);
      }
      if (!response.ok) throw new Error(((data as any)?.error || `HTTP ${response.status}`) + (raw ? ` - ${raw.slice(0, 120)}` : ""));
  // Coerce Spans into an array so UI logic can always treat `result` as an array
  const spans = Array.isArray(data?.Spans) ? data.Spans : (data?.Spans ? [data.Spans] : []);
  if (data?.Spans && !Array.isArray(data.Spans)) console.warn('Dev Logic App returned non-array Spans, coercing to array', data.Spans);
  setResult(spans as any);
      const dc = data?.DataCenter || facilityCodeA;
      setRackDC(dc);
      setRackUrl(getRackElevationUrl(dc));
    } catch (err: any) {
      setError(err?.message || "Unknown error occurred.");
    } finally { setLoading(false); setSearchDone(true); }
  };

  // Stage 2: send email via backend (currently not implemented server-side)
  const validateCompose = (): boolean => {
    const errors: Record<string, string> = {};
    if (!startDate) errors.startDate = "Required";
    if (!startTime) errors.startTime = "Required";
    if (!endDate) errors.endDate = "Required";
    if (!endTime) errors.endTime = "Required";
    if (!location?.trim()) errors.location = "Required";
    if (!isp?.trim()) errors.isp = "Required";
    if (!ispTicket?.trim()) errors.ispTicket = "Required";
    if (!(impactExpected === true || impactExpected === false)) errors.impactExpected = "Required";
    if (!cc?.trim()) errors.cc = "Required";
    setFieldErrors(errors);
    return Object.keys(errors).length === 0;
  };

  const handleSend = async () => {
    const email = (() => {
      try {
        return localStorage.getItem("loggedInEmail") || "";
      } catch {
        return "";
      }
    })();
    logAction(email, "Send VSO Dev Maintenance Email", {
      facilityCodeA,
      spans: spansComma,
      notificationType,
    });

    setShowValidation(true);
    if (!validateCompose()) { window.scrollTo({ top: 0, behavior: 'smooth' }); return; }
    setSendError(null); setSendSuccess(null); setSendLoading(true);
    try {
      const startList: string[] = []; const endList: string[] = [];
      const startPrimary = formatUtcString(startDate, startTime); if (startPrimary) startList.push(startPrimary);
      const endPrimary = formatUtcString(endDate, endTime); if (endPrimary) endList.push(endPrimary);
      additionalWindows.forEach((w) => {
        const s = formatUtcString(w.startDate, w.startTime); const e = formatUtcString(w.endDate, w.endTime);
        if (s) startList.push(s); if (e) endList.push(e);
      });
      const payload = {
        FacilityCodeA: facilityCodeA || "",
        Diversity: (diversity || "").toString(),
        SpliceRackA: spliceRackA || "",
        Stage: "Email_Template",
        CC: cc || "",
        Subject: subject || "",
        CircuitIds: spansComma || "",
        StartDatetime: startList.join(', '),
        EndDatetime: endList.join(', '),
        LatLong: (lat && lng) ? `${lat},${lng}` : "",
        NotificationType: notificationType || "",
        MaintenanceReason: maintenanceReason || "",
        Location: location || "",
        ISP: isp || "",
        ISPTicket: ispTicket || "",
        ImpactExpected: impactExpected ? "True" : "False",
      } as const;
      const resp = await fetch(`${API_BASE}/vso2`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
      const text = await resp.text();
      if (resp.status === 501) {
        setSendSuccess("Backend email stage is not implemented yet. This verified the UI up to send; calendar will be updated locally.");
        setShowSendSuccessDialog(true);
      } else if (!resp.ok) {
        throw new Error(text || `HTTP ${resp.status}`);
      } else {
        setSendSuccess("Email request submitted to backend.");
        setShowSendSuccessDialog(true);
      }
      // Add events locally so user can test calendar flow
      const dcCode = rackDC || facilityCodeA;
      const makeAllDayRange = (d: Date | null) => { if (!d) return null; const start = new Date(d.getFullYear(), d.getMonth(), d.getDate()); const end = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1); return { start, end }; };
      const newEvents: VsoCalendarEvent[] = [];
      const spansShort = (() => selectedSpans.length <= 3 ? selectedSpans.join(", ") : `${selectedSpans.slice(0,3).join(", ")} (+${selectedSpans.length-3} more)`)();
      const title = `Fiber Maintenance ${dcCode || ""} - ${spansShort || "Spans"}`.trim();
      const primary = makeAllDayRange(startDate);
      if (primary) newEvents.push({ id: `vso-${Date.now()}-0-${Math.random().toString(36).slice(2,6)}`, title, start: primary.start, end: primary.end, status: "Draft", summary: maintenanceReason.slice(0,160), maintenanceReason, dcCode: dcCode || undefined, spans: [...selectedSpans], startTimeUtc: startTime, endTimeUtc: endTime, subject, notificationType, location, isp, ispTicket, impactExpected });
      additionalWindows.forEach((w, i) => { const r = makeAllDayRange(w.startDate); if (!r) return; newEvents.push({ id: `vso-${Date.now()}-${i+1}-${Math.random().toString(36).slice(2,6)}`, title, start: r.start, end: r.end, status: "Draft", summary: maintenanceReason.slice(0,160), maintenanceReason, dcCode: dcCode || undefined, spans: [...selectedSpans], startTimeUtc: w.startTime, endTimeUtc: w.endTime, subject, notificationType, location, isp, ispTicket, impactExpected }); });
      setVsoEvents((prev) => [...prev, ...newEvents]);
      if (!calendarDate && (startDate || additionalWindows[0]?.startDate)) { const d = startDate || additionalWindows[0]?.startDate || null; if (d) setCalendarDate(new Date(d.getFullYear(), d.getMonth(), 1)); }
    } catch (e: any) {
      setSendError(e?.message || "Failed to send email.");
    } finally { setSendLoading(false); }
  };

  // UI helpers copied from original file
  const getStatusClass = (status?: string) => {
    const t = (status || "").toLowerCase();
    if (t.includes("inproduction") || t === "in production" || t === "production") return "good";
    if (t.includes("decom") || t.includes("retired") || t.includes("outofservice") || t.includes("out of service") || t.includes("warning")) return "warning";
    return "accent";
  };
  const getDiversityClass = (div?: string) => {
    const t = (div || "").toLowerCase().trim();
    if (t.includes("east 1")) return "accent";
    if (t.includes("east 2")) return "good";
    if (t === "south") return "accent";
    if (t === "y") return "accent";
    if (t.includes("west 1")) return "danger";
    if (t.includes("west 2")) return "warning";
    if (t === "north") return "danger";
    if (t === "z") return "danger";
    if (t.startsWith("east")) return "accent";
    if (t.startsWith("west")) return "danger";
    return "accent";
  };

  // Build time options
  const timeOptions: IDropdownOption[] = useMemo(() => { const opts: IDropdownOption[] = []; for (let h=0; h<24; h++){ for (let m=0; m<60; m+=30){ const hh = h.toString().padStart(2,"0"); const mm = m.toString().padStart(2,"0"); const text = `${hh}:${mm}`; opts.push({ key: text, text }); } } return opts; }, []);

  // Prefill subject when entering compose
  useEffect(() => { if (!composeOpen) return; if (!subject || subject.trim().length===0){ const region = rackDC || facilityCodeA || "Region"; setSubject(`[${region}] Maintenance scheduled in <${region}> Contractor:  Lead Engineer:`); } }, [composeOpen, subject, rackDC, facilityCodeA]);

  // Keep CC in sync with detected user
  useEffect(() => { if (composeOpen && !cc && userEmail) setCc(userEmail); }, [composeOpen, userEmail, cc]);

  return (
    <div className="main-content fade-in">
      <div className="vso-form-container glow" style={{ width: "80%", maxWidth: 1000 }}>
        <div className="banner-title">
          <span className="title-text">Fiber VSO Assistant (Backend)</span>
          <span className="title-sub">Search uses the new Azure Functions endpoint.</span>
        </div>

        {!composeOpen && (
          <>
            <Text styles={{ root: { color: "#ccc", fontSize: 15, fontWeight: 500 } }}>Data Center <span style={{ color: "#ff4d4d" }}>*</span></Text>
            <ComboBox
              placeholder="Type or select a Data Center"
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
              onChange={(_, option, index, value) => {
                const typed = (value || "").toString().toLowerCase();
                const found = datacenterOptions.find((d) => {
                  const keyStr = d.key?.toString().toLowerCase();
                  const textStr = d.text?.toString().toLowerCase();
                  return textStr === typed || keyStr === typed;
                });
                if (option) {
                  const selectedKey = option.key?.toString() ?? "";
                  if (selectedKey === facilityCodeA) setFacilityCodeA(""); else setFacilityCodeA(selectedKey);
                } else if (found) {
                  setFacilityCodeA(found.key.toString());
                } else {
                  setFacilityCodeA("");
                }
              }}
              onPendingValueChanged={(option, index, value) => setDcSearch(value || "")}
              onMenuDismiss={() => setDcSearch("")}
            />

            <Text styles={{ root: { color: "#ccc", fontSize: 15, fontWeight: 500, marginTop: 10 } }}>Diversity Path <span className="optional-text">(Optional)</span></Text>
            <Dropdown
              placeholder=""
              options={diversityOptions}
              calloutProps={{ className: 'combo-dark-callout' }}
              styles={diversityDropdownStyles}
              selectedKey={diversity === undefined || diversity === "" ? undefined : diversity}
              onChange={(_, option) => {
                if (!option) return;
                const nextKey = option.key?.toString() ?? "";
                if (nextKey === "") { setDiversity(undefined); return; }
                if ((diversity || "") === nextKey) setDiversity(undefined); else setDiversity(nextKey);
              }}
            />

            <Text styles={{ root: { color: "#ccc", fontSize: 15, fontWeight: 500 } }}>Splice Rack A <span className="optional-text">(Optional)</span></Text>
            <TextField placeholder="e.g. AM111" onChange={(_, v) => setSpliceRackA(v || undefined)} styles={textFieldStyles} />

            <div className="form-buttons" style={{ marginTop: 16 }}>
              <button className="submit-btn" onClick={handleSubmit}>Submit</button>
            </div>

            {loading && <Spinner label="Loading results..." size={SpinnerSize.medium} styles={{ root: { marginTop: 15 } }} />}
            {error && (<MessageBar messageBarType={MessageBarType.error} isMultiline={false}>{error}</MessageBar>)}
            {!loading && !error && searchDone && result.length === 0 && (
              <div className="notice-banner warning"><span className="banner-icon">!</span><div className="banner-text">There were no results for the selections you made. Try adjusting your search.</div></div>
            )}

            {result.length > 0 && (
              <div className="table-container" style={{ marginTop: 14 }}>
                <div style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", marginBottom: 8 }}>
                  {(() => {
                    const resolvedDC = rackDC || facilityCodeA;
                    const href = rackUrl || getRackElevationUrl(resolvedDC);
                    return href ? (
                      <div style={{ textAlign: "center", flex: 1 }}>
                        <a href={href} target="_blank" rel="noopener noreferrer" className="rack-btn slim" style={{ display: "inline-block", minWidth: 320 }}>
                          {`View Rack Elevation - ${resolvedDC}`}
                        </a>
                      </div>
                    ) : null;
                  })()}
                  <div style={{ textAlign: "right" }}>
                    <button className="sleek-btn optical" onClick={() => setShowAll(!showAll)}>{showAll ? "Show Only Production" : "Show All Spans"}</button>
                  </div>
                </div>

                <table className="data-table compact">
                  <thead>
                    <tr>
                      <th></th>
                      <th onClick={() => handleSort("diversity")}>Diversity {sortBy === "diversity" && (sortDir === "asc" ? "▲" : "▼")}</th>
                      <th onClick={() => handleSort("span")}>Span ID {sortBy === "span" && (sortDir === "asc" ? "▲" : "▼")}</th>
                      <th onClick={() => handleSort("idf")}>IDF {sortBy === "idf" && (sortDir === "asc" ? "▲" : "▼")}</th>
                      <th onClick={() => handleSort("splice")}>Splice Rack {sortBy === "splice" && (sortDir === "asc" ? "▲" : "▼")}</th>
                      <th onClick={() => handleSort("scope")}>Scope {sortBy === "scope" && (sortDir === "asc" ? "▲" : "▼")}</th>
                      <th onClick={() => handleSort("status")}>Status {sortBy === "status" && (sortDir === "asc" ? "▲" : "▼")}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedResults.map((row, i) => (
                      <tr key={i} className={selectedSpans.includes(row.SpanID) ? "highlight-row" : ""} onClick={() => toggleSelectSpan(row.SpanID)} style={{ cursor: "pointer" }}>
                        <td><Checkbox checked={selectedSpans.includes(row.SpanID)} onChange={() => toggleSelectSpan(row.SpanID)} /></td>
                        <td><span className={`status-label ${getDiversityClass(row.Diversity)}`}>{row.Diversity}</span></td>
                        <td>
                          <a href={row.OpticalLink} target="_blank" rel="noopener noreferrer" className="uid-click" onClick={(e) => e.stopPropagation()}>
                            {row.SpanID}
                          </a>
                        </td>
                        <td>{row.IDF_A}</td>
                        <td>{row.SpliceRackA}</td>
                        <td>{row.WiringScope}</td>
                        <td>
                          {(() => {
                            const stateVal = (((row as any).State || '') as string).toLowerCase();
                            const isNew = stateVal === 'new';
                            const display = isNew ? 'New' : (row.Status || '');
                            return <span className={`status-label ${getStatusClass(display)}`}>{display}</span>;
                          })()}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div style={{ textAlign: "center", marginTop: 12 }}>
                  <PrimaryButton text={`Continue (${selectedSpans.length} selected)`} disabled={selectedSpans.length === 0} onClick={() => setComposeOpen(true)} />
                </div>
              </div>
            )}
          </>
        )}

        {composeOpen && (
          <div className="table-container compose-container" style={{ marginTop: 16 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
              <IconButton className="back-button" iconProps={{ iconName: 'ChevronLeft' }} title="Back" ariaLabel="Back" onClick={() => setComposeOpen(false)} />
              <div className="section-title" style={{ margin: 0 }}>Compose Maintenance Email</div>
            </div>

            {sendError && (<MessageBar messageBarType={MessageBarType.error} isMultiline={false}>{sendError}</MessageBar>)}

            <Dialog hidden={!showSendSuccessDialog} className="dialog-send-success" onDismiss={() => setShowSendSuccessDialog(false)} dialogContentProps={{ type: DialogType.normal, title: 'Email Stage', subText: sendSuccess || 'Done.' }} modalProps={{ isBlocking: false }}>
              <DialogFooter>
                <PrimaryButton text="Start Over" onClick={() => { setComposeOpen(false); setSelectedSpans([]); setSubject(""); setNotificationType("New Maintenance Scheduled"); setLocation(""); setLat(""); setLng(""); setIsp(""); setIspTicket(""); setMaintenanceReason(""); setImpactExpected(true); setStartDate(null); setStartTime("00:00"); setEndDate(null); setEndTime("00:00"); setAdditionalWindows([]); setCc(""); setFieldErrors({}); setShowValidation(false); setSendError(null); setSendSuccess(null); setShowSendSuccessDialog(false); }} />
                <DefaultButton text="Close" onClick={() => setShowSendSuccessDialog(false)} />
              </DialogFooter>
            </Dialog>

            <div style={{ display: 'flex', gap: 12, alignItems: 'flex-start' }}>
              <div style={{ flex: 1 }}>
                <div style={{ display: 'flex', gap: 12, alignItems: 'center' }}>
                  <div style={{ flex: 2 }}>
                    <TextField label="Subject" placeholder="[region] Maintenance scheduled in <region> Contractor:  Lead Engineer:" value={subject} onChange={(_, v) => setSubject(v || "")} styles={textFieldStyles} required errorMessage={showValidation ? fieldErrors.subject : undefined} />
                  </div>
                  <div style={{ flex: 1 }}>
                    <TextField label="CC" placeholder="name@contoso.com" value={cc} onChange={(_, v) => setCc(v || "")} styles={textFieldStyles} required errorMessage={showValidation ? fieldErrors.cc : undefined} />
                  </div>
                  <div style={{ width: 320, flexShrink: 0 }}>
                    <Dropdown label="Notification Type" options={[{ key: "New Maintenance Scheduled", text: "New Maintenance Scheduled" }, { key: "Rescheduled", text: "Rescheduled" }, { key: "Maintenance Cancelled", text: "Maintenance Cancelled" }, { key: "Maintenance Reminder", text: "Maintenance Reminder" }]} selectedKey={notificationType} onChange={(_, opt) => opt && setNotificationType(opt.key.toString())} styles={dropdownStyles} required />
                  </div>
                </div>

                <div className="compose-datetime-row">
                  <div className="dt-field">
                    <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>Start Date</Text>
                    <DatePicker placeholder="Select start date" value={startDate || undefined} onSelectDate={(d) => setStartDate(d || null)} styles={datePickerStyles} isRequired aria-label="Start Date" />
                  </div>
                  <div className="dt-time">
                    <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>Start Time</Text>
                    <Dropdown options={timeOptions} selectedKey={startTime} onChange={(_, opt) => opt && setStartTime(opt.key.toString())} styles={timeDropdownStyles} required errorMessage={showValidation ? fieldErrors.startTime : undefined} />
                  </div>
                  <div className="dt-field">
                    <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>End Date</Text>
                    <DatePicker placeholder="Select end date" value={endDate || undefined} onSelectDate={(d) => setEndDate(d || null)} styles={datePickerStyles} isRequired aria-label="End Date" />
                  </div>
                  <div className="dt-time">
                    <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>End Time</Text>
                    <Dropdown options={timeOptions} selectedKey={endTime} onChange={(_, opt) => opt && setEndTime(opt.key.toString())} styles={timeDropdownStyles} required errorMessage={showValidation ? fieldErrors.endTime : undefined} />
                  </div>
                  <div className="dt-actions">
                    <button type="button" className="tiny-icon-btn add-window-btn" aria-label="Add Window" onClick={() => setAdditionalWindows((w) => [...w, { startDate: null, startTime: "00:00", endDate: null, endTime: "00:00" }])} title="Add Window"><span className="glyph">+</span></button>
                  </div>
                </div>

                {additionalWindows.map((w, i) => (
                  <div key={i} className="compose-datetime-row additional-window">
                    <div className="dt-field">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>Start Date</Text>
                      <DatePicker placeholder="Select start date" value={w.startDate || undefined} onSelectDate={(d) => setAdditionalWindows((arr) => { const next = [...arr]; next[i] = { ...next[i], startDate: d || null }; return next; })} styles={datePickerStyles} />
                    </div>
                    <div className="dt-time">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>Start Time</Text>
                      <Dropdown options={timeOptions} selectedKey={w.startTime} onChange={(_, opt) => opt && setAdditionalWindows((arr) => { const next = [...arr]; next[i] = { ...next[i], startTime: opt.key.toString() }; return next; })} styles={timeDropdownStyles} />
                    </div>
                    <div className="dt-field">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>End Date</Text>
                      <DatePicker placeholder="Select end date" value={w.endDate || undefined} onSelectDate={(d) => setAdditionalWindows((arr) => { const next = [...arr]; next[i] = { ...next[i], endDate: d || null }; return next; })} styles={datePickerStyles} />
                    </div>
                    <div className="dt-time">
                      <Text styles={{ root: { color: "#ccc", fontSize: 12, fontWeight: 600 } }}>End Time</Text>
                      <Dropdown options={timeOptions} selectedKey={w.endTime} onChange={(_, opt) => opt && setAdditionalWindows((arr) => { const next = [...arr]; next[i] = { ...next[i], endTime: opt.key.toString() }; return next; })} styles={timeDropdownStyles} />
                    </div>
                    <div className="dt-actions">
                      <button type="button" className="tiny-icon-btn remove-window-btn" aria-label={`Remove window ${i + 2}`} title="Remove Window" onClick={() => setAdditionalWindows((arr) => arr.filter((_, idx) => idx !== i))}><span className="glyph">−</span></button>
                    </div>
                  </div>
                ))}

                <div style={{ display: 'flex', gap: 12, marginTop: 8 }}>
                  <div style={{ flex: 1 }}>
                    <TextField label="Location" value={location} onChange={(_, v) => setLocation(v || "")} styles={textFieldStyles} required errorMessage={showValidation ? fieldErrors.location : undefined} />
                  </div>
                  <div style={{ flex: 1 }}>
                    <TextField label="Latitude" value={lat} onChange={(_, v) => setLat((v || '').trim())} styles={textFieldStyles} />
                  </div>
                  <div style={{ flex: 1 }}>
                    <TextField label="Longitude" value={lng} onChange={(_, v) => setLng((v || '').trim())} styles={textFieldStyles} />
                  </div>
                </div>

                <div style={{ display: 'flex', gap: 12, marginTop: 8, alignItems: 'flex-end' }}>
                  <div style={{ flex: 1 }}>
                    <TextField label="ISP" value={isp} onChange={(_, v) => setIsp(v || "")} styles={textFieldStyles} required errorMessage={showValidation ? fieldErrors.isp : undefined} />
                  </div>
                  <div style={{ flex: 1 }}>
                    <TextField label="ISP Ticket / Change ID" value={ispTicket} onChange={(_, v) => setIspTicket(v || "")} styles={textFieldStyles} required errorMessage={showValidation ? fieldErrors.ispTicket : undefined} />
                  </div>
                  <div style={{ width: 200 }}>
                    <Dropdown label="Impact Expected" options={[{ key: "true", text: "Yes/True" }, { key: "false", text: "No/False" }]} selectedKey={impactExpected ? "true" : "false"} onChange={(_, opt) => opt && setImpactExpected(opt.key === "true")} styles={dropdownStyles} required errorMessage={showValidation ? fieldErrors.impactExpected : undefined} />
                  </div>
                </div>

                <div style={{ display: 'flex', gap: 8, marginTop: 8, alignItems: 'center' }}>
                  {lat && lng ? (<a className="uid-click" href={`https://www.bing.com/maps?q=${encodeURIComponent(lat+','+lng)}`} target="_blank" rel="noopener noreferrer">Open Map</a>) : null}
                </div>
              </div>
            </div>

            <div style={{ marginTop: 10 }}>
              <div className="reason-wrapper" style={{ position: 'relative' }}>
                <TextField className="reason-field" label="Reason for Maintenance" multiline autoAdjustHeight value={maintenanceReason} onChange={(_, v) => setMaintenanceReason(v || "")} styles={{ ...textFieldStyles, field: { ...(textFieldStyles as any).field, minHeight: 220, paddingBottom: 28 }, fieldGroup: { ...(textFieldStyles as any).fieldGroup, height: 'auto' } }} maxLength={2000} aria-label="Reason for Maintenance" required errorMessage={showValidation ? fieldErrors.maintenanceReason : undefined} />
                <div className="reason-counter" aria-hidden style={{ position: 'absolute', right: 10, bottom: 8, fontSize: 12 }}>{`${maintenanceReason.length}/2000`}</div>
              </div>
            </div>

            <div className="section-title" style={{ marginTop: 6 }}>Email Body Preview</div>
            <div style={{ background: "#0f0f0f", border: "1px solid #333", borderRadius: 8, padding: 12, whiteSpace: "pre-wrap", color: "#dfefff" }}>
              {(() => {
                const startList: string[] = []; const endList: string[] = [];
                const s1 = formatUtcString(startDate, startTime); const e1 = formatUtcString(endDate, endTime);
                if (s1) startList.push(s1); if (e1) endList.push(e1);
                additionalWindows.forEach((w) => { const s = formatUtcString(w.startDate, w.startTime); const e = formatUtcString(w.endDate, w.endTime); if (s) startList.push(s); if (e) endList.push(e); });
                const impactStr = impactExpected ? "Yes/True" : "No/False";
                const parts: string[] = [
                  `To: opticaldri@microsoft.com`,
                  `From: Fibervsoassistant@microsoft.com`,
                  `CC: ${cc || ""}`,
                  `Subject: ${subject}`,
                  ``,
                  `----------------------------------------`,
                  `CircuitIds: ${spansComma}`,
                  `StartDatetime: ${startList.join(', ')}`,
                  `EndDatetime: ${endList.join(', ')}`,
                  `NotificationType: ${notificationType}`,
                  `MaintenanceReason: ${maintenanceReason}`,
                  `Location: ${location}`,
                  `ISP: ${isp}`,
                  `ISPTicket: ${ispTicket}`,
                  `ImpactExpected: ${impactStr}`,
                ];
                return parts.join("\n");
              })()}
            </div>

            <div style={{ display: "flex", justifyContent: "space-between", alignItems: 'center', marginTop: 12 }}>
              <button className="sleek-btn danger" onClick={() => setComposeOpen(false)}>Back</button>
              <button className="sleek-btn wan" disabled={selectedSpans.length === 0 || sendLoading} onClick={handleSend}>{sendLoading ? "Sending..." : "Confirm & Send"}</button>
            </div>
          </div>
        )}

        <hr />
        <div className="disclaimer">This page is a backend-tied clone of VSO Assistant for testing the new Azure Functions API. Email sending is not implemented server-side yet; the preview and calendar flow are still available for validation.</div>
      </div>

      <VSOCalendar
        events={vsoEvents}
        date={calendarDate || undefined}
        onNavigate={(d) => { setCalendarDate(d); try { localStorage.setItem("vsoCalendarDate", d.toISOString()); } catch {} }}
        onEventClick={(ev) => { /* no-op for now */ }}
      />
    </div>
  );
};

export default VSOAssistantDev;
