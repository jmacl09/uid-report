import React from "react";

/**
 * ThemedProgressBar
 * - Simulates determinate progress over an expected duration (default ~30s)
 * - When `complete` becomes true, it smoothly accelerates to 100% and calls `onDone`
 * - Visual style matches app's blue theme; place inside a centered container for best look
 */
export default function ThemedProgressBar({
  active,
  expectedMs = 30000,
  complete = false,
  label = "Fetching data…",
  onDone,
  style,
}: {
  active: boolean;
  expectedMs?: number;
  complete?: boolean;
  label?: string;
  onDone?: () => void;
  style?: React.CSSProperties;
}) {
  const [progress, setProgress] = React.useState(0);
  const rafRef = React.useRef<number | null>(null);
  const tickRef = React.useRef<number | null>(null);
  const startRef = React.useRef<number>(0);
  const lastRef = React.useRef<number>(0);

  const clamp = (v: number, a = 0, b = 100) => Math.max(a, Math.min(b, v));
  const easeOutQuad = (t: number) => 1 - (1 - t) * (1 - t);

  // Baseline progression loop while active
  React.useEffect(() => {
    if (!active) {
      // reset and cleanup when not active
      if (rafRef.current) cancelAnimationFrame(rafRef.current);
      if (tickRef.current) window.clearInterval(tickRef.current);
      rafRef.current = null;
      tickRef.current = null;
      setProgress(0);
      return;
    }

    startRef.current = performance.now();
    lastRef.current = startRef.current;
    setProgress(2); // quick nudge so users see motion right away

    // Gentle interval to bump progress toward ~95% over expectedMs
    tickRef.current = window.setInterval(() => {
      const now = performance.now();
      const elapsed = now - startRef.current;
      const ratio = clamp(elapsed / expectedMs, 0, 1);
      // Target caps at ~94-96% so the final jump feels meaningful
      const target = 2 + easeOutQuad(ratio) * 94;
      // Add small micro-variance to make it feel alive
      const jitter = (Math.random() - 0.5) * 0.4; // +/-0.2%
      setProgress((p) => clamp(Math.max(p, target + jitter), 0, 96));
      lastRef.current = now;
    }, 250) as unknown as number;

    return () => {
      if (tickRef.current) window.clearInterval(tickRef.current);
      tickRef.current = null;
    };
  }, [active, expectedMs]);

  // When signaled complete, animate to 100% quickly then call onDone
  const didFinishRef = React.useRef(false);
  React.useEffect(() => {
    if (!active) return;
    if (!complete) return;
    if (didFinishRef.current) return;

    didFinishRef.current = true;
    if (tickRef.current) window.clearInterval(tickRef.current);

    const start = performance.now();
    const startVal = Math.min( Math.max(progress, 60), 98 ); // ensure we don't jump backwards or finish from too low
    const duration = 700; // ms for the final sweep

    const step = (t: number) => {
      const dt = t - start;
      const r = clamp(dt / duration, 0, 1);
      const eased = easeOutQuad(r);
      const next = startVal + (100 - startVal) * eased;
      setProgress(clamp(next));
      if (r < 1) {
        rafRef.current = requestAnimationFrame(step);
      } else {
        // small delay to let users see it hit 100%
        window.setTimeout(() => {
          onDone && onDone();
          // reset local flags for next run
          didFinishRef.current = false;
          setProgress(0);
        }, 250);
      }
    };

    rafRef.current = requestAnimationFrame(step);

    return () => {
      if (rafRef.current) cancelAnimationFrame(rafRef.current);
      rafRef.current = null;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [complete, active]);

  if (!active) return null;

  return (
    <div className="themed-progress" style={style} aria-label={label} aria-valuemin={0} aria-valuemax={100} aria-valuenow={Math.round(progress)} role="progressbar">
      <div className="themed-progress-track">
        <div className="themed-progress-bar" style={{ width: `${clamp(progress)}%` }} />
      </div>
      <div className="themed-progress-label">{label}</div>
    </div>
  );
}
