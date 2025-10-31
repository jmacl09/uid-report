import React from 'react';

type CapacityCircleProps = {
  main: string; // main metric e.g. "800K"
  sub?: string; // sub label e.g. "8 links"
  size?: number; // diameter in px (optional, responsive by default)
  className?: string;
};

// Reusable CapacityCircle component styled to match Optical360 dashboard theme.
// Uses inline styles so it works without Tailwind. Exports a centered, responsive
// circular indicator with neon blue and green accents and a subtle animated glow.
export default function CapacityCircle({ main, sub, size = 140, className }: CapacityCircleProps) {
  // Match WF Finished green
  const green = '#00c853';
  // Use a slightly darker blue for the ring so it reads stronger against the dark background
  // (darker than the bright banner blue):
  const themeBlue = '#0078d4';
  const diameter = Math.max(80, Math.min(size, 360));
  const inner = Math.round(diameter * 0.72); // inner face diameter
  // ringThickness computed previously is not used after design update; remove to avoid lint warning

  const containerStyle: React.CSSProperties = {
    width: diameter,
    height: diameter,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    position: 'relative',
    boxSizing: 'border-box',
    // nudge the whole circle up so it doesn't sit too low in cards/layouts
    transform: 'translateY(-12%)',
    // allow the halo and circle to render outside the parent so it isn't clipped
    overflow: 'visible',
  };

  // removed translucent halo per request; ring will be a solid color to match banners

  const ringStyle: React.CSSProperties = {
    position: 'absolute',
    inset: 0,
    borderRadius: '50%',
    // solid ring color (no gradient) to match brand but slightly darker
    backgroundColor: themeBlue,
    // stronger crisp shadow and faint aura so the ring feels luminous
    boxShadow: `0 8px 20px ${themeBlue}55, inset 0 0 10px ${themeBlue}44`,
    zIndex: 1,
  };

  const haloStyle: React.CSSProperties = {
    position: 'absolute',
    left: -Math.round(diameter * 0.12),
    top: -Math.round(diameter * 0.12),
    right: -Math.round(diameter * 0.12),
    bottom: -Math.round(diameter * 0.12),
    borderRadius: '50%',
    // larger soft glow using the same darker blue
    background: `radial-gradient(ellipse at center, ${themeBlue}33, ${themeBlue}11 40%, transparent 70%)`,
    filter: 'blur(14px)',
    pointerEvents: 'none',
    zIndex: 0,
    animation: 'capacityGlow 3.6s ease-in-out infinite',
  };

  const innerStyle: React.CSSProperties = {
    width: inner,
    height: inner,
    borderRadius: '50%',
    background: 'linear-gradient(180deg, rgba(7,20,36,0.98), rgba(0,16,22,0.95))',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    zIndex: 3,
    boxShadow: 'inset 0 -6px 18px rgba(0,0,0,0.6)',
  };
  // compute a dynamic font size for the main metric so longer numbers shrink to fit
  const sanitizedMainChars = main ? main.replace(/[^A-Za-z0-9]/g, '') : '';
  const mainCharCount = Math.max(1, sanitizedMainChars.length);
  const computedMainFont = Math.max(
    Math.round(inner * 0.22),
    Math.min(Math.round(inner * 0.5), Math.round((inner * 0.9) / mainCharCount))
  );

  const mainStyle: React.CSSProperties = {
    fontSize: computedMainFont,
    lineHeight: 1,
    fontWeight: 900,
    color: green,
    // stronger glow so the metric is the clear focus
    textShadow: `0 4px 14px ${green}55, 0 2px 6px rgba(0,0,0,0.6)`,
    letterSpacing: -0.5,
    whiteSpace: 'nowrap',
    display: 'block',
    maxWidth: inner * 0.9,
    // subtle pulse to match WF Finished vibe
    animation: 'textGreenPulse 1.8s ease-in-out infinite',
  };
  const labelStyle: React.CSSProperties = {
    fontSize: Math.round(inner * 0.12),
    color: '#dfefff',
    marginTop: 8,
    fontWeight: 700,
    opacity: 0.95,
    textTransform: 'uppercase',
    letterSpacing: 1,
  };
  const subStyle: React.CSSProperties = {
    fontSize: Math.round(inner * 0.095),
    color: '#cfe8ff',
    marginTop: 6,
    fontWeight: 700,
    opacity: 0.9,
  };

  // Responsive: allow scaling with container using max-width:100% when size is large
  const wrapper: React.CSSProperties = {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '100%',
    overflow: 'visible',
    // add top padding to the component's bounding box so parents don't clip the halo/top
    paddingTop: Math.round(diameter * 0.18),
  };

  return (
    <div style={wrapper} className={className}>
      <div style={containerStyle} aria-hidden={false}>
        <div style={haloStyle} />
        <div style={ringStyle} />
        <div style={innerStyle}>
          <div style={{ textAlign: 'center' }}>
            <div style={mainStyle}>{main}</div>
            {/* clear label so 'Capacity' is visible even if the arc is subtle */}
            <div style={labelStyle}>Capacity</div>
            {sub && <div style={subStyle}>{sub}</div>}
          </div>
        </div>

        <style>{`
          @keyframes capacityGlow {
            0% { transform: scale(0.995); opacity: 0.9 }
            50% { transform: scale(1.01); opacity: 1 }
            100% { transform: scale(0.995); opacity: 0.9 }
          }
          @keyframes textGreenPulse {
            0%, 100% { opacity: 0.92; text-shadow: 0 2px 6px rgba(0,0,0,0.6), 0 0 12px rgba(0,200,83,0.35) }
            50% { opacity: 1; text-shadow: 0 2px 6px rgba(0,0,0,0.6), 0 0 20px rgba(0,200,83,0.55) }
          }
        `}</style>
      </div>
    </div>
  );
}
