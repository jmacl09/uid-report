import React from 'react';

type LinksCircleProps = {
  lines: [string, string, string, string]; // [count, 'x', '100G', 'Links']
  size?: number; // diameter in px
  className?: string;
};

// Circle that visually matches CapacityCircle but renders four stacked lines
// with emphasis on 1st and 3rd lines (count and per-link rate).
export default function LinksCircle({ lines, size = 140, className }: LinksCircleProps) {
  const [l0, l1, l2, l3] = lines;
  // Match WF Finished green
  const green = '#00c853';
  const themeBlue = '#0078d4';

  const diameter = Math.max(80, Math.min(size, 360));
  const inner = Math.round(diameter * 0.72);

  const containerStyle: React.CSSProperties = {
    width: diameter,
    height: diameter,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    position: 'relative',
    boxSizing: 'border-box',
    transform: 'translateY(-12%)',
    overflow: 'visible',
  };

  const ringStyle: React.CSSProperties = {
    position: 'absolute',
    inset: 0,
    borderRadius: '50%',
    backgroundColor: themeBlue,
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

  // Dynamic font sizing helpers (fit within inner circle width)
  const fitFont = (text: string, base: number) => {
    const clean = (text || '').replace(/[^A-Za-z0-9]/g, '');
    const len = Math.max(1, clean.length);
    // make sure it never gets absurdly large and shrinks for longer tokens
    return Math.max(Math.round(inner * base * 0.6), Math.min(Math.round(inner * base), Math.round((inner * 0.9) / len)));
  };

  const mainStyle = (text: string): React.CSSProperties => ({
    // Even smaller to ensure comfortable fit for short values like "4" and "100G"
    fontSize: fitFont(text, 0.26),
    lineHeight: 1,
    fontWeight: 900,
    color: green,
    textShadow: `0 4px 14px ${green}55, 0 2px 6px rgba(0,0,0,0.6)`,
    letterSpacing: -0.5,
    whiteSpace: 'nowrap',
    animation: 'textGreenPulse 1.8s ease-in-out infinite',
  });

  const lightStyle = (ratio = 0.16): React.CSSProperties => ({
    fontSize: Math.round(inner * ratio),
    color: '#cfe8ff',
    fontWeight: 700,
    opacity: 0.95,
  });

  const wrapper: React.CSSProperties = {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '100%',
    overflow: 'visible',
    paddingTop: Math.round(diameter * 0.18),
  };

  return (
    <div style={wrapper} className={className}>
      <div style={containerStyle}>
        <div style={haloStyle} />
        <div style={ringStyle} />
        <div style={innerStyle}>
          <div style={{ textAlign: 'center', display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
            <div style={mainStyle(l0)}>{l0}</div>
            <div style={lightStyle(0.13)}>{l1}</div>
            <div style={mainStyle(l2)}>{l2}</div>
            <div style={lightStyle(0.12)}>{l3}</div>
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
