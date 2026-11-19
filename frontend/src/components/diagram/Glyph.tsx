// Simple line icons for each component type (stroke uses currentColor).
export function Glyph({ type, size = 26 }: { type: string; size?: number }) {
  const common = { width: size, height: size, viewBox: "0 0 32 32", fill: "none", stroke: "currentColor", strokeWidth: 2, strokeLinejoin: "round" as const, strokeLinecap: "round" as const, "aria-hidden": true };
  switch (type) {
    case "collector":
      return <svg {...common}><path d="M4 24 L10 8 H28 L22 24 Z" /><path d="M7 16 H25 M13 8 L9 24 M19 8 L15 24 M25 8 L21 24" strokeWidth={1.2} /></svg>;
    case "heat_exchanger":
      return <svg {...common}><circle cx="16" cy="16" r="11" /><path d="M6 16 C10 10, 12 22, 16 16 S22 10, 26 16" /></svg>;
    case "tank":
      return <svg {...common}><rect x="9" y="4" width="14" height="24" rx="6" /><path d="M9 16 H23" strokeWidth={1.2} /></svg>;
    case "fan_coil":
      return <svg {...common}><rect x="4" y="6" width="24" height="20" rx="2" /><circle cx="16" cy="16" r="6" /><path d="M16 10 V22 M10 16 H22" strokeWidth={1.2} /></svg>;
    case "direct_emitter":
      return <svg {...common}><path d="M4 22 C8 22, 8 12, 12 12 S16 22, 20 22 S24 12, 28 12" /><path d="M4 27 H28" /></svg>;
    case "house":
      return <svg {...common}><path d="M4 15 L16 5 L28 15" /><path d="M8 13 V27 H24 V13" /><rect x="14" y="19" width="5" height="8" /></svg>;
    case "dhw_load":
      return <svg {...common}><path d="M8 10 H20 C24 10, 24 14, 24 16" /><path d="M8 6 V14" /><path d="M24 20 C22 24, 22 26, 24 27 C26 26, 26 24, 24 20 Z" /></svg>;
    default:
      return <svg {...common}><rect x="6" y="6" width="20" height="20" rx="3" /></svg>;
  }
}
