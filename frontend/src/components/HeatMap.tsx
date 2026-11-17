import { useEffect, useMemo, useRef, useState } from "react";
import type { SimulationResult } from "../api";
import { MONTHS, fmt } from "../api";

// Perceptually ordered ramp (viridis-like) from low to high.
const RAMP = ["#440154", "#3b528b", "#21918c", "#5ec962", "#fde725"];

function color(t: number) {
  const x = Math.max(0, Math.min(1, t)) * (RAMP.length - 1);
  const i = Math.min(Math.floor(x), RAMP.length - 2);
  const f = x - i;
  const a = parseInt(RAMP[i].slice(1), 16), b = parseInt(RAMP[i + 1].slice(1), 16);
  const ch = (s: number) => Math.round(((a >> s) & 255) * (1 - f) + ((b >> s) & 255) * f);
  return `rgb(${ch(16)},${ch(8)},${ch(0)})`;
}

export function HeatMap({ result }: { result: SimulationResult }) {
  const [key, setKey] = useState(result.series.find((s) => s.key === "Qu")?.key ?? result.series[0].key);
  const [hover, setHover] = useState<{ day: number; hour: number; value: number } | null>(null);
  const canvas = useRef<HTMLCanvasElement>(null);
  const values = result.hourly?.[key];
  const [min, max] = useMemo(() => {
    if (!values) return [0, 1];
    let lo = Infinity, hi = -Infinity;
    for (const v of values) { lo = Math.min(lo, v); hi = Math.max(hi, v); }
    return [lo, hi === lo ? lo + 1 : hi];
  }, [values]);

  useEffect(() => {
    const ctx = canvas.current?.getContext("2d");
    if (!ctx || !values) return;
    for (let d = 0; d < 365; d++) {
      for (let h = 0; h < 24; h++) {
        ctx.fillStyle = color((values[d * 24 + h] - min) / (max - min));
        ctx.fillRect(d * 2, h * 8, 2, 8);
      }
    }
  }, [values, min, max]);

  if (!values) return <p>Hourly data was not included in this run.</p>;
  const info = result.series.find((s) => s.key === key)!;
  const unit = info.kind === "temperature" ? "C" : info.kind === "irradiance" ? "kJ/hr·m²" : "kJ/hr";

  return (
    <div>
      <div className="toolbar">
        <label className="inline">Series
          <select value={key} onChange={(e) => setKey(e.target.value)}>
            {result.series.map((s) => <option key={s.key} value={s.key}>{s.key} – {s.label}</option>)}
          </select>
        </label>
        <span className="hint">
          {hover ? `Day ${hover.day + 1}, hour ${hover.hour + 1}: ${fmt(hover.value, 1)} ${unit}` : "Hover for values"}
        </span>
      </div>
      <div className="heatmap">
        <div className="hm-hours">{[1, 6, 12, 18, 24].map((h) => <span key={h} style={{ top: (h - 1) * 8 }}>{h}</span>)}</div>
        <div>
          <canvas ref={canvas} width={730} height={192}
            onMouseMove={(e) => {
              const r = e.currentTarget.getBoundingClientRect();
              const day = Math.floor(((e.clientX - r.left) / r.width) * 365);
              const hour = Math.floor(((e.clientY - r.top) / r.height) * 24);
              if (day >= 0 && day < 365 && hour >= 0 && hour < 24) setHover({ day, hour, value: values[day * 24 + hour] });
            }}
            onMouseLeave={() => setHover(null)} />
          <div className="hm-months">{MONTHS.map((m) => <span key={m}>{m}</span>)}</div>
        </div>
      </div>
      <div className="legend-bar">
        <span>{fmt(min, 1)}</span>
        <div style={{ background: `linear-gradient(to right, ${RAMP.join(",")})` }} />
        <span>{fmt(max, 1)} {unit}</span>
      </div>
      <p className="hint">Each column is a day of the year; each row is an hour (1 = midnight–1 am).</p>
    </div>
  );
}
