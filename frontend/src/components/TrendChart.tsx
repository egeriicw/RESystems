import { useMemo, useState } from "react";
import { CartesianGrid, Legend, Line, LineChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import type { SimulationResult } from "../api";
import { MONTHS, fmt } from "../api";

const COLORS = ["#2563eb", "#ea580c", "#16a34a", "#9333ea", "#dc2626", "#0891b2", "#ca8a04", "#db2777", "#4b5563", "#65a30d", "#7c3aed"];
type Resolution = "monthly" | "daily" | "hourly";

function dayLabel(d: number) {
  const date = new Date(Date.UTC(2001, 0, 1 + d));
  return `${MONTHS[date.getUTCMonth()]} ${date.getUTCDate()}`;
}

export function TrendChart({ result }: { result: SimulationResult }) {
  const [resolution, setResolution] = useState<Resolution>("monthly");
  const energy = result.series.filter((s) => s.kind !== "temperature").map((s) => s.key);
  const [selected, setSelected] = useState<string[]>(energy.filter((k) => k !== "It" && k !== "Ih").slice(0, 3));
  const [startDay, setStartDay] = useState(0);
  const hourlyAvailable = result.hourly !== null;

  const data = useMemo(() => {
    if (resolution === "monthly") return result.monthly.map((r) => ({ ...r, x: MONTHS[r.month - 1] }));
    if (resolution === "daily") {
      return result.daily[result.series[0].key].map((_, d) => {
        const row: Record<string, number | string> = { x: dayLabel(d) };
        for (const s of result.series) row[s.key] = result.daily[s.key][d];
        return row;
      });
    }
    const rows = [];
    for (let h = startDay * 24; h < Math.min((startDay + 7) * 24, 8760); h++) {
      const row: Record<string, number | string> = { x: `${dayLabel(Math.floor(h / 24))} ${(h % 24) + 1}:00` };
      for (const s of result.series) row[s.key] = result.hourly![s.key][h];
      rows.push(row);
    }
    return rows;
  }, [result, resolution, startDay]);

  const tempKeys = selected.filter((k) => result.series.find((s) => s.key === k)?.kind === "temperature");
  const valueKeys = selected.filter((k) => !tempKeys.includes(k));
  const unit = resolution === "hourly" ? "kJ/hr" : "kJ/day";

  return (
    <div>
      <div className="toolbar">
        <div className="segmented">
          {(["monthly", "daily", "hourly"] as Resolution[]).map((r) => (
            <button key={r} className={r === resolution ? "on" : ""} disabled={r === "hourly" && !hourlyAvailable}
              onClick={() => setResolution(r)}>{r}</button>
          ))}
        </div>
        {resolution === "hourly" && (
          <label className="inline">Week starting {dayLabel(startDay)}
            <input type="range" min={0} max={358} value={startDay} onChange={(e) => setStartDay(Number(e.target.value))} />
          </label>
        )}
      </div>
      <div className="series-picker">
        {result.series.map((s, i) => {
          const on = selected.includes(s.key);
          return (
            <button key={s.key} className={on ? "chip on" : "chip"} style={on ? { background: COLORS[i % COLORS.length], borderColor: COLORS[i % COLORS.length] } : {}}
              title={`${s.label} (${s.unit})`}
              onClick={() => setSelected(on ? selected.filter((k) => k !== s.key) : [...selected, s.key])}>
              {s.key}
            </button>
          );
        })}
      </div>
      <ResponsiveContainer width="100%" height={360}>
        <LineChart data={data} margin={{ top: 10, right: 20, bottom: 10, left: 10 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="var(--grid)" />
          <XAxis dataKey="x" minTickGap={30} tick={{ fontSize: 11 }} />
          {valueKeys.length > 0 && (
            <YAxis yAxisId="e" tickFormatter={(v) => fmt(v)} tick={{ fontSize: 11 }} width={70}
              label={{ value: unit, angle: -90, position: "insideLeft", fontSize: 11 }} />
          )}
          {tempKeys.length > 0 && (
            <YAxis yAxisId="t" orientation="right" tick={{ fontSize: 11 }}
              label={{ value: "°C", angle: 90, position: "insideRight", fontSize: 11 }} />
          )}
          <Tooltip formatter={(v: number) => fmt(v, 1)} />
          <Legend />
          {result.series.map((s, i) => selected.includes(s.key) && (
            <Line key={s.key} yAxisId={s.kind === "temperature" ? "t" : "e"} dataKey={s.key} name={s.label}
              stroke={COLORS[i % COLORS.length]} dot={resolution === "monthly"} strokeWidth={resolution === "monthly" ? 2 : 1.2}
              isAnimationActive={false} />
          ))}
        </LineChart>
      </ResponsiveContainer>
      <p className="hint">Radiation (It, Ih) is per m² of collector: kJ/day·m² for daily/monthly values, kJ/hr·m² hourly.</p>
    </div>
  );
}
