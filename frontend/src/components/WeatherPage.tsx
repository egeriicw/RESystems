import { useState } from "react";
import { Bar, CartesianGrid, ComposedChart, Legend, Line, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import type { WeatherSummary } from "../api";
import { MONTHS, api, fmt } from "../api";

export function WeatherPage({ weather, onChange }: { weather: WeatherSummary[]; onChange: () => void }) {
  const [selected, setSelected] = useState(weather[0]?.id);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [synth, setSynth] = useState({ name: "My site", latitude: 39.9, longitude: -84.2, time_zone: -5, mean_temp: 11.5, temp_amplitude: 13 });
  const current = weather.find((w) => w.id === selected) ?? weather[0];

  const run = async (fn: () => Promise<WeatherSummary | void>) => {
    setBusy(true); setError(null);
    try {
      const w = await fn();
      onChange();
      if (w) setSelected(w.id);
    } catch (e) { setError((e as Error).message); } finally { setBusy(false); }
  };

  return (
    <div className="weather-page">
      <section className="card">
        <h2>Weather datasets</h2>
        <table>
          <thead><tr><th>Name</th><th>Source</th><th>Lat</th><th>Lon</th><th>TZ</th><th /></tr></thead>
          <tbody>
            {weather.map((w) => (
              <tr key={w.id} className={w.id === current?.id ? "selected" : ""} onClick={() => setSelected(w.id)}>
                <td>{w.name}</td><td>{w.source}</td><td>{fmt(w.latitude, 2)}</td><td>{fmt(w.longitude, 2)}</td><td>{w.time_zone}</td>
                <td>{w.id !== "sample" && (
                  <button className="link" onClick={(e) => { e.stopPropagation(); run(() => api.deleteWeather(w.id)); }}>remove</button>
                )}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {error && <p className="error">{error}</p>}
      </section>

      <div className="two-col">
        <section className="card">
          <h3>Upload TMY2 file</h3>
          <p className="hint">NREL Typical Meteorological Year 2 (.tm2) files: one header line plus 8,760 hourly records.
            The same fixed columns as the original program are read: global horizontal radiation, dry-bulb and dew-point temperature.</p>
          <input type="file" accept=".tm2,.txt" disabled={busy}
            onChange={(e) => { const f = e.target.files?.[0]; if (f) run(() => api.uploadWeather(f)); e.target.value = ""; }} />
        </section>
        <section className="card">
          <h3>Generate synthetic year</h3>
          <div className="grid-2">
            {(Object.keys(synth) as (keyof typeof synth)[]).map((k) => (
              <label className="field" key={k}><span>{k.replace("_", " ")}</span>
                <input type={k === "name" ? "text" : "number"} value={synth[k]} step="any"
                  onChange={(e) => setSynth({ ...synth, [k]: k === "name" ? e.target.value : Number(e.target.value) })} />
              </label>
            ))}
          </div>
          <button className="primary" disabled={busy} onClick={() => run(() => api.syntheticWeather(synth))}>Generate</button>
        </section>
      </div>

      {current && (
        <section className="card">
          <h3>{current.name} – monthly climate</h3>
          <ResponsiveContainer width="100%" height={280}>
            <ComposedChart data={current.monthly.map((m) => ({ ...m, x: MONTHS[m.month - 1] }))}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--grid)" />
              <XAxis dataKey="x" />
              <YAxis yAxisId="i" tickFormatter={(v) => fmt(v)} width={70} />
              <YAxis yAxisId="t" orientation="right" />
              <Tooltip formatter={(v: number) => fmt(v, 1)} />
              <Legend />
              <Bar isAnimationActive={false} yAxisId="i" dataKey="Ih" name="Horizontal radiation (kJ/day·m²)" fill="#f59e0b" />
              <Line isAnimationActive={false} yAxisId="t" dataKey="Ta" name="Mean temperature (°C)" stroke="#2563eb" strokeWidth={2} />
            </ComposedChart>
          </ResponsiveContainer>
        </section>
      )}
    </div>
  );
}
