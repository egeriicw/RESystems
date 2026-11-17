import { useEffect, useState } from "react";
import { CartesianGrid, Legend, Line, LineChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import type { ModelInfo, Run, Scenario } from "../api";
import { MONTHS, db, fmt, pct, shortDate } from "../api";

export function ScenarioList({ scenarios, models, currentId, compare, onCompareChange, onOpen, onRun, onDelete, onCompare }: {
  scenarios: Scenario[]; models: ModelInfo[]; currentId: number | null; compare: number[];
  onCompareChange: (ids: number[]) => void; onOpen: (s: Scenario) => void; onRun: (s: Scenario) => void;
  onDelete: (s: Scenario) => void; onCompare: () => void;
}) {
  const title = (id: string) => models.find((m) => m.id === id)?.title ?? id;
  if (scenarios.length === 0) {
    return <p className="hint">No saved scenarios yet. Set up the inputs, give the scenario a name in the side panel and choose “Save &amp; run”.</p>;
  }
  return (
    <div className="table-wrap">
      <div className="toolbar">
        <button className="primary" disabled={compare.length !== 2} onClick={onCompare}>Compare selected</button>
        <span className="hint">Tick two scenarios that have been run to compare their latest results.</span>
      </div>
      <table>
        <thead>
          <tr><th /><th className="left">Scenario</th><th className="left">Site</th><th className="left">Model</th><th>Efficiency</th><th>Solar fraction</th><th>Last run</th><th /></tr>
        </thead>
        <tbody>
          {scenarios.map((s) => {
            const run = s.latest_run;
            const ticked = run !== null && compare.includes(run.id);
            return (
              <tr key={s.id} className={s.id === currentId ? "selected" : ""}>
                <td><input type="checkbox" aria-label={`Compare ${s.name}`} disabled={!run || (!ticked && compare.length >= 2)} checked={ticked}
                  onChange={() => run && onCompareChange(ticked ? compare.filter((x) => x !== run.id) : [...compare, run.id])} /></td>
                <td className="left"><button className="link" onClick={() => onOpen(s)}>{s.name}</button></td>
                <td className="left">{s.site_name ?? "—"}</td>
                <td className="left">{title(s.model)}</td>
                <td>{pct(run?.metrics.collector_efficiency)}</td>
                <td>{pct(run?.metrics.solar_fraction)}</td>
                <td>{run ? shortDate(run.created_at) : "not run"}</td>
                <td>
                  <button className="link" onClick={() => onRun(s)}>Run</button>{" · "}
                  <button className="link danger" onClick={() => window.confirm(`Delete scenario “${s.name}” and its runs?`) && onDelete(s)}>Delete</button>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

export function Compare({ runIds }: { runIds: number[] }) {
  const [runs, setRuns] = useState<Run[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [key, setKey] = useState<string>("");

  useEffect(() => {
    setRuns(null); setError(null);
    if (runIds.length !== 2) return;
    Promise.all(runIds.map((id) => db.run(id, false))).then(setRuns).catch((e) => setError(e.message));
  }, [runIds]);

  if (runIds.length !== 2) return <p className="hint">Tick two scenarios in the Scenarios view, then choose “Compare selected”.</p>;
  if (error) return <p className="error">{error}</p>;
  if (!runs) return <p className="hint">Loading…</p>;

  const [a, b] = runs;
  const common = a.result.series.filter((s) => b.result.series.some((t) => t.key === s.key));
  const active = common.find((s) => s.key === key) ?? common.find((s) => s.key === "Qts") ?? common[0];
  const data = a.result.monthly.map((row, i) => ({ x: MONTHS[i], A: row[active.key], B: b.result.monthly[i][active.key] }));
  const allKeys = [...a.result.series, ...b.result.series.filter((s) => !a.result.series.some((t) => t.key === s.key))];
  const digits = (kind: string) => (kind === "temperature" ? 1 : 0);

  return (
    <div>
      <div className="compare-heads">
        {runs.map((r, i) => (
          <div key={r.id} className={`card run-${i}`}>
            <span className="tag">{i === 0 ? "A" : "B"}</span>
            <strong>{r.scenario_name}</strong>
            <small className="sub">{r.title} · {r.weather_name} · {shortDate(r.created_at)}</small>
            <div className="pair">
              <div><span>Efficiency</span><b>{pct(r.metrics.collector_efficiency)}</b></div>
              <div><span>Solar fraction</span><b>{pct(r.metrics.solar_fraction)}</b></div>
              <div><span>Collector area</span><b>{fmt(r.result.inputs.collector_area_m2, 2)} m²</b></div>
            </div>
          </div>
        ))}
      </div>

      <section className="card">
        <div className="toolbar">
          <label className="inline">Monthly series
            <select value={active.key} onChange={(e) => setKey(e.target.value)}>
              {common.map((s) => <option key={s.key} value={s.key}>{s.key} – {s.label}</option>)}
            </select>
          </label>
          <span className="hint">{active.unit}</span>
        </div>
        <ResponsiveContainer width="100%" height={300}>
          <LineChart data={data} margin={{ top: 10, right: 20, bottom: 0, left: 10 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--grid)" />
            <XAxis dataKey="x" tick={{ fontSize: 11 }} />
            <YAxis tickFormatter={(v) => fmt(v)} tick={{ fontSize: 11 }} width={70} />
            <Tooltip formatter={(v: number) => fmt(v, 1)} />
            <Legend />
            <Line dataKey="A" name={`A · ${a.scenario_name}`} stroke="#2563eb" strokeWidth={2} isAnimationActive={false} />
            <Line dataKey="B" name={`B · ${b.scenario_name}`} stroke="#ea580c" strokeWidth={2} isAnimationActive={false} />
          </LineChart>
        </ResponsiveContainer>
      </section>

      <section className="card table-wrap">
        <h3>Annual averages</h3>
        <table>
          <thead><tr><th>Series</th><th>Unit</th><th>A</th><th>B</th><th>B − A</th></tr></thead>
          <tbody>
            {allKeys.map((s) => {
              const va = a.annual[s.key], vb = b.annual[s.key];
              const both = va !== undefined && vb !== undefined;
              return (
                <tr key={s.key}>
                  <td>{s.label} <small>({s.key})</small></td>
                  <td>{s.unit}</td>
                  <td>{fmt(va, digits(s.kind))}</td>
                  <td>{fmt(vb, digits(s.kind))}</td>
                  <td>{both ? `${vb - va > 0 ? "+" : ""}${fmt(vb - va, digits(s.kind))}` : "—"}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </section>
    </div>
  );
}
