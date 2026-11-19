import { useEffect, useMemo, useState } from "react";
import { CartesianGrid, Legend, Line, LineChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import type { ComponentSpec, DiagramIssue, ParamSpec, Params, SimulationResult } from "../../api";
import { MONTHS, fmt } from "../../api";
import { Glyph } from "./Glyph";
import { ParamField, parseParams } from "./ParamField";
import { KIND_COLORS, KIND_LABELS, fmtValue } from "./context";

export interface EditTarget {
  kind: "node" | "edge";
  id: string;
  title: string;          // component type title, or "Wire"
  subtitle?: string;      // e.g. "Collector → Heat exchanger"
  ctype?: string;         // component type, for nodes
  wireKind?: string;      // port kind, for wires
  label?: string;         // editable name, for nodes
  spec?: ComponentSpec;   // for nodes
  paramSpecs: ParamSpec[];
  params: Params;
}

const COLORS = ["#2563eb", "#ea580c", "#16a34a", "#9333ea"];

export function EditBox({ target, result, stale, issues, rerun, onSave, onDelete, onClose }: {
  target: EditTarget; result: SimulationResult | null; stale: boolean; issues: DiagramIssue[]; rerun: boolean;
  onSave: (params: Params, label?: string) => void; onDelete: () => void; onClose: () => void;
}) {
  const [tab, setTab] = useState<"params" | "results">("params");
  const [label, setLabel] = useState(target.label ?? "");
  const [draft, setDraft] = useState<Record<string, unknown>>(() => {
    const d: Record<string, unknown> = {};
    for (const p of target.paramSpecs) d[p.key] = target.params[p.key] ?? p.default;
    return d;
  });
  const { values, errors } = useMemo(() => parseParams(target.paramSpecs, draft), [target.paramSpecs, draft]);
  const valid = Object.keys(errors).length === 0;
  const mine = issues.filter((i) => (target.kind === "node" ? i.node_id : i.edge_id) === target.id);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") onClose(); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);

  const save = () => { if (valid) onSave(values, target.kind === "node" ? label.trim() : undefined); };

  return (
    <div className="modal-backdrop" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="modal" role="dialog" aria-modal="true" aria-labelledby="editbox-title">
        <header className="modal-head">
          <span className="modal-icon" style={target.wireKind ? { color: KIND_COLORS[target.wireKind] } : undefined}>
            {target.ctype ? <Glyph type={target.ctype} size={30} /> : <span className="wire-swatch" style={{ background: KIND_COLORS[target.wireKind ?? ""] }} />}
          </span>
          <div>
            <h3 id="editbox-title">{target.kind === "node" ? (label || target.title) : target.title}</h3>
            <small className="sub">{target.subtitle ?? target.spec?.description}</small>
          </div>
          <button className="icon-btn" aria-label="Close" onClick={onClose}>×</button>
        </header>

        <div className="segmented modal-tabs">
          <button className={tab === "params" ? "on" : ""} onClick={() => setTab("params")}>Parameters</button>
          <button className={tab === "results" ? "on" : ""} onClick={() => setTab("results")}>Results</button>
        </div>

        <div className="modal-body">
          {mine.length > 0 && (
            <ul className="issue-list">{mine.map((i, k) => <li key={k}>{i.message}</li>)}</ul>
          )}
          {tab === "params" && (
            <>
              {target.kind === "node" && (
                <label className="field"><span>Name on diagram</span>
                  <input type="text" maxLength={60} value={label} placeholder={target.title} onChange={(e) => setLabel(e.target.value)} />
                </label>
              )}
              {target.paramSpecs.map((p) => (
                <ParamField key={p.key} spec={p} value={draft[p.key]} error={errors[p.key]}
                  onChange={(v) => setDraft({ ...draft, [p.key]: v })} />
              ))}
              {target.paramSpecs.length === 0 && (
                <p className="hint">
                  {target.kind === "edge"
                    ? "This wire has no settings of its own; what flows through it is set by the components at each end."
                    : "This component has no settings of its own."}
                </p>
              )}
            </>
          )}
          {tab === "results" && <Results target={target} result={result} stale={stale} />}
        </div>

        <footer className="modal-foot">
          <button className="link danger" onClick={onDelete}>{target.kind === "node" ? "Delete component" : "Delete wire"}</button>
          <span className="spacer" />
          <button className="secondary" onClick={onClose}>Cancel</button>
          <button className="primary" disabled={!valid} onClick={save}>{rerun ? "Save & rerun" : "Save"}</button>
        </footer>
      </div>
    </div>
  );
}

function Results({ target, result, stale }: { target: EditTarget; result: SimulationResult | null; stale: boolean }) {
  if (!result?.diagram) return <p className="hint">Run the simulation to see results for this {target.kind === "node" ? "component" : "wire"}.</p>;
  const staleNote = stale && <p className="note">These results are from before your latest change.</p>;

  if (target.kind === "edge") {
    const carried = result.diagram.edges[target.id];
    if (!carried) return <p className="hint">This wire wasn't part of the last run.</p>;
    return (
      <>
        {staleNote}
        <div className="kv"><span>{carried.label}</span><strong>{fmt(carried.value)} {carried.unit}</strong></div>
        <p className="hint">Average per day over the year.{target.wireKind && ` Carries ${KIND_LABELS[target.wireKind]?.toLowerCase()}.`}</p>
      </>
    );
  }

  const annual = result.diagram.nodes[target.id];
  if (!annual || !target.spec) return <p className="hint">This component wasn't part of the last run.</p>;
  const outputs = target.spec.outputs;
  const energy = outputs.filter((o) => o.kind === "energy");
  const data = result.monthly.map((r) => {
    const row: Record<string, number | string> = { x: MONTHS[(r.month as number) - 1] };
    for (const o of outputs) row[o.key] = r[`${target.id}:${o.key}`];
    return row;
  });
  const chartKeys = energy.length ? energy : outputs;
  return (
    <>
      {staleNote}
      {outputs.map((o) => (
        <div className="kv" key={o.key}><span>{o.label}</span><strong>{fmtValue(annual[o.key], o.unit)}</strong></div>
      ))}
      <p className="hint">Annual averages ({energy.length ? "energy per day" : "means"}).</p>
      <ResponsiveContainer width="100%" height={200}>
        <LineChart data={data} margin={{ top: 10, right: 10, bottom: 0, left: 0 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="var(--grid)" />
          <XAxis dataKey="x" tick={{ fontSize: 10 }} />
          <YAxis tickFormatter={(v) => fmt(v)} tick={{ fontSize: 10 }} width={60} />
          <Tooltip formatter={(v: number) => fmt(v, 1)} />
          <Legend wrapperStyle={{ fontSize: 11 }} />
          {chartKeys.map((o, i) => (
            <Line key={o.key} dataKey={o.key} name={`${o.label} (${o.unit})`} stroke={COLORS[i % COLORS.length]} dot={false} isAnimationActive={false} />
          ))}
        </LineChart>
      </ResponsiveContainer>
    </>
  );
}
