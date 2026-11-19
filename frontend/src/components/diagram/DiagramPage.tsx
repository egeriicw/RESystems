import { useCallback, useEffect, useMemo, useRef, useState, type DragEvent } from "react";
import {
  Background, Controls, MiniMap, ReactFlow, ReactFlowProvider, addEdge, useEdgesState, useNodesState, useReactFlow,
  type Connection, type IsValidConnection,
} from "@xyflow/react";
import "@xyflow/react/dist/style.css";
import type {
  CollectorParams, ComponentSpec, Diagram, DiagramIssue, DiagramTemplate, Params, Scenario, SimulationResult,
  Site, SystemParams, WeatherSummary,
} from "../../api";
import { ApiError, api, db, diagramApi } from "../../api";
import { HeatMap } from "../HeatMap";
import { Metrics, ResultsTable } from "../ResultsTable";
import { TrendChart } from "../TrendChart";
import { ComponentNode, type ComponentFlowNode } from "./ComponentNode";
import { EditBox, type EditTarget } from "./EditBox";
import { Glyph } from "./Glyph";
import { WireEdge, type WireFlowEdge } from "./WireEdge";
import { Ctx, KIND_COLORS, KIND_LABELS, type DiagramCtx } from "./context";

const nodeTypes = { component: ComponentNode };
const edgeTypes = { wire: WireEdge };
const DRAG_TYPE = "application/x-resystems-component";

export interface OpenRequest { scenario: Scenario; nonce: number }

// ---------- conversions between the canvas and the API's diagram format

function toDiagram(nodes: ComponentFlowNode[], edges: WireFlowEdge[]): Diagram {
  return {
    nodes: nodes.map((n) => ({ id: n.id, type: n.data.ctype, label: n.data.label, position: { x: Math.round(n.position.x), y: Math.round(n.position.y) }, params: n.data.params })),
    edges: edges.map((e) => ({ id: e.id, source: e.source, source_handle: e.sourceHandle ?? "", target: e.target, target_handle: e.targetHandle ?? "", params: e.data?.params ?? {} })),
  };
}

function fromDiagram(d: Diagram, specs: Record<string, ComponentSpec>): { nodes: ComponentFlowNode[]; edges: WireFlowEdge[] } {
  const types = Object.fromEntries(d.nodes.map((n) => [n.id, n.type]));
  return {
    nodes: d.nodes.map((n) => ({ id: n.id, type: "component", position: n.position, data: { ctype: n.type, label: n.label, params: n.params } })),
    edges: d.edges.map((e) => ({
      id: e.id, type: "wire", source: e.source, sourceHandle: e.source_handle, target: e.target, targetHandle: e.target_handle,
      data: { params: e.params, kind: specs[types[e.source]]?.ports.find((p) => p.id === e.source_handle)?.kind ?? "" },
    })),
  };
}

/** What the simulation depends on: everything except positions and names. */
function signature(d: Diagram) {
  return JSON.stringify([d.nodes.map((n) => [n.id, n.type, n.params]), d.edges.map((e) => [e.id, e.source, e.source_handle, e.target, e.target_handle, e.params])]);
}

/** A result narrowed to one component (keys "id:Out" shown as "Out") or to the system totals. */
function subResult(r: SimulationResult, prefix: string | null): SimulationResult {
  const pick = (key: string) => (prefix === null ? !key.includes(":") : key.startsWith(`${prefix}:`));
  const short = (key: string) => (prefix === null ? key : key.slice(prefix.length + 1));
  const series = r.series.filter((s) => pick(s.key)).map((s) => ({ ...s, key: short(s.key), label: prefix === null ? s.label : s.label.split(" – ").slice(1).join(" – ") || s.label }));
  const remap = (obj: Record<string, unknown>) => {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(obj)) if (k === "month" || pick(k)) out[k === "month" ? k : short(k)] = v;
    return out;
  };
  return {
    ...r, series,
    monthly: r.monthly.map((m) => remap(m) as Record<string, number>),
    annual: remap(r.annual) as Record<string, number>,
    daily: remap(r.daily) as Record<string, number[]>,
    hourly: r.hourly ? (remap(r.hourly) as Record<string, number[]>) : null,
  };
}

function nextId(type: string, taken: Set<string>) {
  for (let k = 1; ; k++) if (!taken.has(`${type}-${k}`)) return `${type}-${k}`;
}

// ---------- page

export function DiagramPage(props: {
  weather: WeatherSummary[]; sites: Site[]; scenarios: Scenario[]; onSaved: () => Promise<unknown>; open: OpenRequest | null;
}) {
  return <ReactFlowProvider><Editor {...props} /></ReactFlowProvider>;
}

function Editor({ weather, sites, scenarios, onSaved, open }: {
  weather: WeatherSummary[]; sites: Site[]; scenarios: Scenario[]; onSaved: () => Promise<unknown>; open: OpenRequest | null;
}) {
  const rf = useReactFlow<ComponentFlowNode, WireFlowEdge>();
  const [specs, setSpecs] = useState<Record<string, ComponentSpec>>({});
  const [templates, setTemplates] = useState<DiagramTemplate[]>([]);
  const [defaults, setDefaults] = useState<{ collector: CollectorParams; system: SystemParams } | null>(null);
  const [nodes, setNodes, onNodesChange] = useNodesState<ComponentFlowNode>([]);
  const [edges, setEdges, onEdgesChange] = useEdgesState<WireFlowEdge>([]);
  const [issues, setIssues] = useState<DiagramIssue[]>([]);
  const [result, setResult] = useState<SimulationResult | null>(null);
  const [resultSig, setResultSig] = useState<string | null>(null);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [autoRun, setAutoRun] = useState(true);
  const [editing, setEditing] = useState<{ kind: "node" | "edge"; id: string } | null>(null);
  const [name, setName] = useState("");
  const [siteId, setSiteId] = useState<number | null>(null);
  const [weatherId, setWeatherId] = useState("sample");
  const [scenarioId, setScenarioId] = useState<number | null>(null);
  const [savedSig, setSavedSig] = useState<string | null>(null);
  const [tab, setTab] = useState<"summary" | "trends" | "heatmap">("summary");
  const [focus, setFocus] = useState<string | null>(null);
  const requestNo = useRef(0);

  const diagram = useMemo(() => toDiagram(nodes, edges), [nodes, edges]);
  const sig = useMemo(() => signature(diagram), [diagram]);
  const runSig = `${sig}|${weatherId}`;
  const stale = result !== null && resultSig !== runSig;
  const saveSig = `${JSON.stringify(diagram)}|${name}|${siteId}|${weatherId}`;
  const dirty = savedSig !== null && savedSig !== saveSig;

  const load = useCallback((d: Diagram, sp: Record<string, ComponentSpec>) => {
    const { nodes: n, edges: e } = fromDiagram(d, sp);
    setNodes(n); setEdges(e); setResult(null); setResultSig(null); setEditing(null); setFocus(null);
    // Fitting an empty diagram would zoom right in, so start blank canvases at 100%.
    setTimeout(() => (n.length ? rf.fitView({ padding: 0.2, maxZoom: 1.1, duration: 0 }) : rf.setViewport({ x: 0, y: 0, zoom: 1 })), 50);
  }, [rf, setNodes, setEdges]);

  useEffect(() => {
    Promise.all([diagramApi.library(), diagramApi.templates(), api.defaults()]).then(([lib, tpls, defs]) => {
      const sp = Object.fromEntries(lib.components.map((c) => [c.type, c]));
      setSpecs(sp); setTemplates(tpls); setDefaults(defs);
      if (!open) {
        const start = tpls.find((t) => t.id === "dhw") ?? tpls[0];
        load(start.diagram, sp); setName(start.title);
      }
    }).catch((e) => setError(`Cannot load the component library: ${e.message}`));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Open a saved diagram scenario (from the Scenarios list or Sites page).
  useEffect(() => {
    if (!open || !open.scenario.diagram || Object.keys(specs).length === 0) return;
    const s = open.scenario;
    load(s.diagram!, specs);
    setName(s.name); setSiteId(s.site_id); setWeatherId(s.effective_weather_id); setScenarioId(s.id);
    setSavedSig(`${JSON.stringify(s.diagram)}|${s.name}|${s.site_id}|${s.effective_weather_id}`);
  }, [open, specs, load]);

  const run = useCallback(async (d: Diagram, wid: string, rsig: string) => {
    const no = ++requestNo.current;
    setRunning(true); setError(null);
    try {
      const r = await diagramApi.simulate(wid, d);
      if (no !== requestNo.current) return;
      setResult(r); setResultSig(rsig); setIssues([]);
    } catch (e) {
      if (no !== requestNo.current) return;
      if (e instanceof ApiError && e.issues.length) setIssues(e.issues);
      else setError((e as Error).message);
    } finally {
      if (no === requestNo.current) setRunning(false);
    }
  }, []);

  // Validate after every change that affects the physics; rerun automatically once the diagram is complete.
  useEffect(() => {
    if (Object.keys(specs).length === 0) return;
    const d = diagram;
    const handle = setTimeout(async () => {
      try {
        const { issues: found } = await diagramApi.validate(d);
        setIssues(found);
        if (found.length === 0 && autoRun) run(d, weatherId, runSig);
      } catch (e) { setError((e as Error).message); }
    }, 350);
    return () => clearTimeout(handle);
    // Positions and names don't affect results, so they are left out on purpose.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [runSig, autoRun, specs]);

  // ---------- editing

  const addComponent = (type: string, position?: { x: number; y: number }) => {
    const taken = new Set(nodes.map((n) => n.id));
    const id = nextId(type, taken);
    const spec = specs[type];
    const pane = document.querySelector(".diagram-canvas")?.getBoundingClientRect();
    const center = pane ? rf.screenToFlowPosition({ x: pane.left + pane.width / 2, y: pane.top + pane.height / 2 }) : { x: 0, y: 0 };
    // Lay new components out on a grid around the middle of the view so they don't stack.
    const k = nodes.length;
    const offset = { x: ((k % 3) - 1) * 270, y: (Math.floor(k / 3) % 3) * 140 - 140 };
    const params: Params = Object.fromEntries(spec.params.map((p) => [p.key, p.default]));
    setNodes((ns) => [...ns, { id, type: "component", position: position ?? { x: center.x - 95 + offset.x, y: center.y - 35 + offset.y }, data: { ctype: type, label: "", params } }]);
  };

  const portOf = useCallback((nodeId: string, handle: string | null) => {
    const n = rf.getNode(nodeId) as ComponentFlowNode | undefined;
    return n ? specs[n.data.ctype]?.ports.find((p) => p.id === handle) : undefined;
  }, [rf, specs]);

  const isValidConnection: IsValidConnection<WireFlowEdge> = useCallback((c) => {
    if (c.source === c.target) return false;
    const sp = portOf(c.source, c.sourceHandle ?? null), tp = portOf(c.target, c.targetHandle ?? null);
    return !!sp && !!tp && sp.kind === tp.kind;
  }, [portOf]);

  const onConnect = useCallback((c: Connection) => {
    const sp = portOf(c.source, c.sourceHandle), tp = portOf(c.target, c.targetHandle);
    if (!sp || !tp) return;
    const params = Object.fromEntries([...sp.wire_params, ...tp.wire_params].map((p) => [p.key, p.default]));
    const id = `w-${Date.now().toString(36)}${Math.floor(Math.random() * 1e4).toString(36)}`;
    setEdges((es) => addEdge({ ...c, id, type: "wire", data: { params, kind: sp.kind } }, es));
  }, [portOf, setEdges]);

  const onDrop = (e: DragEvent) => {
    e.preventDefault();
    const type = e.dataTransfer.getData(DRAG_TYPE);
    if (type && specs[type]) addComponent(type, rf.screenToFlowPosition({ x: e.clientX - 90, y: e.clientY - 30 }));
  };

  const editTarget: EditTarget | null = useMemo(() => {
    if (!editing) return null;
    if (editing.kind === "node") {
      const n = nodes.find((x) => x.id === editing.id);
      if (!n) return null;
      const spec = specs[n.data.ctype];
      return { kind: "node", id: n.id, title: spec.title, ctype: n.data.ctype, label: n.data.label, spec, paramSpecs: spec.params, params: n.data.params };
    }
    const e = edges.find((x) => x.id === editing.id);
    if (!e) return null;
    const sp = portOf(e.source, e.sourceHandle ?? null), tp = portOf(e.target, e.targetHandle ?? null);
    const nameOf = (id: string) => { const n = nodes.find((x) => x.id === id); return n ? (n.data.label || specs[n.data.ctype]?.title) : id; };
    return {
      kind: "edge", id: e.id, title: `Wire – ${KIND_LABELS[e.data?.kind ?? ""] ?? "connection"}`,
      subtitle: `${nameOf(e.source)} (${sp?.label.toLowerCase()}) → ${nameOf(e.target)} (${tp?.label.toLowerCase()})`,
      wireKind: e.data?.kind, paramSpecs: [...(sp?.wire_params ?? []), ...(tp?.wire_params ?? [])], params: e.data?.params ?? {},
    };
  }, [editing, nodes, edges, specs, portOf]);

  const saveEdit = (params: Params, label?: string) => {
    if (!editing) return;
    if (editing.kind === "node") {
      setNodes((ns) => ns.map((n) => (n.id === editing.id ? { ...n, data: { ...n.data, params, label: label ?? n.data.label } } : n)));
    } else {
      setEdges((es) => es.map((e) => (e.id === editing.id ? { ...e, data: { ...e.data!, params } } : e)));
    }
    setEditing(null);  // the change effect above reruns the simulation when auto-rerun is on
  };

  const deleteEditing = () => {
    if (!editing) return;
    if (editing.kind === "node") {
      setNodes((ns) => ns.filter((n) => n.id !== editing.id));
      setEdges((es) => es.filter((e) => e.source !== editing.id && e.target !== editing.id));
      if (focus === editing.id) setFocus(null);
    } else {
      setEdges((es) => es.filter((e) => e.id !== editing.id));
    }
    setEditing(null);
  };

  const runNow = () => run(diagram, weatherId, runSig);

  const loadTemplate = (id: string) => {
    const t = id === "blank" ? { title: "New system", diagram: { nodes: [], edges: [] } } : templates.find((x) => x.id === id);
    if (!t) return;
    if (nodes.length && !window.confirm("Replace the current diagram?")) return;
    load(t.diagram, specs); setName(t.title); setScenarioId(null); setSavedSig(null); setIssues([]);
  };

  const openSaved = (id: number) => {
    const s = scenarios.find((x) => x.id === id);
    if (!s?.diagram) return;
    if (dirty && !window.confirm("Discard unsaved changes to this diagram?")) return;
    load(s.diagram, specs);
    setName(s.name); setSiteId(s.site_id); setWeatherId(s.effective_weather_id); setScenarioId(s.id);
    setSavedSig(`${JSON.stringify(s.diagram)}|${s.name}|${s.site_id}|${s.effective_weather_id}`);
  };

  const chooseSite = (id: number | null) => {
    setSiteId(id);
    const site = sites.find((s) => s.id === id);
    if (site?.weather_id) setWeatherId(site.weather_id);
  };

  const save = async (asNew: boolean) => {
    if (!defaults) return;
    setRunning(true); setError(null);
    try {
      const site = sites.find((x) => x.id === siteId);
      const body = {
        name: name.trim(), site_id: siteId, weather_id: site && site.weather_id === weatherId ? null : weatherId,
        model: "diagram" as const, collector: defaults.collector, system: defaults.system, diagram,
      };
      const saved = scenarioId !== null && !asNew ? await db.updateScenario(scenarioId, body) : await db.createScenario(body);
      setScenarioId(saved.id);
      setSavedSig(saveSig);
      if (issues.length === 0) {
        const runResult = await db.runScenario(saved.id);
        setResult(runResult.result); setResultSig(runSig);
      }
      await onSaved();
    } catch (e) {
      if (e instanceof ApiError && e.issues.length) setIssues(e.issues);
      else setError((e as Error).message);
    } finally { setRunning(false); }
  };

  const ctx: DiagramCtx = useMemo(() => ({ specs, result, issues, stale, open: setEditing }), [specs, result, issues, stale]);
  const categories = useMemo(() => {
    const groups: Record<string, ComponentSpec[]> = {};
    for (const c of Object.values(specs)) (groups[c.category] ??= []).push(c);
    return Object.entries(groups);
  }, [specs]);
  const diagramScenarios = scenarios.filter((s) => s.model === "diagram");
  const shown = result ? subResult(result, focus) : null;
  const nodeName = (id: string) => { const n = nodes.find((x) => x.id === id); return n ? (n.data.label || specs[n.data.ctype]?.title || id) : id; };

  return (
    <Ctx.Provider value={ctx}>
      <div className="layout diagram-layout">
        <aside className="sidebar diagram-side">
          <strong className="side-title">Component library</strong>
          <p className="hint">Click to add, or drag onto the canvas. Wire matching coloured ports together.</p>
          {categories.map(([cat, list]) => (
            <div key={cat} className="palette-group">
              <small className="sub">{cat}</small>
              {list.map((c) => (
                <button key={c.type} className="palette-item" draggable title={c.description}
                  onDragStart={(e) => { e.dataTransfer.setData(DRAG_TYPE, c.type); e.dataTransfer.effectAllowed = "move"; }}
                  onClick={() => addComponent(c.type)}>
                  <Glyph type={c.type} size={20} /> {c.title}
                </button>
              ))}
            </div>
          ))}
          <div className="legend-kinds">
            {Object.entries(KIND_LABELS).map(([k, l]) => (
              <span key={k}><i style={{ background: KIND_COLORS[k] }} />{l}</span>
            ))}
          </div>

          <div className="sidebar-run">
            <label className="field"><span>Start from template</span>
              <select value="" onChange={(e) => loadTemplate(e.target.value)}>
                <option value="" disabled>Choose…</option>
                <option value="blank">Blank canvas</option>
                {templates.map((t) => <option key={t.id} value={t.id}>{t.title}</option>)}
              </select>
            </label>
            {diagramScenarios.length > 0 && (
              <label className="field"><span>Open saved diagram</span>
                <select value="" onChange={(e) => openSaved(Number(e.target.value))}>
                  <option value="" disabled>Choose…</option>
                  {diagramScenarios.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                </select>
              </label>
            )}
          </div>

          <div className="sidebar-run">
            <label className="field"><span>Weather</span>
              <select value={weatherId} onChange={(e) => setWeatherId(e.target.value)}>
                {weather.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
              </select>
            </label>
            <label className="inline"><input type="checkbox" checked={autoRun} onChange={(e) => setAutoRun(e.target.checked)} /> Rerun after every change</label>
            <button className="primary" disabled={running || issues.length > 0 || nodes.length === 0} onClick={runNow}>
              {running ? "Running…" : "Run simulation"}
            </button>
          </div>

          <div className="sidebar-run">
            <strong className="side-title">{scenarioId ? "Saved diagram" : "Save diagram"}{dirty && <em className="dirty"> · unsaved changes</em>}</strong>
            <label className="field"><span>Name</span>
              <input type="text" value={name} maxLength={120} onChange={(e) => setName(e.target.value)} />
            </label>
            <label className="field"><span>Site</span>
              <select value={siteId ?? ""} onChange={(e) => chooseSite(e.target.value === "" ? null : Number(e.target.value))}>
                <option value="">No site</option>
                {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </label>
            <button className="secondary" disabled={running || !name.trim() || !defaults} onClick={() => save(false)}>
              {scenarioId ? "Update & run" : "Save & run"}
            </button>
            {scenarioId && <button className="link" disabled={running || !name.trim()} onClick={() => save(true)}>Save as a new diagram</button>}
            <p className="hint">Saved diagrams appear under Simulator → Scenarios and can be compared.</p>
          </div>
        </aside>

        <main className="content diagram-main">
          {error && <p className="error">{error}</p>}
          <div className="diagram-status">
            {issues.length > 0 ? (
              <div className="issues-box">
                <strong>{issues.length === 1 ? "1 thing to fix before running" : `${issues.length} things to fix before running`}</strong>
                <ul>
                  {issues.map((i, k) => (
                    <li key={k}>
                      {(i.node_id || i.edge_id) ? (
                        <button className="link" onClick={() => setEditing({ kind: i.node_id ? "node" : "edge", id: (i.node_id ?? i.edge_id)! })}>
                          {i.node_id ? nodeName(i.node_id) : "Wire"}
                        </button>
                      ) : null}{(i.node_id || i.edge_id) && ": "}{i.message}
                    </li>
                  ))}
                </ul>
              </div>
            ) : (
              <span className={`status-pill${running ? " busy" : stale ? " stale" : result ? " ok" : ""}`}>
                {running ? "Simulating…" : stale ? "Changed since last run" : result ? "Results up to date" : "Ready to run"}
              </span>
            )}
          </div>

          <div className="diagram-canvas" onDragOver={(e) => { e.preventDefault(); e.dataTransfer.dropEffect = "move"; }} onDrop={onDrop}>
            <ReactFlow<ComponentFlowNode, WireFlowEdge>
              nodes={nodes} edges={edges} nodeTypes={nodeTypes} edgeTypes={edgeTypes}
              onNodesChange={onNodesChange} onEdgesChange={onEdgesChange} onConnect={onConnect}
              isValidConnection={isValidConnection}
              onNodeClick={(_, n) => setEditing({ kind: "node", id: n.id })}
              onEdgeClick={(_, e) => setEditing({ kind: "edge", id: e.id })}
              defaultEdgeOptions={{ type: "wire" }} colorMode="system" fitView minZoom={0.2} maxZoom={2}
              deleteKeyCode={["Delete", "Backspace"]} proOptions={{ hideAttribution: false }}>
              <Background gap={20} />
              <Controls />
              <MiniMap pannable zoomable nodeStrokeWidth={3} />
            </ReactFlow>
          </div>

          {result && shown && (
            <section className="card diagram-results">
              <div className="toolbar">
                <div className="segmented">
                  {(["summary", "trends", "heatmap"] as const).map((t) => (
                    <button key={t} className={tab === t ? "on" : ""} onClick={() => setTab(t)}>{t === "heatmap" ? "Heat map" : t}</button>
                  ))}
                </div>
                <label className="inline">Show
                  <select value={focus ?? ""} onChange={(e) => setFocus(e.target.value || null)}>
                    <option value="">System totals</option>
                    {nodes.filter((n) => result.diagram?.nodes[n.id]).map((n) => <option key={n.id} value={n.id}>{nodeName(n.id)}</option>)}
                  </select>
                </label>
                {stale && <span className="hint">Showing the last run; the diagram has changed since.</span>}
              </div>
              {tab === "summary" && (<><Metrics result={result} />{result.notes.map((n) => <p key={n} className="note">{n}</p>)}<ResultsTable result={shown} /></>)}
              {tab === "trends" && <TrendChart key={`${focus}-${resultSig}`} result={shown} />}
              {tab === "heatmap" && <HeatMap key={`${focus}-${resultSig}`} result={shown} />}
            </section>
          )}
        </main>
      </div>

      {editTarget && (
        <EditBox key={`${editTarget.kind}-${editTarget.id}`} target={editTarget} result={result} stale={stale} issues={issues} rerun={autoRun}
          onSave={saveEdit} onDelete={deleteEditing} onClose={() => setEditing(null)} />
      )}
    </Ctx.Provider>
  );
}
