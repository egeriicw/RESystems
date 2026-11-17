import { useCallback, useEffect, useState } from "react";
import type {
  CollectorParams, ModelInfo, ModelType, Scenario, SimulationResult, Site, SystemParams, WeatherSummary,
} from "./api";
import { api, db } from "./api";
import { HeatMap } from "./components/HeatMap";
import { AboutPage, HelpPage } from "./components/InfoPages";
import { CollectorForm, ModelPicker, SystemForm } from "./components/ParamsForm";
import { Metrics, ResultsTable } from "./components/ResultsTable";
import { Compare, ScenarioList } from "./components/Scenarios";
import { SitesPage } from "./components/SitesPage";
import { TrendChart } from "./components/TrendChart";
import { WeatherPage } from "./components/WeatherPage";

type Page = "simulator" | "sites" | "weather" | "about" | "help";
type View = "inputs" | "table" | "trends" | "heatmap" | "scenarios" | "compare";

const VIEWS: { id: View; label: string; needsResult?: boolean }[] = [
  { id: "inputs", label: "Inputs" },
  { id: "table", label: "Monthly table", needsResult: true },
  { id: "trends", label: "Trends", needsResult: true },
  { id: "heatmap", label: "Heat map", needsResult: true },
  { id: "scenarios", label: "Scenarios" },
  { id: "compare", label: "Compare" },
];
const RESULT_VIEWS: View[] = ["table", "trends", "heatmap"];

function hasInvalid(obj: object): boolean {
  return Object.values(obj).some((v) => (typeof v === "number" ? !Number.isFinite(v) : v && typeof v === "object" ? hasInvalid(v) : false));
}

export default function App() {
  const [page, setPage] = useState<Page>("simulator");
  const [view, setView] = useState<View>("inputs");
  const [models, setModels] = useState<ModelInfo[]>([]);
  const [weather, setWeather] = useState<WeatherSummary[]>([]);
  const [sites, setSites] = useState<Site[]>([]);
  const [scenarios, setScenarios] = useState<Scenario[]>([]);
  const [weatherId, setWeatherId] = useState("sample");
  const [model, setModel] = useState<ModelType>("dhw");
  const [collector, setCollector] = useState<CollectorParams | null>(null);
  const [system, setSystem] = useState<SystemParams | null>(null);
  const [scenarioId, setScenarioId] = useState<number | null>(null);
  const [scenarioName, setScenarioName] = useState("");
  const [siteId, setSiteId] = useState<number | null>(null);
  const [result, setResult] = useState<SimulationResult | null>(null);
  const [resultLabel, setResultLabel] = useState<string | null>(null);
  const [compare, setCompare] = useState<number[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [running, setRunning] = useState(false);

  const fail = (e: unknown) => setError((e as Error).message);
  const loadWeather = useCallback(() => api.weather().then(setWeather).catch(fail), []);
  const loadSaved = useCallback(
    () => Promise.all([db.sites(), db.scenarios()]).then(([s, sc]) => { setSites(s); setScenarios(sc); }).catch(fail), []);

  useEffect(() => {
    Promise.all([api.models(), api.defaults()])
      .then(([m, d]) => { setModels(m); setCollector(d.collector); setSystem(d.system); })
      .catch((e) => setError(`Cannot reach the API: ${e.message}`));
    loadWeather();
    loadSaved();
  }, [loadWeather, loadSaved]);

  useEffect(() => {
    if (weather.length && !weather.some((w) => w.id === weatherId)) setWeatherId(weather[0].id);
  }, [weather, weatherId]);

  // Drop comparison picks whose runs no longer exist as a scenario's latest run.
  useEffect(() => {
    setCompare((c) => c.filter((id) => scenarios.some((s) => s.latest_run?.id === id)));
  }, [scenarios]);

  const invalid = !collector || !system || hasInvalid(collector) || hasInvalid(system);

  const busy = async (fn: () => Promise<void>) => {
    setRunning(true); setError(null);
    try { await fn(); } catch (e) { fail(e); } finally { setRunning(false); }
  };

  const show = (r: SimulationResult, label: string | null) => {
    setResult(r); setResultLabel(label);
    if (!RESULT_VIEWS.includes(view)) setView("table");
  };

  const runAdHoc = () => busy(async () => {
    show(await api.simulate({ weather_id: weatherId, model, collector: collector!, system: system! }), null);
  });

  const saveAndRun = (asNew: boolean) => busy(async () => {
    // If the chosen weather is the site's own, store null so the scenario keeps
    // following the site's weather when that changes.
    const site = sites.find((x) => x.id === siteId);
    const weather_id = site && site.weather_id === weatherId ? null : weatherId;
    const body = { name: scenarioName.trim(), site_id: siteId, weather_id, model, collector: collector!, system: system! };
    const saved = scenarioId !== null && !asNew ? await db.updateScenario(scenarioId, body) : await db.createScenario(body);
    setScenarioId(saved.id);
    const run = await db.runScenario(saved.id);
    show(run.result, `Saved as “${saved.name}”`);
    await loadSaved();
  });

  const openScenario = (s: Scenario) => {
    setScenarioId(s.id); setScenarioName(s.name); setSiteId(s.site_id);
    setModel(s.model); setCollector(s.collector); setSystem(s.system);
    setWeatherId(s.effective_weather_id);
    setPage("simulator");
    if (s.latest_run) {
      busy(async () => { const run = await db.run(s.latest_run!.id); setResult(run.result); setResultLabel(`Stored run of “${s.name}”`); setView("table"); });
    } else {
      setResult(null); setView("inputs");
    }
  };

  const chooseSite = (id: number | null) => {
    setSiteId(id);
    const site = sites.find((s) => s.id === id);
    if (!site) return;
    if (site.weather_id) setWeatherId(site.weather_id);
    if (site.ua_house != null && system) setSystem({ ...system, ua_house: site.ua_house, ua_components: null });
  };

  const newScenarioForSite = (site: Site) => {
    setScenarioId(null); setScenarioName(`${site.name} scenario`); setResult(null);
    setSiteId(site.id);
    if (site.weather_id) setWeatherId(site.weather_id);
    if (site.ua_house != null && system) setSystem({ ...system, ua_house: site.ua_house, ua_components: null });
    setPage("simulator"); setView("inputs");
  };

  const currentScenario = scenarios.find((s) => s.id === scenarioId);

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">☀ Renewable Energy Systems</div>
        <nav>
          {(["simulator", "sites", "weather", "about", "help"] as Page[]).map((p) => (
            <button key={p} className={p === page ? "on" : ""} onClick={() => setPage(p)}>{p}</button>
          ))}
        </nav>
      </header>

      {page === "simulator" && (
        <div className="layout">
          <aside className="sidebar">
            {VIEWS.map((v) => (
              <button key={v.id} className={v.id === view ? "on" : ""} disabled={v.needsResult && !result}
                onClick={() => setView(v.id)}>{v.label}{v.id === "scenarios" && scenarios.length > 0 && ` (${scenarios.length})`}</button>
            ))}
            <div className="sidebar-run">
              <label className="field"><span>Weather</span>
                <select value={weatherId} onChange={(e) => setWeatherId(e.target.value)}>
                  {weather.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
                </select>
              </label>
              <button className="primary" disabled={running || invalid} onClick={runAdHoc}>
                {running ? "Running…" : "Run simulation"}
              </button>
              {invalid && collector && <p className="error small">Fill in every input.</p>}
            </div>
            <div className="sidebar-run">
              <strong className="side-title">{currentScenario ? "Saved scenario" : "Save as scenario"}</strong>
              <label className="field"><span>Name</span>
                <input type="text" value={scenarioName} maxLength={120} placeholder="e.g. 8 collectors, 60° tilt"
                  onChange={(e) => setScenarioName(e.target.value)} />
              </label>
              <label className="field"><span>Site</span>
                <select value={siteId ?? ""} onChange={(e) => chooseSite(e.target.value === "" ? null : Number(e.target.value))}>
                  <option value="">No site</option>
                  {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                </select>
              </label>
              <button className="secondary" disabled={running || invalid || scenarioName.trim() === ""} onClick={() => saveAndRun(false)}>
                {currentScenario ? "Update & run" : "Save & run"}
              </button>
              {currentScenario && (
                <button className="link" disabled={running || invalid || scenarioName.trim() === ""} onClick={() => saveAndRun(true)}>
                  Save as a new scenario
                </button>
              )}
              <p className="hint">Saved runs are stored and can be compared. “Run simulation” does not store anything.</p>
            </div>
          </aside>

          <main className="content">
            {error && <p className="error">{error}</p>}
            {result && RESULT_VIEWS.includes(view) && (
              <>
                <h2>{result.title}</h2>
                {resultLabel && <p className="hint above">{resultLabel}</p>}
                <Metrics result={result} />
                {result.notes.map((n) => <p key={n} className="note">{n}</p>)}
              </>
            )}
            {view === "inputs" && collector && system && (
              <div className="inputs">
                <div>
                  <ModelPicker models={models} value={model} onChange={setModel} />
                  <CollectorForm value={collector} onChange={setCollector} />
                </div>
                <div><SystemForm model={model} value={system} onChange={setSystem} /></div>
              </div>
            )}
            {result && view === "table" && <section className="card"><ResultsTable result={result} /></section>}
            {result && view === "trends" && <section className="card"><TrendChart key={result.model} result={result} /></section>}
            {result && view === "heatmap" && <section className="card"><HeatMap key={result.model} result={result} /></section>}
            {view === "scenarios" && (
              <section className="card">
                <h2>Saved scenarios</h2>
                <ScenarioList scenarios={scenarios} models={models} currentId={scenarioId} compare={compare}
                  onCompareChange={setCompare} onOpen={openScenario} onCompare={() => setView("compare")}
                  onRun={(s) => busy(async () => { await db.runScenario(s.id); await loadSaved(); })}
                  onDelete={(s) => busy(async () => {
                    await db.deleteScenario(s.id);
                    if (s.id === scenarioId) setScenarioId(null);
                    await loadSaved();
                  })} />
              </section>
            )}
            {view === "compare" && <><h2>Compare runs</h2><Compare runIds={compare} /></>}
          </main>
        </div>
      )}
      {page === "sites" && (
        <main className="content wide">
          <SitesPage sites={sites} weather={weather} scenarios={scenarios} onChange={loadSaved}
            onNewScenario={newScenarioForSite} onOpenScenario={openScenario} />
        </main>
      )}
      {page === "weather" && <main className="content wide"><WeatherPage weather={weather} onChange={() => { loadWeather(); loadSaved(); }} /></main>}
      {page === "about" && <main className="content wide"><AboutPage /></main>}
      {page === "help" && <main className="content wide"><HelpPage /></main>}
    </div>
  );
}
