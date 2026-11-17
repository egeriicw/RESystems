import { useCallback, useEffect, useState } from "react";
import type { CollectorParams, ModelInfo, ModelType, SimulationResult, SystemParams, WeatherSummary } from "./api";
import { api } from "./api";
import { HeatMap } from "./components/HeatMap";
import { AboutPage, HelpPage } from "./components/InfoPages";
import { CollectorForm, ModelPicker, SystemForm } from "./components/ParamsForm";
import { Metrics, ResultsTable } from "./components/ResultsTable";
import { TrendChart } from "./components/TrendChart";
import { WeatherPage } from "./components/WeatherPage";

type Page = "simulator" | "weather" | "about" | "help";
type View = "inputs" | "table" | "trends" | "heatmap";

const VIEWS: { id: View; label: string }[] = [
  { id: "inputs", label: "Inputs" },
  { id: "table", label: "Monthly table" },
  { id: "trends", label: "Trends" },
  { id: "heatmap", label: "Heat map" },
];

function hasInvalid(obj: object): boolean {
  return Object.values(obj).some((v) => (typeof v === "number" ? !Number.isFinite(v) : v && typeof v === "object" ? hasInvalid(v) : false));
}

export default function App() {
  const [page, setPage] = useState<Page>("simulator");
  const [view, setView] = useState<View>("inputs");
  const [models, setModels] = useState<ModelInfo[]>([]);
  const [weather, setWeather] = useState<WeatherSummary[]>([]);
  const [weatherId, setWeatherId] = useState("sample");
  const [model, setModel] = useState<ModelType>("dhw");
  const [collector, setCollector] = useState<CollectorParams | null>(null);
  const [system, setSystem] = useState<SystemParams | null>(null);
  const [result, setResult] = useState<SimulationResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [running, setRunning] = useState(false);

  const loadWeather = useCallback(() => api.weather().then(setWeather).catch((e) => setError(e.message)), []);

  useEffect(() => {
    Promise.all([api.models(), api.defaults()])
      .then(([m, d]) => { setModels(m); setCollector(d.collector); setSystem(d.system); })
      .catch((e) => setError(`Cannot reach the API: ${e.message}`));
    loadWeather();
  }, [loadWeather]);

  useEffect(() => {
    if (weather.length && !weather.some((w) => w.id === weatherId)) setWeatherId(weather[0].id);
  }, [weather, weatherId]);

  const run = async () => {
    if (!collector || !system) return;
    setRunning(true); setError(null);
    try {
      setResult(await api.simulate({ weather_id: weatherId, model, collector, system }));
      if (view === "inputs") setView("table");
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setRunning(false);
    }
  };

  const invalid = !collector || !system || hasInvalid(collector) || hasInvalid(system);

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">☀ Renewable Energy Systems</div>
        <nav>
          {(["simulator", "weather", "about", "help"] as Page[]).map((p) => (
            <button key={p} className={p === page ? "on" : ""} onClick={() => setPage(p)}>{p}</button>
          ))}
        </nav>
      </header>

      {page === "simulator" && (
        <div className="layout">
          <aside className="sidebar">
            {VIEWS.map((v) => (
              <button key={v.id} className={v.id === view ? "on" : ""} disabled={v.id !== "inputs" && !result}
                onClick={() => setView(v.id)}>{v.label}</button>
            ))}
            <div className="sidebar-run">
              <label className="field"><span>Weather</span>
                <select value={weatherId} onChange={(e) => setWeatherId(e.target.value)}>
                  {weather.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
                </select>
              </label>
              <button className="primary" disabled={running || invalid} onClick={run}>
                {running ? "Running…" : "Run simulation"}
              </button>
              {invalid && collector && <p className="error small">Fill in every input.</p>}
            </div>
          </aside>

          <main className="content">
            {error && <p className="error">{error}</p>}
            {result && view !== "inputs" && (
              <>
                <h2>{result.title}</h2>
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
          </main>
        </div>
      )}
      {page === "weather" && <main className="content wide"><WeatherPage weather={weather} onChange={loadWeather} /></main>}
      {page === "about" && <main className="content wide"><AboutPage /></main>}
      {page === "help" && <main className="content wide"><HelpPage /></main>}
    </div>
  );
}
