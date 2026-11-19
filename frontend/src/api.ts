export type ModelType = "q_useful" | "house_heating" | "dhw" | "house_fan_coil" | "house_fan_coil_dhw";

export interface SeriesInfo { key: string; label: string; kind: "energy" | "temperature" | "irradiance"; unit: string }
export interface ModelInfo { id: ModelType; title: string; series: SeriesInfo[] }

export interface CollectorParams {
  tilt: number; azimuth: number; ground_reflectance: number; fr_ta: number; fr_ul: number;
  num_glazings: number; num_collectors: number; area_per_collector: number;
}
export interface SystemParams {
  collector_inlet_temp: number; flow_per_collector_gpm: number; storage_initial_temp: number;
  storage_mass: number; storage_max_temp: number; collector_hx_effectiveness: number; ua_house: number;
  ua_components: { walls: number; windows: number; ceiling: number; infiltration: number } | null;
  internal_gains: number; room_temp: number; supply_air_temp: number; fan_coil_effectiveness: number;
  fan_coil_capacitance: number; ground_water_temp: number; hot_water_supply_temp: number;
  heating_months: number[]; hot_water_profile: number[];
}
export interface WeatherSummary {
  id: string; name: string; source: string; latitude: number; longitude: number; time_zone: number;
  hours: number; monthly: { month: number; Ta: number; Ih: number }[];
}
export interface SimulationResult {
  model: ModelType | "diagram"; title: string;
  weather: { name: string; source: string; latitude: number; longitude: number; time_zone: number };
  inputs: { collector_area_m2: number; house_ua: number };
  series: SeriesInfo[];
  monthly: Record<string, number>[];
  annual: Record<string, number>;
  // system_efficiency is absent from runs stored before it was added.
  metrics: { collector_efficiency?: number | null; system_efficiency?: number | null; solar_fraction?: number | null };
  daily: Record<string, number[]>;
  hourly: Record<string, number[]> | null;
  notes: string[];
  // Diagram runs: annual outputs per component, energy carried per wire.
  diagram?: { nodes: Record<string, Record<string, number>>; edges: Record<string, { label: string; value: number; unit: string }> } | null;
}

export class ApiError extends Error {
  constructor(message: string, public status: number, public issues: DiagramIssue[] = []) { super(message); }
}

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(`/api${path}`, init);
  if (!res.ok) {
    let detail = res.statusText;
    let issues: DiagramIssue[] = [];
    try {
      const body = await res.json();
      if (typeof body.detail === "string") detail = body.detail;
      else if (body.detail?.issues) { detail = body.detail.message; issues = body.detail.issues; }
      else detail = JSON.stringify(body.detail);
    } catch { /* not JSON */ }
    throw new ApiError(detail, res.status, issues);
  }
  return res.status === 204 ? (undefined as T) : res.json();
}

const json = (body: unknown): RequestInit => ({
  method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
});

export const api = {
  models: () => request<ModelInfo[]>("/models"),
  defaults: () => request<{ collector: CollectorParams; system: SystemParams }>("/defaults"),
  weather: () => request<WeatherSummary[]>("/weather"),
  uploadWeather: (file: File) => {
    const form = new FormData();
    form.append("file", file);
    return request<WeatherSummary>("/weather/upload", { method: "POST", body: form });
  },
  syntheticWeather: (body: Record<string, unknown>) => request<WeatherSummary>("/weather/synthetic", json(body)),
  deleteWeather: (id: string) => request<void>(`/weather/${id}`, { method: "DELETE" }),
  simulate: (body: { weather_id: string; model: ModelType; collector: CollectorParams; system: SystemParams }) =>
    request<SimulationResult>("/simulate", json(body)),
};

export const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

export function fmt(v: number | null | undefined, digits = 0): string {
  if (v === null || v === undefined || Number.isNaN(v)) return "—";
  return v.toLocaleString(undefined, { minimumFractionDigits: digits, maximumFractionDigits: digits });
}

// ---- Sites, scenarios and stored runs ----

export interface SiteInput {
  name: string; latitude: number; longitude: number; time_zone?: number | null; address?: string | null;
  building_type?: string | null; floor_area_m2?: number | null; year_built?: number | null;
  occupants?: number | null; ua_house?: number | null; notes?: string | null; weather_id?: string | null;
}
export interface Site extends SiteInput {
  id: number; time_zone: number; created_at: string; scenario_count: number;
  weather_name: string | null; weather_distance_km: number | null;
}
export interface ScenarioInput {
  name: string; site_id: number | null; weather_id: string | null; model: ModelType | "diagram";
  collector: CollectorParams; system: SystemParams; diagram?: Diagram | null;
}
export interface RunSummary {
  id: number; scenario_id: number; created_at: string; title: string; weather_name: string;
  metrics: SimulationResult["metrics"]; annual: Record<string, number>;
}
export interface Scenario extends ScenarioInput {
  id: number; effective_weather_id: string; site_name: string | null; created_at: string; updated_at: string; latest_run: RunSummary | null;
}
export interface Run extends RunSummary { scenario_name: string; result: SimulationResult }

const send = (method: string, body: unknown): RequestInit => ({ ...json(body), method });

export const db = {
  sites: () => request<Site[]>("/sites"),
  createSite: (body: SiteInput) => request<Site>("/sites", json(body)),
  updateSite: (id: number, body: Partial<SiteInput>) => request<Site>(`/sites/${id}`, send("PATCH", body)),
  deleteSite: (id: number) => request<void>(`/sites/${id}`, { method: "DELETE" }),
  scenarios: () => request<Scenario[]>("/scenarios"),
  createScenario: (body: ScenarioInput) => request<Scenario>("/scenarios", json(body)),
  updateScenario: (id: number, body: ScenarioInput) => request<Scenario>(`/scenarios/${id}`, send("PUT", body)),
  deleteScenario: (id: number) => request<void>(`/scenarios/${id}`, { method: "DELETE" }),
  runScenario: (id: number) => request<Run>(`/scenarios/${id}/run`, { method: "POST" }),
  run: (id: number, hourly = true) => request<Run>(`/runs/${id}?hourly=${hourly}`),
};

export const pct = (v?: number | null) => (v === null || v === undefined ? "—" : `${(v * 100).toFixed(1)}%`);
export const shortDate = (iso: string) => new Date(iso).toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" });

// ---- One-line diagrams ----

export interface ParamSpec {
  key: string; label: string; type: "number" | "integer" | "select" | "profile" | "months";
  unit: string | null; default: unknown; min: number | null; max: number | null; step: number | null;
  options: { value: string | number; label: string }[] | null; help: string | null;
}
export interface PortSpec {
  id: string; label: string; direction: "in" | "out"; kind: string; max_connections: number | null; wire_params: ParamSpec[];
}
export interface ComponentSpec {
  type: string; title: string; category: string; description: string; ports: PortSpec[]; params: ParamSpec[];
  outputs: { key: string; label: string; kind: string; unit: string }[]; summary: string;
}
export interface ComponentLibrary { port_kinds: string[]; components: ComponentSpec[] }

export type Params = Record<string, unknown>;
export interface DiagramNode { id: string; type: string; label: string; position: { x: number; y: number }; params: Params }
export interface DiagramEdge { id: string; source: string; source_handle: string; target: string; target_handle: string; params: Params }
export interface Diagram { nodes: DiagramNode[]; edges: DiagramEdge[] }
export interface DiagramIssue { message: string; node_id: string | null; edge_id: string | null; param: string | null }
export interface DiagramTemplate { id: string; title: string; diagram: Diagram }

export const diagramApi = {
  library: () => request<ComponentLibrary>("/components"),
  templates: () => request<DiagramTemplate[]>("/diagram/templates"),
  validate: (diagram: Diagram) => request<{ issues: DiagramIssue[] }>("/diagram/validate", json({ diagram })),
  simulate: (weather_id: string, diagram: Diagram) => request<SimulationResult>("/diagram/simulate", json({ weather_id, diagram })),
};
