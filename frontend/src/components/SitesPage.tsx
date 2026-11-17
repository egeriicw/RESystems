import { useEffect, useState } from "react";
import { CircleMarker, MapContainer, TileLayer, Tooltip, useMap, useMapEvents } from "react-leaflet";
import "leaflet/dist/leaflet.css";
import type { Scenario, Site, SiteInput, WeatherSummary } from "../api";
import { db, fmt, pct } from "../api";

type Draft = { [K in keyof SiteInput]-?: string };
const EMPTY: Draft = {
  name: "", latitude: "", longitude: "", time_zone: "", address: "", building_type: "", floor_area_m2: "",
  year_built: "", occupants: "", ua_house: "", notes: "", weather_id: "",
};
const NUMERIC: (keyof SiteInput)[] = ["latitude", "longitude", "time_zone", "floor_area_m2", "year_built", "occupants", "ua_house"];
const FIELDS: { key: keyof SiteInput; label: string; unit?: string }[] = [
  { key: "name", label: "Name" },
  { key: "address", label: "Address" },
  { key: "latitude", label: "Latitude", unit: "°N" },
  { key: "longitude", label: "Longitude", unit: "°E" },
  { key: "time_zone", label: "Time zone", unit: "hrs from UTC" },
  { key: "building_type", label: "Building type" },
  { key: "floor_area_m2", label: "Floor area", unit: "m²" },
  { key: "year_built", label: "Year built" },
  { key: "occupants", label: "Occupants" },
  { key: "ua_house", label: "Heat loss (UA)", unit: "kJ/hr·C" },
];

function toDraft(s: Site): Draft {
  const d = { ...EMPTY };
  for (const k of Object.keys(EMPTY) as (keyof SiteInput)[]) d[k] = s[k] === null || s[k] === undefined ? "" : String(s[k]);
  return d;
}

function toInput(d: Draft): SiteInput {
  const out: Record<string, unknown> = {};
  for (const k of Object.keys(EMPTY) as (keyof SiteInput)[]) {
    const v = d[k].trim();
    out[k] = v === "" ? null : NUMERIC.includes(k) ? Number(v) : v;
  }
  return out as unknown as SiteInput;
}

function ClickToPlace({ onPick }: { onPick: (lat: number, lon: number) => void }) {
  useMapEvents({ click: (e) => onPick(e.latlng.lat, e.latlng.lng) });
  return null;
}

function FlyTo({ lat, lon }: { lat: number; lon: number }) {
  const map = useMap();
  useEffect(() => { map.setView([lat, lon], Math.max(map.getZoom(), 9)); }, [map, lat, lon]);
  return null;
}

export function SitesPage({ sites, weather, scenarios, onChange, onNewScenario, onOpenScenario }: {
  sites: Site[]; weather: WeatherSummary[]; scenarios: Scenario[]; onChange: () => void;
  onNewScenario: (site: Site) => void; onOpenScenario: (s: Scenario) => void;
}) {
  const [selectedId, setSelectedId] = useState<number | "new" | null>(sites[0]?.id ?? null);
  const [draft, setDraft] = useState<Draft>(EMPTY);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const selected = typeof selectedId === "number" ? sites.find((s) => s.id === selectedId) : undefined;

  useEffect(() => { if (selected) setDraft(toDraft(selected)); }, [selected]);
  useEffect(() => { if (selectedId === null && sites.length) setSelectedId(sites[0].id); }, [sites, selectedId]);

  const lat = Number(draft.latitude), lon = Number(draft.longitude);
  const hasPoint = draft.latitude !== "" && draft.longitude !== "" && Number.isFinite(lat) && Number.isFinite(lon);
  const valid = draft.name.trim() !== "" && hasPoint && Math.abs(lat) <= 90 && Math.abs(lon) <= 180;

  const act = async (fn: () => Promise<Site | void>) => {
    setBusy(true); setError(null);
    try {
      const s = await fn();
      onChange();
      setSelectedId(s ? s.id : null);
    } catch (e) { setError((e as Error).message); } finally { setBusy(false); }
  };

  const save = () => act(() => {
    const body = toInput(draft);
    if (selectedId === "new") return db.createSite(body);
    // Leaving these blank means "keep the current value" when editing.
    if (body.time_zone === null) delete body.time_zone;
    return db.updateSite(selectedId as number, body);
  });

  const pick = (la: number, lo: number) => {
    if (selectedId === null) { setSelectedId("new"); setDraft({ ...EMPTY, latitude: la.toFixed(4), longitude: lo.toFixed(4) }); }
    else setDraft((d) => ({ ...d, latitude: la.toFixed(4), longitude: lo.toFixed(4) }));
  };

  const siteScenarios = selected ? scenarios.filter((s) => s.site_id === selected.id) : [];
  const center: [number, number] = sites.length ? [sites[0].latitude, sites[0].longitude] : [39.9, -84.2];

  return (
    <div className="sites-page">
      <section className="card map-card">
        <div className="toolbar">
          <h2>Sites</h2>
          <button className="primary" onClick={() => { setSelectedId("new"); setDraft(EMPTY); }}>New site</button>
          <span className="hint">Click the map to set the location of the site being edited.</span>
        </div>
        <MapContainer center={center} zoom={sites.length ? 8 : 5} className="map">
          <TileLayer url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
            attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors' />
          <ClickToPlace onPick={pick} />
          {hasPoint && <FlyTo lat={lat} lon={lon} />}
          {weather.map((w) => (
            <CircleMarker key={w.id} center={[w.latitude, w.longitude]} radius={5}
              pathOptions={{ color: "#2563eb", fillOpacity: 0.3, weight: 1 }}>
              <Tooltip>Weather: {w.name}</Tooltip>
            </CircleMarker>
          ))}
          {sites.map((s) => (
            <CircleMarker key={s.id} center={[s.latitude, s.longitude]} radius={9}
              pathOptions={{ color: "#b45309", fillColor: "#f59e0b", fillOpacity: s.id === selectedId ? 0.95 : 0.55, weight: s.id === selectedId ? 3 : 1 }}
              eventHandlers={{ click: () => setSelectedId(s.id) }}>
              <Tooltip>{s.name}</Tooltip>
            </CircleMarker>
          ))}
          {hasPoint && (selectedId === "new" || (selected && (lat !== selected.latitude || lon !== selected.longitude))) && (
            <CircleMarker center={[lat, lon]} radius={9} pathOptions={{ color: "#dc2626", dashArray: "4", fillOpacity: 0.1 }}>
              <Tooltip>Unsaved location</Tooltip>
            </CircleMarker>
          )}
        </MapContainer>
        <p className="hint"><span className="dot site" /> site &nbsp; <span className="dot wx" /> weather dataset</p>
      </section>

      <div className="two-col">
        <section className="card">
          <h3>All sites</h3>
          {sites.length === 0 && <p className="hint">No sites yet. Click “New site”, then click the map.</p>}
          <table>
            <tbody>
              {sites.map((s) => (
                <tr key={s.id} className={s.id === selectedId ? "selected clickable" : "clickable"} onClick={() => setSelectedId(s.id)}>
                  <td><strong>{s.name}</strong><small className="sub">{s.building_type ?? "—"} · {s.floor_area_m2 ? `${fmt(s.floor_area_m2)} m²` : "area not set"}</small></td>
                  <td>{s.scenario_count} scenario{s.scenario_count === 1 ? "" : "s"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </section>

        {selectedId !== null && (
          <section className="card">
            <h3>{selectedId === "new" ? "New site" : selected?.name}</h3>
            {FIELDS.map((f) => (
              <label className="field" key={f.key}><span>{f.label}{f.unit && <em> ({f.unit})</em>}</span>
                <input type={NUMERIC.includes(f.key) ? "number" : "text"} step="any" value={draft[f.key]}
                  placeholder={f.key === "time_zone" ? "from longitude" : ""}
                  onChange={(e) => setDraft({ ...draft, [f.key]: e.target.value })} />
              </label>
            ))}
            <label className="field"><span>Weather dataset</span>
              <select value={draft.weather_id} onChange={(e) => setDraft({ ...draft, weather_id: e.target.value })}>
                <option value="">{selectedId === "new" ? "Nearest (automatic)" : "None"}</option>
                {weather.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
              </select>
            </label>
            {selected?.weather_distance_km != null && (
              <p className="hint">{selected.weather_name} is {fmt(selected.weather_distance_km, 1)} km from this site.</p>
            )}
            <label className="field notes"><span>Notes</span>
              <textarea rows={2} value={draft.notes} onChange={(e) => setDraft({ ...draft, notes: e.target.value })} />
            </label>
            {error && <p className="error">{error}</p>}
            <div className="actions">
              <button className="primary" disabled={busy || !valid} onClick={save}>{selectedId === "new" ? "Create site" : "Save changes"}</button>
              {selected && <button className="secondary" onClick={() => onNewScenario(selected)}>New scenario for this site</button>}
              {selected && (
                <button className="link danger" disabled={busy} onClick={() => {
                  if (window.confirm(`Delete ${selected.name} and its ${selected.scenario_count} scenario(s)?`)) act(() => db.deleteSite(selected.id));
                }}>Delete site</button>
              )}
            </div>

            {selected && (
              <>
                <h4>Scenarios</h4>
                {siteScenarios.length === 0 && <p className="hint">No scenarios saved for this site.</p>}
                <table>
                  <tbody>
                    {siteScenarios.map((sc) => (
                      <tr key={sc.id}>
                        <td><button className="link" onClick={() => onOpenScenario(sc)}>{sc.name}</button></td>
                        <td>{sc.latest_run ? `solar fraction ${pct(sc.latest_run.metrics.solar_fraction)}` : "not run"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </>
            )}
          </section>
        )}
      </div>
    </div>
  );
}
