import type { CollectorParams, ModelInfo, ModelType, SystemParams } from "../api";
import { MONTHS } from "../api";

type NumKey<T> = { [K in keyof T]: T[K] extends number ? K : never }[keyof T];

interface Field<T> { key: NumKey<T>; label: string; unit?: string; step?: number; models?: ModelType[] }

const HOUSE: ModelType[] = ["house_heating", "house_fan_coil", "house_fan_coil_dhw"];
const FAN: ModelType[] = ["house_fan_coil", "house_fan_coil_dhw"];
const DHW: ModelType[] = ["dhw", "house_fan_coil_dhw"];
const STORAGE: ModelType[] = ["house_heating", "dhw", "house_fan_coil", "house_fan_coil_dhw"];
const HX: ModelType[] = ["dhw", "house_fan_coil", "house_fan_coil_dhw"];

const COLLECTOR_FIELDS: Field<CollectorParams>[] = [
  { key: "tilt", label: "Tilt (β)", unit: "deg" },
  { key: "azimuth", label: "Azimuth (γ, 0 = south)", unit: "deg" },
  { key: "ground_reflectance", label: "Ground reflectance (ρg)", step: 0.05 },
  { key: "fr_ta", label: "FR(τα)", step: 0.005 },
  { key: "fr_ul", label: "FR·UL", unit: "kJ/hr·m²·C", step: 0.1 },
  { key: "num_glazings", label: "Glazings (1–2)" },
  { key: "num_collectors", label: "Number of collectors" },
  { key: "area_per_collector", label: "Area per collector", unit: "m²", step: 0.01 },
];

const SYSTEM_FIELDS: { group: string; models: ModelType[]; fields: Field<SystemParams>[] }[] = [
  { group: "Collector fluid", models: ["q_useful"], fields: [
    { key: "collector_inlet_temp", label: "Fluid entering collector", unit: "C" },
  ] },
  { group: "Collector loop", models: ["house_heating", ...HX], fields: [
    { key: "flow_per_collector_gpm", label: "Flow per collector", unit: "gpm", step: 0.01, models: ["house_heating", ...FAN] },
    { key: "collector_hx_effectiveness", label: "Collector HX effectiveness", step: 0.05, models: HX },
  ] },
  { group: "Thermal storage", models: STORAGE, fields: [
    { key: "storage_initial_temp", label: "Initial temperature", unit: "C" },
    { key: "storage_mass", label: "Water mass", unit: "kg" },
    { key: "storage_max_temp", label: "Max temperature", unit: "C" },
  ] },
  { group: "Load – house", models: HOUSE, fields: [
    { key: "ua_house", label: "House UA", unit: "kJ/hr·C" },
    { key: "internal_gains", label: "Internal gains", unit: "kJ/hr" },
    { key: "room_temp", label: "Set point", unit: "C" },
    { key: "supply_air_temp", label: "Min tank temp for fan coil", unit: "C", models: FAN },
  ] },
  { group: "Fan heat exchanger", models: FAN, fields: [
    { key: "fan_coil_effectiveness", label: "Effectiveness", step: 0.05 },
    { key: "fan_coil_capacitance", label: "Capacitance rate", unit: "kJ/hr·C" },
  ] },
  { group: "Load – water", models: DHW, fields: [
    { key: "ground_water_temp", label: "Ground water temperature", unit: "C" },
    { key: "hot_water_supply_temp", label: "Supply water temperature", unit: "C" },
  ] },
];

function NumberInput({ label, unit, value, step, onChange }: {
  label: string; unit?: string; value: number; step?: number; onChange: (v: number) => void;
}) {
  return (
    <label className="field">
      <span>{label}{unit && <em> ({unit})</em>}</span>
      <input type="number" value={Number.isFinite(value) ? value : ""} step={step ?? "any"}
        onChange={(e) => onChange(e.target.value === "" ? NaN : Number(e.target.value))} />
    </label>
  );
}

export function ModelPicker({ models, value, onChange }: {
  models: ModelInfo[]; value: ModelType; onChange: (m: ModelType) => void;
}) {
  return (
    <fieldset>
      <legend>Model type</legend>
      {models.map((m) => (
        <label key={m.id} className="radio">
          <input type="radio" name="model" checked={value === m.id} onChange={() => onChange(m.id)} />
          {m.title}
        </label>
      ))}
    </fieldset>
  );
}

export function CollectorForm({ value, onChange }: { value: CollectorParams; onChange: (v: CollectorParams) => void }) {
  return (
    <fieldset>
      <legend>Collectors & placement</legend>
      {COLLECTOR_FIELDS.map((f) => (
        <NumberInput key={f.key} label={f.label} unit={f.unit} step={f.step} value={value[f.key]}
          onChange={(v) => onChange({ ...value, [f.key]: v })} />
      ))}
      <p className="hint">Total area: {(value.num_collectors * value.area_per_collector).toFixed(2)} m². The original program used 2 collectors for hot water and 8 for house heating.</p>
    </fieldset>
  );
}

export function SystemForm({ model, value, onChange }: {
  model: ModelType; value: SystemParams; onChange: (v: SystemParams) => void;
}) {
  const useComponents = value.ua_components !== null;
  const ua = value.ua_components ?? { walls: 0, windows: 0, ceiling: 0, infiltration: 0 };
  return (
    <>
      {SYSTEM_FIELDS.filter((g) => g.models.includes(model)).map((g) => (
        <fieldset key={g.group}>
          <legend>{g.group}</legend>
          {g.group === "Load – house" && (
            <div className="toggle">
              <label className="radio"><input type="radio" checked={!useComponents}
                onChange={() => onChange({ ...value, ua_components: null })} />Total UA</label>
              <label className="radio"><input type="radio" checked={useComponents}
                onChange={() => onChange({ ...value, ua_components: ua })} />Individual UA values</label>
            </div>
          )}
          {g.fields.filter((f) => !f.models || f.models.includes(model)).map((f) => {
            if (f.key === "ua_house" && useComponents) {
              return (["walls", "windows", "ceiling", "infiltration"] as const).map((k) => (
                <NumberInput key={k} label={`UA ${k}`} unit="kJ/hr·C" value={ua[k]}
                  onChange={(v) => onChange({ ...value, ua_components: { ...ua, [k]: v } })} />
              ));
            }
            return (
              <NumberInput key={f.key} label={f.label} unit={f.unit} step={f.step} value={value[f.key]}
                onChange={(v) => onChange({ ...value, [f.key]: v })} />
            );
          })}
          {g.group === "Load – house" && (
            <div className="months">
              <span>Heating months</span>
              <div>
                {MONTHS.map((m, i) => {
                  const on = value.heating_months.includes(i + 1);
                  return (
                    <button type="button" key={m} className={on ? "chip on" : "chip"} onClick={() => onChange({
                      ...value,
                      heating_months: on ? value.heating_months.filter((x) => x !== i + 1) : [...value.heating_months, i + 1],
                    })}>{m}</button>
                  );
                })}
              </div>
            </div>
          )}
        </fieldset>
      ))}
    </>
  );
}
