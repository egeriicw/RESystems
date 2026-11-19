import type { ParamSpec } from "../../api";
import { MONTHS } from "../../api";

/** One editable parameter, rendered from its schema. Numbers are kept as strings while typing. */
export function ParamField({ spec, value, onChange, error }: {
  spec: ParamSpec; value: unknown; onChange: (v: unknown) => void; error?: string;
}) {
  const label = <span>{spec.label}{spec.unit && <em> ({spec.unit})</em>}</span>;
  const hint = (spec.help || error) && <small className={error ? "field-error" : "field-help"}>{error ?? spec.help}</small>;

  if (spec.type === "select") {
    return (
      <div className="pfield">
        <label className="field">{label}
          <select value={String(value)} onChange={(e) => {
            const opt = spec.options!.find((o) => String(o.value) === e.target.value)!;
            onChange(opt.value);
          }}>
            {spec.options!.map((o) => <option key={String(o.value)} value={String(o.value)}>{o.label}</option>)}
          </select>
        </label>{hint}
      </div>
    );
  }
  if (spec.type === "months") {
    const months = (value as number[]) ?? [];
    return (
      <div className="pfield">
        <div className="months"><span>{spec.label}</span>
          <div>{MONTHS.map((m, i) => {
            const on = months.includes(i + 1);
            return <button type="button" key={m} className={on ? "chip on" : "chip"} aria-pressed={on}
              onClick={() => onChange(on ? months.filter((x) => x !== i + 1) : [...months, i + 1].sort((a, b) => a - b))}>{m}</button>;
          })}</div>
        </div>{hint}
      </div>
    );
  }
  if (spec.type === "profile") {
    const values = (value as (number | string)[]) ?? [];
    return (
      <div className="pfield">
        <span className="profile-title">{label}</span>
        <div className="profile-grid">
          {values.map((v, i) => (
            <label key={i}><small>{i + 1}</small>
              <input type="number" step="any" min={0} value={String(v)} aria-label={`Hour ${i + 1}`}
                onChange={(e) => onChange(values.map((x, k) => (k === i ? e.target.value : x)))} />
            </label>
          ))}
        </div>{hint}
      </div>
    );
  }
  return (
    <div className="pfield">
      <label className="field">{label}
        <input type="number" step={spec.step ?? "any"} min={spec.min ?? undefined} max={spec.max ?? undefined}
          value={value === undefined || value === null ? "" : String(value)} onChange={(e) => onChange(e.target.value)} />
      </label>{hint}
    </div>
  );
}

/** Convert edited strings back to numbers and check them against the schema. */
export function parseParams(specs: ParamSpec[], draft: Record<string, unknown>): { values: Record<string, unknown>; errors: Record<string, string> } {
  const values: Record<string, unknown> = {};
  const errors: Record<string, string> = {};
  const num = (v: unknown) => (typeof v === "number" ? v : v === "" || v === null || v === undefined ? NaN : Number(v));
  for (const p of specs) {
    const v = draft[p.key];
    if (p.type === "number" || p.type === "integer") {
      const n = num(v);
      if (!Number.isFinite(n)) errors[p.key] = "Enter a number";
      else if (p.type === "integer" && !Number.isInteger(n)) errors[p.key] = "Enter a whole number";
      else if (p.min !== null && n < p.min) errors[p.key] = `At least ${p.min}`;
      else if (p.max !== null && n > p.max) errors[p.key] = `At most ${p.max}`;
      values[p.key] = n;
    } else if (p.type === "profile") {
      const arr = ((v as unknown[]) ?? []).map(num);
      if (arr.length !== 24 || arr.some((x) => !Number.isFinite(x) || x < 0)) errors[p.key] = "Enter 24 values, each 0 or more";
      values[p.key] = arr;
    } else {
      values[p.key] = v;
    }
  }
  return { values, errors };
}
