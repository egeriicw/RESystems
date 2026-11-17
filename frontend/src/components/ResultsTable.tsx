import type { SimulationResult } from "../api";
import { MONTHS, fmt } from "../api";

export function ResultsTable({ result }: { result: SimulationResult }) {
  const cols = result.series;
  const digits = (kind: string) => (kind === "temperature" ? 1 : 0);
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Month</th>
            {cols.map((c) => <th key={c.key} title={c.label}>{c.key}<small>{c.unit}</small></th>)}
          </tr>
        </thead>
        <tbody>
          {result.monthly.map((row) => (
            <tr key={row.month}>
              <td>{MONTHS[row.month - 1]}</td>
              {cols.map((c) => <td key={c.key}>{fmt(row[c.key], digits(c.kind))}</td>)}
            </tr>
          ))}
        </tbody>
        <tfoot>
          <tr>
            <td>Avg.</td>
            {cols.map((c) => <td key={c.key}>{fmt(result.annual[c.key], digits(c.kind))}</td>)}
          </tr>
        </tfoot>
      </table>
      <p className="hint">Energy and radiation columns are average daily totals for each month; temperatures are means.</p>
    </div>
  );
}

export function Metrics({ result }: { result: SimulationResult }) {
  const pct = (v?: number | null) => (v === null || v === undefined ? "—" : `${(v * 100).toFixed(1)}%`);
  return (
    <div className="metrics">
      <div><span>Collector efficiency</span><strong>{pct(result.metrics.collector_efficiency)}</strong></div>
      {"solar_fraction" in result.metrics && (
        <div><span>Solar fraction</span><strong>{pct(result.metrics.solar_fraction)}</strong></div>
      )}
      <div><span>Collector area</span><strong>{fmt(result.inputs.collector_area_m2, 2)} m²</strong></div>
      <div><span>Weather</span><strong>{result.weather.name}</strong></div>
    </div>
  );
}
