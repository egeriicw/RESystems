import { createContext, useContext } from "react";
import type { ComponentSpec, DiagramIssue, Params, SimulationResult } from "../../api";

export const KIND_COLORS: Record<string, string> = {
  solar_loop: "#ea580c",
  heating_loop: "#dc2626",
  hot_water: "#0891b2",
  heat: "#9333ea",
  electric: "#ca8a04",
};
export const KIND_LABELS: Record<string, string> = {
  solar_loop: "Collector loop",
  heating_loop: "Heating water",
  hot_water: "Hot water",
  heat: "Heat to space",
  electric: "Electric (future)",
};

export interface NodeData extends Record<string, unknown> { ctype: string; label: string; params: Params }
export interface EdgeData extends Record<string, unknown> { params: Params; kind: string }

export interface DiagramCtx {
  specs: Record<string, ComponentSpec>;
  result: SimulationResult | null;
  issues: DiagramIssue[];
  stale: boolean; // the result is from before the latest edit
  open: (target: { kind: "node" | "edge"; id: string }) => void;
}

export const Ctx = createContext<DiagramCtx | null>(null);
export const useDiagram = () => useContext(Ctx)!;

export function fmtValue(v: number | undefined, unit: string) {
  if (v === undefined || v === null || Number.isNaN(v)) return "—";
  if (unit === "C") return `${v.toFixed(1)} °C`;
  if (Math.abs(v) >= 1000) return `${(v / 1000).toLocaleString(undefined, { maximumFractionDigits: 1 })} MJ/day`;
  return `${v.toLocaleString(undefined, { maximumFractionDigits: 0 })} ${unit}`;
}
