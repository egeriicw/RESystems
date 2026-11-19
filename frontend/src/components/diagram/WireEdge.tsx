import { BaseEdge, EdgeLabelRenderer, getSmoothStepPath, type Edge, type EdgeProps } from "@xyflow/react";
import { KIND_COLORS, fmtValue, useDiagram, type EdgeData } from "./context";

export type WireFlowEdge = Edge<EdgeData, "wire">;

export function WireEdge({ id, data, sourceX, sourceY, targetX, targetY, sourcePosition, targetPosition, selected }: EdgeProps<WireFlowEdge>) {
  const { result, issues, stale, open } = useDiagram();
  const [path, lx, ly] = getSmoothStepPath({ sourceX, sourceY, targetX, targetY, sourcePosition, targetPosition, borderRadius: 10 });
  const color = KIND_COLORS[data?.kind ?? ""] ?? "#64748b";
  const carried = result?.diagram?.edges?.[id];
  const problem = issues.find((i) => i.edge_id === id);
  return (
    <>
      <BaseEdge id={id} path={path} interactionWidth={18}
        style={{ stroke: problem ? "#dc2626" : color, strokeWidth: selected ? 4 : 2.5, strokeDasharray: problem ? "6 4" : undefined }} />
      <EdgeLabelRenderer>
        {(carried || problem) && (
          <button className={`wire-label nodrag nopan${problem ? " issue" : ""}${stale ? " stale" : ""}`}
            style={{ transform: `translate(-50%, -50%) translate(${lx}px, ${ly}px)`, borderColor: problem ? "#dc2626" : color }}
            title={problem?.message ?? carried?.label} onClick={() => open({ kind: "edge", id })}>
            {problem ? "⚠" : fmtValue(carried!.value, carried!.unit)}
          </button>
        )}
      </EdgeLabelRenderer>
    </>
  );
}
