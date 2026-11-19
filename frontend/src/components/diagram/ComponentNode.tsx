import { Handle, Position, type Node, type NodeProps } from "@xyflow/react";
import { Glyph } from "./Glyph";
import { KIND_COLORS, fmtValue, useDiagram, type NodeData } from "./context";

export type ComponentFlowNode = Node<NodeData, "component">;

export function ComponentNode({ id, data, selected }: NodeProps<ComponentFlowNode>) {
  const { specs, result, issues, stale } = useDiagram();
  const spec = specs[data.ctype];
  if (!spec) return <div className="cnode error">Unknown component “{data.ctype}”</div>;
  const ins = spec.ports.filter((p) => p.direction === "in");
  const outs = spec.ports.filter((p) => p.direction === "out");
  const problems = issues.filter((i) => i.node_id === id);
  const summary = result?.diagram?.nodes?.[id]?.[spec.summary];
  const summaryOut = spec.outputs.find((o) => o.key === spec.summary);
  const rows = Math.max(ins.length, outs.length, 1);
  // Several outputs get labels, so start their handles below the title row.
  const top = outs.length > 1 ? 50 : 32;

  return (
    <div className={`cnode${selected ? " selected" : ""}${problems.length ? " has-issue" : ""}`}
      style={{ minHeight: top + 32 + (rows - 1) * 22, paddingBottom: outs.length > 1 ? 8 + (outs.length - 1) * 22 : undefined }} title={problems.map((p) => p.message).join("\n") || spec.description}>
      <div className="cnode-head">
        <span className="cnode-icon"><Glyph type={data.ctype} /></span>
        <div>
          <div className="cnode-title">{data.label || spec.title}</div>
          {data.label && <div className="cnode-type">{spec.title}</div>}
        </div>
      </div>
      {summaryOut && summary !== undefined && (
        <div className={`cnode-value${stale ? " stale" : ""}`}>
          <span>{summaryOut.label}</span> {fmtValue(summary, summaryOut.unit)}
        </div>
      )}
      {problems.length > 0 && <div className="cnode-issue">⚠ {problems[0].message}{problems.length > 1 ? ` (+${problems.length - 1})` : ""}</div>}
      {ins.map((p, i) => (
        <Handle key={p.id} id={p.id} type="target" position={Position.Left}
          style={{ top: top + i * 22, background: KIND_COLORS[p.kind] }} title={`${p.label} (in)`} />
      ))}
      {outs.map((p, i) => (
        <Handle key={p.id} id={p.id} type="source" position={Position.Right}
          style={{ top: top + i * 22, background: KIND_COLORS[p.kind] }} title={`${p.label} (out)`} />
      ))}
      {outs.length > 1 && outs.map((p, i) => (
        <span key={p.id} className="port-label right" style={{ top: top - 6 + i * 22 }}>{p.label}</span>
      ))}
    </div>
  );
}
