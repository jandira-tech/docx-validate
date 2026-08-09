import type { ValidationIssue } from "../../../lib/types";
import { counterKey, severityClassFor } from "./docx-diagnostics";
import type { DocxSemanticInventory, InventorySeverityClass } from "./docx-diagnostics";

export interface DocxInventoryDelta {
  key: string;
  path: string;
  category: string;
  label: string;
  unit: string;
  before: number;
  after: number;
}

export interface DocxInventoryDiff {
  added: DocxInventoryDelta[];
  removed: DocxInventoryDelta[];
  changed: DocxInventoryDelta[];
  unchangedCount: number;
}

function sortDeltas(a: DocxInventoryDelta, b: DocxInventoryDelta): number {
  return (
    a.path.localeCompare(b.path) ||
    a.category.localeCompare(b.category) ||
    a.label.localeCompare(b.label)
  );
}

export function diffDocxInventories(
  before: DocxSemanticInventory,
  after: DocxSemanticInventory,
): DocxInventoryDiff {
  const added: DocxInventoryDelta[] = [];
  const removed: DocxInventoryDelta[] = [];
  const changed: DocxInventoryDelta[] = [];
  let unchangedCount = 0;

  for (const b of before.counters.values()) {
    const key = counterKey(b.path, b.category, b.label, b.unit);
    const a = after.counters.get(key);
    const afterCount = a?.count ?? 0;
    const delta: DocxInventoryDelta = {
      key,
      path: b.path,
      category: b.category,
      label: b.label,
      unit: b.unit,
      before: b.count,
      after: afterCount,
    };
    if (afterCount === 0) removed.push(delta);
    else if (afterCount !== b.count) changed.push(delta);
    else unchangedCount += 1;
  }
  for (const a of after.counters.values()) {
    const key = counterKey(a.path, a.category, a.label, a.unit);
    if (before.counters.has(key)) continue;
    added.push({
      key,
      path: a.path,
      category: a.category,
      label: a.label,
      unit: a.unit,
      before: 0,
      after: a.count,
    });
  }
  added.sort(sortDeltas);
  removed.sort(sortDeltas);
  changed.sort(sortDeltas);
  return { added, removed, changed, unchangedCount };
}

type Direction = "decrease" | "increase";

interface TierResult {
  severity: ValidationIssue["severity"];
  code: string;
}

export function severityFor(category: string, direction: Direction): TierResult {
  const cls: InventorySeverityClass = severityClassFor(category);
  switch (cls) {
    case "content":
      return direction === "decrease"
        ? { severity: "error", code: "inventory-content-loss" }
        : { severity: "warning", code: "inventory-content-added" };
    case "table-shape":
    case "section-geometry":
    case "image-shape":
      return { severity: "warning", code: "inventory-shape-change" };
    case "formatting":
      return direction === "decrease"
        ? { severity: "warning", code: "inventory-formatting-drift" }
        : { severity: "info", code: "inventory-formatting-drift" };
    case "atomic-marks":
      return direction === "decrease"
        ? { severity: "warning", code: "inventory-mark-drift" }
        : { severity: "info", code: "inventory-mark-drift" };
    case "bookkeeping":
      return direction === "decrease"
        ? { severity: "warning", code: "inventory-bookkeeping-drift" }
        : { severity: "info", code: "inventory-bookkeeping-drift" };
  }
}

function directionOf(d: DocxInventoryDelta): Direction {
  return d.after < d.before ? "decrease" : "increase";
}

export function inventoryDiffToIssues(diff: DocxInventoryDiff): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  const emit = (d: DocxInventoryDelta, direction: Direction): void => {
    const tier = severityFor(d.category, direction);
    issues.push({
      severity: tier.severity,
      code: tier.code,
      path: d.path,
      message: `${d.category} '${d.label}': ${d.before} → ${d.after} (${d.unit}).`,
    });
  };
  for (const d of diff.removed) emit(d, "decrease");
  for (const d of diff.added) emit(d, "increase");
  for (const d of diff.changed) emit(d, directionOf(d));
  return issues;
}

const SEV_PREFIX: Record<ValidationIssue["severity"], string> = {
  error: "🔴",
  warning: "🟠",
  info: "⚪",
};

function renderSection(title: string, deltas: DocxInventoryDelta[]): string[] {
  if (deltas.length === 0) return [];
  const lines = [`## ${title}`, ""];
  let lastPath = "";
  for (const d of deltas) {
    if (d.path !== lastPath) {
      lines.push(`### ${d.path}`);
      lastPath = d.path;
    }
    // directionOf classifies removed (→decrease), added (→increase), and
    // changed deltas alike, so the per-line severity is always correct.
    const tier = severityFor(d.category, directionOf(d));
    lines.push(
      `- ${SEV_PREFIX[tier.severity]} ${d.category} \`${d.label}\`: ${d.before} → ${d.after} (${d.unit})`,
    );
  }
  lines.push("");
  return lines;
}

const SHAPE_CATEGORIES = new Set(["table shape", "section geometry", "image shape"]);

function renderReshaped(diff: DocxInventoryDiff): string[] {
  // Display-only pairing: within each (path, category) of a shape class, zip
  // removed shape-keys with added shape-keys. No identity guarantee.
  const groups = new Map<string, { removed: DocxInventoryDelta[]; added: DocxInventoryDelta[] }>();
  const push = (d: DocxInventoryDelta, side: "removed" | "added"): void => {
    if (!SHAPE_CATEGORIES.has(d.category)) return;
    const k = `${d.path}\u0000${d.category}`;
    const g = groups.get(k) ?? { removed: [], added: [] };
    g[side].push(d);
    groups.set(k, g);
  };
  for (const d of diff.removed) push(d, "removed");
  for (const d of diff.added) push(d, "added");
  const lines: string[] = [];
  for (const [k, g] of [...groups.entries()].sort(([a], [b]) => a.localeCompare(b))) {
    if (g.removed.length === 0 || g.added.length === 0) continue;
    const [pathValue, category] = k.split("\u0000");
    if (lines.length === 0) lines.push("## Reshaped", "");
    lines.push(`### ${pathValue} — ${category}`);
    const n = Math.max(g.removed.length, g.added.length);
    for (let i = 0; i < n; i += 1) {
      lines.push(`- 🟠 \`${g.removed[i]?.label ?? "—"}\` → \`${g.added[i]?.label ?? "—"}\``);
    }
  }
  if (lines.length > 0) lines.push("");
  return lines;
}

export function formatInventoryDiffMarkdown(diff: DocxInventoryDiff): string {
  const issues = inventoryDiffToIssues(diff);
  const errorCount = issues.filter((i) => i.severity === "error").length;
  const warnCount = issues.filter((i) => i.severity === "warning").length;
  const lines: string[] = ["# DOCX inventory diff", ""];
  lines.push(
    `**Summary:** ${diff.added.length} added, ${diff.removed.length} removed, ${diff.changed.length} changed, ${diff.unchangedCount} unchanged; ${errorCount} error / ${warnCount} warn`,
    "",
  );
  lines.push(...renderReshaped(diff));
  lines.push(...renderSection("Removed", diff.removed));
  lines.push(...renderSection("Added", diff.added));
  lines.push(...renderSection("Changed", diff.changed));
  return lines.join("\n");
}
