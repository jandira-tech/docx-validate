/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Named checks. Each check has a stable id, a description, the package
 * formats it applies to, the checks it depends on, and a runner that reads a
 * shared {@link ValidationContext}. The registry validates the definitions
 * once and resolves selections to an ordered list that includes every
 * prerequisite.
 */

import type { PackageFormat, ValidationIssue } from "./types";
import type { ValidationContext } from "./validation-context";

/**
 * Every declared group name, used by `--checks` and `--list-checks`. A group
 * with no registered check yet is still a known name: resolving it selects
 * nothing instead of raising {@link UnknownCheckError}, so a preset can name a
 * whole family (say `content-types`) before its checks land in the registry.
 */
export const CHECK_GROUPS = [
    "package",
    "xml",
    "content-types",
    "relationships",
    "tables",
    "drawings",
    "story",
    "comments",
    "tracked-changes",
    "math",
    "settings",
    "schema",
] as const;

/** A named check group. Derived from {@link CHECK_GROUPS}, so the list and the type cannot drift. */
export type CheckGroup = (typeof CHECK_GROUPS)[number];

export type CheckDefinition = {
    readonly id: string;
    readonly description: string;
    readonly groups: readonly CheckGroup[];
    readonly formats: readonly PackageFormat[];
    /** Checks that must pass first. Each must be registered before this one and apply to every format this check applies to. */
    readonly requires?: readonly string[];
    /**
     * A failure excludes the parts its error findings name from its
     * dependents, instead of skipping them.
     */
    readonly gatesByPart?: boolean;
    /** Needs the packed archive bytes; unavailable for unpacked directories. */
    readonly archiveOnly?: boolean;
    readonly run: (ctx: ValidationContext) => readonly ValidationIssue[] | Promise<readonly ValidationIssue[]>;
};

/** A selection named checks (or groups) the registry does not know. */
export class UnknownCheckError extends Error {
    public readonly ids: readonly string[];

    public constructor(ids: readonly string[]) {
        super(`Unknown check id(s): ${ids.join(", ")}`);
        this.name = "UnknownCheckError";
        this.ids = ids;
    }
}

export class CheckRegistry {
    private readonly definitions: readonly CheckDefinition[];
    private readonly byId = new Map<string, CheckDefinition>();

    public constructor(definitions: readonly CheckDefinition[]) {
        for (const def of definitions) {
            if (this.byId.has(def.id)) {
                throw new Error(`Check registry: duplicate check id "${def.id}"`);
            }
            for (const prerequisite of def.requires ?? []) {
                const prior = this.byId.get(prerequisite);
                if (prior === undefined) {
                    const later = definitions.some((other) => other.id === prerequisite);
                    throw new Error(
                        later
                            ? `Check registry: "${def.id}" requires "${prerequisite}", which must be registered before it`
                            : `Check registry: "${def.id}" requires unknown check "${prerequisite}"`,
                    );
                }
                // A prerequisite that never runs for one of this check's
                // formats would leave the check permanently skipped for that
                // format, silently reporting a valid run with nothing
                // executed (issue #299) — a configuration error, like an
                // unknown id.
                const uncovered = def.formats.filter((format) => !prior.formats.includes(format));
                if (uncovered.length > 0) {
                    throw new Error(
                        `Check registry: "${def.id}" requires "${prerequisite}", which does not apply to its format(s): ${uncovered.join(", ")}`,
                    );
                }
            }
            this.byId.set(def.id, def);
        }
        this.definitions = definitions;
    }

    /** Registered ids in registration order. */
    public get ids(): readonly string[] {
        return this.definitions.map((def) => def.id);
    }

    public get(id: string): CheckDefinition | undefined {
        return this.byId.get(id);
    }

    /** Definitions applicable to `format` (all definitions when omitted), in registration order. */
    public list(format?: PackageFormat): readonly CheckDefinition[] {
        return format === undefined ? this.definitions : this.definitions.filter((def) => def.formats.includes(format));
    }

    /**
     * Resolve check ids and group names to the ids to run, prerequisites
     * included, in registration order. Throws {@link UnknownCheckError} for any
     * name that is neither a check id nor a group with at least one check.
     *
     * With `format`, a group contributes only its checks for that format, and
     * an explicitly named check of another format is a configuration error.
     */
    public resolve(selection: readonly string[], format?: PackageFormat): readonly string[] {
        if (selection.length === 0) {
            throw new Error("Check selection is empty: name at least one check id or group");
        }
        const applies = (def: CheckDefinition): boolean => format === undefined || def.formats.includes(format);
        const wanted = new Set<string>();
        const unknown: string[] = [];
        const inapplicable: string[] = [];
        for (const name of selection) {
            const def = this.byId.get(name);
            if (def) {
                if (applies(def)) {
                    wanted.add(name);
                } else {
                    inapplicable.push(name);
                }
                continue;
            }
            const members = this.definitions.filter((member) => (member.groups as readonly string[]).includes(name));
            if (members.length === 0 && !(CHECK_GROUPS as readonly string[]).includes(name)) {
                unknown.push(name);
            }
            for (const member of members) {
                if (applies(member)) {
                    wanted.add(member.id);
                }
            }
        }
        if (unknown.length > 0) {
            throw new UnknownCheckError(unknown);
        }
        if (inapplicable.length > 0) {
            throw new Error(inapplicable.map((id) => `Check ${id} does not apply to ${format ?? "this"} packages`).join("; "));
        }
        // Prerequisites are registered before their dependents, so walking the
        // definitions backwards visits every dependent before what it needs.
        for (let i = this.definitions.length - 1; i >= 0; i -= 1) {
            const def = this.definitions[i];
            if (def && wanted.has(def.id)) {
                for (const prerequisite of def.requires ?? []) {
                    wanted.add(prerequisite);
                }
            }
        }
        return this.ids.filter((id) => wanted.has(id));
    }
}
