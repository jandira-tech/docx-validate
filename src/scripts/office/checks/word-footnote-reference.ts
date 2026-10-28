/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * A w:footnoteReference whose w:id is not a w:footnote w:id makes
 * Word offer recovery (fixture 1829). One such reference is enough.
 * Removing it opens clean once text-box deleted text is also gone.
 * A reference that names a present footnote opens clean. Duplicate
 * footnote ids are not this defect. The same mechanism judges endnotes:
 * Word session 58 (s58a) offers recovery for a reference to endnote 1 when
 * the part holds only endnote 5; the matching-id calibration opens clean.
 *
 * The defined notes come only from the parts a `footnotes`
 * (respectively `endnotes`) relationship of the main document, or of a
 * reached glossary, names, whatever the part is called: Word session 60
 * (s60a) opened clean a used notes part stored as `word/notes-foot.xml`
 * and `word/notes-end.xml` and related by its relationship, which the
 * name-only reading rejected (#793), and session 61 (s61a, s61b) offered
 * recovery for a body reference whose note 1 sat only in a
 * `word/footnotes.xml` / `word/endnotes.xml` no relationship names — alone
 * or beside a related part holding note 5 — while the same unrelated part
 * with nothing referencing it opened clean: Word defines notes only from
 * the parts their relationships name, so the former basename reading is
 * gone (#797). But a part the content-types stream types
 * `application/octet-stream` defines no note, whatever it is called:
 * session 63 (s63c) offered recovery for a body footnote reference whose
 * note sat in `word/notes.bin` typed octet-stream by its Override, and
 * session 64 (s64b) repeated that refusal for the same row and for
 * `word/notes.xml` and `word/footnotes.xml` alike — the type, not the name,
 * makes Word refuse the part (#829); its endnotes twin is unprobed. The
 * media type decides, parameters stripped and ASCII-lowercased as every
 * content-type judgment resolves it; case and parameter spellings on a
 * notes part, and every other unrecognised type, are unprobed. The
 * relationship walks are ungated by usage, so a related-but-unused part
 * defines its ids too — this set only adds definitions, never a rejection.
 * A `<kind>s` relationship target that escapes the package resolves by
 * stopping its `..` segments at the package root, as RFC 3986's
 * remove_dot_segments does: session 63 (s63c) opened `word/`'s
 * `../../footnotes.xml` clean with the note at `/footnotes.xml` while
 * word-valid rejected it, and session 64 (s64a) confirmed the reading — the
 * `../../../footnotes.xml` twin and the endnotes twin `../../endnotes.xml`
 * opened clean with their notes at the root, while the same escaping target
 * with the note only in `word/footnotes.xml` got the recovery offer and
 * stays rejected (#830). An External `<kind>s` relationship names no part,
 * so a body reference whose note sits in an unreferenced
 * `word/footnotes.xml` blocks, matching session 63 (s63c)'s recovery offer.
 * Every relationship the read model walks resolves an escaping target the
 * same way, as Word read the styles, header, customXml, comments and
 * officeDocument climbs of sessions 75 to 78 (#986). References read
 * through the main document use only its own relationships' definitions:
 * session 63 (s63c) offered recovery for a body footnote reference whose
 * note only a reached glossary's notes part defines, while the main-related matching
 * calibration opened clean (#828). The endnotes twin is unprobed; it
 * follows #793's shared footnote-to-endnote reading. References read
 * through a reached glossary keep the union of both sets of definitions;
 * that reading is unprobed.
 *
 * The references are read only from the story roots, as the read model
 * (`wordReadTypedTargets`) resolves them, wherever they are stored: every
 * main document and reached glossary (segment `officeDocument`). The
 * name-based story list this reading replaced missed a main document
 * stored as `word/main.xml`, whose dangling footnote and endnote references
 * Word session 62 refused (s62a), and read a `word/header9.xml` no
 * relationship reaches, whose dangling footnote reference opened clean
 * (s62d). References in a read header or footer are not judged either:
 * session 62 (s62d, s62e) opened clean dangling endnote and footnote
 * references in a read header (#806), and session 63 (s63d) opened the
 * read-footer twins clean (#826). Nor is a reference inside a notes part
 * or a comments part judged: session 16 (s16a) had already opened an
 * endnote reference inside a footnote clean, and sessions 65 and 66 opened
 * clean every nesting they stored — footnote 1 holding a footnote
 * reference to 99 and endnote 1 an endnote reference to 99, in-package or
 * behind the escaping target (s65a), endnote 1 a footnote reference to 99
 * (s66a), and a comment a footnote reference to 99 or to 1 only a reached
 * glossary's notes part defines, the comments part read by the main
 * document or by a glossary (s66a, s65d) — so session 16's special case is
 * subsumed and gone (#852). A comments part both the main document and a
 * glossary read is refused, but by the shared reach, not the reference
 * (s66a; #850). An endnote reference in a comments part, and a note
 * reference in a glossary's own notes part, are unprobed.
 *
 * When the package view cannot be read, or names no stored main document,
 * the read model and the relationship walks all return nothing, and a
 * dangling reference would pass unjudged beside the package defect's own
 * findings (#816). Both note checks then fall back to the reading they used
 * before #800 — references from the name-based story list but its headers,
 * footers, comments and notes parts, definitions from the parts named
 * `<kind>s.xml` wherever stored — as `comments.ts` falls
 * back to a by-name lookup when no relationship names the main document.
 * When the package view can be read, nothing changes. No Word session has
 * shown a note reference in such a package; the reading is unprobed, and the
 * package defect's own findings still block the full run.
 */

import path from "node:path";

import type { CheckDefinition } from "../../../lib/check-registry";
import {
    canonicalPartName,
    contentTypeMediaType,
    findContentTypesStream,
    parseContentTypes,
    resolveContentType,
    type ContentTypeTable,
} from "../../../lib/content-types";
import type { ValidationContext } from "../../../lib/validation-context";
import type { Profile, ValidationIssue } from "../../../lib/types";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";
import { glossaryTypedRelationshipTargets, wordReadTypedTargets, wordTypedRelationshipTargets } from "../../../lib/word-read-parts";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

/**
 * The media type that makes Word refuse a notes part entirely: session 63
 * (s63c) and session 64 (s64b) offered recovery for a body footnote
 * reference whose note sat in a part typed `application/octet-stream` —
 * `word/notes.bin`, `word/notes.xml` or `word/footnotes.xml` alike — so
 * such a part defines no note (#829).
 */
const OCTET_STREAM_MEDIA_TYPE = "application/octet-stream";

/**
 * A w:id is ST_DecimalNumber (xsd:integer), and Word compares note ids by
 * integer value: session 62 (s62b) opened clean a reference spelled `01` or
 * `+1` against a note spelled `1`, each PDF the size of its matching
 * calibration's (#801). An id whose whole string is an optionally signed
 * digit run therefore compares as its BigInt value — `-0`, `+0` and `00`
 * all equal `0`. Any other spelling compares as its raw string, exactly as
 * before: whitespace-padded, empty and non-integer ids are unprobed. The
 * issue message keeps the reference's stored spelling.
 */
const noteIdKey = (id: string): string | bigint => (/^[+-]?[0-9]+$/.test(id) ? BigInt(id) : id);

const noteReferenceMissingCheck = (kind: "footnote" | "endnote"): CheckDefinition => ({
    id: `word-${kind}-reference-missing`,
    description: `${kind === "footnote" ? "Footnote" : "Endnote"} references name ${kind}s present in the package.`,
    groups: ["story"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];

        // A notes part the content-types stream types octet-stream defines
        // no note, whatever it is called (session 63 s63c, session 64 s64b,
        // #829); a missing or unparseable stream leaves its own findings to
        // block the run, and every other typing keeps defining its notes.
        const typesStream = findContentTypesStream(ctx.partNames);
        const typesDom = typesStream === undefined ? undefined : ctx.effectiveXml(typesStream);
        const table: ContentTypeTable | undefined = typesDom?.ok ? parseContentTypes(typesDom.document) : undefined;
        const defined = new Set<string | bigint>();
        const collect = (part: string): void => {
            const resolved = table === undefined ? null : resolveContentType(table, part);
            if (resolved !== null && contentTypeMediaType(resolved.contentType) === OCTET_STREAM_MEDIA_TYPE) {
                return;
            }
            const parsed = ctx.effectiveXml(part);
            if (!parsed?.ok) {
                return;
            }
            const dom = parsed.document;
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                const notes = getElementsByTagNameNSAll(dom, ns, kind);
                for (const entry of notes) {
                    const id = entry?.getAttributeNS(ns, "id");
                    if (id) {
                        defined.add(noteIdKey(id));
                    }
                }
            }
        };
        // The references one part holds, each judged against its definitions.
        const judged = (xmlFile: string, definitions: ReadonlySet<string | bigint>, union?: ReadonlySet<string | bigint>): void => {
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                return;
            }
            const dom = parsed.document;
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                const refs = getElementsByTagNameNSAll(dom, ns, `${kind}Reference`);
                for (const entry of refs) {
                    const id = entry?.getAttributeNS(ns, "id");
                    const key = id ? noteIdKey(id) : undefined;
                    if (key === undefined || definitions.has(key)) {
                        continue;
                    }
                    const glossaryOnly = union?.has(key) === true;
                    issues.push({
                        severity: "error",
                        message:
                            `${kind === "footnote" ? "A" : "An"} ${kind} reference uses id ${id}, and ` +
                            (glossaryOnly
                                ? `the notes parts the main document relates define no w:${kind} with that id; only a reached glossary's notes part does. `
                                : `this package has no w:${kind} with that id. `) +
                            "Microsoft Word offers to recover the file. " +
                            (glossaryOnly
                                ? `Remove the w:${kind}Reference, or add the note to a notes part the main document relates.`
                                : `Remove the w:${kind}Reference, or add the matching ${kind}.`),
                        path: xmlFile,
                        code: `word-${kind}-reference-missing`,
                    });
                }
            }
        };
        // A main-document entry joins the read model whenever the package
        // view can be read and its `officeDocument` relationship names a
        // stored part, so its absence means neither holds: every walk below
        // and the read model itself return nothing, and a dangling reference
        // would pass unjudged beside the package defect's own findings
        // (#816). Both note checks then fall back to the reading they used
        // before #800, as `comments.ts` falls back to a by-name lookup when
        // no relationship names the main document: definitions from the
        // parts named `<kind>s.xml` wherever stored, references from the
        // name-based story list but its headers, footers, comments and
        // notes parts. No Word session has shown a note reference in such a
        // package; the reading is unprobed.
        const readTargets = await wordReadTypedTargets(ctx, profile);
        if (!readTargets.some((target) => target.segment === "officeDocument")) {
            for (const part of ctx.partNames) {
                if (path.posix.basename(part) === `${kind}s.xml`) {
                    collect(part);
                }
            }
            for (const xmlFile of ctx.storyParts()) {
                // A header or footer is not judged here either (#806, #826),
                // nor a comments or notes part (sessions 65 and 66, #852).
                if (/^(?:(?:header|footer)\d*|comments|footnotes|endnotes)\.xml$/i.test(path.posix.basename(xmlFile))) {
                    continue;
                }
                judged(xmlFile, defined);
            }
            return issues;
        }
        // A note is defined only by the part its relationship names,
        // whatever the part is called, from the main document and from a
        // reached glossary alike (session 60 s60a, #793); a part no
        // relationship names defines nothing, even under the standard
        // names (session 61 s61a, s61b, #797). The walks are ungated by
        // usage, so a related-but-unused part defines its ids too, as this
        // set only adds definitions, never a rejection. A relationship
        // target that escapes the package names the part its `..` segments
        // stop at, at the package root, as RFC 3986's remove_dot_segments
        // stops them (sessions 63 s63c, 64 s64a and 74 s74a; #830, #986).
        const segment = `${kind}s`;
        for (const target of await wordTypedRelationshipTargets(ctx, new Set([segment]), profile)) {
            collect(target.part);
        }
        // Main-document references cannot use a glossary's definitions
        // (session 63 s63c, #828), and a reference only the union
        // satisfies says so in its message (#838). Glossary references
        // keep the union, whose reading remains unprobed.
        const mainDefined = new Set(defined);
        for (const target of await glossaryTypedRelationshipTargets(ctx, new Set([segment]), profile)) {
            collect(target.part);
        }
        // The references are read only from the story roots, as the read
        // model resolves them, wherever they are stored: every main document
        // and reached glossary. A story-named part no relationship reaches
        // is no longer read for references (session 62 s62d opened one
        // holding a dangling reference clean), and a renamed main document
        // now is (s62a refused its dangling references). Read headers,
        // footers, comments and notes parts are excluded too: session 62
        // (s62d, s62e) opened a read header's dangling endnote and footnote
        // references clean (#806), session 63 (s63d) the read-footer twins
        // (#826), and sessions 65 and 66 the notes-part and comments
        // nestings they stored (#852). Filter before deduping so another
        // segment's reach cannot hide a story root.
        const seen = new Set<string>();
        for (const target of readTargets) {
            if (target.segment !== "officeDocument") {
                continue;
            }
            // A part reached as a story root twice is judged once.
            const canonical = canonicalPartName(target.part);
            if (seen.has(canonical)) {
                continue;
            }
            seen.add(canonical);
            judged(target.part, target.via === "glossary" ? defined : mainDefined, defined);
        }

        return issues;
    },
});

export const wordFootnoteReferenceMissingCheck = noteReferenceMissingCheck("footnote");
export const wordEndnoteReferenceMissingCheck = noteReferenceMissingCheck("endnote");
