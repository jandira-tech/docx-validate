/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { runChecks } from "../src/lib/check-scheduler";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { parseXml } from "../src/lib/xml-helpers";
import { isWordBlockingIssue } from "../src/scripts/office/validators/word-blocking";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
const CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex";
const document = (body: string) => `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`;
const comments = (body = '<w:comment w:id="0"><w:p w14:paraId="11111111"/></w:comment>', namespace = W) =>
    `<w:comments xmlns:w="${namespace}" xmlns:w14="${W14}">${body}</w:comments>`;
const ids = (body: string) => `<c:commentsIds xmlns:c="${CID}">${body}</c:commentsIds>`;
const extensible = (body: string) => `<c:commentsExtensible xmlns:c="${CEX}">${body}</c:commentsExtensible>`;
const extended = (body: string) => `<c:commentsEx xmlns:c="${W15}">${body}</c:commentsEx>`;
const run = (parts: Readonly<Record<string, string>>) =>
    runChecks(ValidationContext.fromEntries(parts), createDefaultCheckRegistry(), { format: "docx", select: ["comments"] });
// Fails loudly rather than yielding no findings when a check is not registered.
const runRegistered = async (id: string, parts: Readonly<Record<string, string>>) => {
    const check = createDefaultCheckRegistry().get(id);
    if (!check) throw new Error(`${id} is not registered`);
    return check.run(ValidationContext.fromEntries(parts));
};
const markerIssues = (parts: Readonly<Record<string, string>>) => runRegistered("comment-markers", parts);
const threadIssues = (parts: Readonly<Record<string, string>>) => runRegistered("comment-threading", parts);

describe("comment reference facts (#784)", () => {
    it.each([false, true])("distinguishes a missing reference from range-only markers (reference=%s)", async (reference) => {
        const issues = await markerIssues({
            "word/document.xml": document(
                `<w:commentRangeStart w:id="9"/><w:commentRangeEnd w:id="9"/>${reference ? '<w:commentReference w:id="9"/>' : ""}`,
            ),
            "word/comments.xml": comments(""),
        });
        expect(issues).toHaveLength(1);
        const issue = issues[0]!;
        expect(issue.code).toBe("comment-marker-missing");
        expect(isWordBlockingIssue(issue)).toBe(reference);
        issue.message = reference ? "a range marker" : "w:commentReference names a missing comment";
        expect(isWordBlockingIssue(issue)).toBe(reference);
        expect(isWordBlockingIssue({ ...issue })).toBe(false);
    });
});

// All package bytes in this suite are supplied in memory; no filesystem loaders or mocks.
describe("registered comment checks", () => {
    it("registers DOCX checks outside the Word rules, without suppressing legacy parse findings", () => {
        const registry = createDefaultCheckRegistry();
        expect(registry.resolve(["comments"], "docx")).toEqual(["comment-markers", "comment-threading"]);
        expect(registry.resolve(["comments"], "pptx")).toEqual([]);
        for (const id of ["comment-markers", "comment-threading"]) {
            expect(registry.get(id)?.requires ?? []).toEqual([]);
            expect(registry.get(id)?.reportsParseFailures).toBe(true);
        }
    });

    it("reports the same marker messages, paths, order and native severities entirely in memory", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document('<w:commentRangeEnd w:id="10"/><w:commentRangeStart w:id="2"/><w:commentReference w:id="9"/>'),
            "word/comments.xml": comments(""),
        });
        const result = await runChecks(ctx, createDefaultCheckRegistry(), { format: "docx", select: ["comments"] });
        expect(result.issues.filter((issue) => !issue.code?.startsWith("comment-thread"))).toEqual([
            {
                severity: "error",
                path: "word/document.xml",
                code: "comment-orphan-end",
                message: 'commentRangeEnd id="10" has no matching commentRangeStart',
            },
            {
                severity: "error",
                path: "word/document.xml",
                code: "comment-orphan-start",
                message: 'commentRangeStart id="2" has no matching commentRangeEnd',
            },
            ...["2", "9", "10"].map((id) => ({
                severity: "error",
                path: "word/document.xml",
                code: "comment-marker-missing",
                message: `marker id="${id}" references non-existent comment`,
            })),
        ]);
        expect(result.executed).toEqual(["comment-markers", "comment-threading"]);
        expect(ctx.parseCount("word/document.xml")).toBe(1);
        expect(ctx.readCount("word/comments.xml")).toBe(1);
    });

    it.each<Readonly<Record<string, string>>>([{}, { "word/document.xml": document("<w:p/>") }, { "word/comments.xml": comments() }])(
        "accepts absent optional parts: %j",
        async (parts) => {
            expect((await run(parts)).issues).toEqual([]);
        },
    );

    it("keeps matched markers and ignores empty marker/comment ids", async () => {
        expect(
            await markerIssues({
                "word/document.xml": document(
                    '<w:commentRangeStart w:id="0"/><w:commentRangeEnd w:id="0"/><w:commentReference/><w:commentReference w:id="0"/>',
                ),
                "word/comments.xml": comments('<w:comment/><w:comment w:id="0"/>'),
            }),
        ).toEqual([]);
    });

    it("finds OPC-equivalent part names and retains their actual paths", async () => {
        const result = await run({ "WORD/%64ocument.XML": document('<w:commentRangeStart w:id="3"/>'), "WORD/COMMENTS.XML": comments("") });
        expect(result.issues).toContainEqual(expect.objectContaining({ code: "comment-marker-missing", path: "WORD/%64ocument.XML" }));
    });

    it.each(["word/document.xml", "word/comments.xml"])("preserves marker parse diagnostics for %s", async (part) => {
        const malformed = "<broken";
        let message = "";
        try {
            parseXml(malformed);
        } catch (error) {
            message = (error as Error).message;
        }
        expect(await markerIssues({ "word/document.xml": document(""), "word/comments.xml": comments(), [part]: malformed })).toEqual([
            { code: "comment-marker-parse", severity: "error", path: part, message: `Error parsing XML: ${message}` },
        ]);
    });

    it.each(["word/comments.xml", "word/commentsIds.xml", "word/commentsExtensible.xml", "word/commentsExtended.xml"])(
        "preserves threading parse diagnostics for %s",
        async (part) => {
            const malformed = "<broken";
            const parsed = ValidationContext.fromEntries({ [part]: malformed }).xml(part);
            expect(await threadIssues({ "word/comments.xml": comments(), [part]: malformed })).toEqual([
                {
                    code: "comment-thread-parse",
                    severity: "error",
                    path: part,
                    message: `Error parsing XML: ${parsed && !parsed.ok ? parsed.message : ""}`,
                },
            ]);
        },
    );

    it("leaves document parse failure to the marker check while threading continues", async () => {
        const result = await run({
            "word/document.xml": "<broken",
            "word/comments.xml": comments(),
            "word/commentsIds.xml": ids("<c:commentId/>"),
        });
        expect(result.issues.map((issue) => issue.code)).toEqual([
            "comment-marker-parse",
            "comment-thread-commentid-missing-paraid",
            "comment-thread-commentid-missing-durableid",
        ]);
    });

    it.each([W, "http://purl.oclc.org/ooxml/wordprocessingml/main"])(
        "matches every comment paragraph under %s and tolerates duplicate/empty comment ids",
        async (namespace) => {
            const body =
                '<w:comment/><w:comment w:id="0"><w:p/><w:p w14:paraId="11111111"/><w:p w14:paraId="22222222"/></w:comment><w:comment w:id="0"/>';
            expect(
                await threadIssues({
                    "word/comments.xml": comments(body, namespace),
                    "word/commentsIds.xml": ids('<c:commentId c:paraId="22222222" c:durableId="33333333"/>'),
                    "word/commentsExtended.xml": extended('<c:commentEx c:paraId="22222222"/>'),
                }),
            ).toEqual([]);
        },
    );

    it("reports all three mismatched counts", async () => {
        expect(
            (await threadIssues({ "word/document.xml": document(""), "word/comments.xml": comments() })).map((issue) => issue.message),
        ).toEqual([
            "commentRangeStart count (0) does not match comment count in comments.xml (1)",
            "commentRangeEnd count (0) does not match comment count in comments.xml (1)",
            "commentReference count (0) does not match comment count in comments.xml (1)",
        ]);
    });

    it("accepts matching marker counts", async () => {
        expect(
            await threadIssues({
                "word/document.xml": document("<w:commentRangeStart/><w:commentRangeEnd/><w:commentReference/>"),
                "word/comments.xml": comments(),
            }),
        ).toEqual([]);
    });

    it("reports orphan, missing and duplicate comment ids and durable ids", async () => {
        const result = await threadIssues({
            "word/comments.xml": comments(),
            "word/commentsIds.xml": ids(
                '<c:commentId c:paraId="BAD" c:durableId="A"/><c:commentId c:paraId="BAD" c:durableId="A"/><c:commentId/>',
            ),
            "word/commentsExtensible.xml": extensible(
                '<c:commentExtensible/><c:commentExtensible c:durableId="B"/><c:commentExtensible c:durableId="B"/><c:commentExtensible c:durableId="A"/>',
            ),
        });
        expect(result.map((issue) => issue.code)).toEqual([
            "comment-thread-commentid-paraid-orphan",
            "comment-thread-commentid-paraid-orphan",
            "comment-thread-commentid-missing-paraid",
            "comment-thread-commentid-missing-durableid",
            "comment-thread-commentid-duplicate-paraid",
            "comment-thread-commentid-duplicate-durableid",
            "comment-thread-durableid-missing",
            "comment-thread-durableid-orphan",
            "comment-thread-durableid-orphan",
            "comment-thread-durableid-duplicate",
        ]);
        expect(result.every((issue) => issue.severity === "error")).toBe(true);
        expect(result.slice(0, 6).every((issue) => issue.path === "word/commentsIds.xml")).toBe(true);
        expect(result.slice(6).every((issue) => issue.path === "word/commentsExtensible.xml")).toBe(true);
    });

    it("reports extension duplicates, unmatched paragraphs and orphan parents with legacy severity", async () => {
        const result = await threadIssues({
            "word/comments.xml": comments(
                '<w:comment w:id="0"><w:p w14:paraId="11111111"/></w:comment><w:comment w:id="1"><w:p/></w:comment>',
            ),
            "word/commentsExtended.xml": extended(
                '<c:commentEx/><c:commentEx c:paraId="BAD" c:paraIdParent="PARENT"/><c:commentEx c:paraId="BAD" c:paraIdParent="BAD"/>',
            ),
        });
        expect(result.map((issue) => [issue.code, issue.severity])).toEqual([
            ["comment-thread-duplicate-paraid", "error"],
            ["comment-thread-paraid-missing", "warning"],
            ["comment-thread-paraid-orphan", "error"],
            ["comment-thread-orphan-parent", "error"],
        ]);
    });
});
