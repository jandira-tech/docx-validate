/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * ZIP-envelope checks. They read the original archive bytes, so they are
 * archive-only: an unpacked directory has no envelope to inspect.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { countZipBytesAfterEocd, zipEocdCommentOverrun } from "../zip-bytes-after-eocd";

export const zipBytesAfterEocdCheck: CheckDefinition = {
    id: "zip-bytes-after-eocd",
    description: "No undeclared bytes follow the ZIP end-of-central-directory record (Word offers recovery).",
    groups: ["package"],
    formats: ["docx", "pptx", "xlsx"],
    archiveOnly: true,
    run: (ctx): ValidationIssue[] => {
        const trailing = ctx.archiveBytes ? countZipBytesAfterEocd(ctx.archiveBytes) : 0;
        if (trailing === 0) return [];
        return [
            {
                severity: "error",
                code: "zip-bytes-after-eocd",
                message: `ZIP end of central directory is followed by ${trailing} undeclared byte(s). Word offers to recover the document. Remove the bytes after that record; a comment whose length is declared in the record is not this defect.`,
            },
        ];
    },
};

export const zipEocdCommentOverrunCheck: CheckDefinition = {
    id: "zip-eocd-comment-overrun",
    description: "The ZIP end-of-central-directory comment length fits inside the file (Word offers recovery).",
    groups: ["package"],
    formats: ["docx", "pptx", "xlsx"],
    archiveOnly: true,
    run: (ctx): ValidationIssue[] => {
        const overrun = ctx.archiveBytes ? zipEocdCommentOverrun(ctx.archiveBytes) : undefined;
        if (!overrun) return [];
        return [
            {
                severity: "error",
                code: "zip-eocd-comment-overrun",
                message: `ZIP end-of-central-directory comment length is ${overrun.declared}, but only ${overrun.remaining} byte(s) remain in the file. Word offers to recover the document. Set the comment length to the bytes that are present.`,
            },
        ];
    },
};
