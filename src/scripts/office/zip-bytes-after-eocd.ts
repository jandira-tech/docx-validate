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
 * Envelope defects of the ZIP end-of-central-directory record that make
 * Microsoft Word offer to recover the document.
 *
 * `countZipBytesAfterEocd` counts bytes after that record that its comment
 * length does not cover. A comment whose length is stored in the record is
 * not that defect. A later end-of-central-directory record at the end of the
 * file does not hide bytes that follow an earlier record.
 *
 * `zipEocdCommentOverrun` reports a comment length that claims more bytes
 * than remain in the file. Setting the length to the number of bytes actually
 * present is not that defect.
 */

import {
    EOCD_BYTES,
    EOCD_FIXED_LENGTH,
    EOCD_SIGNATURE,
    hasSignature,
    isStructuralEocd,
    MAX_COMMENT_LENGTH,
    readU16,
    readU32,
    ZIP64_LOCATOR_LENGTH,
    ZIP64_LOCATOR_SIGNATURE,
    ZIP64_OFFSET,
} from "../../lib/zip-central-directory";

export type ZipEocdCommentOverrun = {
    readonly declared: number;
    readonly remaining: number;
};

/**
 * The nearest structural end-of-central-directory record whose comment length
 * runs past the end of the file. A record whose comment fits, including a
 * record followed by undeclared bytes, is not an overrun.
 */
export const zipEocdCommentOverrun = (buf: Uint8Array): ZipEocdCommentOverrun | undefined => {
    if (buf.length < EOCD_FIXED_LENGTH) return undefined;
    const earliest = Math.max(0, buf.length - (EOCD_FIXED_LENGTH + MAX_COMMENT_LENGTH));
    for (let offset = buf.length - EOCD_FIXED_LENGTH; offset >= earliest; offset -= 1) {
        if (readU32(buf, offset) !== EOCD_SIGNATURE) continue;
        if (!isStructuralEocd(buf, offset)) continue;
        const declared = readU16(buf, offset + 20);
        const remaining = buf.length - (offset + EOCD_FIXED_LENGTH);
        if (declared > remaining) return { declared, remaining };
        return undefined;
    }
    return undefined;
};

const recordEnd = (buf: Uint8Array, offset: number): number => offset + EOCD_FIXED_LENGTH + readU16(buf, offset + 20);

/** The central directory ends where this end-of-central-directory record begins. */
const abutsCentralDirectory = (buf: Uint8Array, offset: number): boolean => {
    const centralDirectoryOffset = readU32(buf, offset + 16);
    if (centralDirectoryOffset === ZIP64_OFFSET) {
        return hasSignature(buf, offset - ZIP64_LOCATOR_LENGTH, ZIP64_LOCATOR_SIGNATURE);
    }
    return centralDirectoryOffset + readU32(buf, offset + 12) === offset;
};

/**
 * Bytes after an earlier structural record that a later record at EOF would
 * otherwise hide. The earlier record has to sit directly after its central
 * directory, which is the layout Word treats as the end of the archive.
 */
const bytesHiddenByLaterEocd = (buf: Uint8Array): number => {
    const lastStart = buf.length - EOCD_FIXED_LENGTH;
    // Search signatures in native code rather than reading every byte in JS.
    // Keep checking earlier records even when the last EOCD ends at EOF:
    // concatenated archives have that layout too (Word still recovers them).
    const bytes = Buffer.from(buf.buffer, buf.byteOffset, buf.byteLength);
    for (let offset = bytes.indexOf(EOCD_BYTES); offset >= 0 && offset < lastStart; offset = bytes.indexOf(EOCD_BYTES, offset + 1)) {
        if (offset + EOCD_FIXED_LENGTH > buf.length) continue;
        const end = recordEnd(buf, offset);
        if (end >= buf.length) continue;
        if (!isStructuralEocd(buf, offset)) continue;
        if (!abutsCentralDirectory(buf, offset)) continue;
        return buf.length - end;
    }
    return 0;
};

/** Count of bytes after the end-of-central-directory record past its declared comment. */
export const countZipBytesAfterEocd = (buf: Uint8Array): number => {
    if (buf.length < EOCD_FIXED_LENGTH) return 0;
    const earliest = Math.max(0, buf.length - (EOCD_FIXED_LENGTH + MAX_COMMENT_LENGTH));
    for (let offset = buf.length - EOCD_FIXED_LENGTH; offset >= earliest; offset -= 1) {
        if (readU32(buf, offset) !== EOCD_SIGNATURE) continue;
        const commentLength = readU16(buf, offset + 20);
        const end = offset + EOCD_FIXED_LENGTH + commentLength;
        if (end > buf.length) continue;
        if (!isStructuralEocd(buf, offset)) continue;
        const trailing = buf.length - end;
        if (trailing > 0) return trailing;
        return bytesHiddenByLaterEocd(buf);
    }
    return bytesHiddenByLaterEocd(buf);
};
