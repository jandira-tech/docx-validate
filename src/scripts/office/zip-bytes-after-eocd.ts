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
 * not that defect.
 *
 * `zipEocdCommentOverrun` reports a comment length that claims more bytes
 * than remain in the file. Setting the length to the number of bytes actually
 * present is not that defect.
 */

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
const ZIP64_LOCATOR_SIGNATURE = 0x07064b50;
const EOCD_FIXED_LENGTH = 22;
const ZIP64_LOCATOR_LENGTH = 20;
const MAX_COMMENT_LENGTH = 0xffff;
const ZIP64_OFFSET = 0xffffffff;

const readU16 = (buf: Uint8Array, offset: number): number => buf[offset]! | (buf[offset + 1]! << 8);

const readU32 = (buf: Uint8Array, offset: number): number =>
    (buf[offset]! | (buf[offset + 1]! << 8) | (buf[offset + 2]! << 16) | (buf[offset + 3]! << 24)) >>> 0;

const hasSignature = (buf: Uint8Array, offset: number, signature: number): boolean => {
    if (offset < 0 || offset + 4 > buf.length) return false;
    return readU32(buf, offset) === signature;
};

const isStructuralEocd = (buf: Uint8Array, offset: number): boolean => {
    const centralDirectoryOffset = readU32(buf, offset + 16);
    if (hasSignature(buf, centralDirectoryOffset, CENTRAL_DIRECTORY_SIGNATURE)) return true;
    const centralDirectorySize = readU32(buf, offset + 12);
    const entryCount = readU16(buf, offset + 10);
    if (centralDirectorySize === 0 && entryCount === 0 && centralDirectoryOffset === offset) return true;
    if (centralDirectoryOffset === ZIP64_OFFSET) {
        return hasSignature(buf, offset - ZIP64_LOCATOR_LENGTH, ZIP64_LOCATOR_SIGNATURE);
    }
    return false;
};

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

/** Count of bytes after the end-of-central-directory record past its declared comment. */
export const countZipBytesAfterEocd = (buf: Uint8Array): number => {
    if (buf.length < EOCD_FIXED_LENGTH) return 0;
    const earliest = Math.max(0, buf.length - (EOCD_FIXED_LENGTH + MAX_COMMENT_LENGTH));
    for (let offset = buf.length - EOCD_FIXED_LENGTH; offset >= earliest; offset -= 1) {
        if (readU32(buf, offset) !== EOCD_SIGNATURE) continue;
        const commentLength = readU16(buf, offset + 20);
        const recordEnd = offset + EOCD_FIXED_LENGTH + commentLength;
        if (recordEnd > buf.length) continue;
        if (!isStructuralEocd(buf, offset)) continue;
        return buf.length - recordEnd;
    }
    return 0;
};
