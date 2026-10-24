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
 * Byte-level reads of the ZIP envelope: the end-of-central-directory record
 * and the central directory's stored entry names.
 *
 * The structure helpers (`isStructuralEocd`, the readers and signatures) back
 * the envelope checks in `scripts/office/zip-bytes-after-eocd.ts`.
 *
 * `zipCentralDirectoryEntries` reads the names as stored, in central-directory
 * order, keeping duplicates. Loading loses them: JSZip keeps only the last
 * entry under a given name and rewrites `.`/`..`/empty segments before that
 * (`utils.resolve`), so the stored-name diagnosis and the item-name refusal
 * need this view (#411). Everything fails soft: a malformed envelope or
 * central directory yields `undefined`, and callers keep the loaded view's
 * names — a broken envelope is already diagnosed by the ZIP envelope checks.
 */

export const EOCD_SIGNATURE = 0x06054b50;
export const EOCD_BYTES = Buffer.from([0x50, 0x4b, 0x05, 0x06]);
export const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
export const ZIP64_LOCATOR_SIGNATURE = 0x07064b50;
export const EOCD_FIXED_LENGTH = 22;
export const ZIP64_LOCATOR_LENGTH = 20;
export const MAX_COMMENT_LENGTH = 0xffff;
export const ZIP64_OFFSET = 0xffffffff;

const ZIP64_EOCD_SIGNATURE = 0x06064b50;
const ZIP64_EOCD_FIXED_LENGTH = 56;
const ZIP64_COUNT = 0xffff;
const CENTRAL_HEADER_FIXED_LENGTH = 46;
const UTF8_NAME_FLAG = 0x0800;
const DOS_DIRECTORY_ATTRIBUTE = 0x10;

export const readU16 = (buf: Uint8Array, offset: number): number => buf[offset]! + buf[offset + 1]! * 0x100;

export const readU32 = (buf: Uint8Array, offset: number): number =>
    buf[offset]! + buf[offset + 1]! * 0x100 + buf[offset + 2]! * 0x10000 + buf[offset + 3]! * 0x1000000;

/** Low 8 bytes as a number; only archives whose offsets fit in doubles qualify anyway. */
const readU64 = (buf: Uint8Array, offset: number): number => readU32(buf, offset) + readU32(buf, offset + 4) * 0x100000000;

export const hasSignature = (buf: Uint8Array, offset: number, signature: number): boolean => {
    if (offset < 0 || offset + 4 > buf.length) {
        return false;
    }
    return readU32(buf, offset) === signature;
};

export const isStructuralEocd = (buf: Uint8Array, offset: number): boolean => {
    const centralDirectoryOffset = readU32(buf, offset + 16);
    if (hasSignature(buf, centralDirectoryOffset, CENTRAL_DIRECTORY_SIGNATURE)) {
        return true;
    }
    const centralDirectorySize = readU32(buf, offset + 12);
    const entryCount = readU16(buf, offset + 10);
    if (centralDirectorySize === 0 && entryCount === 0 && centralDirectoryOffset === offset) {
        return true;
    }
    if (centralDirectoryOffset === ZIP64_OFFSET) {
        return hasSignature(buf, offset - ZIP64_LOCATOR_LENGTH, ZIP64_LOCATOR_SIGNATURE);
    }
    return false;
};

/**
 * The offset of the last occurrence of `signature` in the whole buffer, or
 * `undefined`. This is JSZip's `lastIndexOfSignature`: the search stops at the
 * first match from the end and judges nothing about it, so an archive can
 * carry any amount of undeclared data after its record (#447) and a later
 * signature, structural or not, is the one read (#462).
 */
const lastSignatureOffset = (buf: Uint8Array, signature: number): number | undefined => {
    for (let offset = buf.length - 4; offset >= 0; offset -= 1) {
        if (buf[offset] === 0x50 && readU32(buf, offset) === signature) {
            return offset;
        }
    }
    return undefined;
};

/** IBM CP437, bytes 0x80–0xFF: the encoding APPNOTE assigns to names without the UTF-8 flag. */
const CP437_HIGH =
    "ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒáíóúñÑªº¿⌐¬½¼¡«»░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀αßΓπΣσµτΦΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙·√ⁿ²■ ";

const utf8 = new TextDecoder();

/**
 * The stored name's decode: UTF-8 when general-purpose flag bit 11 is set,
 * CP437 otherwise. ASCII names decode identically under both, which is the
 * ordinary case and what `unsafeOriginalName` holds. JSZip decodes every name
 * as UTF-8 unless the caller overrides it, so for a non-ASCII name stored
 * without the flag this name can differ from the loaded entry's — the
 * central-directory bytes are authoritative for the stored spelling. The
 * Info-ZIP unicode-path extra field is not consulted.
 */
const decodeName = (buf: Uint8Array, offset: number, length: number, flags: number): string => {
    const bytes = buf.subarray(offset, offset + length);
    // General-purpose flag bit 11 marks the name as UTF-8.
    if (flags % 0x1000 >= UTF8_NAME_FLAG) {
        return utf8.decode(bytes);
    }
    let name = "";
    for (const byte of bytes) {
        name += byte < 0x80 ? String.fromCharCode(byte) : CP437_HIGH[byte - 0x80]!;
    }
    return name;
};

export type ZipStoredEntry = {
    /** The item's stored name, before any normalization or resolution. */
    readonly name: string;
    /** A folder entry: a trailing `/` or the MS-DOS directory attribute, as JSZip marks `dir`. */
    readonly directory: boolean;
};

/**
 * The central directory's stored entry names, in order, duplicates kept, or
 * `undefined` when the envelope or directory cannot be read. The entry count
 * and declared size must match what the walk finds — a directory that lies is
 * not read at all.
 *
 * The directory is the one JSZip reads (`zipEntries.js` `readEndOfCentral`),
 * so the names describe the archive the loaded view holds (#462): the record
 * is the last EOCD signature, and the bytes before the archive — the record's
 * offset minus where its directory says it should sit — rebase every offset
 * (`reader.zero`). A prepended archive is skipped, a fake signature behind the
 * record yields JSZip's empty directory, and a record whose directory would
 * end past it is refused, as JSZip refuses it. ZIP64 follows JSZip too: the
 * last locator, its record or else the last ZIP64 record, and an expected
 * position that counts the locator and the ZIP64 record.
 */
export const zipCentralDirectoryEntries = (buf: Uint8Array): readonly ZipStoredEntry[] | undefined => {
    const eocd = lastSignatureOffset(buf, EOCD_SIGNATURE);
    if (eocd === undefined || eocd + EOCD_FIXED_LENGTH > buf.length) {
        return undefined;
    }
    let offset = readU32(buf, eocd + 16);
    let size = readU32(buf, eocd + 12);
    let count = readU16(buf, eocd + 10);
    const zip64 =
        readU16(buf, eocd + 4) === ZIP64_COUNT ||
        readU16(buf, eocd + 6) === ZIP64_COUNT ||
        readU16(buf, eocd + 8) === ZIP64_COUNT ||
        count === ZIP64_COUNT ||
        size === ZIP64_OFFSET ||
        offset === ZIP64_OFFSET;
    // Where the record sits when nothing precedes the archive.
    let expected = offset + size;
    if (zip64) {
        const locator = lastSignatureOffset(buf, ZIP64_LOCATOR_SIGNATURE);
        if (locator === undefined || locator + ZIP64_LOCATOR_LENGTH > buf.length) {
            return undefined;
        }
        let record: number | undefined = readU64(buf, locator + 8);
        if (!hasSignature(buf, record, ZIP64_EOCD_SIGNATURE)) {
            record = lastSignatureOffset(buf, ZIP64_EOCD_SIGNATURE);
        }
        if (record === undefined || record + ZIP64_EOCD_FIXED_LENGTH > buf.length) {
            return undefined;
        }
        count = readU64(buf, record + 32);
        size = readU64(buf, record + 40);
        offset = readU64(buf, record + 48);
        // The locator, then the ZIP64 record: 12 leading bytes plus its declared size.
        expected = offset + size + ZIP64_LOCATOR_LENGTH + 12 + readU64(buf, record + 4);
    }
    const zero = eocd - expected;
    if (zero < 0) {
        return undefined;
    }
    offset += zero;
    const end = offset + size;
    if (end > buf.length || end > eocd) {
        return undefined;
    }
    const entries: ZipStoredEntry[] = [];
    while (offset + CENTRAL_HEADER_FIXED_LENGTH <= end) {
        if (!hasSignature(buf, offset, CENTRAL_DIRECTORY_SIGNATURE)) {
            return undefined;
        }
        const flags = readU16(buf, offset + 8);
        const nameLength = readU16(buf, offset + 28);
        const extraLength = readU16(buf, offset + 30);
        const commentLength = readU16(buf, offset + 32);
        const attributes = readU32(buf, offset + 38);
        const nameEnd = offset + CENTRAL_HEADER_FIXED_LENGTH + nameLength;
        if (nameEnd > end) {
            return undefined;
        }
        const name = decodeName(buf, offset + CENTRAL_HEADER_FIXED_LENGTH, nameLength, flags);
        // A folder entry ends in `/` or sets the MS-DOS directory bit, as
        // JSZip marks `dir`.
        entries.push({ name, directory: name.endsWith("/") || attributes % 0x20 >= DOS_DIRECTORY_ATTRIBUTE });
        offset = nameEnd + extraLength + commentLength;
    }
    if (offset !== end || entries.length !== count) {
        return undefined;
    }
    return entries;
};
