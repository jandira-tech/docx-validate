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
 * need this view (#411). The walk mirrors JSZip's `readCentralDir`: the
 * declared size only feeds the rebase, records are read while the signature
 * matches, and a count mismatch is tolerated unless no record is found (#478).
 * `zipBytesBeforeArchive` measures the prefix in front of the archive JSZip
 * loads through the same shared read (#464). Everything fails soft: a
 * directory JSZip's reader would throw on yields `undefined`, and callers keep
 * the loaded view's names — a broken envelope is already diagnosed by the ZIP
 * envelope checks.
 */

export const EOCD_SIGNATURE = 0x06054b50;
export const EOCD_BYTES = Buffer.from([0x50, 0x4b, 0x05, 0x06]);
export const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
export const LOCAL_HEADER_SIGNATURE = 0x04034b50;
export const ZIP64_LOCATOR_SIGNATURE = 0x07064b50;
export const EOCD_FIXED_LENGTH = 22;
export const ZIP64_LOCATOR_LENGTH = 20;
export const MAX_COMMENT_LENGTH = 0xffff;
export const ZIP64_OFFSET = 0xffffffff;

const ZIP64_EOCD_SIGNATURE = 0x06064b50;
const ZIP64_EOCD_FIXED_LENGTH = 56;
const ZIP64_COUNT = 0xffff;
/** `utils.MAX_VALUE_32BITS`: a saturated 4-byte field as `readJszipInt` reads it. */
const JSZIP_MAX_32 = -1;
const CENTRAL_HEADER_FIXED_LENGTH = 46;
const LOCAL_HEADER_FIXED_LENGTH = 30;
const LOCAL_NAME_LENGTH_OFFSET = 26;
const LOCAL_EXTRA_LENGTH_OFFSET = 28;
const COMPRESSION_METHOD_OFFSET = 10;
/** The compression methods JSZip knows (`compressions.js`): store and deflate. */
const STORE_METHOD = 0x0000;
const DEFLATE_METHOD = 0x0008;
const ZIP64_EXTRA_FIELD = 0x0001;
const UTF8_NAME_FLAG = 0x0800;
const DOS_DIRECTORY_ATTRIBUTE = 0x10;

export const readU16 = (buf: Uint8Array, offset: number): number => buf[offset]! + buf[offset + 1]! * 0x100;

export const readU32 = (buf: Uint8Array, offset: number): number =>
    buf[offset]! + buf[offset + 1]! * 0x100 + buf[offset + 2]! * 0x10000 + buf[offset + 3]! * 0x1000000;

/**
 * A 4- or 8-byte field as JSZip's `DataReader.readInt` reads it. Its loop
 * accumulates with a 32-bit `<<`, which keeps the low 32 bits — the first four
 * bytes, little-endian — as a signed number (`utils.MAX_VALUE_32BITS` is -1).
 * The central-directory read mirrors this so that it selects and walks the
 * directory JSZip does, not the one the field's value names (#470).
 */
const readJszipInt = (buf: Uint8Array, offset: number): number => {
    const low = readU32(buf, offset);
    return low >= 0x80000000 ? low - 0x100000000 : low;
};

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

/** Decode APPNOTE's default ZIP name/comment encoding; shared with JSZip's loader. */
export const decodeCp437 = (bytes: Uint8Array): string => {
    let name = "";
    for (const byte of bytes) {
        name += byte < 0x80 ? String.fromCharCode(byte) : CP437_HIGH[byte - 0x80]!;
    }
    return name;
};

/**
 * The stored name's decode: UTF-8 when general-purpose flag bit 11 is set,
 * CP437 otherwise. ASCII names decode identically under both, which is the
 * ordinary case and what `unsafeOriginalName` holds. The validator's JSZip
 * loader shares decodeCp437 for flag-less names (#480); the central-directory
 * bytes are authoritative for the stored spelling. The Info-ZIP unicode-path
 * extra field is not consulted here.
 */
const decodeName = (buf: Uint8Array, offset: number, length: number, flags: number): string => {
    const bytes = buf.subarray(offset, offset + length);
    // General-purpose flag bit 11 marks the name as UTF-8.
    if (flags % 0x1000 >= UTF8_NAME_FLAG) {
        return utf8.decode(bytes);
    }
    return decodeCp437(bytes);
};

export type ZipStoredEntry = {
    /** The item's stored name, before any normalization or resolution. */
    readonly name: string;
    /** A folder entry: a trailing `/` or the MS-DOS directory attribute, as JSZip marks `dir`. */
    readonly directory: boolean;
};

/**
 * The central directory's stored entry names, in order, duplicates kept, or
 * `undefined` when the envelope or directory cannot be read — exactly where
 * JSZip's reader throws, an entry's local part included: `load` reads it
 * after the directory walk, and where that read throws the package never
 * opens (#555). A directory whose declared size or count does not
 * match the records walked is read anyway, as JSZip reads it (#478).
 */
export const zipCentralDirectoryEntries = (buf: Uint8Array): readonly ZipStoredEntry[] | undefined =>
    readCentralDirectoryView(buf)?.records.map((record) => ({ name: record.name, directory: record.directory }));

/** One central-directory record's stored fields, with its raw byte spans. */
export type ZipCentralRecord = {
    /** The item's stored name, decoded as {@link zipCentralDirectoryEntries} does. */
    readonly name: string;
    /** The name's raw bytes between the fixed record header and the extra field. */
    readonly nameBytes: Uint8Array;
    /** The extra-field bytes stored in this record. */
    readonly extraBytes: Uint8Array;
    /** The comment bytes stored in this record. */
    readonly commentBytes: Uint8Array;
    /** The record's offset in the buffer (its signature). */
    readonly recordOffset: number;
    /** The local header offset as stored: relative to the archive, so add the view's `zero`. */
    readonly localHeaderOffset: number;
    /** The compressed size the record stores. */
    readonly compressedSize: number;
    /** The disk the record says the entry starts on: 0 for a one-file archive. */
    readonly diskNumberStart: number;
    /** A folder entry: a trailing `/` or the MS-DOS directory attribute, as JSZip marks `dir`. */
    readonly directory: boolean;
};

/** The archive's central directory and EOCD, located and walked. */
export type CentralDirectoryView = {
    readonly records: readonly ZipCentralRecord[];
    /** The end-of-central-directory record's offset. */
    readonly eocdOffset: number;
    /** One past the EOCD's declared comment, which lies within the buffer. */
    readonly eocdEnd: number;
    /** Where the central directory starts in the buffer. */
    readonly cdOffset: number;
    /**
     * One past the walked central directory: where the first non-header
     * signature sits, not the declared end. A directory whose declared size
     * spans more than its records ends before it (#478).
     */
    readonly cdEnd: number;
    /**
     * The bytes before the archive, which every stored offset is relative to:
     * JSZip's `reader.zero`. Zero unless something was prepended.
     */
    readonly zero: number;
    /** The directory was reached through ZIP64 structures. */
    readonly zip64: boolean;
};

/**
 * The central directory's records, in order, duplicates kept, plus the
 * envelope positions a rewrite needs — or `undefined` under the same
 * conditions {@link zipCentralDirectoryEntries} reports none.
 *
 * The directory is the one JSZip reads (`zipEntries.js` `readEndOfCentral`),
 * so the records describe the archive the loaded view holds (#462): the record
 * is the last EOCD signature, and the bytes before the archive — the record's
 * offset minus where its directory says it should sit — rebase every offset
 * (`reader.zero`). A prepended archive is skipped, a fake signature behind the
 * record yields JSZip's empty directory, and a record whose directory would
 * end past it is refused, as JSZip refuses it. The walk is JSZip's
 * `readCentralDir`: records are read while the central-header signature
 * matches, a record whose 46-byte fixed header the buffer cuts off refuses
 * as `readCentralPart`'s bounded reads refuse (#498), a record whose
 * declared name, extra or comment fields would run past the buffer refuses
 * as `readData` refuses, a field inside the extra region that declares more
 * data than the buffer holds refuses as `readExtraFields`' own `readData`
 * refuses (#525), fewer than four bytes left
 * for the next signature read refuses too, and the declared count only
 * refuses when no record at all is found (#478). The walk ends with the
 * `readLocalFiles` pass JSZip runs after it: a record whose local part
 * `localHeaderOffsetOf` models a throw on refuses the view, as `load`
 * refuses (#524, #555). ZIP64 follows JSZip too: the
 * last locator, its record or else the last ZIP64 record, and an expected
 * position that counts the locator and the ZIP64 record. Every 4- and 8-byte
 * field is read with JSZip's 32-bit arithmetic (`readJszipInt`), and what
 * JSZip throws on reads as nothing: a negative offset or size, a locator
 * offset outside the buffer, a multi-disk locator, ZIP64 extensible data, an
 * EOCD comment past the buffer (#470).
 */
export const readCentralDirectoryView = (buf: Uint8Array): CentralDirectoryView | undefined => {
    const eocd = lastSignatureOffset(buf, EOCD_SIGNATURE);
    // JSZip reads the record's comment too, and throws when it runs past the buffer.
    if (eocd === undefined || eocd + EOCD_FIXED_LENGTH > buf.length || eocd + EOCD_FIXED_LENGTH + readU16(buf, eocd + 20) > buf.length) {
        return undefined;
    }
    let offset = readJszipInt(buf, eocd + 16);
    let size = readJszipInt(buf, eocd + 12);
    let count = readU16(buf, eocd + 10);
    const zip64 =
        readU16(buf, eocd + 4) === ZIP64_COUNT ||
        readU16(buf, eocd + 6) === ZIP64_COUNT ||
        readU16(buf, eocd + 8) === ZIP64_COUNT ||
        count === ZIP64_COUNT ||
        size === JSZIP_MAX_32 ||
        offset === JSZIP_MAX_32;
    // Where the record sits when nothing precedes the archive.
    let expected = offset + size;
    if (zip64) {
        const locator = lastSignatureOffset(buf, ZIP64_LOCATOR_SIGNATURE);
        // JSZip refuses a locator that counts more than one disk.
        if (locator === undefined || locator + ZIP64_LOCATOR_LENGTH > buf.length || readJszipInt(buf, locator + 16) > 1) {
            return undefined;
        }
        let record: number | undefined = readJszipInt(buf, locator + 8);
        // JSZip tests the signature there through `setIndex`, which throws
        // for an index outside the buffer.
        if (record < 0 || record + 4 > buf.length) {
            return undefined;
        }
        if (!hasSignature(buf, record, ZIP64_EOCD_SIGNATURE)) {
            record = lastSignatureOffset(buf, ZIP64_EOCD_SIGNATURE);
        }
        if (record === undefined || record + ZIP64_EOCD_FIXED_LENGTH > buf.length) {
            return undefined;
        }
        const recordSize = readJszipInt(buf, record + 4);
        // Extensible data past the fixed 44 bytes sends JSZip's reading loop,
        // which never advances its counter, to the end of the buffer.
        if (recordSize > ZIP64_EOCD_FIXED_LENGTH - 12) {
            return undefined;
        }
        count = readJszipInt(buf, record + 32);
        size = readJszipInt(buf, record + 40);
        offset = readJszipInt(buf, record + 48);
        // The locator, then the ZIP64 record: 12 leading bytes plus its declared size.
        expected = offset + size + ZIP64_LOCATOR_LENGTH + 12 + recordSize;
    }
    // JSZip throws on a negative index — a negative offset or a negative
    // declared size, which only feeds the rebase — and a declared count it
    // reads as -1 is refused only when the walk finds no record (#492).
    const zero = eocd - expected;
    if (zero < 0 || offset < 0 || size < 0) {
        return undefined;
    }
    offset += zero;
    const cdOffset = offset;
    // The walk is `readCentralDir`: records are read while the central-header
    // signature matches, and the declared size does not bound them — it only
    // fed the rebase. Every read is bounded by the buffer as JSZip's reader
    // bounds it: `readAndCheckSignature` throws when fewer than four bytes
    // remain, and `readData`/`skip`/`setIndex` throw when the record's fields
    // would run past the buffer.
    const records: ZipCentralRecord[] = [];
    for (;;) {
        if (offset + 4 > buf.length) {
            return undefined;
        }
        if (!hasSignature(buf, offset, CENTRAL_DIRECTORY_SIGNATURE)) {
            break;
        }
        // The record's fixed header is `readCentralPart`'s `readInt`/`skip`
        // sequence: it throws when the 46 bytes would pass the buffer (#498).
        if (offset + CENTRAL_HEADER_FIXED_LENGTH > buf.length) {
            return undefined;
        }
        const flags = readU16(buf, offset + 8);
        const nameLength = readU16(buf, offset + 28);
        const extraLength = readU16(buf, offset + 30);
        const commentLength = readU16(buf, offset + 32);
        const attributes = readU32(buf, offset + 38);
        const nameStart = offset + CENTRAL_HEADER_FIXED_LENGTH;
        const nameEnd = nameStart + nameLength;
        const recordEnd = nameEnd + extraLength + commentLength;
        if (recordEnd > buf.length) {
            return undefined;
        }
        // `readCentralPart` refuses the encrypted bit inside the same walk.
        if (flags % 2 === 1) {
            return undefined;
        }
        // `readExtraFields` walks the extra region field by field, and each
        // field's declared length is bounds-checked against the whole buffer
        // by `readData` — a field inside a fitting region can declare more
        // data than the buffer holds, and JSZip throws there (#525).
        const extraEnd = nameEnd + extraLength;
        for (let cursor = nameEnd; cursor + 4 < extraEnd;) {
            const length = readU16(buf, cursor + 2);
            if (cursor + 4 + length > buf.length) {
                return undefined;
            }
            cursor += 4 + length;
        }
        const name = decodeName(buf, nameStart, nameLength, flags);
        records.push({
            name,
            nameBytes: buf.subarray(nameStart, nameEnd),
            extraBytes: buf.subarray(nameEnd, nameEnd + extraLength),
            commentBytes: buf.subarray(nameEnd + extraLength, recordEnd),
            recordOffset: offset,
            localHeaderOffset: readU32(buf, offset + 42),
            compressedSize: readU32(buf, offset + 20),
            diskNumberStart: readU16(buf, offset + 34),
            // A folder entry ends in `/` or sets the MS-DOS directory bit, as
            // JSZip marks `dir`.
            directory: name.endsWith("/") || attributes % 0x20 >= DOS_DIRECTORY_ATTRIBUTE,
        });
        offset = recordEnd;
    }
    // JSZip tolerates a count mismatch, but a count that names records where
    // none are found is "Corrupted zip or bug" — nothing is read.
    if (count !== 0 && records.length === 0) {
        return undefined;
    }
    // `readLocalFiles` follows the walk (`zipEntries.js` `load`): JSZip seeks
    // to every record's local header and reads its local part, so a record
    // whose local part its reader throws on refuses the view too — these are
    // the names of a package `load` never returns (#524, #555).
    for (const record of records) {
        if (localHeaderOffsetOf(buf, zero, record) === undefined) {
            return undefined;
        }
    }
    return {
        records,
        eocdOffset: eocd,
        eocdEnd: eocd + EOCD_FIXED_LENGTH + readU16(buf, eocd + 20),
        cdOffset,
        cdEnd: offset,
        zero,
        zip64,
    };
};

/**
 * The record's local-header offset as JSZip resolves it (`readCentralPart`
 * plus `parseZIP64ExtraField`): the base field read as `readJszipInt`, or —
 * when it saturates to -1 — the 8-byte value the record's ZIP64 extra field
 * (0x0001) stores for it, after the slots any saturated uncompressed and
 * compressed sizes consume first, in that order. `undefined` where JSZip's
 * `load` throws: no 0x0001 field for a saturated value, a value too short
 * for the read, a size still -1 after resolution, a negative effective
 * offset, no local signature at the resolved position, or a local part
 * `readLocalPart` cannot finish — its fixed 30-byte header, local name or
 * extra field past the buffer, a compressed stream past the buffer or one
 * whose final read `checkIndex` puts below zero (a resolved compressed size
 * that reads negative — the 32-bit field 0x80000000, or a ZIP64 value below
 * zero; a small negative size whose end stays at or above zero still reads),
 * or a compression method JSZip does not know (#490, #524, #555).
 */
const localHeaderOffsetOf = (buf: Uint8Array, zero: number, record: ZipCentralRecord): number | undefined => {
    const nameEnd = record.recordOffset + CENTRAL_HEADER_FIXED_LENGTH + record.nameBytes.length;
    const extraEnd = nameEnd + record.extraBytes.length;
    let zip64: Uint8Array | undefined;
    for (let cursor = nameEnd; cursor + 4 < extraEnd;) {
        const id = readU16(buf, cursor);
        const length = readU16(buf, cursor + 2);
        // `readData` throws when a declared field passes the buffer's end.
        if (cursor + 4 + length > buf.length) {
            return undefined;
        }
        if (id === ZIP64_EXTRA_FIELD) {
            zip64 = buf.subarray(cursor + 4, cursor + 4 + length);
        }
        cursor += 4 + length;
    }
    let uncompressed = readJszipInt(buf, record.recordOffset + 24);
    let compressed = readJszipInt(buf, record.recordOffset + 20);
    let offset = readJszipInt(buf, record.recordOffset + 42);
    if (uncompressed === JSZIP_MAX_32 || compressed === JSZIP_MAX_32 || offset === JSZIP_MAX_32) {
        if (zip64 === undefined) {
            return undefined;
        }
        const field = zip64;
        let cursor = 0;
        // `extraReader.readInt(8)`: the low 32 bits, signed; a value too
        // short for the read throws.
        const take = (): number | undefined => {
            if (cursor + 8 > field.length) {
                return undefined;
            }
            const value = readJszipInt(field, cursor);
            cursor += 8;
            return value;
        };
        if (uncompressed === JSZIP_MAX_32) {
            const value = take();
            if (value === undefined) {
                return undefined;
            }
            uncompressed = value;
        }
        if (compressed === JSZIP_MAX_32) {
            const value = take();
            if (value === undefined) {
                return undefined;
            }
            compressed = value;
        }
        if (offset === JSZIP_MAX_32) {
            const value = take();
            if (value === undefined) {
                return undefined;
            }
            offset = value;
        }
    }
    // `readLocalPart` refuses a size still -1; `readLocalFiles`' `setIndex`
    // refuses a negative offset.
    if (uncompressed === JSZIP_MAX_32 || compressed === JSZIP_MAX_32 || offset < 0) {
        return undefined;
    }
    // `readLocalFiles` seeks to `zero + offset` and reads the local signature.
    if (!hasSignature(buf, zero + offset, LOCAL_HEADER_SIGNATURE)) {
        return undefined;
    }
    // `readLocalPart` reads on: the fixed 30-byte header plus the local name
    // and extra field, then the central-declared compressed stream — each
    // bounds-checked against `zero + position` by `checkIndex`, which throws
    // on both sides: past the buffer, and below zero, where a resolved
    // compressed size that reads negative sends the archive-relative index of
    // the final `readData(compressedSize)` (#555) — and the compression
    // method must be one JSZip knows — store (0) or deflate (8), read from
    // the central record (`compressions.js`) (#524).
    const header = zero + offset;
    if (header + LOCAL_HEADER_FIXED_LENGTH > buf.length) {
        return undefined;
    }
    const nameAndExtra = readU16(buf, header + LOCAL_NAME_LENGTH_OFFSET) + readU16(buf, header + LOCAL_EXTRA_LENGTH_OFFSET);
    const dataEnd = offset + LOCAL_HEADER_FIXED_LENGTH + nameAndExtra + compressed;
    if (dataEnd < 0 || zero + dataEnd > buf.length) {
        return undefined;
    }
    const method = readU16(buf, record.recordOffset + COMPRESSION_METHOD_OFFSET);
    if (method !== STORE_METHOD && method !== DEFLATE_METHOD) {
        return undefined;
    }
    return offset;
};

/**
 * The bytes in front of the archive JSZip loads, or `undefined` where its
 * reader throws — through `readLocalPart`'s own bounded reads, not just the
 * signature check (#464, #490, #524). JSZip reads each entry's local header at
 * `zero + localHeaderOffset` (`dataReader`'s `byteAt` adds the rebase), the
 * offset resolved as {@link localHeaderOffsetOf} resolves it — the ZIP64
 * extra field's value where the base field is saturated — so the archive it
 * loads begins at the lowest of those positions, which covers a
 * self-extractor's absolute offsets (`zero` 0, first offset raised) and a
 * prepended archive (`zero` raised, offsets relative). An empty directory
 * names no local header; the rebase alone is the prefix.
 */
export const zipBytesBeforeArchive = (buf: Uint8Array): number | undefined => {
    const view = readCentralDirectoryView(buf);
    if (view === undefined) {
        return undefined;
    }
    let lowest: number | undefined = undefined;
    for (const record of view.records) {
        // The view's own `readLocalFiles` pass has refused any record whose
        // local part its reader throws on, so this never refuses; the guard
        // keeps the measure's contract local.
        const offset = localHeaderOffsetOf(buf, view.zero, record);
        if (offset === undefined) {
            return undefined;
        }
        lowest = lowest === undefined ? offset : Math.min(lowest, offset);
    }
    const prefix = lowest === undefined ? view.zero : view.zero + lowest;
    return Number.isFinite(prefix) ? prefix : undefined;
};
