/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Word reads each ZIP entry's central-directory name and ignores the local
 * file header's name and the Info-ZIP unicode-path extra field (0x7075).
 * JSZip names its entries by the local name and prefers the central record's
 * 0x7075, so archive bytes are normalized before they are loaded: the package
 * view and extraction are keyed by the names Word reads (#445). The
 * normalization never decides whether a package opens: only an archive JSZip
 * loads as given is re-keyed (#468).
 */

import JSZip from "jszip";

import {
    hasSignature,
    readCentralDirectoryView,
    readU16,
    ZIP64_LOCATOR_LENGTH,
    ZIP64_LOCATOR_SIGNATURE,
    ZIP64_OFFSET,
    type ZipCentralRecord,
} from "./zip-central-directory";

const LOCAL_HEADER_SIGNATURE = 0x04034b50;
const LOCAL_HEADER_FIXED_LENGTH = 30;
const CENTRAL_HEADER_FIXED_LENGTH = 46;
const UNICODE_PATH_FIELD = 0x7075;
const EOCD_CD_SIZE_OFFSET = 12;
const EOCD_CD_OFFSET_OFFSET = 16;
const LOCAL_NAME_LENGTH_OFFSET = 26;
const LOCAL_EXTRA_LENGTH_OFFSET = 28;
const CENTRAL_EXTRA_LENGTH_OFFSET = 30;
const CENTRAL_LOCAL_OFFSET_OFFSET = 42;

/** One central record's local header and compressed stream, located. */
type LocalRecord = {
    /** The local header's offset in the archive (its signature). */
    readonly header: number;
    /** Where the stored name starts: the header plus its fixed 30 bytes. */
    readonly nameStart: number;
    /** The stored name's raw bytes. */
    readonly nameBytes: Uint8Array;
    /** The local record's extra-field bytes. */
    readonly extraBytes: Uint8Array;
    /** Where the compressed stream starts. */
    readonly dataStart: number;
    /** One past the compressed stream the central record declares. */
    readonly dataEnd: number;
};

/**
 * The entry's local record, or `undefined` when it cannot be read: no local
 * signature, or the name, extra field or central-declared stream overruns
 * `limit` (the central directory's start). The stored offset is relative to
 * the archive, so `zero` — the bytes before it — rebases it, as JSZip does.
 */
const localLayout = (bytes: Uint8Array, record: ZipCentralRecord, limit: number, zero: number): LocalRecord | undefined => {
    const header = record.localHeaderOffset + zero;
    if (!hasSignature(bytes, header, LOCAL_HEADER_SIGNATURE)) {
        return undefined;
    }
    const nameStart = header + LOCAL_HEADER_FIXED_LENGTH;
    if (nameStart > limit) {
        return undefined;
    }
    const nameEnd = nameStart + readU16(bytes, header + LOCAL_NAME_LENGTH_OFFSET);
    const dataStart = nameEnd + readU16(bytes, header + LOCAL_EXTRA_LENGTH_OFFSET);
    const dataEnd = dataStart + record.compressedSize;
    if (dataEnd > limit) {
        return undefined;
    }
    return {
        header,
        nameStart,
        nameBytes: bytes.subarray(nameStart, nameEnd),
        extraBytes: bytes.subarray(nameEnd, dataStart),
        dataStart,
        dataEnd,
    };
};

const sameBytes = (a: Uint8Array, b: Uint8Array): boolean => a.length === b.length && a.every((byte, i) => byte === b[i]);

/**
 * The extra field without its Info-ZIP unicode path (0x7075): JSZip prefers
 * the central record's over the stored name, and Word ignores it. A
 * truncated extra field keeps its unparseable tail verbatim, as JSZip's
 * reader does.
 */
const withoutUnicodePathField = (extra: Uint8Array): Uint8Array => {
    const out: number[] = [];
    let position = 0;
    while (position + 4 <= extra.length) {
        const id = readU16(extra, position);
        const length = readU16(extra, position + 2);
        const end = position + 4 + length;
        if (end > extra.length) {
            break;
        }
        if (id !== UNICODE_PATH_FIELD) {
            for (let i = position; i < end; i += 1) {
                out.push(extra[i]!);
            }
        }
        position = end;
    }
    for (let i = position; i < extra.length; i += 1) {
        out.push(extra[i]!);
    }
    return Uint8Array.from(out);
};

/**
 * The archive bytes to feed `JSZip.loadAsync` when the package view must be
 * keyed by the names Word reads — the central directory's (#445). A packed
 * input whose local names agree with the central directory, and whose central
 * records carry no 0x7075 unicode-path field, comes back unchanged; any other
 * is rebuilt in memory: every local header stores the central name, the
 * central records drop their 0x7075 fields, and the central records' offsets
 * and the EOCD's directory fields follow the rebuild. JSZip skips the local
 * extra field (`zipEntry.js` `readLocalPart`), so it is kept as stored (#469).
 *
 * The directory is the one JSZip reads — the last EOCD signature, every
 * offset rebased past the bytes before the archive (#467) — and the rebuilt
 * offsets stay relative to the archive, so a prefix stays a prefix. Every
 * other byte — the prefix, data descriptors, gaps between entries, the EOCD
 * comment, trailing bytes — is kept verbatim.
 *
 * Archives the rewrite cannot honor are returned as given, so their findings
 * stay JSZip's: a central directory `readCentralDirectoryView` cannot parse,
 * ZIP64 or multi-disk structure, bytes between the directory and the EOCD, a
 * local header missing or overlapping another entry's data, or a compressed
 * stream that overruns the directory.
 */
export const centralKeyedArchive = (bytes: Uint8Array): Uint8Array => {
    const view = readCentralDirectoryView(bytes);
    // Bytes between the directory and the EOCD are what JSZip reads as a
    // prefix, so a view with them is never built; a rebuild that moved the
    // EOCD up to the directory would drop them (#468).
    if (view === undefined || view.zip64 || view.cdEnd !== view.eocdOffset) {
        return bytes;
    }
    // A ZIP64 locator record sits between the directory and the EOCD; leave
    // its real offsets to JSZip rather than writing a locator for a directory
    // that moved.
    if (hasSignature(bytes, view.eocdOffset - ZIP64_LOCATOR_LENGTH, ZIP64_LOCATOR_SIGNATURE)) {
        return bytes;
    }
    const locals: LocalRecord[] = [];
    let changed = false;
    for (const record of view.records) {
        // Saturated fields mean the real value lives in a ZIP64 extra field
        // the rewrite does not move. Multi-disk entries likewise keep the
        // bytes they came with.
        if (record.localHeaderOffset === ZIP64_OFFSET || record.compressedSize === ZIP64_OFFSET || record.diskNumberStart !== 0) {
            return bytes;
        }
        const local = localLayout(bytes, record, view.cdOffset, view.zero);
        if (local === undefined) {
            return bytes;
        }
        locals.push(local);
        if (
            !sameBytes(local.nameBytes, record.nameBytes) ||
            withoutUnicodePathField(record.extraBytes).length !== record.extraBytes.length
        ) {
            changed = true;
        }
    }
    if (!changed) {
        return bytes;
    }
    const order = locals.map((local, index) => ({ local, index })).sort((a, b) => a.local.header - b.local.header || a.index - b.index);
    for (let i = 1; i < order.length; i += 1) {
        // Two records pointing at one local record, or overlapping streams:
        // the rebuild cannot order their bytes, so the archive stays as is.
        if (order[i]!.local.header === order[i - 1]!.local.header || order[i]!.local.header < order[i - 1]!.local.dataEnd) {
            return bytes;
        }
    }
    const parts: Uint8Array[] = [];
    let outLength = 0;
    const push = (part: Uint8Array): void => {
        parts.push(part);
        outLength += part.length;
    };
    let cursor = 0;
    const newLocalOffset: number[] = [];
    for (const { local, index } of order) {
        const record = view.records[index]!;
        // Verbatim bytes up to this local header: the prefix, other junk, and
        // the previous entry's data descriptor when flag bit 3 records one.
        push(bytes.subarray(cursor, local.header));
        // Stored relative to the archive, as the original stored it.
        newLocalOffset[index] = outLength - view.zero;
        const header = Buffer.from(bytes.subarray(local.header, local.nameStart));
        header.writeUInt16LE(record.nameBytes.length, LOCAL_NAME_LENGTH_OFFSET);
        push(header);
        push(record.nameBytes);
        push(local.extraBytes);
        push(bytes.subarray(local.dataStart, local.dataEnd));
        cursor = local.dataEnd;
    }
    // The bytes between the last stream and the directory, verbatim.
    push(bytes.subarray(cursor, view.cdOffset));
    const cdOffset = outLength;
    for (const [index, record] of view.records.entries()) {
        const strippedExtra = withoutUnicodePathField(record.extraBytes);
        const header = Buffer.from(bytes.subarray(record.recordOffset, record.recordOffset + CENTRAL_HEADER_FIXED_LENGTH));
        header.writeUInt16LE(strippedExtra.length, CENTRAL_EXTRA_LENGTH_OFFSET);
        header.writeUInt32LE(newLocalOffset[index]!, CENTRAL_LOCAL_OFFSET_OFFSET);
        push(header);
        push(record.nameBytes);
        push(strippedExtra);
        push(record.commentBytes);
    }
    const eocd = Buffer.from(bytes.subarray(view.eocdOffset, view.eocdEnd));
    eocd.writeUInt32LE(outLength - cdOffset, EOCD_CD_SIZE_OFFSET);
    eocd.writeUInt32LE(cdOffset - view.zero, EOCD_CD_OFFSET_OFFSET);
    push(eocd);
    push(bytes.subarray(view.eocdEnd));
    return Buffer.concat(parts.map((part) => Buffer.from(part.buffer, part.byteOffset, part.byteLength)));
};

/**
 * `JSZip.loadAsync` over the central-keyed view of `bytes`. The bytes as
 * given are loaded first, so an archive JSZip refuses is refused with JSZip's
 * error: re-keying never turns a package that does not open into one that
 * does (#468).
 */
export const loadCentralNamedArchive = async (bytes: Uint8Array): Promise<JSZip> => {
    const zip = await JSZip.loadAsync(bytes);
    const keyed = centralKeyedArchive(bytes);
    return keyed === bytes ? zip : JSZip.loadAsync(keyed);
};
