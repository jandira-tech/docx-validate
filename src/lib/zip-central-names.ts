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
    decodeCp437,
    hasSignature,
    readCentralDirectoryView,
    readU16,
    ZIP64_LOCATOR_LENGTH,
    ZIP64_LOCATOR_SIGNATURE,
    ZIP64_OFFSET,
    type ZipCentralRecord,
    LOCAL_HEADER_SIGNATURE,
    readU32,
    type CentralDirectoryView,
} from "./zip-central-directory";

const LOCAL_HEADER_FIXED_LENGTH = 30;
const CENTRAL_HEADER_FIXED_LENGTH = 46;
const UNICODE_PATH_FIELD = 0x7075;
const EOCD_CD_SIZE_OFFSET = 12;
const EOCD_CD_OFFSET_OFFSET = 16;
const LOCAL_NAME_LENGTH_OFFSET = 26;
const LOCAL_EXTRA_LENGTH_OFFSET = 28;
const CENTRAL_EXTRA_LENGTH_OFFSET = 30;
const ZIP64_EOCD_SIGNATURE = 0x06064b50;
const ZIP64_EOCD_FIXED_LENGTH = 56;
const ZIP64_EOCD_SIZE_OFFSET = 4;
const ZIP64_EOCD_CD_SIZE_OFFSET = 40;
const ZIP64_EOCD_CD_OFFSET_OFFSET = 48;
const ZIP64_LOCATOR_RECORD_OFFSET_OFFSET = 8;
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
 * The last occurrence of `signature`, scanning as the directory reader's
 * `lastIndexOfSignature` does.
 */
const lastOffsetOf = (bytes: Uint8Array, signature: number): number | undefined => {
    for (let offset = bytes.length - 4; offset >= 0; offset -= 1) {
        if (bytes[offset] === 0x50 && readU32(bytes, offset) === signature) {
            return offset;
        }
    }
    return undefined;
};

/** The ZIP64 end-of-central-directory record and locator, located. */
type Zip64Tail = {
    /** The record's offset in the archive (its signature). */
    readonly record: number;
    /** The locator's offset in the archive (its signature). */
    readonly locator: number;
};

/**
 * The ZIP64 tail the rewrite can move: the record and the locator JSZip
 * reads — the last of each signature, found as JSZip finds them, the record
 * named by the locator's offset field (`isSignature`, whose `zero` is still 0
 * there) or, failing that, the last record signature — adjacent, the record
 * before the locator, the locator before the EOCD, and the record declaring
 * its fixed 44 bytes. Anything else — extensible data, a record away from its
 * locator, a later signature in a comment — is left to JSZip as given, since
 * the rebuild moves only a tail it can re-point at itself (#479).
 */
const zip64Tail = (bytes: Uint8Array, view: CentralDirectoryView): Zip64Tail | undefined => {
    const locator = view.eocdOffset - ZIP64_LOCATOR_LENGTH;
    if (locator < 0 || lastOffsetOf(bytes, ZIP64_LOCATOR_SIGNATURE) !== locator) {
        return undefined;
    }
    const named = readU32(bytes, locator + ZIP64_LOCATOR_RECORD_OFFSET_OFFSET);
    const record = hasSignature(bytes, named, ZIP64_EOCD_SIGNATURE) ? named : lastOffsetOf(bytes, ZIP64_EOCD_SIGNATURE);
    if (record === undefined || !hasSignature(bytes, record, ZIP64_EOCD_SIGNATURE)) {
        return undefined;
    }
    // A record with extensible data, or short of the fixed 44 bytes JSZip
    // reads at fixed offsets, does not move.
    if (readU32(bytes, record + ZIP64_EOCD_SIZE_OFFSET) !== ZIP64_EOCD_FIXED_LENGTH - 12) {
        return undefined;
    }
    // The record ends exactly where the locator begins: JSZip's rebase
    // arithmetic (`expectedEndOfCentralDirOffset`) assumes the two, then the
    // EOCD, are all that follows the directory.
    if (record + ZIP64_EOCD_FIXED_LENGTH !== locator || record < view.cdEnd) {
        return undefined;
    }
    return { record, locator };
};

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
 * A ZIP64 package is rebuilt too (#479): its record and locator — the last of
 * each signature, the record adjacent to the locator, the locator to the EOCD
 * — move with the directory, and the record's directory size and offset and
 * the locator's record offset are rebased like the 32-bit EOCD's fields,
 * counting any bytes a declared size spans past its records (#478). The
 * 32-bit EOCD keeps its saturated directory fields as stored: JSZip reads
 * them only to decide to follow the ZIP64 structures.
 *
 * Archives the rewrite cannot honor are returned as given, so their findings
 * stay JSZip's: a central directory `readCentralDirectoryView` cannot parse,
 * a ZIP64 tail it cannot move (a record with extensible data or short of its
 * fixed 44 bytes, or one not adjacent to its locator), an entry whose
 * local-header offset or compressed size lives in a ZIP64 extra field, bytes
 * between the walked directory's end and the EOCD, a local header missing or
 * overlapping another entry's data, or a compressed stream that overruns the
 * directory.
 */
export const centralKeyedArchive = (bytes: Uint8Array): Uint8Array => {
    const view = readCentralDirectoryView(bytes);
    if (view === undefined) {
        return bytes;
    }
    let zip64: Zip64Tail | undefined;
    if (view.zip64) {
        // A ZIP64 record and locator sit between the directory and the EOCD;
        // the rewrite moves them with the directory, rebasing their offsets
        // (#479).
        zip64 = zip64Tail(bytes, view);
        if (zip64 === undefined) {
            return bytes;
        }
    } else {
        // Bytes between the walked directory's end (`cdEnd`, #478) and the
        // EOCD are what JSZip reads past the last record — a signature record,
        // a hidden directory — so a view with them is never rebuilt; moving
        // the EOCD up to the directory would drop them (#468).
        if (view.cdEnd !== view.eocdOffset) {
            return bytes;
        }
        // A ZIP64 locator record sits between the directory and the EOCD;
        // leave its real offsets to JSZip rather than writing a locator for a
        // directory that moved.
        if (hasSignature(bytes, view.eocdOffset - ZIP64_LOCATOR_LENGTH, ZIP64_LOCATOR_SIGNATURE)) {
            return bytes;
        }
    }
    const locals: LocalRecord[] = [];
    let changed = false;
    for (const record of view.records) {
        // Saturated fields mean the real value lives in a ZIP64 extra field
        // the rewrite does not move; such entries keep their bytes. Of the
        // multi-disk fields, this record's `diskNumberStart` is the only one
        // consulted here, and the locator's disk count is refused in
        // `readCentralDirectoryView`; the EOCD's own disk fields are read only
        // to detect ZIP64, as JSZip reads them (#482).
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
    if (zip64 === undefined) {
        const eocd = Buffer.from(bytes.subarray(view.eocdOffset, view.eocdEnd));
        eocd.writeUInt32LE(outLength - cdOffset, EOCD_CD_SIZE_OFFSET);
        eocd.writeUInt32LE(cdOffset - view.zero, EOCD_CD_OFFSET_OFFSET);
        push(eocd);
    } else {
        // The bytes between the walked directory and the ZIP64 record — a
        // declared size spanning past its records (#478) — stay, and the
        // record keeps counting them: JSZip locates the EOCD from these
        // fields (`expectedEndOfCentralDirOffset`), so they must keep naming
        // a position exactly the record and the locator (76 bytes) before the
        // EOCD, or every offset would re-base past the difference.
        push(bytes.subarray(view.cdEnd, zip64.record));
        const zip64RecordOffset = outLength;
        const record = Buffer.from(bytes.subarray(zip64.record, zip64.locator));
        record.writeBigUInt64LE(BigInt(zip64RecordOffset - cdOffset), ZIP64_EOCD_CD_SIZE_OFFSET);
        record.writeBigUInt64LE(BigInt(cdOffset - view.zero), ZIP64_EOCD_CD_OFFSET_OFFSET);
        push(record);
        const locator = Buffer.from(bytes.subarray(zip64.locator, view.eocdOffset));
        locator.writeBigUInt64LE(BigInt(zip64RecordOffset - view.zero), ZIP64_LOCATOR_RECORD_OFFSET_OFFSET);
        push(locator);
        // The 32-bit EOCD keeps its saturated directory fields as stored:
        // JSZip reads them only to decide to follow the ZIP64 structures,
        // whose offsets are the rebased ones (#479).
        push(bytes.subarray(view.eocdOffset, view.eocdEnd));
    }
    push(bytes.subarray(view.eocdEnd));
    return Buffer.concat(parts.map((part) => Buffer.from(part.buffer, part.byteOffset, part.byteLength)));
};

/**
 * `JSZip.loadAsync` over the central-keyed view of `bytes`. The bytes as
 * given are loaded first, so an archive JSZip refuses is refused with JSZip's
 * error: re-keying never turns a package that does not open into one that
 * does (#468).
 * Both loads share the stored-name reader's CP437 fallback (#480). JSZip
 * still handles bit-11 UTF-8 and valid central Unicode Path fields before
 * calling it; re-keying deliberately strips the latter (#445). The callback
 * also decodes ZIP comments, which the validator does not judge.
 */
export const loadCentralNamedArchive = async (bytes: Uint8Array): Promise<JSZip> => {
    // JSZip supplies Uint8Array on Node/Bun; its types also list string[].
    const options = { decodeFileName: (name: Uint8Array | string[]) => decodeCp437(name as Uint8Array) };
    const zip = await JSZip.loadAsync(bytes, options);
    const keyed = centralKeyedArchive(bytes);
    return keyed === bytes ? zip : JSZip.loadAsync(keyed, options);
};
