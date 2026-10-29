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
    EOCD_FIXED_LENGTH,
} from "./zip-central-directory";

const LOCAL_HEADER_FIXED_LENGTH = 30;
const EOCD_DISK_ENTRIES_OFFSET = 8;
const EOCD_TOTAL_ENTRIES_OFFSET = 10;
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
const ZIP64_EOCD_DISK_ENTRIES_OFFSET = 24;
const ZIP64_EOCD_TOTAL_ENTRIES_OFFSET = 32;
const ZIP64_EOCD_CD_SIZE_OFFSET = 40;
const ZIP64_EOCD_CD_OFFSET_OFFSET = 48;
const ZIP64_LOCATOR_RECORD_OFFSET_OFFSET = 8;
const CENTRAL_LOCAL_OFFSET_OFFSET = 42;
/** A saturated 16-bit field, as a ZIP64 archive stores it in the 32-bit EOCD. */
const ZIP64_COUNT = 0xffff;

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
 * its fixed 44 bytes. Anything else — extensible data, or a record away from
 * its locator — is left to JSZip as given, since the rebuild moves only a
 * tail it can re-point at itself (#479). The tail never looks past the named
 * record: a later record signature rides in the EOCD comment, which moves
 * verbatim and cannot hijack the load — the rewrite names the moved record
 * absolutely, so JSZip's signature check never falls back (#521).
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
 * — move with the directory, the record's directory size and offset rebased
 * like the 32-bit EOCD's fields, counting any bytes a declared size spans
 * past its records (#478), and the locator's record offset kept absolute:
 * JSZip reads that field before its rebase, so a relative value would send
 * its signature check to the last-signature fallback, which a moved EOCD
 * comment can win (#521). The 32-bit EOCD keeps its saturated directory
 * fields as stored: JSZip reads them only to decide to follow the ZIP64
 * structures.
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
        // The locator's record offset names the record's absolute position in
        // the buffer, not a rebased one: JSZip tests it before `reader.zero`
        // is assigned, so a relative value would miss and send the load to
        // the last-signature fallback, which a re-keyed EOCD comment can win
        // (#521).
        const locator = Buffer.from(bytes.subarray(zip64.locator, view.eocdOffset));
        locator.writeBigUInt64LE(BigInt(zip64RecordOffset), ZIP64_LOCATOR_RECORD_OFFSET_OFFSET);
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
 * The archive bytes without the central records `drop` matches, so the rest
 * of a package can be read when one entry is refused (#629). JSZip keys its
 * view by resolved name, and a refused `word/x/../document.xml` stored after
 * `word/document.xml` replaces it there; leaving the refused record out of the
 * directory brings back the entry it hid, where dropping it from the loaded
 * view would lose both. `drop` sees every record, folder records included.
 *
 * Everything before the directory stays verbatim: the prefix, every local
 * record (a dropped one stays, but no record names it any more) and the gaps
 * between them. So do the EOCD's comment and any bytes after it. The kept
 * records keep their bytes and offsets, and the EOCD's entry counts and
 * directory size follow them; its directory offset is unchanged, so the bytes
 * before the archive keep rebasing every offset as before (#467).
 *
 * A ZIP64 archive is rewritten too, whenever the tail `centralKeyedArchive`
 * moves (#479): the record and the locator follow the directory it shrinks —
 * the bytes between the walked directory and the record stay verbatim, as
 * `centralKeyedArchive` keeps them — the record's entry counts (both) and
 * directory size follow the kept records, its directory offset is unchanged,
 * and the locator's record offset names the record's new absolute position
 * (#521). The 32-bit EOCD's counts and directory size are rewritten only
 * where they are not saturated; a saturated field defers to the ZIP64 record,
 * whose readers use it, and stays as stored.
 *
 * The bytes come back as given when no record matches, and `undefined` when
 * the rewrite cannot honor the envelope: a directory
 * `readCentralDirectoryView` cannot parse, a ZIP64 tail it cannot move (a
 * record with extensible data or short of its fixed 44 bytes, or one not
 * adjacent to its locator), or — outside ZIP64 — bytes between the walked
 * directory and the EOCD, which moving the EOCD up would drop (#468, #478).
 */
export const archiveWithoutEntries = (bytes: Uint8Array, drop: (record: ZipCentralRecord) => boolean): Uint8Array | undefined => {
    const view = readCentralDirectoryView(bytes);
    if (view === undefined) {
        return undefined;
    }
    // A ZIP64 package moves its tail with the directory, as `centralKeyedArchive`
    // moves it (#479); the non-ZIP64 bytes between the walked directory and the
    // EOCD stay refused, which moving the EOCD up would drop (#468).
    const zip64 = view.zip64 ? zip64Tail(bytes, view) : undefined;
    if ((view.zip64 && zip64 === undefined) || (!view.zip64 && view.cdEnd !== view.eocdOffset)) {
        return undefined;
    }
    const kept = view.records.filter((record) => !drop(record));
    if (kept.length === view.records.length) {
        return bytes;
    }
    const records = kept.map((record) =>
        bytes.subarray(
            record.recordOffset,
            record.recordOffset +
                CENTRAL_HEADER_FIXED_LENGTH +
                record.nameBytes.length +
                record.extraBytes.length +
                record.commentBytes.length,
        ),
    );
    // What the moved tail counts as the directory: the kept records, plus the
    // bytes a ZIP64 record's declared size spans past them (#478).
    const between = zip64 === undefined ? new Uint8Array() : bytes.subarray(view.cdEnd, zip64.record);
    const directorySize = records.reduce((size, record) => size + record.length, 0) + between.length;
    const eocd = Buffer.from(bytes.subarray(view.eocdOffset, view.eocdOffset + EOCD_FIXED_LENGTH));
    if (zip64 === undefined) {
        eocd.writeUInt16LE(kept.length, EOCD_DISK_ENTRIES_OFFSET);
        eocd.writeUInt16LE(kept.length, EOCD_TOTAL_ENTRIES_OFFSET);
        eocd.writeUInt32LE(directorySize, EOCD_CD_SIZE_OFFSET);
        const parts = [bytes.subarray(0, view.cdOffset), ...records, eocd, bytes.subarray(view.eocdOffset + EOCD_FIXED_LENGTH)];
        return Buffer.concat(parts.map((part) => Buffer.from(part.buffer, part.byteOffset, part.byteLength)));
    }
    // The ZIP64 record keeps its directory offset: the prefix is verbatim and
    // the directory does not move, so the stored (rebased) value stays true.
    const tailRecord = Buffer.from(bytes.subarray(zip64.record, zip64.record + ZIP64_EOCD_FIXED_LENGTH));
    tailRecord.writeBigUInt64LE(BigInt(kept.length), ZIP64_EOCD_DISK_ENTRIES_OFFSET);
    tailRecord.writeBigUInt64LE(BigInt(kept.length), ZIP64_EOCD_TOTAL_ENTRIES_OFFSET);
    tailRecord.writeBigUInt64LE(BigInt(directorySize), ZIP64_EOCD_CD_SIZE_OFFSET);
    // The locator's record offset names the record's absolute position in the
    // buffer, not a rebased one: JSZip tests it before `reader.zero` is
    // assigned, so a relative value would miss (#521).
    const recordOffset = view.cdOffset + directorySize;
    const tailLocator = Buffer.from(bytes.subarray(zip64.locator, zip64.locator + ZIP64_LOCATOR_LENGTH));
    tailLocator.writeBigUInt64LE(BigInt(recordOffset), ZIP64_LOCATOR_RECORD_OFFSET_OFFSET);
    // A saturated field defers to the ZIP64 record, whose readers use it; an
    // unsaturated one keeps naming the subset's own counts and size.
    if (readU16(eocd, EOCD_DISK_ENTRIES_OFFSET) !== ZIP64_COUNT) {
        eocd.writeUInt16LE(kept.length, EOCD_DISK_ENTRIES_OFFSET);
    }
    if (readU16(eocd, EOCD_TOTAL_ENTRIES_OFFSET) !== ZIP64_COUNT) {
        eocd.writeUInt16LE(kept.length, EOCD_TOTAL_ENTRIES_OFFSET);
    }
    if (readU32(eocd, EOCD_CD_SIZE_OFFSET) !== ZIP64_OFFSET) {
        eocd.writeUInt32LE(directorySize, EOCD_CD_SIZE_OFFSET);
    }
    const tailParts = [
        bytes.subarray(0, view.cdOffset),
        ...records,
        between,
        tailRecord,
        tailLocator,
        eocd,
        bytes.subarray(view.eocdOffset + EOCD_FIXED_LENGTH),
    ];
    return Buffer.concat(tailParts.map((part) => Buffer.from(part.buffer, part.byteOffset, part.byteLength)));
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
