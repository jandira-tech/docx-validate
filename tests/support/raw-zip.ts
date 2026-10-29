/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Write a ZIP archive byte by byte. JSZip cannot produce some envelopes a
 * real archiver can — two central-directory entries under the identical
 * stored name, or a name with a `..` segment — so tests build them here.
 * Entries are stored (method 0, no compression) in the order given.
 */

const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n += 1) {
        let c = n;
        for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
        table[n] = c >>> 0;
    }
    return table;
})();

const crc32 = (data: Uint8Array): number => {
    let c = 0xffffffff;
    for (const byte of data) c = CRC_TABLE[(c ^ byte) & 0xff]! ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
};

/**
 * The Info-ZIP unicode-path extra field (0x7075): version 1, the CRC-32 of
 * the local header's name bytes, then the UTF-8 unicode name. JSZip honours
 * the central record's field when its CRC matches the local name it read, and
 * prefers it over that name; Word ignores it (#445).
 */
const unicodePathField = (localNameBytes: Uint8Array, unicodeName: string): Uint8Array => {
    const name = encoder.encode(unicodeName);
    const field = Buffer.alloc(9 + name.length);
    field.writeUInt16LE(0x7075, 0);
    field.writeUInt16LE(5 + name.length, 2);
    field.writeUInt8(1, 4);
    field.writeUInt32LE(crc32(localNameBytes), 5);
    name.forEach((byte, i) => {
        field.writeUInt8(byte, 9 + i);
    });
    return field;
};

export interface RawZipEntry {
    /** The item's central-directory name, stored verbatim (UTF-8 encoded, or `nameBytes` raw). */
    readonly name: string;
    /** Overrides the UTF-8 encoding of `name`, e.g. for CP437 specimens. */
    readonly nameBytes?: Uint8Array;
    /** Local-header name; defaults to the central name. Real archivers can store a different name here, and Word reads only the central one (#445). */
    readonly localName?: string;
    /** Overrides the UTF-8 encoding of `localName`. */
    readonly localNameBytes?: Uint8Array;
    /** Writes an Info-ZIP unicode-path extra field (0x7075) naming this unicode path, in both the local and the central record. */
    readonly unicodePath?: string;
    readonly content?: string | Uint8Array;
    /** Set general-purpose bit 11, marking `name` as UTF-8. */
    readonly utf8?: boolean;
    /** Set the MS-DOS directory attribute bit, as a folder entry does. */
    readonly directory?: boolean;
    /**
     * Store the sentinel 0xFFFFFFFF in this central record's local-header
     * offset, with the real offset in a ZIP64 (0x0001) extra field in the
     * central record, as an entry of an archive past 4 GiB does. JSZip loads
     * the entry through the extra field; the re-keying rewrite bails on the
     * sentinel (#479).
     */
    readonly zip64LocalOffset?: boolean;
}

const encoder = new TextEncoder();

/** The ZIP64 (0x0001) extra field holding the local-header offset, as an archive past 4 GiB stores per entry (#479). */
const zip64OffsetField = (offset: number): Buffer => {
    const field = Buffer.alloc(12);
    field.writeUInt16LE(0x0001, 0);
    field.writeUInt16LE(8, 2);
    field.writeBigUInt64LE(BigInt(offset), 4);
    return field;
};

/**
 * The archive bytes: local headers, then the central directory, then the
 * EOCD record. `zip64` emits a ZIP64 EOCD record and locator — the EOCD's
 * central-directory fields saturated — around the same directory, so readers
 * exercise the ZIP64 path without a multi-gigabyte file. `cdSlack` (ZIP64
 * only) stores junk between the directory and the ZIP64 record that the
 * record's declared directory size counts, as a size spanning past its
 * records does (#478).
 */
export const buildRawZip = (
    entries: readonly RawZipEntry[],
    options: { readonly zip64?: boolean; readonly cdSlack?: number } = {},
): Buffer => {
    const local: Buffer[] = [];
    const central: Buffer[] = [];
    let offset = 0;
    for (const entry of entries) {
        const nameBytes = entry.nameBytes ?? encoder.encode(entry.name);
        const localNameBytes = entry.localNameBytes ?? (entry.localName !== undefined ? encoder.encode(entry.localName) : nameBytes);
        const data = typeof entry.content === "string" ? encoder.encode(entry.content) : (entry.content ?? new Uint8Array());
        const crc = crc32(data);
        const flags = entry.utf8 === true ? 0x0800 : 0;
        const attributes = entry.directory === true ? 0x10 : 0;
        const unicodePath = entry.unicodePath === undefined ? new Uint8Array() : unicodePathField(localNameBytes, entry.unicodePath);
        const centralExtra =
            entry.zip64LocalOffset === true
                ? Buffer.concat([Buffer.from(unicodePath), zip64OffsetField(offset)])
                : Buffer.from(unicodePath);

        const lh = Buffer.alloc(30);
        lh.writeUInt32LE(0x04034b50, 0);
        lh.writeUInt16LE(20, 4);
        lh.writeUInt16LE(flags, 6);
        lh.writeUInt32LE(crc, 14);
        lh.writeUInt32LE(data.length, 18);
        lh.writeUInt32LE(data.length, 22);
        lh.writeUInt16LE(localNameBytes.length, 26);
        lh.writeUInt16LE(unicodePath.length, 28);
        local.push(lh, Buffer.from(localNameBytes), Buffer.from(unicodePath), Buffer.from(data));

        const ch = Buffer.alloc(46);
        ch.writeUInt32LE(0x02014b50, 0);
        ch.writeUInt16LE(20, 4);
        ch.writeUInt16LE(20, 6);
        ch.writeUInt16LE(flags, 8);
        ch.writeUInt32LE(crc, 16);
        ch.writeUInt32LE(data.length, 20);
        ch.writeUInt32LE(data.length, 24);
        ch.writeUInt16LE(nameBytes.length, 28);
        ch.writeUInt16LE(centralExtra.length, 30);
        ch.writeUInt32LE(attributes, 38);
        ch.writeUInt32LE(entry.zip64LocalOffset === true ? 0xffffffff : offset, 42);
        central.push(ch, Buffer.from(nameBytes), centralExtra);

        offset += 30 + localNameBytes.length + unicodePath.length + data.length;
    }
    const cd = Buffer.concat(central);
    const zip64 = options.zip64 === true;
    const slack = Math.max(0, options.cdSlack ?? 0);
    const records: Buffer[] = [...local, cd];
    if (slack > 0) {
        records.push(Buffer.alloc(slack, 0x5a));
    }
    if (zip64) {
        const eocd64 = Buffer.alloc(56);
        eocd64.writeUInt32LE(0x06064b50, 0);
        eocd64.writeBigUInt64LE(44n, 4);
        eocd64.writeUInt16LE(20, 12);
        eocd64.writeUInt16LE(45, 14);
        eocd64.writeBigUInt64LE(BigInt(entries.length), 24);
        eocd64.writeBigUInt64LE(BigInt(entries.length), 32);
        eocd64.writeBigUInt64LE(BigInt(cd.length + slack), 40);
        eocd64.writeBigUInt64LE(BigInt(offset), 48);
        const locator = Buffer.alloc(20);
        locator.writeUInt32LE(0x07064b50, 0);
        locator.writeBigUInt64LE(BigInt(offset + cd.length + slack), 8);
        locator.writeUInt32LE(1, 16);
        records.push(eocd64, locator);
    }
    const eocd = Buffer.alloc(22);
    eocd.writeUInt32LE(0x06054b50, 0);
    eocd.writeUInt16LE(zip64 ? 0xffff : entries.length, 8);
    eocd.writeUInt16LE(zip64 ? 0xffff : entries.length, 10);
    eocd.writeUInt32LE(zip64 ? 0xffffffff : cd.length, 12);
    eocd.writeUInt32LE(zip64 ? 0xffffffff : offset, 16);
    records.push(eocd);
    return Buffer.concat(records);
};

const EOCD_RECORD_BYTES = Buffer.from([0x50, 0x4b, 0x05, 0x06]);

/**
 * A buffer whose whole archive is a central-directory signature at offset 0
 * and an end-of-central-directory record declaring one record there, padded
 * with zeros to `length`: `length` decides how much of the record's 46-byte
 * fixed header the buffer holds (#498).
 */
export const truncatedCentralStub = (length: number): Buffer => {
    const bytes = Buffer.alloc(length);
    bytes.writeUInt32LE(0x02014b50, 0);
    bytes.writeUInt32LE(0x06054b50, 4);
    bytes.writeUInt16LE(1, 12);
    bytes.writeUInt16LE(1, 14);
    bytes.writeUInt32LE(4, 16);
    bytes.writeUInt32LE(0, 20);
    return bytes;
};

/**
 * Rewrite the first `count` central records' local-header-offset field to the
 * ZIP64 spelling: a saturated 0xFFFFFFFF base field plus a 0x0001 extra field
 * holding `value` — the stored offset by default — in `fieldLength` bytes.
 * `sizes` also saturates the record's size fields, which JSZip resolves from
 * the same extra field first, in order (#490).
 */
export const saturateLocalHeaderOffsets = (
    zip: Uint8Array,
    options: { readonly count?: number; readonly sizes?: boolean; readonly fieldLength?: number; readonly value?: bigint } = {},
): Buffer => {
    const buf = Buffer.from(zip);
    const eocd = buf.lastIndexOf(EOCD_RECORD_BYTES);
    const count = buf.readUInt16LE(eocd + 10);
    const cdOffset = buf.readUInt32LE(eocd + 16);
    const cdEnd = eocd;
    const records: Buffer[] = [];
    const patches = Math.min(options.count ?? count, count);
    let offset = cdOffset;
    for (let i = 0; i < count; i += 1) {
        if (buf.readUInt32LE(offset) !== 0x02014b50) throw new Error("bad central header");
        const nameLength = buf.readUInt16LE(offset + 28);
        const extraLength = buf.readUInt16LE(offset + 30);
        const commentLength = buf.readUInt16LE(offset + 32);
        const end = offset + 46 + nameLength + extraLength + commentLength;
        const record = Buffer.from(buf.subarray(offset, end));
        if (i < patches) {
            // The ZIP64 extra field's values in APPNOTE order: the saturated
            // fields only — uncompressed size, compressed size, offset.
            const valueParts: Buffer[] = [];
            if (options.sizes === true) {
                const sizes = Buffer.alloc(16);
                sizes.writeBigUInt64LE(BigInt.asUintN(64, BigInt(record.readUInt32LE(24))), 0);
                sizes.writeBigUInt64LE(BigInt.asUintN(64, BigInt(record.readUInt32LE(20))), 8);
                valueParts.push(sizes);
                record.writeUInt32LE(0xffffffff, 20);
                record.writeUInt32LE(0xffffffff, 24);
            }
            const raw = Buffer.alloc(8);
            raw.writeBigUInt64LE(BigInt.asUintN(64, options.value ?? BigInt(record.readUInt32LE(42))), 0);
            valueParts.push(raw.subarray(0, options.fieldLength ?? 8));
            const value = Buffer.concat(valueParts);
            const field = Buffer.alloc(4 + value.length);
            field.writeUInt16LE(0x0001, 0);
            field.writeUInt16LE(value.length, 2);
            value.copy(field, 4);
            record.writeUInt32LE(0xffffffff, 42);
            record.writeUInt16LE(extraLength + field.length, 30);
            const insert = 46 + nameLength;
            const grown = Buffer.concat([record.subarray(0, insert), field, record.subarray(insert)]);
            records.push(grown);
        } else {
            records.push(record);
        }
        offset = end;
    }
    const central = Buffer.concat(records);
    const tail = Buffer.from(buf.subarray(cdEnd));
    tail.writeUInt32LE(central.length, 12);
    return Buffer.concat([buf.subarray(0, cdOffset), central, tail]);
};

/**
 * Raise every stored offset by `prefix` bytes, as a self-extractor writes a
 * ZIP behind its stub: the EOCD's central-directory offset and each central
 * record's local-header offset become absolute (#464). Not for ZIP64
 * archives — their EOCD record counts are saturated.
 */
export const absoluteOffsetsAfter = (zip: Uint8Array, prefix: number): Buffer => {
    const buf = Buffer.from(zip);
    const eocd = buf.lastIndexOf(EOCD_RECORD_BYTES);
    const count = buf.readUInt16LE(eocd + 10);
    let offset = buf.readUInt32LE(eocd + 16);
    buf.writeUInt32LE(offset + prefix, eocd + 16);
    for (let i = 0; i < count; i += 1) {
        if (buf.readUInt32LE(offset) !== 0x02014b50) throw new Error("bad central header");
        buf.writeUInt32LE(buf.readUInt32LE(offset + 42) + prefix, offset + 42);
        offset += 46 + buf.readUInt16LE(offset + 28) + buf.readUInt16LE(offset + 30) + buf.readUInt16LE(offset + 32);
    }
    return buf;
};
