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

export interface RawZipEntry {
    /** The item name, stored verbatim (UTF-8 encoded, or `nameBytes` raw). */
    readonly name: string;
    /** Overrides the UTF-8 encoding of `name`, e.g. for CP437 specimens. */
    readonly nameBytes?: Uint8Array;
    readonly content?: string | Uint8Array;
    /** Set general-purpose bit 11, marking `name` as UTF-8. */
    readonly utf8?: boolean;
    /** Set the MS-DOS directory attribute bit, as a folder entry does. */
    readonly directory?: boolean;
}

const encoder = new TextEncoder();

/**
 * The archive bytes: local headers, then the central directory, then the
 * EOCD record. `zip64` emits a ZIP64 EOCD record and locator — the EOCD's
 * central-directory fields saturated — around the same directory, so readers
 * exercise the ZIP64 path without a multi-gigabyte file.
 */
export const buildRawZip = (entries: readonly RawZipEntry[], options: { readonly zip64?: boolean } = {}): Buffer => {
    const local: Buffer[] = [];
    const central: Buffer[] = [];
    let offset = 0;
    for (const entry of entries) {
        const nameBytes = entry.nameBytes ?? encoder.encode(entry.name);
        const data = typeof entry.content === "string" ? encoder.encode(entry.content) : (entry.content ?? new Uint8Array());
        const crc = crc32(data);
        const flags = entry.utf8 === true ? 0x0800 : 0;
        const attributes = entry.directory === true ? 0x10 : 0;

        const lh = Buffer.alloc(30);
        lh.writeUInt32LE(0x04034b50, 0);
        lh.writeUInt16LE(20, 4);
        lh.writeUInt16LE(flags, 6);
        lh.writeUInt32LE(crc, 14);
        lh.writeUInt32LE(data.length, 18);
        lh.writeUInt32LE(data.length, 22);
        lh.writeUInt16LE(nameBytes.length, 26);
        local.push(lh, Buffer.from(nameBytes), Buffer.from(data));

        const ch = Buffer.alloc(46);
        ch.writeUInt32LE(0x02014b50, 0);
        ch.writeUInt16LE(20, 4);
        ch.writeUInt16LE(20, 6);
        ch.writeUInt16LE(flags, 8);
        ch.writeUInt32LE(crc, 16);
        ch.writeUInt32LE(data.length, 20);
        ch.writeUInt32LE(data.length, 24);
        ch.writeUInt16LE(nameBytes.length, 28);
        ch.writeUInt32LE(attributes, 38);
        ch.writeUInt32LE(offset, 42);
        central.push(ch, Buffer.from(nameBytes));

        offset += 30 + nameBytes.length + data.length;
    }
    const cd = Buffer.concat(central);
    const zip64 = options.zip64 === true;
    const records: Buffer[] = [...local, cd];
    if (zip64) {
        const eocd64 = Buffer.alloc(56);
        eocd64.writeUInt32LE(0x06064b50, 0);
        eocd64.writeBigUInt64LE(44n, 4);
        eocd64.writeUInt16LE(20, 12);
        eocd64.writeUInt16LE(45, 14);
        eocd64.writeBigUInt64LE(BigInt(entries.length), 24);
        eocd64.writeBigUInt64LE(BigInt(entries.length), 32);
        eocd64.writeBigUInt64LE(BigInt(cd.length), 40);
        eocd64.writeBigUInt64LE(BigInt(offset), 48);
        const locator = Buffer.alloc(20);
        locator.writeUInt32LE(0x07064b50, 0);
        locator.writeBigUInt64LE(BigInt(offset + cd.length), 8);
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
