/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * ZIP64 envelopes whose 8-byte fields JSZip reads differently from their
 * value. JSZip's `DataReader.readInt` accumulates with a 32-bit `<<`, so an
 * 8-byte field reads as its signed low 32 bits (#470).
 */

import { buildRawZip, type RawZipEntry } from "./raw-zip";

const ZIP64_RECORD_LENGTH = 56;

const centralDirectoryOf = (zip: Buffer): Buffer => zip.subarray(zip.readUInt32LE(zip.length - 22 + 16), zip.length - 22);
const localsOf = (zip: Buffer): Buffer => zip.subarray(0, zip.readUInt32LE(zip.length - 22 + 16));

const zip64Record = (count: number, size: number, offset: bigint): Buffer => {
    const record = Buffer.alloc(ZIP64_RECORD_LENGTH);
    record.writeUInt32LE(0x06064b50, 0);
    record.writeBigUInt64LE(44n, 4);
    record.writeUInt16LE(45, 12);
    record.writeUInt16LE(45, 14);
    record.writeBigUInt64LE(BigInt(count), 24);
    record.writeBigUInt64LE(BigInt(count), 32);
    record.writeBigUInt64LE(BigInt(size), 40);
    record.writeBigUInt64LE(offset, 48);
    return record;
};

const locatorAt = (offset: bigint): Buffer => {
    const locator = Buffer.alloc(20);
    locator.writeUInt32LE(0x07064b50, 0);
    locator.writeBigUInt64LE(offset, 8);
    locator.writeUInt32LE(1, 16);
    return locator;
};

const zip64Eocd = (): Buffer => {
    const eocd = Buffer.alloc(22);
    eocd.writeUInt32LE(0x06054b50, 0);
    eocd.writeUInt16LE(0xffff, 8);
    eocd.writeUInt16LE(0xffff, 10);
    eocd.writeUInt32LE(0xffffffff, 12);
    eocd.writeUInt32LE(0xffffffff, 16);
    return eocd;
};

/**
 * Review AE's specimen. A ZIP64 record at offset 0 describes archive A's
 * entries plus `z-clean.xml`; the last ZIP64 record describes `z-clean.xml`
 * alone; the locator stores 0x1_0000_0000. JSZip wraps that to 0, finds a
 * record signature there and loads A plus `z-clean.xml`. An exact read is out
 * of range and would fall back to the last record.
 */
export const wrappedLocatorArchive = (entries: readonly RawZipEntry[]): Buffer => {
    const a = buildRawZip(entries);
    const locals = localsOf(a);
    const central = centralDirectoryOf(a);
    const z = buildRawZip([{ name: "z-clean.xml", content: "z" }]);
    const zLocal = z.subarray(0, 30 + "z-clean.xml".length + 1);
    const zCentral = Buffer.from(centralDirectoryOf(z));
    // Stored relative to the archive, which starts after the first record.
    zCentral.writeUInt32LE(locals.length, 42);
    const count = entries.length + 1;
    const centralStart = ZIP64_RECORD_LENGTH + locals.length + zLocal.length;
    const first = zip64Record(count, central.length + zCentral.length, BigInt(centralStart - ZIP64_RECORD_LENGTH));
    const last = zip64Record(1, zCentral.length, BigInt(centralStart + central.length));
    return Buffer.concat([first, locals, zLocal, central, zCentral, last, locatorAt(0x1_0000_0000n), zip64Eocd()]);
};

/** A ZIP64 archive whose record's directory offset carries 2^32 on top: JSZip wraps it away. */
export const wrappedOffsetArchive = (entries: readonly RawZipEntry[]): Buffer => {
    const zip = Buffer.from(buildRawZip(entries, { zip64: true }));
    const record = zip.lastIndexOf(Buffer.from([0x50, 0x4b, 0x06, 0x06]));
    zip.writeBigUInt64LE(zip.readBigUInt64LE(record + 48) + 0x1_0000_0000n, record + 48);
    return zip;
};

/** A ZIP64 archive whose locator points past the buffer once JSZip wraps it: JSZip throws. */
export const outOfRangeLocatorArchive = (entries: readonly RawZipEntry[]): Buffer => {
    const zip = Buffer.from(buildRawZip(entries, { zip64: true }));
    const locator = zip.lastIndexOf(Buffer.from([0x50, 0x4b, 0x06, 0x07]));
    zip.writeBigUInt64LE(0x7fff_fff0n, locator + 8);
    return zip;
};

/**
 * A ZIP64 archive whose record declares 8 bytes of extensible data, present
 * and accounted for in every offset. JSZip's extensible-data loop never
 * advances its index, so it reads until the buffer ends and throws.
 */
export const extensibleDataArchive = (entries: readonly RawZipEntry[]): Buffer => {
    const zip = Buffer.from(buildRawZip(entries, { zip64: true }));
    const record = zip.lastIndexOf(Buffer.from([0x50, 0x4b, 0x06, 0x06]));
    zip.writeBigUInt64LE(52n, record + 4);
    return Buffer.concat([zip.subarray(0, record + ZIP64_RECORD_LENGTH), Buffer.alloc(8), zip.subarray(record + ZIP64_RECORD_LENGTH)]);
};

/** A ZIP64 archive whose locator counts two disks: JSZip refuses multi-volume archives. */
export const multiVolumeArchive = (entries: readonly RawZipEntry[]): Buffer => {
    const zip = Buffer.from(buildRawZip(entries, { zip64: true }));
    const locator = zip.lastIndexOf(Buffer.from([0x50, 0x4b, 0x06, 0x07]));
    zip.writeUInt32LE(2, locator + 16);
    return zip;
};
