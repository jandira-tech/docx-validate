import { readFileSync } from "node:fs";
import { countZipBytesAfterEocd } from "../../src/scripts/office/zip-bytes-after-eocd";
import { zipBytesBeforeArchive } from "../../src/lib/zip-central-directory";
const buf = new Uint8Array(readFileSync(process.argv[2]!));
console.log("countZipBytesAfterEocd:", countZipBytesAfterEocd(buf));
console.log("zipBytesBeforeArchive:", zipBytesBeforeArchive(buf));
