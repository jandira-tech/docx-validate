import { readFileSync } from "node:fs";
import { countZipBytesAfterEocd } from "../../src/scripts/office/zip-bytes-after-eocd";
import { zipBytesBeforeArchive } from "../../src/lib/zip-central-directory";
const buf = new Uint8Array(readFileSync("tmp/8ab8b602-verify/selfcat.docx"));
console.log("afterEocd:", countZipBytesAfterEocd(buf));
console.log("beforeArchive:", zipBytesBeforeArchive(buf));
