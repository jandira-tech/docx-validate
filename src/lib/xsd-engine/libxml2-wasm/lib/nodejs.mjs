/**
 * A side module for Node.js users only.
 *
 * ```ts
 * import { <symbol> } from 'libxml2-wasm/lib/nodejs.mjs';
 * ```
 *
 * @packageDocumentation
 */
import fs from 'node:fs';
import { xmlRegisterInputProvider } from './libxml2.mjs';
function filePath(filename) {
    try {
        const url = new URL(filename);
        if (url.protocol !== 'file:') {
            return null;
        }
        /* c8 ignore next 3, Windows only, won't be covered on other platforms */
        if (url.pathname.startsWith('/') && process.platform === 'win32') {
            return url.pathname.slice(1);
        }
        return url.pathname;
    }
    catch {
        return filename;
    }
}
function fileExists(filename) {
    const filepath = filePath(filename);
    return filepath != null && fs.existsSync(filepath);
}
/**
 * The virtual IO input providers for file operations in Node.js,
 * utilizing the `node:fs` module.
 *
 * These providers support both file paths (e.g., `path/to/file.xml`)
 * and file URLs (e.g., `file:///path/to/file.xml`).
 *
 * @see {@link libxml2-wasm!xmlRegisterInputProvider}
 */
export const fsInputProviders = {
    match(filename) {
        return fileExists(filename);
    },
    open(filename) {
        const filepath = filePath(filename);
        if (filepath == null) {
            return undefined;
        }
        try {
            return fs.openSync(filepath, 'r');
        }
        catch {
            return undefined;
        }
    },
    read(fd, buf) {
        try {
            return fs.readSync(fd, buf, 0, buf.byteLength, null);
        }
        catch {
            return -1;
        }
    },
    close(fd) {
        try {
            fs.closeSync(fd);
        }
        catch {
            // ignore exception in closing file
        }
        return true;
    },
};
/**
 * Register {@link fsInputProviders}.
 *
 * @see {@link libxml2-wasm!xmlRegisterInputProvider}
 */
export function xmlRegisterFsInputProviders() {
    return xmlRegisterInputProvider(fsInputProviders);
}
/**
 * Synchronously save the {@link XmlDocument} to a file.
 * @param doc The XmlDocument to be saved.
 * @param fd The file descriptor returned by `fs.open` or `fs.openSync`, etc.
 * @param options Options for saving.
 */
export function saveDocSync(doc, fd, options) {
    const handler = {
        fd,
        write(buf) {
            return fs.writeSync(this.fd, buf);
        },
        close() {
            fs.closeSync(this.fd);
            return true;
        },
    };
    doc.save(handler, options);
}