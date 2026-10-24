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
 * Make one entry of a DEFLATE-compressed archive unreadable: overwrite the
 * first byte of its compressed stream with a reserved block type, so the
 * archive still opens but reading that entry fails.
 */
export const poisonEntry = (archive: Uint8Array, name: string): Uint8Array => {
    const bytes = Uint8Array.from(archive);
    const view = new DataView(bytes.buffer);
    const encoded = new TextEncoder().encode(name);
    for (let at = 0; at + 30 <= bytes.length; at++) {
        if (view.getUint32(at, true) !== 0x04034b50) {
            continue;
        }
        const nameLength = view.getUint16(at + 26, true);
        const extraLength = view.getUint16(at + 28, true);
        if (nameLength === encoded.length && encoded.every((byte, i) => bytes[at + 30 + i] === byte)) {
            bytes[at + 30 + nameLength + extraLength] = 0x07;
            return bytes;
        }
    }
    throw new Error(`no local header for ${name}`);
};
