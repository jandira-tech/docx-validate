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
 * Reading and writing XML parts in the encoding they are stored in.
 *
 * A part's bytes decide its encoding — the byte order mark first, then the
 * XML declaration — and an editing step must put the part back the way it
 * found it: a UTF-16 part stays UTF-16 with its BOM, a UTF-8 part stays
 * UTF-8, written without a BOM as before. Reading a part as a UTF-8 string turns a UTF-16 part into
 * NUL-interleaved garbage, and writing that garbage back corrupts the part
 * (#866).
 *
 * The decoding rules live in `xml-helpers.ts`; this module adds the file
 * system around them.
 */

import { promises as fs } from "node:fs";

import {
    decodeXmlPart,
    encodeXmlPart,
    serializeXml,
    xmlDeclarationEncoding,
    type DecodedXmlPart,
    type XmlPartEncoding,
} from "./xml-helpers";

export type { DecodedXmlPart, XmlPartEncoding };

/**
 * Read an XML part and decode it.
 *
 * Returns `undefined` when the part declares an encoding this project cannot
 * decode: the caller must then leave the part's bytes untouched rather than
 * write back text it decoded lossily.
 */
export const readXmlPart = async (file: string): Promise<DecodedXmlPart | undefined> => decodeXmlPart(await fs.readFile(file));

/**
 * Write a DOM back to the part it came from, in the part's own encoding.
 *
 * A UTF-16 part is written back as UTF-16, with its BOM if it had one, and a
 * UTF-8 part as UTF-8. The declaration's `encoding` pseudo-attribute names
 * the encoding the bytes are written in when the writer adds the declaration
 * or rewrites one that has it; a declaration without one keeps none (#888,
 * #894). A part whose encoding cannot be decoded is left untouched — writing
 * text decoded from it would corrupt the part — and `false` is returned, so
 * a caller does not count or report a repair that never reached disk (#880).
 * A written part returns `true`.
 *
 * `part` is the encoding the caller decoded the part in. A caller that holds
 * only a DOM may leave it out: the part's current bytes then decide the
 * encoding, so the part is still written back the way it is stored.
 */
export const writeXmlPart = async (file: string, dom: Document, part?: XmlPartEncoding): Promise<boolean> => {
    const encoding = part ?? decodeXmlPart(await fs.readFile(file));
    if (encoding === undefined) {
        return false;
    }
    await fs.writeFile(file, encodeXmlPart(serializeXml(dom, xmlDeclarationEncoding(encoding)), encoding));
    return true;
};

/**
 * Write text back to the part it came from, in the part's own encoding.
 *
 * The text is written verbatim: the caller owns the XML it produces, and
 * `part` is the encoding it decoded the part in.
 */
export const writeXmlText = async (file: string, text: string, part: XmlPartEncoding): Promise<void> => {
    await fs.writeFile(file, encodeXmlPart(text, part));
};
