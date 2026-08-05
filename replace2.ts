import * as fs from 'fs';

const filePath = 'src/scripts/office/validators/docx.ts';
let code = fs.readFileSync(filePath, 'utf-8');

const search = `import { makeSelect, parseXml, serializeXml } from "../../../lib/xml-helpers";`;

const replace = `import { parseXml, serializeXml } from "../../../lib/xml-helpers";`;

code = code.replace(search, replace);
fs.writeFileSync(filePath, code);
