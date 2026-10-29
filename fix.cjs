const fs = require('fs');

let content = fs.readFileSync('tests/docx-direct-directory-stored-names.integration.test.ts', 'utf8');
content = content.replace(/import \{ normalizePartName \} from "\.\.\/src\/lib\/validation-context";/, 'import { normalizePartName } from "../src/lib/types";');
fs.writeFileSync('tests/docx-direct-directory-stored-names.integration.test.ts', content);
