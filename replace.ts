import * as fs from 'fs';

const filePath = 'src/scripts/office/validators/docx.ts';
let code = fs.readFileSync(filePath, 'utf-8');

const search = `    async validateDeletions(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const $$ = makeSelect();
        for (const xmlFile of this.documentXmlFiles()) {`;

const replace = `    async validateDeletions(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        for (const xmlFile of this.documentXmlFiles()) {`;

code = code.replace(search, replace);
fs.writeFileSync(filePath, code);
