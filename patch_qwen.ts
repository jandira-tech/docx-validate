import fs from 'fs';
let content = fs.readFileSync('.github/workflows/qwen-review.yml', 'utf-8');
content = content.replace(/DASHSCOPE_API_KEY: \$\{\{ secrets\.DASHSCOPE_API_KEY \}\}/g, 'DASHSCOPE_API_KEY: ${{ secrets.DASHSCOPE_API_KEY || secrets.ALIBABA_CODING_PLAN }}');
fs.writeFileSync('.github/workflows/qwen-review.yml', content);
