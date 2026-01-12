const fs = require('fs');
const path = require('path');

const templatesDir = path.join(__dirname, '..', 'templates');
const outputPath = path.join(__dirname, '..', 'js', 'template-store.js');

console.log(`Scanning templates in: ${templatesDir}`);

if (!fs.existsSync(templatesDir)) {
    console.error('Error: templates directory not found.');
    process.exit(1);
}

const files = fs.readdirSync(templatesDir).filter(f => f.endsWith('.xlsx'));
if (files.length === 0) {
    console.error('No .xlsx files found in templates directory.');
    process.exit(1);
}

const templatesStore = {};

files.forEach(file => {
    // Expected format: template_8.xlsx -> key: 8
    const match = file.match(/template_(\d+)\.xlsx/);
    if (match) {
        const cavities = match[1];
        const filePath = path.join(templatesDir, file);
        const buffer = fs.readFileSync(filePath);
        const base64 = buffer.toString('base64');
        templatesStore[cavities] = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64}`;
        console.log(`Packed template for ${cavities} cavities (${(base64.length / 1024).toFixed(2)} KB)`);
    }
});

const jsContent = `/**
 * Built-in Excel Template Storage (Multi-Template)
 * Generated at: ${new Date().toISOString()}
 */

var SPC_TEMPLATES = ${JSON.stringify(templatesStore, null, 2)};
`;

fs.writeFileSync(outputPath, jsContent);
console.log(`Success! All templates baked into ${outputPath}`);
