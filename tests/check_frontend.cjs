const fs = require('node:fs');

const html = fs.readFileSync('static/index.html', 'utf8');
const match = html.match(/<script>([\s\S]*?)<\/script>/);
if (!match) throw new Error('Inline application script was not found');
new Function(match[1]);
console.log('Inline application JavaScript parses successfully');
