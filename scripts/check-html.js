const fs = require('fs');
const html = fs.readFileSync('index.html', 'utf8');
const matches = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)];
for (const [i, m] of matches.entries()) {
  new Function(m[1]);
  console.log(`script ${i + 1} ok`);
}
