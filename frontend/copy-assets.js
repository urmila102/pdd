const fs = require('fs');
const path = require('path');

const files = ['index.html', 'style.css', 'app.js'];
const destDir = path.join(__dirname, 'www');

if (!fs.existsSync(destDir)) {
  fs.mkdirSync(destDir);
}

files.forEach(file => {
  fs.copyFileSync(path.join(__dirname, file), path.join(destDir, file));
  console.log(`Copied ${file} to www/${file}`);
});
