const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');
const os = require('os');

const src = path.join(__dirname, 'release', 'Somus Capital-win32-x64');
const dest = path.join(os.homedir(), 'Desktop', 'SomusCapital_v2.0.zip');

if (fs.existsSync(dest)) fs.unlinkSync(dest);

const cmd = `powershell -Command "Compress-Archive -Path '${src}\\*' -DestinationPath '${dest}' -Force"`;
console.log('Creating ZIP...');
execSync(cmd, { encoding: 'utf8', timeout: 300000 });

const size = (fs.statSync(dest).size / 1024 / 1024).toFixed(1);
console.log(`Done: ${dest} (${size} MB)`);
