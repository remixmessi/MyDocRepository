import {mkdir, readFile, writeFile} from 'node:fs/promises';
import {spawnSync} from 'node:child_process';
import {dirname, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';

// Run on macOS after replacing the master image; generated files are committed
// so Windows builds do not need sips or iconutil.
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const master = resolve(root, 'assets/branding/mydoc-icon-v1.png');
const resources = resolve(root, 'packages/frontend/apps/desktop/resources');
const publicDir = resolve(root, 'packages/frontend/apps/web/public');
const iconset = resolve(root, 'tmp/app-icons/mydoc.iconset');
for (const path of [resources, publicDir, iconset]) await mkdir(path, {recursive:true});
function run(command, args) {
  const result = spawnSync(command, args, {stdio:'pipe'});
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(result.stderr.toString());
}
function resize(size, destination) {
  run('sips', ['-z', String(size), String(size), master, '--out', destination]);
}
for (const size of [16, 32, 128, 256, 512]) {
  resize(size, resolve(iconset, `icon_${size}x${size}.png`));
  resize(size * 2, resolve(iconset, `icon_${size}x${size}@2x.png`));
}
run('iconutil', ['-c', 'icns', iconset, '-o', resolve(resources, 'icon.icns')]);
resize(1024, resolve(resources, 'icon.png'));
resize(256, resolve(publicDir, 'mydoc-icon.png'));
resize(64, resolve(publicDir, 'favicon.png'));
// ICO supports PNG frames; include small sizes for Explorer and window chrome.
const sizes = [16, 32, 48, 64, 128, 256];
const frames = [];
for (const size of sizes) {
  const path = resolve(root, `tmp/app-icons/ico-${size}.png`);
  resize(size, path);
  frames.push(await readFile(path));
}
const header = Buffer.alloc(6 + sizes.length * 16);
header.writeUInt16LE(1, 2);
header.writeUInt16LE(sizes.length, 4);
let offset = header.length;
sizes.forEach((size, index) => {
  const entry = 6 + index * 16;
  header[entry] = header[entry + 1] = size === 256 ? 0 : size;
  header.writeUInt16LE(1, entry + 4);
  header.writeUInt16LE(32, entry + 6);
  header.writeUInt32LE(frames[index].length, entry + 8);
  header.writeUInt32LE(offset, entry + 12);
  offset += frames[index].length;
});
await writeFile(resolve(resources, 'icon.ico'), Buffer.concat([header, ...frames]));
console.log('Generated macOS, Windows, runtime and web icons.');
