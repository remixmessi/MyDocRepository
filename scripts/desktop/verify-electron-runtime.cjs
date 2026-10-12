const {join} = require('node:path');

module.exports = async context => {
  if (context.electronPlatformName !== 'darwin') return;

  const expected = context.arch === 1 ? 'x64' : context.arch === 3 ? 'arm64' : null;
  if (!expected) return;

  // beforePack exposes the configured source on packager.config. When no
  // custom electronDist is set, electron-builder resolves the matching
  // cached runtime itself and the afterPack bundle check remains authoritative.
  const configured = context.packager.config.electronDist;
  if (typeof configured !== 'string' || !configured.trim()) return;
  const {stat} = await import('node:fs/promises');
  const resolved = require('node:path').resolve(context.packager.projectDir, configured);
  if (!(await stat(resolved)).isDirectory()) return;
  const {machoArchitectures, assertArchitecture} = await import('./mac-architecture.mjs');
  const electronPath = join(resolved, 'Electron.app', 'Contents', 'MacOS', 'Electron');
  const {readFile} = await import('node:fs/promises');
  let actual;
  try {
    actual = machoArchitectures(await readFile(electronPath));
  } catch (error) {
    throw new Error(`无法读取 Electron 运行时：${electronPath}（${error.message}）`);
  }
  assertArchitecture(actual, [expected], electronPath);
  console.log(`已验证 Electron 运行时：目标 ${expected}，实际 ${actual.join('+')}。`);
};
