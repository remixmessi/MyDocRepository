import {spawn, spawnSync} from 'node:child_process';
import {createWriteStream} from 'node:fs';
import {mkdir, rename, rm} from 'node:fs/promises';
import {dirname, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import {pipeline} from 'node:stream/promises';
import {createGzip} from 'node:zlib';

const root = fileURLToPath(new URL('../../', import.meta.url));
const options = {platform: 'linux/amd64', tag: 'mydoc-web:latest', output: undefined};
const args = process.argv.slice(2);
for (let i = 0; i < args.length; i++) {
  if (args[i] === '--help') {
    console.log('Usage: node scripts/docker/export-image.mjs [--platform linux/amd64|linux/arm64] [--tag mydoc-web:latest] [--output path.tar.gz]');
    process.exit(0);
  }
  const key = args[i].replace(/^--/, '');
  if (!args[i].startsWith('--') || !Object.hasOwn(options, key) || !args[i + 1] || args[i + 1].startsWith('--')) {
    throw new Error(`Invalid argument: ${args[i]}. Use --help for usage.`);
  }
  options[key] = args[++i];
}
if (!['linux/amd64', 'linux/arm64'].includes(options.platform)) {
  throw new Error('Build one platform per archive: linux/amd64 or linux/arm64.');
}
const output = resolve(root, options.output ?? `release/docker/mydoc-web-${options.platform.replace('/', '-')}.tar.gz`);
if (!output.endsWith('.tar.gz')) throw new Error('Output must end in .tar.gz.');

function run(commandArgs, capture = false) {
  const result = spawnSync('docker', commandArgs, {
    cwd: root,
    stdio: capture ? ['ignore', 'pipe', 'inherit'] : 'inherit',
    encoding: 'utf8',
  });
  if (result.error) {
    throw new Error(`Cannot run Docker. Install Docker and start its engine first. ${result.error.message}`);
  }
  if (result.status !== 0) throw new Error(`docker ${commandArgs[0]} failed (${result.status}).`);
  return result.stdout?.trim();
}

run(['info'], true);
run(['buildx', 'build', '--platform', options.platform, '--load', '--tag', options.tag, '.']);
const platform = run(['image', 'inspect', '--format', '{{.Os}}/{{.Architecture}}', options.tag], true);
if (platform !== options.platform) throw new Error(`Image platform mismatch: ${platform}.`);
await mkdir(dirname(output), {recursive: true});
const temporary = `${output}.${process.pid}.tmp`;
const save = spawn('docker', ['image', 'save', options.tag], {cwd: root, stdio: ['ignore', 'pipe', 'inherit']});
const saved = new Promise((accept, reject) => {
  save.once('error', reject);
  save.once('close', code => code === 0 ? accept() : reject(new Error(`docker image save failed (${code}).`)));
});
try {
  await Promise.all([saved, pipeline(save.stdout, createGzip(), createWriteStream(temporary, {flags: 'wx'}))]);
  await rename(temporary, output);
} catch (error) {
  save.kill();
  await rm(temporary, {force: true});
  throw error;
}
console.log(`\n1Panel image archive: ${output}\nImage: ${options.tag}\nPlatform: ${platform}`);
