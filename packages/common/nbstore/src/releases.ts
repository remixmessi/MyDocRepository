export type ReleasePlatform = 'darwin' | 'win32' | 'linux';
export type ReleaseArchitecture = 'x64' | 'arm64';
export interface DesktopRelease {
  id: string;
  version: string;
  platform: ReleasePlatform;
  /** Empty for packages uploaded before architecture was recorded. */
  arch: ReleaseArchitecture | '';
  filename: string;
  notes: string;
  bytes: number;
  sha256: string;
  created: number;
  downloadPath: string;
}
export const releasePlatforms: Record<ReleasePlatform, string> = {darwin: 'macOS', win32: 'Windows', linux: 'Linux'};
export function releaseTarget(platform: string, arch: string) {
  const name = releasePlatforms[platform as ReleasePlatform] ?? platform;
  if (!arch) return `${name} · 未标注架构`;
  if (platform === 'darwin' && arch === 'arm64') return 'macOS · Apple 芯片 (arm64)';
  if (platform === 'darwin' && arch === 'x64') return 'macOS · Intel (x64)';
  return `${name} · ${arch}`;
}
export const releaseVersionPattern = '(0|[1-9][0-9]*)\\.(0|[1-9][0-9]*)\\.(0|[1-9][0-9]*)(-[0-9A-Za-z\\-]+(\\.[0-9A-Za-z\\-]+)*)?(\\+[0-9A-Za-z\\-]+(\\.[0-9A-Za-z\\-]+)*)?';
export function validReleaseVersion(value: string) {
  if (value.length > 100 || !new RegExp('^' + releaseVersionPattern + '$').test(value)) return false;
  const pre = value.split('+')[0].split('-').slice(1).join('-');
  return !pre.split('.').some(part => /^0[0-9]+$/.test(part));
}
/** SemVer precedence, ignoring build metadata. Positive means a is newer. */
export function compareReleaseVersions(a: string, b: string) {
  const parse = (value: string) => {
    const [core, ...pre] = value.split('+')[0].split('-');
    return {core: core.split('.').map(BigInt), pre: pre.join('-')};
  };
  const av = parse(a), bv = parse(b);
  for (let i = 0; i < 3; i++) if (av.core[i] !== bv.core[i]) return av.core[i] > bv.core[i] ? 1 : -1;
  if (!av.pre || !bv.pre) return av.pre === bv.pre ? 0 : av.pre ? -1 : 1;
  const ap = av.pre.split('.'), bp = bv.pre.split('.');
  for (let i = 0; i < Math.max(ap.length, bp.length); i++) {
    if (ap[i] === bp[i]) continue;
    if (ap[i] === undefined) return -1;
    if (bp[i] === undefined) return 1;
    const an = /^\d+$/.test(ap[i]), bn = /^\d+$/.test(bp[i]);
    if (an && bn) return BigInt(ap[i]) > BigInt(bp[i]) ? 1 : -1;
    if (an !== bn) return an ? -1 : 1;
    return ap[i] > bp[i] ? 1 : -1;
  }
  return 0;
}
