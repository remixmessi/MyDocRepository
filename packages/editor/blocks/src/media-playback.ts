import type { AssetRecord } from '@mydoc/nbstore';
export type PlaybackKind = 'audio' | 'video';
export interface PlaybackType { kind: PlaybackKind; mime: string }
const extensions: Record<string, PlaybackType> = {
  mp3:{kind:'audio',mime:'audio/mpeg'},wav:{kind:'audio',mime:'audio/wav'},ogg:{kind:'audio',mime:'audio/ogg'},oga:{kind:'audio',mime:'audio/ogg'},opus:{kind:'audio',mime:'audio/ogg'},flac:{kind:'audio',mime:'audio/flac'},m4a:{kind:'audio',mime:'audio/mp4'},aac:{kind:'audio',mime:'audio/aac'},weba:{kind:'audio',mime:'audio/webm'},aif:{kind:'audio',mime:'audio/aiff'},aiff:{kind:'audio',mime:'audio/aiff'},
  mp4:{kind:'video',mime:'video/mp4'},m4v:{kind:'video',mime:'video/mp4'},webm:{kind:'video',mime:'video/webm'},ogv:{kind:'video',mime:'video/ogg'},mov:{kind:'video',mime:'video/quicktime'},avi:{kind:'video',mime:'video/x-msvideo'},mkv:{kind:'video',mime:'video/x-matroska'},
};
export function playbackType(name: string, mime: string, bytes?: Uint8Array): PlaybackType | null {
  const header = bytes ? String.fromCharCode(...bytes.subarray(0, 16)) : '';
  if (header.startsWith('RIFF') && header.slice(8,12) === 'WAVE') return {kind:'audio',mime:'audio/wav'};
  if (header.startsWith('fLaC')) return {kind:'audio',mime:'audio/flac'};
  if (header.startsWith('ID3')) return {kind:'audio',mime:'audio/mpeg'};
  const normalized = mime.trim().toLowerCase();
  if (/^(audio|video)\/[a-z0-9.+-]+(?:\s*;[^\r\n]*)?$/.test(normalized)) return {kind:normalized.startsWith('audio/')?'audio':'video',mime:normalized};
  return extensions[name.split('.').pop()?.toLowerCase() ?? ''] ?? null;
}
export const playbackRates = [.5,.75,1,1.25,1.5,2];
const PLAYING_EVENT = 'mydoc-media-playing';
export interface AttachmentPlayer { dom: HTMLElement; pause(): void; destroy(): void }
export function attachmentPlayer(record: AssetRecord, type: PlaybackType, retry: () => void): AttachmentPlayer {
  let disposed = false;
  const dom = document.createElement('span'); dom.className = 'attachment-player';
  const media = document.createElement(type.kind);
  media.controls = true; media.preload = 'metadata'; media.autoplay = false;
  media.setAttribute('aria-label', `${type.kind==='audio'?'音频':'视频'}播放器 ${record.name}`);
  if (media instanceof HTMLVideoElement) media.playsInline = true;
  const controls = document.createElement('span'); controls.className = 'attachment-player-tools';
  const label = document.createElement('label'); label.append('倍速 ');
  const rate = document.createElement('select'); rate.setAttribute('aria-label', `播放倍速 ${record.name}`); rate.disabled = true;
  for (const value of playbackRates) { const option = document.createElement('option'); option.value = String(value); option.textContent = `${value}×`; rate.append(option); }
  rate.value = '1'; label.append(rate); controls.append(label);
  const status = document.createElement('span'); status.className = 'attachment-player-status'; status.role = 'status'; status.textContent = '正在读取媒体信息…';
  const again = document.createElement('button'); again.type = 'button'; again.textContent = '重试播放'; again.setAttribute('aria-label', `重试播放 ${record.name}`); again.hidden = true; again.addEventListener('click', retry);
  const full = document.createElement('button'); full.type = 'button'; full.textContent = '全屏'; full.setAttribute('aria-label', `视频全屏 ${record.name}`); full.disabled = true;
  if (type.kind === 'video') controls.append(full);
  controls.append(status, again); dom.append(media, controls);
  const url = URL.createObjectURL(new Blob([record.bytes], {type:type.mime}));
  const pause = () => media.pause();
  const playing = () => {
    if (disposed || !dom.isConnected || dom.closest('[hidden]') || document.hidden) { pause(); return; }
    document.dispatchEvent(new CustomEvent(PLAYING_EVENT, {detail:media}));
  };
  const other = (event: Event) => { if ((event as CustomEvent).detail !== media) pause(); };
  const hidden = () => { if (document.hidden) pause(); };
  media.addEventListener('play', playing);
  media.addEventListener('loadedmetadata', () => {
    if (disposed) return;
    status.textContent = ''; rate.disabled = false;
    full.disabled = !document.fullscreenEnabled || !media.requestFullscreen;
  });
  media.addEventListener('error', () => {
    if (disposed) return;
    pause(); rate.disabled = true; full.disabled = true; again.hidden = false;
    status.role = 'alert'; status.textContent = '浏览器无法播放此文件，可能格式不支持或文件损坏；可下载原文件或重试。';
    dom.classList.add('playback-error');
  });
  rate.addEventListener('change', () => { const value = Number(rate.value); if (playbackRates.includes(value)) media.playbackRate = value; });
  full.addEventListener('click', () => { void media.requestFullscreen().catch(() => { status.textContent = '无法进入全屏，请使用视频原生全屏控件。'; }); });
  document.addEventListener(PLAYING_EVENT, other); document.addEventListener('visibilitychange', hidden); window.addEventListener('pagehide', pause);
  media.src = url;
  return {dom, pause, destroy() {
    disposed = true; pause(); media.removeAttribute('src'); media.load();
    document.removeEventListener(PLAYING_EVENT, other); document.removeEventListener('visibilitychange', hidden); window.removeEventListener('pagehide', pause);
    URL.revokeObjectURL(url);
  }};
}
