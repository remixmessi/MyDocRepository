import { useEffect, useRef, useState } from 'react';
import { imageMime, type AssetRecord, type WorkspaceAssets } from '@mydoc/nbstore';

export function ImagePreview({
  assetId,
  src,
  alt,
  name,
  assets,
  onClose,
}: {
  assetId?: string;
  src?: string;
  alt?: string;
  name?: string;
  assets?: WorkspaceAssets;
  onClose: () => void;
}) {
  const dialog = useRef<HTMLDialogElement>(null);
  const [asset, setAsset] = useState<AssetRecord>();
  const [url, setUrl] = useState(src);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(Boolean(assetId));
  const [retry, setRetry] = useState(0);

  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null;
    dialog.current!.showModal();
    return () => {
      dialog.current?.close();
      previous?.isConnected && previous.focus();
    };
  }, []);

  useEffect(() => {
    let active = true;
    let objectUrl: string | undefined;
    setAsset(undefined);
    setError('');
    setLoading(Boolean(assetId));
    setUrl(src);
    if (!assetId) return () => { active = false; };
    if (!assets) {
      setLoading(false);
      setError('图片资源不可用，请重试');
      return () => { active = false; };
    }
    void assets.get(assetId).then(record => {
      if (!active) return;
      if (!record) throw new Error('本地文件缺失，请恢复包含文件的工作区备份');
      const mime = imageMime(record.bytes);
      if (!mime) throw new Error('图片数据无法读取');
      objectUrl = URL.createObjectURL(new Blob([record.bytes], { type: mime }));
      setAsset(record);
      setUrl(objectUrl);
      setLoading(false);
    }).catch(value => {
      if (!active) return;
      setLoading(false);
      setError(value instanceof Error ? value.message : '图片读取失败，请重试');
    });
    return () => {
      active = false;
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    };
  }, [assetId, assets, retry, src]);

  const displayName = asset?.name || name || alt || '图片';
  const download = () => {
    if (!asset || !url) return;
    const link = document.createElement('a');
    link.href = url;
    link.download = asset.name;
    link.click();
  };
  const open = () => {
    if (!url) return;
    if (!asset) { window.open(url, '_blank', 'noopener,noreferrer'); return; }
    const openedUrl = URL.createObjectURL(new Blob([asset.bytes], { type: imageMime(asset.bytes) || asset.mime }));
    window.open(openedUrl, '_blank', 'noopener,noreferrer');
    setTimeout(() => URL.revokeObjectURL(openedUrl), 60000);
  };

  return <dialog
    ref={dialog}
    className="app-dialog image-preview-dialog"
    aria-label="图片预览"
    onCancel={event => { event.preventDefault(); onClose(); }}
    onKeyDown={event => {
      if (event.nativeEvent.isComposing) return;
      if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); onClose(); }
    }}
  >
    <div className="dialog-heading"><h2>图片预览</h2><button type="button" aria-label="关闭图片预览" onClick={onClose}>×</button></div>
    <p className="image-preview-title">{displayName}</p>
    {loading && <p className="image-preview-status" role="status">正在读取图片…</p>}
    {error && <p className="image-preview-error" role="alert">{error} <button type="button" onClick={() => setRetry(value => value + 1)}>重试预览</button></p>}
    {url && !error && <div className="image-preview-stage"><img src={url} alt={alt || displayName} /></div>}
    <div className="dialog-actions"><button type="button" disabled={!asset || !url} onClick={download}>下载原文件</button><button type="button" disabled={!url} onClick={open}>新窗口打开</button></div>
  </dialog>;
}
