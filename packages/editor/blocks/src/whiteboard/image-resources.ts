import {imageMime,type WorkspaceAssets,type AssetRecord} from '@mydoc/nbstore';
export interface ImageResource {url:string;image:HTMLImageElement;record:AssetRecord}
interface Entry {refs:number;url:string;promise:Promise<ImageResource>}
const pools=new WeakMap<WorkspaceAssets,Map<string,Entry>>();
export function acquireImage(assets:WorkspaceAssets,id:string){
  let pool=pools.get(assets);if(!pool){pool=new Map();pools.set(assets,pool);}let entry=pool.get(id);
  if(!entry){entry={refs:0,url:'',promise:null!};const current=entry;entry.promise=(async()=>{const record=await assets.get(id);if(!record)throw new Error('本地图片缺失，请恢复包含文件的工作区备份');const mime=imageMime(record.bytes);if(!mime)throw new Error('图片数据无法读取');if(!current.refs)throw new Error('图片读取已取消');current.url=URL.createObjectURL(new Blob([record.bytes],{type:mime}));const image=new Image();await new Promise<void>((resolve,reject)=>{image.onload=()=>resolve();image.onerror=()=>reject(new Error('图片无法显示，可下载原文件检查'));image.src=current.url;});return{url:current.url,image,record};})();pool.set(id,entry);}
  entry.refs++;let released=false;const current=entry;return {promise:entry.promise,release(){if(released)return;released=true;if(--current.refs===0)queueMicrotask(()=>{if(current.refs)return;if(current.url)URL.revokeObjectURL(current.url);if(pool!.get(id)===current)pool!.delete(id);});},};
}
export function retryImage(assets:WorkspaceAssets,id:string){pools.get(assets)?.delete(id);}
