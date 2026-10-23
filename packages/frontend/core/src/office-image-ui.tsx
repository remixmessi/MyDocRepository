import {useEffect,useState,useRef} from 'react';
import {Lock,LockOpen,Trash2} from 'lucide-react';
import {Y} from '@mydoc/crdt';
import {officeImageMime,createWordImage,OFFICE_IMAGE_COUNT,OFFICE_MEDIA_BYTES,officeImageNodes} from '../../../common/nbstore/src/office-image';
import type {LocalWorkspace} from './workspace';
import {createPresentationObject,presentationSize} from '../../../common/nbstore/src/office-presentation';
import {validWordRevisionMark,wordRevisionColor,wordTrackingAuthor,markWordImageInserted,type WordRevisionMark} from '../../../common/nbstore/src/office-word-revision';
/** A blob URL for a document picture, revoked on change or unmount; `retry` loads it again after an error. */
export function useOfficeImageUrl(workspace:LocalWorkspace,id:string,assetId:string){
 const [state,setState]=useState<{url?:string;error?:string}>({}),[attempt,setAttempt]=useState(0);
 useEffect(()=>{let disposed=false,url:string|undefined;setState({});void workspace.documentAssets(id).get(assetId).then(asset=>{if(!asset)throw new Error('图片文件缺失或无权访问');officeImageMime(asset.bytes);if(disposed)return;url=URL.createObjectURL(new Blob([new Uint8Array(asset.bytes)],{type:asset.mime}));setState({url});}).catch(error=>{if(!disposed)setState({error:error instanceof Error?error.message:'图片加载失败'});});return()=>{disposed=true;if(url)URL.revokeObjectURL(url);};},[workspace,id,assetId,attempt]);
 return {...state,setState,retry:()=>setAttempt(value=>value+1)};
}
export function OfficeImage({workspace,id,assetId,alt='图片'}:{workspace:LocalWorkspace;id:string;assetId:string;alt?:string}){
 const {retry,setState,...state}=useOfficeImageUrl(workspace,id,assetId);
 if(state.error)return <span role="alert">{state.error}<button onClick={retry}>重试图片</button></span>;
 return state.url?<img src={state.url} alt={alt} className="office-image-content" onError={()=>setState({error:'图片无法解码，请重新插入有效图片'})}/>:<span role="status">正在加载图片…</span>;
}
/** A header/footer picture inside the page-margin button: no nested controls, a short text if it cannot load. */
export function WordHeaderPicture({workspace,id,picture}:{workspace:LocalWorkspace;id:string;picture:{assetId:string;width:number;height:number;alt:string}}){
 const {url,error}=useOfficeImageUrl(workspace,id,picture.assetId);
 return <span className="word-margin-picture" style={{width:picture.width,aspectRatio:`${picture.width} / ${picture.height}`}}>{url?<img src={url} alt={picture.alt}/>:<span className="word-margin-placeholder">{error?'图片无法显示':'…'}</span>}</span>;
}
/**
 * Checks a picture file against the document's image limits (count, 8 MB in total) and saves it as a document
 * asset; the size is the natural size scaled down to fit `fit` pixels. The caller inserts the reference.
 */
export async function saveOfficeImage(workspace:LocalWorkspace,id:string,doc:Y.Doc,kind:'word'|'presentation',file:File,fit:number|{width:number;height:number}){
 workspace.assertWritable(id);if(file.name.length>1000||/[\x00-\x1f]/.test(file.name))throw new Error('图片文件名称无效');const bytes=new Uint8Array(await file.arrayBuffer());officeImageMime(bytes);if(officeImageNodes(doc,kind).length>=OFFICE_IMAGE_COUNT)throw new Error('办公图片超过 1000 个');let total=bytes.length;for(const assetId of new Set(officeImageNodes(doc,kind).map(node=>String(node.get('assetId'))))){const existing=await workspace.documentAssets(id).get(assetId);if(!existing)throw new Error('已有图片缺失，请先处理');total+=existing.size;if(total>OFFICE_MEDIA_BYTES)throw new Error('办公图片总计不能超过 8 MB');}
 const box=typeof fit==='number'?{width:fit,height:fit}:fit,blobUrl=URL.createObjectURL(new Blob([bytes]));let width=320,height=240;try{const image=new Image();image.src=blobUrl;await image.decode();const scale=Math.min(1,box.width/image.naturalWidth,box.height/image.naturalHeight);width=image.naturalWidth*scale;height=image.naturalHeight*scale;}finally{URL.revokeObjectURL(blobUrl);}
 const asset=await workspace.documentAssets(id).save(file,true);workspace.assertWritable(id);return {asset,width,height};
}
export function InsertOfficeImage({workspace,id,doc,target,kind,disabled,onPresentationImage}:{workspace:LocalWorkspace;id:string;doc:Y.Doc;target:Y.Array<Y.Map<unknown>|Y.Text>|Y.Array<Y.Map<unknown>>;kind:'word'|'presentation';disabled:boolean;onPresentationImage?:(assetId:string,geometry:{width:number;height:number})=>Promise<void>}){
 const input=useRef<HTMLInputElement>(null),[busy,setBusy]=useState(false),[error,setError]=useState('');
 const insert=async(file:File)=>{setBusy(true);setError('');try{const {asset,width,height}=await saveOfficeImage(workspace,id,doc,kind,file,640);if(kind==='word'){const image=createWordImage(asset.id,width,height,file.name),author=wordTrackingAuthor(doc);if(author)markWordImageInserted(image,author);target.push([image]);}else{if(!target.doc||target._item?.deleted)throw new Error('目标幻灯片已删除，请重新插入');const size=presentationSize(doc),ratio=height/width*size.width/size.height,imageWidth=Math.min(0.5,0.8/ratio);if(onPresentationImage)await onPresentationImage(asset.id,{width:imageWidth,height:imageWidth*ratio});else{const object=createPresentationObject('image');object.set('assetId',asset.id);object.set('width',imageWidth);object.set('height',imageWidth*ratio);target.push([object]);}}}catch(e){setError(e instanceof Error?e.message:'插入失败');}finally{setBusy(false);if(input.current)input.current.value='';}};
 return <span className="office-image-insert"><input ref={input} type="file" accept="image/png,image/jpeg" aria-label="选择办公图片" hidden onChange={e=>{const file=e.target.files?.[0];if(file)void insert(file);}}/><button disabled={disabled||busy} onClick={()=>input.current?.click()}>{busy?'正在插入图片…':'插入图片'}</button>{error&&<span role="alert">{error}</span>}</span>;
}
const MAX_IMAGE_SIZE=10000;
/** Width available to an image: the paper's content box (paper width minus page margins). */
const contentWidth=(figure:HTMLElement)=>{const paper=figure.closest<HTMLElement>('.word-paper')??figure.parentElement;if(!paper)return MAX_IMAGE_SIZE;const style=getComputedStyle(paper);return Math.max(16,paper.clientWidth-parseFloat(style.paddingLeft)-parseFloat(style.paddingRight));};
const clamp=(value:number)=>Math.max(1,Math.min(MAX_IMAGE_SIZE,Math.round(value)));
/**
 * A Word picture: click to select, then drag a corner handle (keeps the aspect ratio) or use the floating picture
 * toolbar above it — size with optional ratio lock, original size, fit to page width, alt text and delete.
 * The toolbar floats over the page instead of sitting in the text flow, so selecting an image never moves content.
 */
export function WordImageEditor({block,workspace,id,disabled,onDelete}:{block:Y.Map<unknown>;workspace:LocalWorkspace;id:string;disabled:boolean;onDelete:()=>void}){
 const figure=useRef<HTMLElement>(null),width=Number(block.get('width')),height=Number(block.get('height')),[locked,setLocked]=useState(true),[drag,setDrag]=useState<{width:number;height:number}>(),[below,setBelow]=useState(false);
 const shown=drag??{width,height},ratio=height>0?width/height:1;
 const write=(next:{width?:number;height?:number;alt?:string})=>{if(disabled)return;workspace.assertWritable(id);if(block._item?.deleted)return;block.doc?.transact(()=>{for(const [key,value]of Object.entries(next))if(block.get(key)!==value)block.set(key,value);});};
 const setSize=(key:'width'|'height',value:number)=>{if(!(value>0&&value<=MAX_IMAGE_SIZE))return;if(!locked){write({[key]:value});return;}write(key==='width'?{width:value,height:clamp(value/ratio)}:{height:value,width:clamp(value*ratio)});};
 const fit=(target:number)=>{const scale=target/width;write({width:clamp(width*scale),height:clamp(height*scale)});};
 const natural=()=>{const image=figure.current?.querySelector('img');if(!image?.naturalWidth)return;const scale=Math.min(1,contentWidth(figure.current!)/image.naturalWidth);write({width:clamp(image.naturalWidth*scale),height:clamp(image.naturalHeight*scale)});};
 // Corner handles resize proportionally; the paper may be zoomed, so pointer deltas are converted back to page pixels.
 const resize=(event:React.PointerEvent<HTMLSpanElement>,corner:'nw'|'ne'|'sw'|'se')=>{
  if(disabled||event.button!==0)return;event.preventDefault();event.stopPropagation();figure.current?.focus();const handle=event.currentTarget;handle.setPointerCapture(event.pointerId);
  const paper=figure.current?.closest<HTMLElement>('.word-paper'),zoom=paper?paper.getBoundingClientRect().width/(paper.offsetWidth||1)||1:1,startX=event.clientX,max=contentWidth(figure.current!),direction=corner.endsWith('e')?1:-1;let latest={width,height};
  const move=(next:PointerEvent)=>{const w=Math.max(16,Math.min(max,width+direction*(next.clientX-startX)/zoom));latest={width:clamp(w),height:clamp(w/ratio)};setDrag(latest);};
  const up=()=>{handle.removeEventListener('pointermove',move);handle.removeEventListener('pointerup',up);handle.removeEventListener('pointercancel',up);setDrag(undefined);if(latest.width!==width)write(latest);};
  handle.addEventListener('pointermove',move);handle.addEventListener('pointerup',up);handle.addEventListener('pointercancel',up);
 };
 const keys=(event:React.KeyboardEvent<HTMLElement>)=>{if(event.target!==event.currentTarget||disabled)return;if(event.key==='Delete'||event.key==='Backspace'){event.preventDefault();onDelete();}};
 const place=()=>{const element=figure.current;if(element)setBelow(element.offsetTop<64);};
 const revision=validWordRevisionMark(block.get('del'))?'del':validWordRevisionMark(block.get('ins'))?'ins':undefined,revisionMark=revision&&block.get(revision) as WordRevisionMark|undefined,itemId=block._item?.id;
 return <figure ref={figure} className="office-word-image" data-word-image={itemId?`${itemId.client}:${itemId.clock}`:undefined} data-revision={revision} style={revisionMark?{['--revision-color' as string]:wordRevisionColor(revisionMark.author)}:undefined} tabIndex={0} aria-label={revision==='del'?'图片（已删除，待审阅）':revision==='ins'?'图片（新插入，待审阅）':'图片'} title="点击选中，拖动角点调整大小" onFocus={place} onKeyDown={keys} data-toolbar={below?'below':undefined}>
  <div className="office-word-image-frame" style={{width:shown.width,aspectRatio:`${shown.width} / ${shown.height}`,maxWidth:'100%'}}><OfficeImage workspace={workspace} id={id} assetId={String(block.get('assetId'))} alt={String(block.get('alt'))}/>
   {!disabled&&(['nw','ne','sw','se'] as const).map(corner=><span key={corner} className="word-image-handle" data-corner={corner} aria-hidden="true" onPointerDown={event=>resize(event,corner)}/>)}
   {drag&&<span className="word-image-size-tip" aria-hidden="true">{drag.width} × {drag.height}</span>}
  </div>
  <fieldset className="word-image-toolbar" role="toolbar" aria-label="图片工具" disabled={disabled} onMouseDown={event=>event.stopPropagation()}>
   <SizeInput label="宽度（像素）" short="宽" value={width} onCommit={value=>setSize('width',value)}/>
   <button type="button" className="word-image-lock" aria-pressed={locked} aria-label="锁定纵横比" title={locked?'已锁定纵横比':'未锁定纵横比'} onClick={()=>setLocked(value=>!value)}>{locked?<Lock size={13} aria-hidden="true"/>:<LockOpen size={13} aria-hidden="true"/>}</button>
   <SizeInput label="高度（像素）" short="高" value={height} onCommit={value=>setSize('height',value)}/>
   <span className="word-image-sep" aria-hidden="true"/>
   <button type="button" onClick={natural}>原始大小</button><button type="button" onClick={()=>fit(contentWidth(figure.current!))}>适应页宽</button>
   <span className="word-image-sep" aria-hidden="true"/>
   <label className="word-image-alt" title="替代文字"><span aria-hidden="true">替代文字</span><input aria-label="替代文字" maxLength={1000} value={String(block.get('alt'))} placeholder="为读屏软件描述图片" onChange={e=>write({alt:e.target.value})}/></label>
   <button type="button" className="word-image-delete" onClick={onDelete}><Trash2 size={13} aria-hidden="true"/>删除图片</button>
  </fieldset>
 </figure>;
}
/** Commits on Enter, blur or arrow-key steps, so typing "200" never writes 2 and 20 first. */
function SizeInput({label,short,value,onCommit}:{label:string;short:string;value:number;onCommit:(value:number)=>void}){
 const [draft,setDraft]=useState<string>();
 const commit=(text:string)=>{const next=Number(text);if(Number.isFinite(next)&&next>0&&next<=MAX_IMAGE_SIZE&&next!==value)onCommit(Math.round(next));};
 return <label className="word-image-size" title={label}><span aria-hidden="true">{short}</span><input aria-label={label} type="number" min="1" max={MAX_IMAGE_SIZE} value={draft??Math.round(value)} onChange={event=>setDraft(event.target.value)} onBlur={event=>{commit(event.target.value);setDraft(undefined);}} onKeyDown={event=>{if(event.key==='Enter'){event.preventDefault();commit(event.currentTarget.value);setDraft(undefined);}if(event.key==='ArrowUp'||event.key==='ArrowDown'){event.preventDefault();const next=Math.round(Number(event.currentTarget.value)||value)+(event.key==='ArrowUp'?1:-1)*(event.shiftKey?10:1);setDraft(undefined);commit(String(next));}if(event.key==='Escape')setDraft(undefined);}}/></label>;
}
