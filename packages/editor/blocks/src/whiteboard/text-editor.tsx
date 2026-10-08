import {fitTextHeight} from './shape-controls';
import {textFont,lineHeight} from './text-style';
import {useEffect,useRef,useState} from 'react';
import type {SurfaceStore,Viewport} from './model';
import {shapeTextLayout,type EditableShape} from './text-layout';
export function ShapeTextEditor({shape,surface,view,size,close}:{shape:EditableShape;surface:SurfaceStore;view:Viewport;size:{width:number;height:number};close:()=>void}){
  const input=useRef<HTMLTextAreaElement>(null),composing=useRef(false);
  const [draft,setDraft]=useState(shape.text),[error,setError]=useState('');
  useEffect(()=>{surface.store.history.stopCapturing();input.current?.focus({preventScroll:true});input.current?.select();input.current?.scrollIntoView({block:'nearest',inline:'nearest'});return()=>surface.store.history.stopCapturing();},[surface]);
  useEffect(()=>{if(!composing.current)setDraft(shape.text);},[shape.text]);
  const measureCanvas=document.createElement('canvas'),ctx=measureCanvas.getContext('2d');if(ctx)ctx.font=textFont(shape);
  const layout=shapeTextLayout({...shape,text:draft},text=>ctx?.measureText(text).width??text.length*16);
  const save=(text:string)=>{try{surface.updateText(shape.id,text,shape.textStyle?.autoHeight?fitTextHeight({...shape,text}).height:undefined);setError('');}catch(err){setDraft(surface.elements.get(shape.id)?.text??'');setError(err instanceof Error?err.message:'无法保存文字');}};
  const left=Math.max(4,Math.min(size.width-156,view.x+shape.x*view.zoom)),top=Math.max(4,Math.min(size.height-40,view.y+shape.y*view.zoom-36));
  return <div className="board-text-editor" role="group" aria-label="画布文字编辑">
    <textarea ref={input} aria-label="白板文字正文" aria-describedby="board-text-help" spellCheck={false} value={draft}
      style={{width:layout.width,height:layout.height, paddingTop:layout.paddingTop,font:textFont(shape),lineHeight:lineHeight(shape)+'px',color:shape.textStyle?.color??'#20242c',textAlign:shape.textStyle?.align??(layout.centered?'center':'left'),transform:`translate(${view.x+layout.x*view.zoom}px,${view.y+layout.y*view.zoom}px) scale(${view.zoom})`}}
      onChange={e=>{setDraft(e.target.value);if(!composing.current)save(e.target.value);}}
      onCompositionStart={()=>{composing.current=true;}} onCompositionEnd={e=>{composing.current=false;save(e.currentTarget.value);}}
      onPointerDown={e=>e.stopPropagation()}
      onKeyDown={e=>{e.stopPropagation();if(e.nativeEvent.isComposing||composing.current)return;if(e.key==='Escape'||((e.metaKey||e.ctrlKey)&&e.key==='Enter')){e.preventDefault();close();}else if((e.metaKey||e.ctrlKey)&&['z','y'].includes(e.key.toLowerCase())){e.preventDefault();if(e.shiftKey||e.key.toLowerCase()==='y')surface.store.history.redo();else surface.store.history.undo();}}}/>
    <div className="board-text-actions" style={{left,top}}><button onClick={close}>完成文字编辑</button><span id="board-text-help">Esc 完成</span></div>
    {error&&<p role="alert" className="board-text-error">{error}</p>}
  </div>;
}
