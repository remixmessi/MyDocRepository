import {useEffect,useState} from 'react';
import {expandFrames,expandGroups,selectionBounds,type SurfaceElement} from './model';
export function GroupSizeProperties({ids,elements,keepRatio,setKeepRatio,apply}:{ids:string[];elements:SurfaceElement[];keepRatio:boolean;setKeepRatio:(value:boolean)=>void;apply:(width:number,height:number)=>void}){
  const members=expandFrames(ids,elements),bounds=selectionBounds(members,elements);
  const memberSet=new Set(members);const locked=elements.some(s=>memberSet.has(s.id)&&(s.kind==='mindmap'||s.kind==='image')),ratio=locked||keepRatio;
  const [width,setWidth]=useState(''),[height,setHeight]=useState('');
  const key=ids.join(',');
  useEffect(()=>{setWidth(String(Math.round((bounds?.width??0)*100)/100));setHeight(String(Math.round((bounds?.height??0)*100)/100));},[key,bounds?.width,bounds?.height]);
  if(!bounds)return null;
  return <div className="whiteboard-selection-tools whiteboard-group-size" aria-label="组合尺寸">
    <label>宽度<input aria-label="组合宽度" type="number" min="1" step="any" value={width} onChange={e=>{setWidth(e.target.value);if(ratio&&Number(e.target.value)>0)setHeight(String(Math.round(Number(e.target.value)*bounds.height/bounds.width*100)/100));}}/></label>
    <label>高度<input aria-label="组合高度" type="number" min="1" step="any" value={height} onChange={e=>{setHeight(e.target.value);if(ratio&&Number(e.target.value)>0)setWidth(String(Math.round(Number(e.target.value)*bounds.width/bounds.height*100)/100));}}/></label>
    <label><input aria-label="保持组合比例" type="checkbox" checked={ratio} disabled={locked} onChange={e=>setKeepRatio(e.target.checked)}/>保持比例</label>
    <button disabled={!(Number(width)>0&&Number(height)>0&&Number.isFinite(Number(width))&&Number.isFinite(Number(height)))} onClick={()=>apply(Number(width),Number(height))}>应用组合尺寸</button>
    <output aria-label="组合范围">{Math.round(bounds.width*100)/100} × {Math.round(bounds.height*100)/100}</output>
    <span>{locked?'含图片或思维导图，保持比例缩放。':'以组合左上角为基准，成员保留最小尺寸。'}文本字号、笔画粗细保持；外部连接端点保留锚定。</span>
  </div>;
}
