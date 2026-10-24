import {useRef,useState} from 'react';
import {Modal} from './dialogs';
import {parsePdfPageRange,validPdfWatermark,type PdfWatermark} from '../../../common/nbstore/src/office-pdf';
/** A drawn signature: strokes of points normalised to 0..1, and the drawing's width/height ratio. */
export type PdfSignature={strokes:number[][];ratio:number};
/** 提取页面: positions such as "1-3,5" of the arranged pages become a new PDF. */
export function PdfExtractDialog({count,initial,busy,onExtract,onClose}:{count:number;initial:string;busy:boolean;onExtract:(positions:number[])=>void;onClose:()=>void}){
 const [value,setValue]=useState(initial),positions=parsePdfPageRange(value,count);
 return <Modal title="提取页面" onClose={()=>{if(!busy)onClose();}}>
  <label className="form-field">页码范围<input aria-label="提取页码范围" value={value} placeholder={`例如 1-3,5（共 ${count} 页）`} onChange={event=>setValue(event.target.value)}/></label>
  <p role="status">{positions?`将提取 ${positions.length} 页，生成新的 PDF 文件（含批注，原件不变）`:'请输入有效的页码，用逗号分隔，连续页用“-”'}</p>
  <div className="dialog-actions"><button disabled={busy} onClick={onClose}>取消</button><button className="primary" disabled={busy||!positions} onClick={()=>positions&&onExtract(positions)}>{busy?'正在提取…':'提取'}</button></div>
 </Modal>;
}
const PRESETS=['机密','内部资料','草稿','样本','严禁复制','CONFIDENTIAL'];
/** 水印 for PDF: text, colour, opacity, size and angle; drawn into every page of exported copies. */
export function PdfWatermarkDialog({current,busy,readOnly,onSave,onClose}:{current?:PdfWatermark;busy:boolean;readOnly:boolean;onSave:(value:PdfWatermark|null)=>void;onClose:()=>void}){
 const [text,setText]=useState(current?.text??PRESETS[0]),[color,setColor]=useState(current?.color??'#C00000'),[opacity,setOpacity]=useState(current?.opacity??.25),[size,setSize]=useState(current?.size??60),[rotation,setRotation]=useState(current?.rotation??45);
 const value:PdfWatermark={text:text.trim(),color:color.toUpperCase(),opacity,size,rotation};
 return <Modal title="PDF 水印" onClose={()=>{if(!busy)onClose();}}>
  <label className="form-field">文字<input aria-label="水印文字" maxLength={50} value={text} onChange={event=>setText(event.target.value)}/></label>
  <div className="word-watermark-presets" role="group" aria-label="常用水印">{PRESETS.map(preset=><button key={preset} type="button" aria-pressed={text===preset} onClick={()=>setText(preset)}>{preset}</button>)}</div>
  <label className="form-field">颜色<input type="color" aria-label="水印颜色" value={color} onChange={event=>setColor(event.target.value.toUpperCase())}/></label>
  <label className="form-field">透明度 {Math.round((1-opacity)*100)}%<input type="range" aria-label="水印透明度" min={0} max={95} step={5} value={Math.round((1-opacity)*100)} onChange={event=>setOpacity(Math.round((1-Number(event.target.value)/100)*100)/100)}/></label>
  <label className="form-field">字号<select aria-label="水印字号" value={size} onChange={event=>setSize(Number(event.target.value))}>{[24,36,48,60,72,96,120,144].map(value=><option key={value} value={value}>{value}</option>)}</select></label>
  <label className="form-field">角度<select aria-label="水印角度" value={rotation} onChange={event=>setRotation(Number(event.target.value))}><option value={45}>斜向上 45°</option><option value={-45}>斜向下 45°</option><option value={0}>水平</option><option value={90}>竖直</option></select></label>
  <p className="dialog-note">水印显示在阅读页面上，导出带批注或扁平化 PDF 时写入每一页；原件不变。</p>
  <div className="dialog-actions">{current&&<button disabled={busy||readOnly} onClick={()=>onSave(null)}>删除水印</button>}<button disabled={busy} onClick={onClose}>取消</button><button className="primary" disabled={busy||readOnly||!validPdfWatermark(value)} onClick={()=>onSave(value)}>确定</button></div>
 </Modal>;
}
/** 签名: draw with mouse, pen or finger; the strokes are then placed on the page as one ink annotation. */
export function PdfSignatureDialog({onDone,onClose}:{onDone:(signature:PdfSignature)=>void;onClose:()=>void}){
 const canvas=useRef<HTMLCanvasElement>(null),[strokes,setStrokes]=useState<number[][]>([]),drawing=useRef<number[]|null>(null),width=480,height=180;
 const point=(event:React.PointerEvent<HTMLCanvasElement>)=>{const rect=event.currentTarget.getBoundingClientRect();return [(event.clientX-rect.left)*width/rect.width,(event.clientY-rect.top)*height/rect.height];};
 const paint=(list:number[][])=>{const context=canvas.current?.getContext('2d');if(!context)return;context.clearRect(0,0,width,height);context.lineWidth=2.5;context.lineCap='round';context.lineJoin='round';context.strokeStyle='#1e2a5a';for(const stroke of list){context.beginPath();for(let i=0;i<stroke.length;i+=2)i?context.lineTo(stroke[i],stroke[i+1]):context.moveTo(stroke[i],stroke[i+1]);context.stroke();}};
 const finish=()=>{const xs=strokes.flatMap(s=>s.filter((_,i)=>i%2===0)),ys=strokes.flatMap(s=>s.filter((_,i)=>i%2===1)),left=Math.min(...xs),top=Math.min(...ys),w=Math.max(1,Math.max(...xs)-left),h=Math.max(1,Math.max(...ys)-top);
  onDone({ratio:w/h,strokes:strokes.map(stroke=>stroke.map((n,i)=>i%2===0?(n-left)/w:(n-top)/h))});};
 return <Modal title="签名" onClose={onClose}>
  <p>在下方区域书写签名，确定后在页面上单击放置。此为手写签名图样，不是数字证书签名。</p>
  <canvas ref={canvas} className="pdf-signature-pad" width={width} height={height} aria-label="签名书写区" role="img" style={{touchAction:'none'}}
   onPointerDown={event=>{event.currentTarget.setPointerCapture(event.pointerId);drawing.current=point(event);}}
   onPointerMove={event=>{if(!drawing.current)return;drawing.current=[...drawing.current,...point(event)];paint([...strokes,drawing.current]);}}
   onPointerUp={()=>{const stroke=drawing.current;drawing.current=null;if(stroke&&stroke.length>=4&&strokes.length<100){const next=[...strokes,stroke];setStrokes(next);paint(next);}else paint(strokes);}}/>
  <p role="status">{strokes.length?`已书写 ${strokes.length} 笔`:'尚未书写'}</p>
  <div className="dialog-actions"><button disabled={!strokes.length} onClick={()=>{setStrokes([]);paint([]);}}>清除</button><button onClick={onClose}>取消</button><button className="primary" disabled={!strokes.length} onClick={finish}>确定并放置</button></div>
 </Modal>;
}
/** Signature strokes placed with their top-left at `at` (PDF points, y up), `width` wide, inside `box`; at most 1000 points. */
export function placePdfSignature(signature:PdfSignature,at:[number,number],box:number[],width=160){
 const height=width/Math.max(.2,signature.ratio),[left,bottom,right,top]=box,x=Math.max(left,Math.min(right-width,at[0])),y=Math.max(bottom,Math.min(top-height,at[1]-height));
 const total=signature.strokes.reduce((n,s)=>n+s.length/2,0),step=Math.max(1,Math.ceil(total/1000));
 const strokes=signature.strokes.map(stroke=>{const kept:number[]=[];for(let i=0;i<stroke.length;i+=2)if((i/2)%step===0||i===stroke.length-2)kept.push(x+stroke[i]*width,y+(1-stroke[i+1])*height);return kept.length>=4?kept:[...kept,...kept];});
 return {x,y,width,height,points:strokes.flat(),strokes:strokes.map(stroke=>stroke.length/2)};
}
