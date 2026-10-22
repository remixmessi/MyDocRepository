import {Y} from '@mydoc/crdt';
import {decodeBinary} from './backup';
import type {OfficeKind} from '../../../common/nbstore/src/office';
const roots:Record<OfficeKind,string[]>={text:['text','textSettings'],word:['paragraphs','pageSettings'],spreadsheet:['sheetColumns','sheet','workbookSheets','sheetOrder','sheetCellTypes','sheetImport','sheetXlsx','sheetFormats','sheetMerges','sheetObjects'],pdf:['pdfFields','annotations'],presentation:['objects','slides','slideOrder','presentation']};
type Snapshot={resource:{kind:OfficeKind;name:string};state:{data:string;sha256:string};original:null|{sha256:string};media?:{id:string;sha256:string}[]};
function canonical(value:unknown):unknown{
 if(value instanceof Y.Text)return {attributes:canonical(value.getAttributes()),runs:canonical(value.toDelta())};
 if(value instanceof Y.Map)return Object.fromEntries([...value.entries()].sort(([a],[b])=>a.localeCompare(b)).map(([key,value])=>[key,canonical(value)]));
 if(value instanceof Y.Array||Array.isArray(value))return [...value].map(canonical);
 if(value&&typeof value==='object')return Object.fromEntries(Object.entries(value).sort(([a],[b])=>a.localeCompare(b)).map(([key,value])=>[key,canonical(value)]));
 return value;
}
export function officeVersionContent(snapshot:Snapshot){
 if(!snapshot||!snapshot.resource||!Object.hasOwn(roots,snapshot.resource.kind)||typeof snapshot.state?.data!=='string'||snapshot.state.data.length>24*1024*1024)throw new Error('版本比较数据无效');
 const doc=new Y.Doc();try{
  Y.applyUpdate(doc,decodeBinary(snapshot.state.data));if(doc.store.pendingStructs||doc.store.pendingDs)throw new Error('版本比较数据缺少依赖');
  const content=Object.fromEntries([...roots[snapshot.resource.kind],'officeComments'].filter(root=>doc.share.has(root)).map(root=>{
   // Yjs decodes a root as AbstractType until its expected shared type is requested.
   const mapRoots=['pdfFields','sheetColumns','textSettings','pageSettings','workbookSheets','sheetCellTypes','sheetImport','sheetXlsx','sheetFormats','sheetMerges','sheetObjects','presentation','slides','officeComments'];
   const value=root==='text'?doc.getText(root):mapRoots.includes(root)?doc.getMap(root):doc.getArray(root);return [root,canonical(value)];
  }));
  return JSON.stringify({name:snapshot.resource.name,content,original:snapshot.original?.sha256??null,media:[...(snapshot.media??[])].map(({id,sha256})=>({id,sha256})).sort((a,b)=>a.id.localeCompare(b.id))},null,2);
 }finally{doc.destroy();}
}
export function compareOfficeVersions(previous:Snapshot,current:Snapshot){
 if(previous.resource.kind!==current.resource.kind)throw new Error('版本文件类型不一致');
 const before=officeVersionContent(previous),after=officeVersionContent(current);if(before===after)return {equal:true,before:'',after:'',removed:0,added:0,truncated:false};
 const left=before.split('\n'),right=after.split('\n');let first=0,endLeft=left.length,endRight=right.length;
 while(first<endLeft&&first<endRight&&left[first]===right[first])first++;
 while(endLeft>first&&endRight>first&&left[endLeft-1]===right[endRight-1]){endLeft--;endRight--;}
 const clip=(lines:string[],end:number)=>lines.slice(Math.max(0,first-3),Math.min(lines.length,end+3)).slice(0,1000).join('\n');
 return {equal:false,before:clip(left,endLeft),after:clip(right,endRight),removed:endLeft-first,added:endRight-first,truncated:endLeft-first>994||endRight-first>994};
}
