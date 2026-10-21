import {useEffect,useState} from 'react';
import * as Y from 'yjs';
import {CLOUD_ORIGIN,cloudJson,CloudError} from '../../../common/nbstore/src/cloud';
import type {LocalWorkspace,OfficePdfDocument} from './workspace';
type Native={key:string;page:number;subtype:string;importable:boolean;reason?:string};
/** Native annotations already in the original: Text/Highlight/Ink can be taken over once as editable annotations. */
export function PdfNativeAnnotations({item,workspace,readOnly}:{item:OfficePdfDocument;workspace:LocalWorkspace;readOnly:boolean}){
 const [natives,setNatives]=useState<Native[]>(),[busy,setBusy]=useState(false),[error,setError]=useState(''),[pending,setPending]=useState<string>(),[,refresh]=useState(0);
 // Background information: start after the page and form panel have their parser slots, and retry while the pool is busy.
 useEffect(()=>{let disposed=false,timer:ReturnType<typeof setTimeout>;const load=(attempt:number)=>{timer=setTimeout(()=>void cloudJson<Native[]>('/office-pdf-native-annotations?'+new URLSearchParams({workspace:workspace.id,id:item.id})).then(value=>{if(!disposed)setNatives(value);}).catch(error=>{if(disposed)return;if(error instanceof CloudError&&error.status===503&&attempt<6)load(attempt+1);else setError(error instanceof Error?error.message:'读取原件批注失败');}),attempt?1500:800);};load(0);return()=>{disposed=true;clearTimeout(timer);};},[workspace,item.id]);
 useEffect(()=>{const update=()=>refresh(value=>value+1);item.doc.on('update',update);return()=>item.doc.off('update',update);},[item.doc]);
 if(!natives?.length&&!error)return null;const imported=item.doc.share.has('pdfSource')&&item.doc.getMap('pdfSource').has('$imported'),importable=natives?.filter(native=>native.importable).length??0,others=(natives?.length??0)-importable;
 const run=async()=>{if(busy||readOnly)return;setBusy(true);setError('');const operationId=pending??crypto.randomUUID();setPending(operationId);try{const result=await cloudJson<{update:string}>('/office-pdf-native-annotations?'+new URLSearchParams({workspace:workspace.id,id:item.id}),'POST',{operationId});Y.applyUpdate(item.doc,Uint8Array.from(atob(result.update),char=>char.charCodeAt(0)),CLOUD_ORIGIN);setPending(undefined);}catch(error){setError(error instanceof Error?error.message:'导入原件批注失败，可重试');}finally{setBusy(false);}};
 return <section className="pdf-native-annotations" aria-label="原件批注"><p>原件含 {natives?.length??0} 个批注：{importable} 个附注/高亮/手绘{imported?'已转为可协作编辑':'可转为可协作编辑'}{others?`，${others} 个其他类型仅保留在副本中，扁平化时按原外观绘制`:''}。{imported?'修改或删除后，导出副本以协作内容为准。':'转换后原件不变；导出副本以协作内容替换这些原生批注。'}</p>
  {error&&<p role="alert">{error}</p>}{(pending||!imported&&importable>0)&&<button disabled={readOnly||busy} onClick={()=>void run()}>{busy?'正在导入…':pending?'重试导入原件批注':`导入 ${importable} 个原件批注以便编辑`}</button>}
  {natives&&natives.some(native=>native.reason)&&<details><summary>未转换的批注</summary><ul>{natives.filter(native=>native.reason).map(native=><li key={native.key}>第 {native.page} 页 · {native.subtype}：{native.reason}</li>)}</ul></details>}
 </section>;
}
