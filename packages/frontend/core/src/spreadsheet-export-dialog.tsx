import {useState} from 'react';
import {Modal} from './dialogs';
import {cloudRequest} from '../../../common/nbstore/src/cloud';
import {downloadOfficeBytes} from './office-download';
import type {LocalWorkspace} from './workspace';
export function SpreadsheetExportDialog({workspace,id,name,kind,onClose}:{workspace:LocalWorkspace;id:string;name:string;kind:'text'|'formula';onClose:()=>void}){
 const [cache,setCache]=useState(false),[busy,setBusy]=useState(false),[error,setError]=useState('');
 const run=async()=>{if(busy)return;setBusy(true);setError('');try{await workspace.prepareOfficeBackup(id);const response=await cloudRequest('/office-export?'+new URLSearchParams({workspace:workspace.id,id,format:'xlsx',scope:kind==='formula'?(cache?'formula-cache-grid':'formula-grid'):'text-grid'}));downloadOfficeBytes(name.replace(/\.[^.]+$/,'')+(kind==='formula'?(cache?'-公式结果副本.xlsx':'-公式副本.xlsx'):'-文本副本.xlsx'),new Uint8Array(await response.arrayBuffer()),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');onClose();}catch(e){setError(e instanceof Error?e.message:'表格导出失败');}finally{setBusy(false);}};
 return <Modal title={kind==='formula'?'导出 XLSX 公式副本':'导出 XLSX 文本副本'} onClose={()=>{if(!busy)onClose();}}><p>{kind==='formula'?'导出当前工作簿中全部已接入工作表的已保存内容。保留文本和普通公式；自动类型中的数字按数值保存，编号请使用文本类型。公式需通过当前基础计算器校验，不支持的公式或错误会阻止导出。':'导出当前工作簿中全部已接入工作表的已保存内容。所有单元格均按文本保存，保留前导零、长编号和公式原文；公式不会计算。'}</p>
  {kind==='formula'&&<><label><input type="checkbox" checked={cache} disabled={busy} onChange={event=>setCache(event.target.checked)}/>写入服务端计算结果</label><p>{cache?'同时写入当前快照的计算结果，保留数值、文本和布尔类型；兼容软件打开后仍会重新计算。':'公式副本不包含计算缓存，打开后由兼容软件重新计算。'}支持已接入工作表之间的引用；共享公式、数组公式或完整 Excel 函数尚未接入。</p></>}
  <p>副本包含已接入的单元格格式、数字/日期格式和合并单元格{kind==='text'?'（文本副本中的日期按显示文字保存）':''}；单元格讨论写为 Excel 批注（同一单元格的讨论与回复合并为一条，已解决的标注“[已解决]”）。不包含边框、主题、图表或协作历史。原件保持不变；需要完整协作内容时可使用办公文件备份。</p>
  {error&&<p role="alert" className="form-error">{error}</p>}
  <div className="dialog-actions"><button disabled={busy} onClick={onClose}>取消</button><button className="primary" disabled={busy} onClick={()=>void run()}>{busy?'正在导出…':kind==='formula'?'下载公式副本':'下载文本副本'}</button></div></Modal>;
}
