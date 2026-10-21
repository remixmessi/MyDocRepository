import {captureOfficeTextAnchor,officeAnchorTarget} from '../../../common/nbstore/src/office-comment-anchor';
import {OfficeComments} from './office-comments-ui';
import {captureTextCursor} from './office-text-presence';
import {useTextPresence} from './use-text-presence';
import {TextPresenceList} from './text-presence-list';
import {TextLineGutter} from './text-line-gutter';
import {textLineStarts,textLinePosition,textCharacterCount} from './office-text-lines';
import {TextFindPanel} from './text-find-panel';
import {useCallback,useEffect,useLayoutEffect,useMemo,useReducer,useRef,useState} from 'react';
import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
import {Download,Undo2,Redo2,Settings2,FileDown} from 'lucide-react';
import type {OfficeTextDocument,LocalWorkspace} from './workspace';
import {SaveStatus} from './text-editor-status';
import {applyOfficeTextEdit,officeTextSettings,setOfficeTextSettings} from './office-text';
import {downloadOfficeBytes} from './office-download';
import {getCloudOfficeOriginal} from '../../../common/nbstore/src/cloud';
import {exportOfficeText,type TextEncoding,type TextLineEnding} from '../../../common/nbstore/src/office';

export function TextDocumentEditor({item,workspace,name}:{item:OfficeTextDocument;workspace:LocalWorkspace;name:string}){
  const [,refresh]=useReducer(n=>n+1,0);
  const [value,setValue]=useState(()=>item.text.toString());
  const [error,setError]=useState('');
  const [settingsOpen,setSettingsOpen]=useState(false),[downloading,setDownloading]=useState(false);
  const [findOpen,setFindOpen]=useState(false),[composing,setComposing]=useState(false);
  const input=useRef<HTMLTextAreaElement>(null),queryInput=useRef<HTMLTextAreaElement>(null);
  const composition=useRef<{doc:Y.Doc;text:Y.Text;before:string}>();
  const history=useRef<Y.UndoManager>();
  const localCursor=useRef<ReturnType<typeof captureTextCursor>>(),pendingCaret=useRef<{start:number;end:number;backward:boolean}>();
  const [wrap,setWrap]=useState(false),[lineNumbers,setLineNumbers]=useState(true),[selection,setSelection]=useState({start:0,end:0,direction:'none' as 'forward'|'backward'|'none'});
  const lineStarts=useMemo(()=>textLineStarts(value),[value]);
  const position=useMemo(()=>textLinePosition(value,lineStarts,selection.direction==='backward'?selection.start:selection.end),[value,lineStarts,selection]);
  const selectedCharacters=useMemo(()=>textCharacterCount(value,selection.start,selection.end),[value,selection]);
  const readPosition=useCallback(()=>{const element=input.current;if(!element)return;const start=element.selectionStart,end=element.selectionEnd,direction=element.selectionDirection;if(!composition.current&&!pendingCaret.current)localCursor.current=captureTextCursor(item.text,start,end,direction==='backward');setSelection(previous=>previous.start===start&&previous.end===end&&previous.direction===direction?previous:{start,end,direction});},[item.text]);
  useLayoutEffect(()=>{const caret=pendingCaret.current;pendingCaret.current=undefined;if(caret)input.current?.setSelectionRange(caret.start,caret.end,caret.backward?'backward':'forward');readPosition();},[value,readPosition]);
  useEffect(()=>{const element=input.current;if(!element)return;element.addEventListener('select',readPosition);return()=>element.removeEventListener('select',readPosition);},[readPosition]);
  const peers=useTextPresence(item.text,input,selection,composing,workspace.sync?.status==='connected');
  const readOnly=workspace.documentReadOnly(item.id);
  const settings=officeTextSettings(item.doc),resource=workspace.officeResources.get(item.id);
  useEffect(()=>{
    const undo=history.current=new Y.UndoManager(item.text,{trackedOrigins:new Set([LOCAL_ORIGIN])});
    const update=(event?:{transaction?:{origin:unknown}})=>{
      if(!composition.current){
        if(event?.transaction?.origin!==LOCAL_ORIGIN&&localCursor.current&&input.current===document.activeElement){const anchor=Y.createAbsolutePositionFromRelativePosition(localCursor.current.anchor,item.doc),head=Y.createAbsolutePositionFromRelativePosition(localCursor.current.head,item.doc);if(anchor?.type===item.text&&head?.type===item.text)pendingCaret.current={start:Math.min(anchor.index,head.index),end:Math.max(anchor.index,head.index),backward:anchor.index>head.index};}
        setValue(item.text.toString());
      }refresh();
    };
    item.text.observe(update);item.doc.getMap('textSettings').observe(update);const off=item.persistence.subscribe(refresh),offSync=workspace.subscribe(refresh);setValue(item.text.toString());refresh();
    return()=>{item.text.unobserve(update);item.doc.getMap('textSettings').unobserve(update);off();offSync();undo.destroy();composition.current?.doc.destroy();composition.current=undefined;};
  },[item,workspace]);
  const commit=(next:string)=>{if(readOnly)return;try{workspace.assertWritable(item.id);applyOfficeTextEdit(item.text,item.text.toString(),next);setValue(item.text.toString());setError('');}catch(error){setError(String(error));}};
  const changeSettings=(value:Parameters<typeof setOfficeTextSettings>[1])=>{if(readOnly)return;try{setOfficeTextSettings(item.doc,value);setError('');}catch(error){setError(String(error));}};
  const exportText=(utf8=false)=>{try{downloadOfficeBytes(name.endsWith('.txt')?name:name+'.txt',exportOfficeText(item.text.toString(),utf8?{...settings,encoding:'utf-8',bom:false}:settings));setError('');}catch(error){setError(error instanceof Error?error.message:'导出失败');}};
  return <section className="text-document-editor" aria-label={name+' 文本编辑器'}>
    <header className="editor-header"><span title={name}>{name} / 文本</span><div className="header-actions"><OfficeComments doc={item.doc} workspace={workspace} id={item.id} captureAnchor={()=>{const el=input.current;if(el&&el.selectionEnd>el.selectionStart)return captureOfficeTextAnchor(item.text,el.selectionStart,item.text,el.selectionEnd,item.text.toString().slice(el.selectionStart,el.selectionEnd));}} onLocate={anchor=>{const target=officeAnchorTarget(item.doc,anchor);if(target?.kind==='text'){input.current?.focus();input.current?.setSelectionRange(target.start,target.end);readPosition();}}}/>
      <SaveStatus persistence={item.persistence} sync={workspace.sync}/>
      <button title="撤销" aria-label="撤销文本" disabled={readOnly||!history.current?.canUndo()} onClick={()=>history.current?.undo()}><Undo2 size={16}/></button>
      <button title="重做" aria-label="重做文本" disabled={readOnly||!history.current?.canRedo()} onClick={()=>history.current?.redo()}><Redo2 size={16}/></button>
      <button aria-label="查找替换" aria-expanded={findOpen} onClick={()=>setFindOpen(value=>!value)}>查找替换</button>
      <button title="文本格式" aria-label="文本格式" aria-expanded={settingsOpen} onClick={()=>setSettingsOpen(value=>!value)}><Settings2 size={16}/></button>
      <button title={`导出 ${settings.encoding.toUpperCase()} TXT`} aria-label="导出 TXT" onClick={()=>exportText()}><Download size={16}/></button>
      {resource?.originalAssetId&&<button title="下载原件" aria-label="下载原件" disabled={downloading} onClick={()=>{setDownloading(true);setError('');void getCloudOfficeOriginal(workspace.id,item.id).then(response=>response.arrayBuffer()).then(buffer=>downloadOfficeBytes(name,new Uint8Array(buffer))).catch(error=>setError(error instanceof Error?error.message:'原件下载失败')).finally(()=>setDownloading(false));}}><FileDown size={16}/></button>}
    </div></header>
    {settingsOpen&&<div className="text-format-controls" role="group" aria-label="文本格式设置">
      <label>编码<select aria-label="文本编码" disabled={readOnly} value={settings.encoding} onChange={event=>changeSettings({encoding:event.target.value as TextEncoding})}><option value="utf-8">UTF-8</option><option value="utf-16le">UTF-16LE</option><option value="utf-16be">UTF-16BE</option></select></label>
      <label>换行<select aria-label="文本换行" disabled={readOnly} value={settings.lineEnding} onChange={event=>changeSettings({lineEnding:event.target.value as TextLineEnding})}><option value="lf">LF</option><option value="crlf">CRLF</option><option value="cr">CR</option></select></label>
      <label><input type="checkbox" aria-label="文本 BOM" checked={settings.bom} disabled={readOnly} onChange={event=>changeSettings({bom:event.target.checked})}/>BOM</label>
      <label><input type="checkbox" aria-label="自动换行" checked={wrap} onChange={event=>setWrap(event.target.checked)}/>自动换行</label>
      <label><input type="checkbox" aria-label="显示行号" checked={lineNumbers&&!wrap} disabled={wrap} onChange={event=>setLineNumbers(event.target.checked)}/>显示行号</label>
      {wrap&&<span>自动换行时隐藏侧栏行号</span>}
      <button onClick={()=>exportText(true)}><Download size={16}/>另存 UTF-8</button>
    </div>}
    {findOpen&&<TextFindPanel text={item.text} input={input} queryInput={queryInput} readOnly={readOnly} composing={composing} assertWritable={()=>workspace.assertWritable(item.id)} boundary={()=>history.current?.stopCapturing()} onClose={()=>setFindOpen(false)}/>}
    <TextPresenceList peers={peers} input={input}/>
    {error&&<p role="alert">{error}</p>}
    <div className="text-document-surface">
    {lineNumbers&&!wrap&&<TextLineGutter input={input} count={lineStarts.length} value={value} peers={composing?[]:peers}/>}
    <textarea wrap={wrap?'soft':'off'} onSelect={readPosition} onClick={readPosition} onKeyUp={readPosition} ref={input} className="text-document-input" aria-label={name} readOnly={readOnly} value={value} spellCheck={false}
      onCompositionStart={()=>{if(readOnly)return;setComposing(true);composition.current?.doc.destroy();const doc=new Y.Doc();const text=doc.getText('text');Y.applyUpdate(doc,Y.encodeStateAsUpdate(item.doc));composition.current={doc,text,before:text.toString()};}}
      onCompositionEnd={event=>{setComposing(false);const pending=composition.current;composition.current=undefined;if(!pending)return;try{if(!readOnly){workspace.assertWritable(item.id);applyOfficeTextEdit(item.text,pending.before,event.currentTarget.value,pending.text);}setValue(item.text.toString());}catch(error){setError(String(error));}finally{pending.doc.destroy();}}}
      onChange={event=>{if(composition.current)setValue(event.target.value);else commit(event.target.value);}}
      onKeyDown={event=>{if(!event.nativeEvent.isComposing&&(event.metaKey||event.ctrlKey)&&['f','h'].includes(event.key.toLowerCase())){event.preventDefault();event.stopPropagation();setFindOpen(true);requestAnimationFrame(()=>queryInput.current?.focus());return;}if(event.nativeEvent.isComposing||!(event.metaKey||event.ctrlKey)||event.key.toLowerCase()!=='z')return;event.preventDefault();if(!readOnly){if(event.shiftKey)history.current?.redo();else history.current?.undo();}}}/>
    </div>
    <div className="text-position-status" role="status" aria-live="off" aria-label="文本位置">第 {position.line} 行，第 {position.column} 列 · 共 {lineStarts.length} 行{selectedCharacters>0&&` · 已选 ${selectedCharacters} 个字符`} · {wrap?'自动换行':'不自动换行'}</div>
  </section>;
}
