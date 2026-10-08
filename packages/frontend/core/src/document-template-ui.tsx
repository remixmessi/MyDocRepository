import { useEffect, useRef, useState } from 'react';
import type { BlockStore } from '@mydoc/editor-framework';
import { readDocument, validSurfaceElement } from '@mydoc/editor-blocks';
import { renderEmbedPreview } from '../../../editor/blocks/src/embed-preview';
import { SurfacePreview } from '../../../editor/blocks/src/surface-preview';
import { elementBounds } from '../../../editor/blocks/src/whiteboard/model';
import { Modal, ResourceDialog } from './dialogs';
import { FolderSelect } from './sidebar';
import type { LocalWorkspace } from './workspace';
import { downloadFile } from './backup';
import { builtinDocumentTemplates, templateStore, disposeTemplate, saveDocumentTemplate, renameDocumentTemplate, deleteDocumentTemplate, instantiateDocumentTemplate, exportDocumentTemplate, parseTemplateBundle, importDocumentTemplate, type DocumentTemplate } from './document-templates';

export function DocumentTemplateDialog({workspace,saveDocId,onClose,onCreate}:{workspace:LocalWorkspace;saveDocId?:string;onClose:()=>void;onCreate:(id:string)=>void}) {
  const [,refresh]=useState(0);useEffect(()=>workspace.subscribe(()=>refresh(n=>n+1)),[workspace]);
  const [selected,setSelected]=useState(builtinDocumentTemplates()[0].id);
  const [query,setQuery]=useState('');const [title,setTitle]=useState('会议记录');const [folder,setFolder]=useState('');
  const [name,setName]=useState('');const [renaming,setRenaming]=useState(false);const [confirm,setConfirm]=useState(false);
  const [busy,setBusy]=useState(false);const [error,setError]=useState('');const fileInput=useRef<HTMLInputElement>(null);
  const all=[...builtinDocumentTemplates(),...[...workspace.templates.values()].sort((a,b)=>b.createdAt-a.createdAt)];
  const entries=all.filter(t=>t.name.toLocaleLowerCase().includes(query.trim().toLocaleLowerCase()));
  const template=entries.find(t=>t.id===selected);const custom=!!template&&workspace.templates.has(template.id);
  useEffect(()=>{if(!entries.some(t=>t.id===selected)){const first=entries[0];setSelected(first?.id??'');setTitle(first?.name??'');setConfirm(false);setRenaming(false);}},[selected,entries.map(t=>t.id).join(',')]);
  const perform=async(action:()=>Promise<void>)=>{if(busy)return;setBusy(true);setError('');try{await action();}catch(e){setError(e instanceof Error?e.message:'模板操作失败，请重试');}finally{setBusy(false);}};
  if(saveDocId)return <ResourceDialog title="保存为文档模板" nameLabel="模板名称" initialName={workspace.documents.get(saveDocId)?.title??''} onClose={onClose} action={async name=>{await saveDocumentTemplate(workspace,saveDocId,name);}}><p className="organization-help">保存当前正文、白板和附件快照。以后修改文档不会改变模板；文件夹、收藏及标签不作为模板内容。</p></ResourceDialog>;
  return <Modal title="文档模板" wide onClose={()=>{if(!busy)onClose();}}>
    <p className="organization-help">选择模板新建文档，或从文档工具栏保存自己的模板。</p>
    <div className="template-toolbar"><label className="sr-only" htmlFor="template-query">搜索文档模板</label><input id="template-query" className="search-input" type="search" placeholder="搜索模板名称" disabled={busy} value={query} onChange={e=>{setQuery(e.target.value);setConfirm(false);setRenaming(false);}}/><button disabled={busy} onClick={()=>fileInput.current?.click()}>导入文档模板</button></div>
    <input ref={fileInput} type="file" hidden aria-label="导入文档模板文件" accept=".json,application/json" disabled={busy} onChange={e=>{const file=e.target.files?.[0];e.target.value='';if(!file)return;void perform(async()=>{if(file.size>50*1024*1024)throw new Error('模板文件超过 50 MB');const t=await importDocumentTemplate(workspace,parseTemplateBundle(await file.text()));setQuery('');setSelected(t.id);setTitle(t.name);setConfirm(false);setRenaming(false);});}}/>
    {busy&&<p role="status">正在处理模板…</p>}{error&&<p className="form-error" role="alert">{error}</p>}
    <div className="trash-layout"><div className="trash-list" aria-label="文档模板列表">{entries.map(t=><button type="button" key={t.id} disabled={busy} aria-pressed={selected===t.id} className={selected===t.id?'selected':''} onClick={()=>{setSelected(t.id);setTitle(t.name);setConfirm(false);setRenaming(false);setError('');}}><strong>{t.name}</strong><small>{workspace.templates.has(t.id)?'自定义模板':'内置模板'}</small></button>)}{!entries.length&&<p className="organization-help">没有匹配的模板，可清除搜索或导入模板。</p>}</div>
      {template&&<div className="trash-detail"><h3>{template.name}</h3><TemplatePreview key={template.id+template.update} workspace={workspace} template={template}/>
        <form onSubmit={e=>{e.preventDefault();void perform(async()=>onCreate(await instantiateDocumentTemplate(workspace,template,title,folder||null)));}}><label className="form-field">新文档标题<input aria-label="模板新文档标题" value={title} maxLength={160} required disabled={busy} onChange={e=>setTitle(e.target.value)}/></label><label className="form-field">所在文件夹<FolderSelect workspace={workspace} value={folder} label="模板新文档文件夹" disabled={busy} onChange={setFolder}/></label><div className="dialog-actions"><button type="submit" className="primary" disabled={busy||!title.trim()}>从模板新建</button></div></form>
        <div className="template-actions"><button disabled={busy} onClick={()=>void perform(async()=>{const bundle=await exportDocumentTemplate(workspace,template);downloadFile(template.name.replace(/[\\/:*?"<>|]/g,'_')+'.文档模板.json',JSON.stringify(bundle),'application/json');})}>导出文档模板</button>{custom&&<><button disabled={busy} onClick={()=>{setName(template.name);setRenaming(true);setConfirm(false);}}>重命名模板</button><button disabled={busy} onClick={()=>{setConfirm(true);setRenaming(false);}}>删除文档模板</button></>}</div>
        {renaming&&<form onSubmit={e=>{e.preventDefault();void perform(async()=>{await renameDocumentTemplate(workspace,template.id,name);setRenaming(false);});}}><label className="form-field">模板名称<input aria-label="重命名模板名称" value={name} maxLength={160} required disabled={busy} onChange={e=>setName(e.target.value)}/></label><div className="dialog-actions"><button type="button" disabled={busy} onClick={()=>setRenaming(false)}>取消</button><button className="primary" disabled={busy||!name.trim()}>保存名称</button></div></form>}
        {confirm&&<div className="organization-warning" role="alert">删除模板“{template.name}”？已创建的文档会保留。<div className="dialog-actions"><button disabled={busy} onClick={()=>setConfirm(false)}>取消</button><button className="danger" disabled={busy} onClick={()=>void perform(async()=>{await deleteDocumentTemplate(workspace,template.id);setConfirm(false);})}>确认删除模板</button></div></div>}
      </div>}
    </div>
  </Modal>;
}
function TemplatePreview({workspace,template}:{workspace:LocalWorkspace;template:DocumentTemplate}){
  const [store,setStore]=useState<BlockStore>();const [error,setError]=useState('');const host=useRef<HTMLDivElement>(null);
  useEffect(()=>{let item:BlockStore|undefined;try{item=templateStore(template);setStore(item);setError('');}catch{setError('模板预览读取失败，请重新导入有效模板。');}return()=>{if(item)disposeTemplate(item);};},[template.id,template.update]);
  useEffect(()=>{if(!store||!host.current)return;const cleanup=renderEmbedPreview(readDocument(store),host.current,workspace.assets,(id,ws)=>ws===workspace.id?workspace.documents.get(id)?.title:undefined,()=>{});host.current.querySelectorAll('button').forEach(b=>b.disabled=true);return cleanup;},[store,workspace]);
  const surface=store?[...store.doc.getMap('surface').values()].filter(validSurfaceElement):[];
  const bounds=surface.map(s=>elementBounds(s,surface));const x=Math.min(...bounds.map(b=>b.x))-20,y=Math.min(...bounds.map(b=>b.y))-20;
  const width=Math.max(...bounds.map(b=>b.x+b.width))+20-x,height=Math.max(...bounds.map(b=>b.y+b.height))+20-y;
  return <>{error&&<p role="alert">{error}</p>}<section className="trash-preview" aria-label="文档模板正文预览"><div ref={host}/><p className="organization-help">只读摘要预览，新建后可编辑完整内容。</p></section>{!!surface.length&&<section className="trash-preview" aria-label="文档模板白板预览"><h4>白板 · {surface.length} 个元素</h4><SurfacePreview assets={workspace.assets} elements={surface} frame={{id:'template-preview',kind:'frame',x,y,width:Math.max(1,width),height:Math.max(1,height),text:'',fill:'#ffffff'}}/></section>}</>;
}
