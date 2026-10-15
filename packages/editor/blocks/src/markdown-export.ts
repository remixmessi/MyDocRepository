import {zipSync,strToU8} from 'fflate';
import {Fragment,type Node} from 'prosemirror-model';
import type {WorkspaceAssets} from '@mydoc/nbstore';
import {markdownSerializer} from './document-format';
const safeName=(value:string)=>value.replace(/[\\/:*?"<>|\u0000-\u001f]/g,'_').replace(/[. ]+$/g,'').slice(0,120)||'未命名文件';
/** Resolve only actual media nodes, so examples in code blocks stay untouched. */
export async function markdownExport(doc:Node,title:string,assets?:Pick<WorkspaceAssets,'get'>,original?:string){
 const ids=new Set<string>();doc.descendants(node=>{if(node.type.name==='attachment')ids.add(node.attrs.assetId);if(node.type.name==='image'&&node.attrs.src.startsWith('asset:'))ids.add(node.attrs.src.slice(6));});
 const filename=safeName(title||'无标题文档');
 if(!ids.size)return {name:filename+'.md',type:'text/markdown;charset=utf-8',bytes:strToU8(original??markdownSerializer.serialize(doc))};
 if(!assets)throw new Error('无法读取附件，请重新打开文档后重试');
 const files:Record<string,Uint8Array>=Object.create(null),paths=new Map<string,string>();let index=0,total=0;
 for(const id of ids){const record=await assets.get(id);if(!record)throw new Error('附件缺失，无法完整导出：'+id);total+=record.bytes.length;if(total>200*1024*1024)throw new Error('附件总大小超过 200 MB，请分批导出');const path='assets/'+(++index)+'-'+safeName(record.name);files[path]=record.bytes;paths.set(id,'./assets/'+encodeURIComponent(path.slice(7)));}
 const convert=(node:Node):Node=>{
  if(node.type.name==='attachment')return doc.type.schema.text(node.attrs.name||'附件',[doc.type.schema.marks.link.create({href:paths.get(node.attrs.assetId)})]);
  if(node.type.name==='image'&&node.attrs.src.startsWith('asset:'))return node.type.create({...node.attrs,src:paths.get(node.attrs.src.slice(6))},node.content,node.marks);
  if(node.isLeaf)return node;const children:Node[]=[];node.forEach(child=>children.push(convert(child)));return node.copy(Fragment.fromArray(children));
 };
 files[filename+'.md']=strToU8(markdownSerializer.serialize(convert(doc)));
 return {name:filename+'.zip',type:'application/zip',bytes:zipSync(files,{level:0})};
}
