import {it,expect,vi} from 'vitest';
import {createServer} from 'node:http';
import * as Y from 'yjs';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {zipSync,strToU8} from 'fflate';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {encodePptxPresentation,validatePresentationUpdate} from '../packages/frontend/apps/web/server/office-presentation';
import {exportPptxPresentation} from '../packages/frontend/apps/web/server/office-presentation-export';
import {createOfficeState} from '../packages/frontend/apps/web/server/office-create';
import {presentationSlides,changePresentationSlide,createPresentationObject,setPresentationNotes,presentationSize} from '../packages/common/nbstore/src/office-presentation';
import {submitPresentationOperation} from '../packages/frontend/apps/web/server/office-presentation-operation';
import {structureRevision} from '../packages/common/nbstore/src/spreadsheet-operation';
import {validateOfficeBackupModel} from '../packages/frontend/apps/web/server/office-backup-model';
import {OfficeParserPool} from '../packages/frontend/apps/web/server/office-parser-process';
const legacy=()=>{const doc=new Y.Doc();Y.applyUpdate(doc,createOfficeState('presentation'));return doc;};
it('多页顺序、对象、备注、尺寸经 PPTX 副本往返，原模型不变',async()=>{
 const doc=legacy(),pool=new OfficeParserPool();try{
  doc.getMap('presentation').set('width',9144000);doc.getMap('presentation').set('height',6858000);doc.getArray<Y.Map<unknown>>('objects').get(0).set('text','中文😀 &<>"\n第二段');setPresentationNotes(doc,'first','备注 &<>\n下一行');
  const id=changePresentationSlide(doc,'duplicate','first'),slide=presentationSlides(doc)[1];slide.objects.get(0).set('text','复制页');for(const type of ['rect','ellipse','line'])slide.objects.push([createPresentationObject(type)]);slide.objects.get(1).set('rotation',45);changePresentationSlide(doc,'up',id);
  const before=Y.encodeStateAsUpdate(doc),bytes=await pool.exportPptx(before),parsed=encodePptxPresentation(bytes);try{const slides=presentationSlides(parsed.doc);expect(slides).toHaveLength(2);expect(slides[0].objects.toArray().map(object=>object.get('type'))).toEqual(['text','rect','ellipse','line']);expect(slides[0].objects.get(1).get('rotation')).toBe(45);expect(slides[0].objects.get(0).get('text')).toBe('复制页');expect(slides[1].objects.get(0).get('text')).toBe('中文😀 &<>"\n第二段');expect(slides[1].notes).toBe('备注 &<>\n下一行');expect(presentationSize(parsed.doc)).toEqual({width:9144000,height:6858000});}finally{parsed.doc.destroy();}expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
  const restored=new Y.Doc();try{Y.applyUpdate(restored,validateOfficeBackupModel('presentation',before));expect(presentationSlides(restored)).toHaveLength(2);}finally{restored.destroy();}
 }finally{doc.destroy();pool.close();}
});
it('按演示关系而非文件名选定顺序，拒绝重复及外部关系',()=>{
 const files={'ppt/presentation.xml':'<p:presentation><p:sldIdLst><p:sldId id="256" r:id="b"/><p:sldId id="257" r:id="a"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>','ppt/_rels/presentation.xml.rels':'<Relationships><Relationship Id="a" Type="office/slide" Target="slides/slide1.xml"/><Relationship Id="b" Type="office/slide" Target="slides/slide2.xml"/></Relationships>','ppt/slides/slide1.xml':'<p:sld><p:sp><a:t>一</a:t></p:sp></p:sld>','ppt/slides/slide2.xml':'<p:sld><p:sp><a:t>二</a:t></p:sp></p:sld>'};
 const encode=(data:Record<string,string>)=>zipSync(Object.fromEntries(Object.entries(data).map(([k,v])=>[k,strToU8(v)])));const parsed=encodePptxPresentation(encode(files));try{expect(presentationSlides(parsed.doc).map(slide=>slide.objects.get(0).get('text'))).toEqual(['二','一']);}finally{parsed.doc.destroy();}
 expect(()=>encodePptxPresentation(encode({...files,'ppt/_rels/presentation.xml.rels':files['ppt/_rels/presentation.xml.rels'].replace('Id="b"','Id="b" TargetMode="External"')}))).toThrow('关系');
});
it('幻灯片结构 CAS、回执幂等及失败回滚，通用更新不能绕过',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-slides-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),doc=legacy();try{
  cloud.database.db.prepare('INSERT INTO documents(scope,id,data) VALUES (?,?,?)').run('cloud:space','slides',Buffer.from(Y.encodeStateAsUpdate(doc)));const operation={operationId:crypto.randomUUID(),action:'add' as const,slide:'first',base:structureRevision(doc)},first=submitPresentationOperation(cloud.database,'space','slides','actor',operation);Y.applyUpdate(doc,first.update);expect(presentationSlides(doc)).toHaveLength(2);expect(submitPresentationOperation(cloud.database,'space','slides','actor',operation).selected).toBe(first.selected);expect(()=>submitPresentationOperation(cloud.database,'space','slides','actor',{...operation,operationId:crypto.randomUUID()})).toThrow('已更新');expect(()=>submitPresentationOperation(cloud.database,'space','slides','actor',{...operation,action:'delete'})).toThrow('编号');
  const before=Y.encodeStateAsUpdate(doc),vector=Y.encodeStateVector(doc);changePresentationSlide(doc,'delete',first.selected);expect(()=>validatePresentationUpdate(before,Y.encodeStateAsUpdate(doc,vector))).toThrow('专用接口');
  const restored=new Y.Doc();try{Y.applyUpdate(restored,validateOfficeBackupModel('presentation',first.update));expect(restored.getMap('presentationOperations').size).toBe(0);}finally{restored.destroy();}
  const latest=new Y.Doc();let latestBase='';try{Y.applyUpdate(latest,first.update);latestBase=structureRevision(latest);}finally{latest.destroy();}cloud.database.db.exec("CREATE TRIGGER fail_slide BEFORE UPDATE ON documents BEGIN SELECT RAISE(ABORT,'write failed'); END");expect(()=>submitPresentationOperation(cloud.database,'space','slides','actor',{...operation,operationId:crypto.randomUUID(),base:latestBase})).toThrow('write failed');expect(cloud.database.load('cloud:space','slides')).toEqual(Buffer.from(first.update));
 }finally{doc.destroy();cloud.close();rmSync(dir,{recursive:true,force:true});}
});
it('非法对象、尺寸、备注、顺序、超量与无法导出的图片明确拒绝',()=>{
 const doc=legacy();try{
  const object=doc.getArray<Y.Map<unknown>>('objects').get(0);object.set('type','image');expect(()=>exportPptxPresentation(Y.encodeStateAsUpdate(doc))).toThrow('图片');expect(()=>validateOfficeBackupModel('presentation',Y.encodeStateAsUpdate(doc))).toThrow('图片');object.set('type','text');object.set('fill','url(x)');expect(()=>validatePresentationUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('颜色');object.set('fill','FFFFFF');doc.getArray('slideOrder').push(['missing']);expect(()=>validatePresentationUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('结构');
 }finally{doc.destroy();}
});

it('PPTX 专用结构接口和导出检查权限，固定快照，异步撤权/删除拒绝响应',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-pptx-http-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',value?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(value?{'Content-Type':'application/json'}:{})},body:value?JSON.stringify(value):undefined});
 try{
  const register=async(name:string)=>{const response=await call('/register','POST',{email:name+'@example.com',name,password:'abcdefghijk'});return {cookie:response.headers.get('set-cookie')!.split(';')[0],id:(await response.json()).user.id};};const owner=await register('pptx-owner'),viewer=await register('pptx-viewer'),space=await(await call('/workspaces','POST',{name:'多页'},owner.cookie)).json(),query='?workspace='+space.id,invite=await(await call('/invite'+query,'POST',{email:'pptx-viewer@example.com',role:'viewer'},owner.cookie)).json();await call('/accept','POST',{token:invite.token},viewer.cookie);
  const resource=await(await call('/office-resources'+query,'POST',{kind:'presentation',name:'演示.pptx',mime:'application/vnd.openxmlformats-officedocument.presentationml.presentation',sourceFormat:'pptx'},owner.cookie)).json(),path='/office-export'+query+'&id='+resource.id+'&format=pptx&scope=slides',operationPath='/office-presentation-operation'+query+'&id='+resource.id;
  expect((await call(path.replace('&scope=slides',''),'GET',undefined,owner.cookie)).status).toBe(400);const response=await call(path,'GET',undefined,viewer.cookie);expect(response.status).toBe(200);expect(response.headers.get('content-type')).toContain('presentationml');expect(response.headers.get('x-mydoc-snapshot')).toMatch(/^[a-f0-9]{64}$/);expect(response.headers.get('cache-control')).toBe('no-store');
  const before=Buffer.from(cloud.database.load('cloud:'+space.id,resource.id)!),doc=new Y.Doc();Y.applyUpdate(doc,before);const input={operationId:crypto.randomUUID(),action:'add',slide:'first',base:structureRevision(doc)};doc.destroy();expect((await call(operationPath,'POST',input,viewer.cookie)).status).toBe(403);expect((await call(operationPath,'POST',input,owner.cookie)).status).toBe(200);const repeat=await call(operationPath,'POST',input,owner.cookie);expect((await repeat.json()).repeated).toBe(true);
  const saved=Buffer.from(cloud.database.load('cloud:'+space.id,resource.id)!);const spy=vi.spyOn(OfficeParserPool.prototype,'exportPptx').mockImplementation(async bytes=>{const next=new Y.Doc();try{Y.applyUpdate(next,saved);next.getArray<Y.Map<unknown>>('objects').get(0).set('text','导出后的修改');cloud.database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(Y.encodeStateAsUpdate(next)),'cloud:'+space.id,resource.id);}finally{next.destroy();}return exportPptxPresentation(bytes);});
  try{const fixed=await call(path,'GET',undefined,owner.cookie),parsed=encodePptxPresentation(new Uint8Array(await fixed.arrayBuffer()));try{expect(presentationSlides(parsed.doc)[0].objects.get(0).get('text')).toBe('');}finally{parsed.doc.destroy();}
   spy.mockImplementation(async bytes=>{cloud.database.db.prepare('INSERT INTO cloud_document_permissions VALUES (?,?,?,?)').run(space.id,resource.id,viewer.id,'denied');return exportPptxPresentation(bytes);});expect((await call(path,'GET',undefined,viewer.cookie)).status).toBe(404);
   spy.mockImplementation(async bytes=>{cloud.database.db.prepare('UPDATE cloud_office_resources SET deleted=1 WHERE workspace=? AND id=?').run(space.id,resource.id);return exportPptxPresentation(bytes);});expect((await call(path,'GET',undefined,owner.cookie)).status).toBe(404);
  }finally{spy.mockRestore();}
 }finally{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('演示对象字段独立修订：同字段竞争与 ABA 拒绝，异字段保留，重试幂等且写失败回滚',async()=>{
 const {presentationObjectIdentity,presentationFieldRevision}=await import('../packages/common/nbstore/src/presentation-field');
 const dir=mkdtempSync(join(tmpdir(),'mydoc-pptx-fields-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),doc=legacy();try{
  cloud.database.append('cloud:space','slides',Buffer.from(Y.encodeStateAsUpdate(doc)));const object=doc.getArray<Y.Map<unknown>>('objects').get(0),identity=presentationObjectIdentity(object);
  const text={action:'setField' as const,operationId:crypto.randomUUID(),slide:'first',object:identity,field:'text' as const,value:'甲确认😀',base:presentationFieldRevision(object,'text')},x={...text,operationId:crypto.randomUUID(),field:'x' as const,value:0.2,base:presentationFieldRevision(object,'x')};
  const result=submitPresentationOperation(cloud.database,'space','slides','actor',text);expect(submitPresentationOperation(cloud.database,'space','slides','actor',text).repeated).toBe(true);expect(()=>submitPresentationOperation(cloud.database,'space','slides','other',text)).toThrow('编号');
  expect(()=>submitPresentationOperation(cloud.database,'space','slides','other',{...text,operationId:crypto.randomUUID(),value:'乙旧草稿'})).toThrow('本人输入已保留');Y.applyUpdate(doc,result.update);
  Y.applyUpdate(doc,submitPresentationOperation(cloud.database,'space','slides','other',x).update);expect(object.get('text')).toBe('甲确认😀');expect(object.get('x')).toBe(0.2);
  const revert={...text,operationId:crypto.randomUUID(),value:'',base:presentationFieldRevision(object,'text')};Y.applyUpdate(doc,submitPresentationOperation(cloud.database,'space','slides','actor',revert).update);expect(()=>submitPresentationOperation(cloud.database,'space','slides','other',{...text,operationId:crypto.randomUUID()})).toThrow('已被修改');
  const before=Y.encodeStateAsUpdate(doc),vector=Y.encodeStateVector(doc);object.set('x',0.7);expect(()=>validatePresentationUpdate(before,Y.encodeStateAsUpdate(doc,vector))).toThrow('专用接口');
  const clean=new Y.Doc();try{Y.applyUpdate(clean,before);const live=clean.getArray<Y.Map<unknown>>('objects').get(0),operation={...x,operationId:crypto.randomUUID(),base:presentationFieldRevision(live,'x'),value:0.9};cloud.database.db.exec("CREATE TRIGGER fail_field BEFORE UPDATE ON documents BEGIN SELECT RAISE(ABORT,'field write failed'); END");const saved=cloud.database.load('cloud:space','slides');expect(()=>submitPresentationOperation(cloud.database,'space','slides','actor',operation)).toThrow('field write failed');expect(cloud.database.load('cloud:space','slides')).toEqual(saved);cloud.database.db.exec('DROP TRIGGER fail_field');clean.getArray('objects').delete(0,1);cloud.database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(Y.encodeStateAsUpdate(clean)),'cloud:space','slides');expect(()=>submitPresentationOperation(cloud.database,'space','slides','actor',operation)).toThrow('已被修改或删除');}finally{clean.destroy();}
 }finally{doc.destroy();cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('演示对象新增/删除走语义接口：幂等重试、删除与字段修改竞争、图片归属及通用更新绕过拒绝',async()=>{
 const {presentationObjectIdentity,presentationFieldRevision,presentationObjectRevision}=await import('../packages/common/nbstore/src/presentation-field');
 const dir=mkdtempSync(join(tmpdir(),'mydoc-pptx-objects-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),doc=legacy();try{
  cloud.database.append('cloud:space','slides',Buffer.from(Y.encodeStateAsUpdate(doc)));const call=(user:string,input:Parameters<typeof submitPresentationOperation>[4])=>submitPresentationOperation(cloud.database,'space','slides',user,input);
  const add={action:'addObject' as const,operationId:crypto.randomUUID(),slide:'first',type:'rect' as const,geometry:{x:0.3,y:0.4}},added=call('actor',add);Y.applyUpdate(doc,added.update);
  const objects=doc.getArray<Y.Map<unknown>>('objects');expect(objects.toArray().map(object=>object.get('type'))).toEqual(['text','rect']);expect(objects.get(1).get('x')).toBe(0.3);expect(added.object).toBe(presentationObjectIdentity(objects.get(1)));
  const retry=call('actor',add);expect(retry.repeated).toBe(true);expect(retry.object).toBe(added.object);Y.applyUpdate(doc,retry.update);expect(objects.length).toBe(2);
  expect(()=>call('actor',{...add,type:'ellipse'})).toThrow('编号');expect(()=>call('actor',{...add,operationId:crypto.randomUUID(),slide:'missing'})).toThrow('幻灯片已不存在');
  for(const bad of [{...add,type:'chart'},{...add,type:'image'},{...add,assetId:'asset'},{...add,geometry:{x:'1'}},{...add,geometry:{width:-1}},{...add,geometry:{rotation:5}},{...add,extra:true}])expect(()=>call('actor',{...bad,operationId:crypto.randomUUID()} as never)).toThrow('无效');
  // Deleting an object another member just edited is rejected rather than discarding that edit.
  const rect=objects.get(1),staleBase=presentationObjectRevision(rect);Y.applyUpdate(doc,call('other',{action:'setField',operationId:crypto.randomUUID(),slide:'first',object:added.object!,field:'fill',value:'FF0000',base:presentationFieldRevision(rect,'fill')}).update);
  expect(()=>call('actor',{action:'deleteObject',operationId:crypto.randomUUID(),slide:'first',object:added.object!,base:staleBase})).toThrow('已被其他成员修改');expect(objects.length).toBe(2);
  const remove={action:'deleteObject' as const,operationId:crypto.randomUUID(),slide:'first',object:added.object!,base:presentationObjectRevision(rect)},removed=call('actor',remove);Y.applyUpdate(doc,removed.update);expect(objects.length).toBe(1);
  expect(doc.getMap('presentationFields').has(added.object+':fill')).toBe(false);expect(call('actor',remove).repeated).toBe(true);expect(()=>call('actor',{...remove,operationId:crypto.randomUUID()})).toThrow('已被其他成员删除');
  // A stale field edit on the deleted object cannot resurrect it.
  expect(()=>call('other',{action:'setField',operationId:crypto.randomUUID(),slide:'first',object:added.object!,field:'x',value:0.5,base:'stale'})).toThrow('已被修改或删除');
  const bypass=(change:(value:Y.Doc)=>void)=>{const copy=new Y.Doc();try{Y.applyUpdate(copy,cloud.database.load('cloud:space','slides')!);const before=Y.encodeStateAsUpdate(copy),vector=Y.encodeStateVector(copy);change(copy);return ()=>validatePresentationUpdate(before,Y.encodeStateAsUpdate(copy,vector));}finally{copy.destroy();}};
  expect(bypass(value=>value.getArray<Y.Map<unknown>>('objects').push([createPresentationObject('rect')]))).toThrow('专用接口');
  expect(bypass(value=>value.getArray('objects').delete(0,1))).toThrow('专用接口');
  Y.applyUpdate(doc,call('actor',{...add,operationId:crypto.randomUUID()}).update);cloud.database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(Y.encodeStateAsUpdate(doc)),'cloud:space','slides');
  expect(bypass(value=>{const array=value.getArray<Y.Map<unknown>>('objects'),moved=array.get(0).clone();array.delete(0,1);array.push([moved]);})).toThrow('专用接口');
  expect(bypass(value=>value.getArray<Y.Map<unknown>>('objects').get(0).set('type','rect'))).toThrow('专用接口');
  expect(bypass(value=>setPresentationNotes(value,'first','允许的备注'))).not.toThrow();
  // Images must belong to this file; failed writes roll back without consuming the operation id.
  const image={action:'addObject' as const,operationId:crypto.randomUUID(),slide:'first',type:'image' as const,assetId:'foreign-asset',geometry:{width:0.4,height:0.3}};
  cloud.database.db.prepare('INSERT INTO cloud_asset_owners VALUES (?,?,?,?)').run('space','foreign-asset','other-file','actor');const saved=cloud.database.load('cloud:space','slides');
  expect(()=>call('actor',image)).toThrow('不属于当前办公文件');expect(cloud.database.load('cloud:space','slides')).toEqual(saved);
  cloud.database.db.exec("CREATE TRIGGER fail_object BEFORE UPDATE ON documents BEGIN SELECT RAISE(ABORT,'object write failed'); END");expect(()=>call('actor',{...add,operationId:crypto.randomUUID()})).toThrow('object write failed');expect(cloud.database.load('cloud:space','slides')).toEqual(saved);
 }finally{doc.destroy();cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('PPTX 占位符按幻灯片→版式→母版继承位置尺寸与填充/边框，按 idx 优先匹配，外部版式关系拒绝',()=>{
 const sp=(ph:string,spPr='',text='')=>`<p:sp><p:nvSpPr><p:cNvPr id="1" name="x"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr>${spPr}</p:spPr><p:txBody><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`,xfrm=(x:number,y:number,w:number,h:number)=>`<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`;
 const files:Record<string,string>={
  'ppt/presentation.xml':'<p:presentation><p:sldIdLst><p:sldId id="256" r:id="s1"/></p:sldIdLst><p:sldSz cx="1000" cy="500"/></p:presentation>','ppt/_rels/presentation.xml.rels':'<Relationships><Relationship Id="s1" Type="office/slide" Target="slides/slide1.xml"/></Relationships>',
  'ppt/slides/slide1.xml':'<p:sld><p:cSld><p:spTree>'+sp('<p:ph type="title"/>','','标题')+sp('<p:ph idx="2"/>','','正文')+sp('<p:ph type="body" idx="9"/>',xfrm(10,20,30,40)+'<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>','自带')+'</p:spTree></p:cSld></p:sld>',
  'ppt/slides/_rels/slide1.xml.rels':'<Relationships><Relationship Id="l" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>',
  'ppt/slideLayouts/slideLayout1.xml':'<p:sldLayout><p:cSld><p:spTree>'+sp('<p:ph type="title"/>',xfrm(100,50,800,100))+sp('<p:ph type="body" idx="1"/>',xfrm(1,1,1,1))+sp('<p:ph type="body" idx="2"/>','')+'</p:spTree></p:cSld></p:sldLayout>',
  'ppt/slideLayouts/_rels/slideLayout1.xml.rels':'<Relationships><Relationship Id="m" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>',
  'ppt/slideMasters/slideMaster1.xml':'<p:sldMaster><p:cSld><p:spTree>'+sp('<p:ph type="title"/>',xfrm(0,0,1,1)+'<a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln>')+sp('<p:ph type="body" idx="1"/>',xfrm(100,200,800,250)+'<a:noFill/>')+'</p:spTree></p:cSld></p:sldMaster>',
 };
 const encode=(data:Record<string,string>)=>zipSync(Object.fromEntries(Object.entries(data).map(([k,v])=>[k,strToU8(v)])));const parsed=encodePptxPresentation(encode(files));
 try{const [title,body,own]=presentationSlides(parsed.doc)[0].objects.toArray().map(object=>object.toJSON());
  expect(title).toMatchObject({text:'标题',x:0.1,y:0.1,width:0.8,height:0.2,fill:'112233',stroke:'445566'});
  // idx 2 has no geometry on the layout, so it falls through to the master body placeholder.
  expect(body).toMatchObject({text:'正文',x:0.1,y:0.4,width:0.8,height:0.5,fill:'FFFFFF'});
  expect(own).toMatchObject({text:'自带',x:0.01,y:0.04,width:0.03,height:0.08,fill:'00FF00'});
 }finally{parsed.doc.destroy();}
 expect(()=>encodePptxPresentation(encode({...files,'ppt/slides/_rels/slide1.xml.rels':files['ppt/slides/_rels/slide1.xml.rels'].replace('/>',' TargetMode="External"/>')}))).toThrow('外部版式');
});

it('对象讨论导出为 PPTX 幻灯片批注：位于对象右上角、回复按顺序、已解决标注，对象已删的讨论留在原幻灯片',async()=>{
 const {readOfficePackage,readOfficeXml}=await import('../packages/frontend/apps/web/server/office-package');
 const {presentationObjectIdentity}=await import('../packages/common/nbstore/src/presentation-field');
 const {COMMENT_UNIT}=await import('../packages/frontend/apps/web/server/office-presentation-export');
 const doc=new Y.Doc();Y.applyUpdate(doc,createOfficeState('presentation'));changePresentationSlide(doc,'add',presentationSlides(doc)[0].id);
 const [first,second]=presentationSlides(doc),object=createPresentationObject('rect');object.set('x',0.25);object.set('y',0.5);object.set('width',0.25);second.objects.push([object]);const gone=createPresentationObject('rect');first.objects.push([gone]);
 const size=presentationSize(doc),message=(id:string,name:string,body:string,created:number)=>({id,author:'u'+id,authorName:name,body,created}),comments=doc.getMap('officeComments');
 comments.set('a',{id:'a',resolved:true,updated:3,anchor:{kind:'presentation',slide:second.id,object:presentationObjectIdentity(object)},messages:[message('m1','王小明','颜色 <太深> & 偏暗',1),message('m2','Lisa Chen','已调整',2)]});
 comments.set('b',{id:'b',resolved:false,updated:4,anchor:{kind:'presentation',slide:first.id,object:presentationObjectIdentity(gone)},messages:[message('m3','王小明','对象删了',4)]});first.objects.delete(first.objects.toArray().indexOf(gone),1);
 try{
  const files=readOfficePackage(exportPptxPresentation(Y.encodeStateAsUpdate(doc))),authors=readOfficeXml(files['ppt/commentAuthors.xml']),one=readOfficeXml(files['ppt/comments/comment1.xml']),two=readOfficeXml(files['ppt/comments/comment2.xml']);
  expect(readOfficeXml(files['ppt/_rels/presentation.xml.rels'])).toContain('commentAuthors.xml');expect(readOfficeXml(files['ppt/slides/_rels/slide2.xml.rels'])).toContain('../comments/comment2.xml');expect(readOfficeXml(files['[Content_Types].xml'])).toContain('presentationml.comments+xml');
  expect(authors).toContain('name="王小明"');expect(authors).toContain('name="Lisa Chen"');expect(authors).toMatch(/name="王小明" initials="王" lastIdx="2"/);
  expect([...two.matchAll(/<p:text>([^<]*)<\/p:text>/g)].map(m=>m[1])).toEqual(['[已解决] 颜色 &lt;太深&gt; &amp; 偏暗','回复：已调整']);
  expect(two).toContain(`<p:pos x="${Math.round(0.5*size.width/COMMENT_UNIT)}" y="${Math.round(0.5*size.height/COMMENT_UNIT)}"/>`);
  expect(one).toContain('<p:text>对象删了</p:text>');expect(one).toContain('<p:pos x="0" y="0"/>');
  const reopened=encodePptxPresentation(exportPptxPresentation(Y.encodeStateAsUpdate(doc)));try{expect(presentationSlides(reopened.doc)).toHaveLength(2);}finally{reopened.doc.destroy();}
 }finally{doc.destroy();}
});

it('文字格式、新形状、层次、背景与切换效果经 PPTX 副本往返，按层次顺序写出',async()=>{
 const {presentationPaintOrder}=await import('../packages/common/nbstore/src/presentation-field');
 const {setPresentationSlideSetting}=await import('../packages/common/nbstore/src/office-presentation');
 const {readOfficePackage,readOfficeXml}=await import('../packages/frontend/apps/web/server/office-package');
 const doc=legacy();try{
  const text=doc.getArray<Y.Map<unknown>>('objects').get(0);text.set('text','标题\n第二行');for(const [key,value]of Object.entries({fontSize:32,bold:true,italic:true,underline:true,color:'C00000',align:'center'}))text.set(key,value);
  const shapes=['roundRect','triangle','diamond','rightArrow','star5'].map(type=>{const object=createPresentationObject(type);object.set('fill','4472C4');doc.getArray<Y.Map<unknown>>('objects').push([object]);return object;});
  // The text box is painted on top of every shape.
  shapes.forEach((object,i)=>object.set('z',i));text.set('z',5);
  setPresentationSlideSetting(doc,['first'],'background','fff2cc');setPresentationSlideSetting(doc,['first'],'transition','push');
  const id=changePresentationSlide(doc,'add','first');expect(presentationSlides(doc)[1]).toMatchObject({background:'FFF2CC',transition:'push'});setPresentationSlideSetting(doc,[id],'transition','zoom');setPresentationSlideSetting(doc,[id],'background',undefined);
  const update=Y.encodeStateAsUpdate(doc);validatePresentationUpdate(null,update);const bytes=exportPptxPresentation(update),files=readOfficePackage(bytes),first=readOfficeXml(files['ppt/slides/slide1.xml']);
  expect(first).toContain('<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>');expect(first).toContain('<p:transition spd="med"><p:push dir="l"/></p:transition>');expect(first).toContain('<a:pPr algn="ctr"/>');expect(first).toMatch(/sz="3200" b="1" i="1" u="sng"><a:solidFill><a:srgbClr val="C00000"\/>/);
  expect(first.indexOf('prst="star5"')).toBeLessThan(first.indexOf('txBox="1"'));expect(readOfficeXml(files['ppt/slides/slide2.xml'])).toContain('<p:zoom/>');expect(readOfficeXml(files['ppt/slides/slide2.xml'])).not.toContain('<p:bg>');
  const parsed=encodePptxPresentation(bytes);try{
   const [one,two]=presentationSlides(parsed.doc);expect(one).toMatchObject({background:'FFF2CC',transition:'push'});expect(two.transition).toBe('zoom');expect(two.background).toBeUndefined();
   const objects=presentationPaintOrder(one.objects.toArray()).map(entry=>entry.object.toJSON());expect(objects.map(object=>object.type)).toEqual(['roundRect','triangle','diamond','rightArrow','star5','text']);
   expect(objects[5]).toMatchObject({text:'标题\n第二行',fontSize:32,bold:true,italic:true,underline:true,color:'C00000',align:'center'});expect(objects[0].fill).toBe('4472C4');
  }finally{parsed.doc.destroy();}
  const restored=new Y.Doc();try{Y.applyUpdate(restored,validateOfficeBackupModel('presentation',update));expect(presentationSlides(restored)[0].background).toBe('FFF2CC');}finally{restored.destroy();}
 }finally{doc.destroy();}
});

it('导入：较新切换效果取 mc:Fallback，未知效果降为淡出，背景按幻灯片→版式→母版继承；空形状保留几何',()=>{
 const encode=(data:Record<string,string>)=>zipSync(Object.fromEntries(Object.entries(data).map(([k,v])=>[k,strToU8(v)])));
 const shape=(prst:string)=>`<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:prstGeom prst="${prst}"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>`;
 const files:Record<string,string>={
  'ppt/presentation.xml':'<p:presentation><p:sldIdLst><p:sldId id="256" r:id="a"/><p:sldId id="257" r:id="b"/><p:sldId id="258" r:id="c"/></p:sldIdLst><p:sldSz cx="1000" cy="500"/></p:presentation>',
  'ppt/_rels/presentation.xml.rels':'<Relationships><Relationship Id="a" Type="office/slide" Target="slides/slide1.xml"/><Relationship Id="b" Type="office/slide" Target="slides/slide2.xml"/><Relationship Id="c" Type="office/slide" Target="slides/slide3.xml"/></Relationships>',
  'ppt/slides/slide1.xml':'<p:sld><p:cSld><p:spTree>'+shape('triangle')+shape('cloud')+'</p:spTree></p:cSld><mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="slow" p14:dur="2000"><p14:vortex dir="r"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:wipe/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>',
  'ppt/slides/slide2.xml':'<p:sld><p:cSld><p:spTree/></p:cSld><p:transition><p14:prism/></p:transition></p:sld>',
  'ppt/slides/slide3.xml':'<p:sld><p:cSld><p:bg><p:bgPr><a:gradFill/></p:bgPr></p:bg><p:spTree/></p:cSld></p:sld>',
  'ppt/slides/_rels/slide1.xml.rels':'<Relationships><Relationship Id="l" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>',
  'ppt/slides/_rels/slide2.xml.rels':'<Relationships><Relationship Id="l" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>',
  'ppt/slides/_rels/slide3.xml.rels':'<Relationships><Relationship Id="l" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>',
  'ppt/slideLayouts/slideLayout1.xml':'<p:sldLayout><p:cSld><p:spTree/></p:cSld></p:sldLayout>',
  'ppt/slideLayouts/_rels/slideLayout1.xml.rels':'<Relationships><Relationship Id="m" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>',
  'ppt/slideMasters/slideMaster1.xml':'<p:sldMaster><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="1f3864"/></a:solidFill></p:bgPr></p:bg><p:spTree/></p:cSld></p:sldMaster>',
 };
 const parsed=encodePptxPresentation(encode(files));try{
  const [one,two,three]=presentationSlides(parsed.doc);expect(one.transition).toBe('wipe');expect(two.transition).toBe('fade');expect(three.transition).toBeUndefined();
  expect(one.background).toBe('1F3864');expect(two.background).toBe('1F3864');expect(three.background).toBeUndefined();
  expect(one.objects.toArray().map(object=>[object.get('type'),object.get('fill')])).toEqual([['triangle','FF0000']]);
 }finally{parsed.doc.destroy();}
});

it('层次调整走语义接口：按层次修订检测竞争，新对象置顶，通用更新不能改层次；文字格式只用于文本框',async()=>{
 const {presentationObjectIdentity,presentationFieldRevision,presentationArrangeRevision,presentationPaintOrder}=await import('../packages/common/nbstore/src/presentation-field');
 const {setPresentationSlideSetting}=await import('../packages/common/nbstore/src/office-presentation');
 const dir=mkdtempSync(join(tmpdir(),'mydoc-pptx-arrange-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),doc=legacy();try{
  cloud.database.append('cloud:space','slides',Buffer.from(Y.encodeStateAsUpdate(doc)));const call=(input:Parameters<typeof submitPresentationOperation>[4])=>{const result=submitPresentationOperation(cloud.database,'space','slides','actor',input);Y.applyUpdate(doc,result.update);return result;};
  for(const type of ['rect','ellipse'] as const)call({action:'addObject',operationId:crypto.randomUUID(),slide:'first',type});
  const objects=doc.getArray<Y.Map<unknown>>('objects'),ids=objects.toArray().map(object=>presentationObjectIdentity(object)),painted=()=>presentationPaintOrder(objects.toArray()).map(entry=>presentationObjectIdentity(entry.object));
  expect(painted()).toEqual(ids);
  const base=presentationArrangeRevision(objects.toArray()),back={action:'arrangeObject' as const,operationId:crypto.randomUUID(),slide:'first',object:ids[2],to:'back' as const,base};call(back);expect(painted()).toEqual([ids[2],ids[0],ids[1]]);
  expect(submitPresentationOperation(cloud.database,'space','slides','actor',back).repeated).toBe(true);expect(()=>call({...back,operationId:crypto.randomUUID(),to:'front'})).toThrow('层次已被其他成员调整');
  call({action:'arrangeObject',operationId:crypto.randomUUID(),slide:'first',object:ids[0],to:'forward',base:presentationArrangeRevision(objects.toArray())});expect(painted()).toEqual([ids[2],ids[1],ids[0]]);
  // Once layers exist, a new object goes on top and ranks stay dense.
  const added=call({action:'addObject',operationId:crypto.randomUUID(),slide:'first',type:'star5'});expect(painted()).toEqual([ids[2],ids[1],ids[0],added.object]);expect(objects.toArray().map(object=>object.get('z')).sort()).toEqual([0,1,2,3]);
  expect(()=>call({action:'arrangeObject',operationId:crypto.randomUUID(),slide:'first',object:ids[0],to:'sideways' as never,base:''})).toThrow('无效');
  const before=Y.encodeStateAsUpdate(doc),vector=Y.encodeStateVector(doc);objects.get(0).set('z',9);expect(()=>validatePresentationUpdate(before,Y.encodeStateAsUpdate(doc,vector))).toThrow('专用接口');objects.get(0).set('z',2);
  const rect=objects.get(1);expect(()=>call({action:'setField',operationId:crypto.randomUUID(),slide:'first',object:ids[1],field:'bold',value:true,base:presentationFieldRevision(rect,'bold')})).toThrow('不是文本框');
  const text=objects.get(0);expect(()=>call({action:'setField',operationId:crypto.randomUUID(),slide:'first',object:ids[0],field:'fontSize',value:2,base:presentationFieldRevision(text,'fontSize')})).toThrow('无效');
  call({action:'setField',operationId:crypto.randomUUID(),slide:'first',object:ids[0],field:'bold',value:true,base:presentationFieldRevision(text,'bold')});call({action:'setField',operationId:crypto.randomUUID(),slide:'first',object:ids[0],field:'align',value:'right',base:presentationFieldRevision(text,'align')});expect(text.get('bold')).toBe(true);expect(text.get('align')).toBe('right');
  // Background and transition are ordinary slide settings; invalid values are rejected.
  const snapshot=Y.encodeStateAsUpdate(doc),at=Y.encodeStateVector(doc);setPresentationSlideSetting(doc,['first'],'transition','fade');expect(()=>validatePresentationUpdate(snapshot,Y.encodeStateAsUpdate(doc,at))).not.toThrow();
  doc.getMap('presentation').set('background','red');expect(()=>validatePresentationUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('演示设置无效');doc.getMap('presentation').delete('background');
  rect.set('fontSize',20);expect(()=>validatePresentationUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('文字格式');
 }finally{doc.destroy();cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('进入动画：按顺序与触发方式分步（单击开新步、同时共享开始、之后顺延），PPTX 写入 p:timing 主序列并能导入还原',async()=>{
 const {presentationBuilds,PRESENTATION_ANIMATION_MS}=await import('../packages/common/nbstore/src/presentation-field');
 const {readOfficePackage,readOfficeXml}=await import('../packages/frontend/apps/web/server/office-package');
 const doc=legacy();try{
  const objects=doc.getArray<Y.Map<unknown>>('objects'),title=objects.get(0);title.set('text','标题');
  const shapes=['rect','ellipse','star5','triangle'].map(type=>{const object=createPresentationObject(type);objects.push([object]);return object;});
  // Order: title (with, on entry) → rect (click) → ellipse (with) → star (after) → triangle (click).
  const set=(object:Y.Map<unknown>,effect:string,trigger:string,order:number)=>{object.set('animation',effect);object.set('animationTrigger',trigger);object.set('animationOrder',order);};
  set(title,'fade','with',1);set(shapes[0],'flyBottom','click',2);set(shapes[1],'zoom','with',3);set(shapes[2],'wipe','after',4);set(shapes[3],'appear','click',5);
  expect(presentationBuilds(objects.toArray()).map(build=>[build.index,build.step,build.delay])).toEqual([[0,0,0],[1,1,0],[2,1,0],[3,1,PRESENTATION_ANIMATION_MS],[4,2,0]]);
  const update=Y.encodeStateAsUpdate(doc);validatePresentationUpdate(null,update);const bytes=exportPptxPresentation(update),slide=readOfficeXml(readOfficePackage(bytes)['ppt/slides/slide1.xml']);
  expect(slide).toContain('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">');expect(slide).toContain('presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="withEffect"');expect(slide).toContain('presetID="2" presetClass="entr" presetSubtype="4" fill="hold" grpId="0" nodeType="clickEffect"');expect(slide).toContain('nodeType="afterEffect"');expect(slide).toContain('<p:attrName>style.visibility</p:attrName>');expect(slide).toContain('<p:bldP spid="2" grpId="0" animBg="1"/>');expect(slide.indexOf('</p:clrMapOvr>')).toBeLessThan(slide.indexOf('<p:timing>'));
  const parsed=encodePptxPresentation(bytes);try{const back=presentationSlides(parsed.doc)[0].objects.toArray();expect(presentationBuilds(back).map(build=>[build.index,build.effect,build.trigger,build.step,build.delay])).toEqual([[0,'fade','with',0,0],[1,'flyBottom','click',1,0],[2,'zoom','with',1,0],[3,'wipe','after',1,PRESENTATION_ANIMATION_MS],[4,'appear','click',2,0]]);}finally{parsed.doc.destroy();}
  title.set('animation','spin');expect(()=>validatePresentationUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('动画');
 }finally{doc.destroy();}
});

it('进入动画：“之后”等待本步已开始的全部动画结束（出现为瞬时），连续“之后”在 PPTX 中各成一组，进入页即播的步骤以 onBegin 启动',async()=>{
 const {presentationBuilds,PRESENTATION_ANIMATION_MS}=await import('../packages/common/nbstore/src/presentation-field');
 const {readOfficePackage,readOfficeXml}=await import('../packages/frontend/apps/web/server/office-package');
 const doc=legacy();try{
  const objects=doc.getArray<Y.Map<unknown>>('objects');objects.get(0).set('text','标题');
  for(const type of ['rect','ellipse','star5'])objects.push([createPresentationObject(type)]);
  const set=(index:number,effect:string,trigger:string,order:number)=>{const object=objects.get(index);object.set('animation',effect);object.set('animationTrigger',trigger);object.set('animationOrder',order);};
  // Step 0: appear (instant) → after starts at once; then with; then two afters in a row each wait for the longest so far.
  set(0,'appear','after',1);set(1,'fade','after',2);set(2,'zoom','with',3);set(3,'wipe','after',4);
  expect(presentationBuilds(objects.toArray()).map(build=>[build.index,build.step,build.delay])).toEqual([[0,0,0],[1,0,0],[2,0,0],[3,0,PRESENTATION_ANIMATION_MS]]);
  set(3,'wipe','click',4);set(2,'zoom','after',3);
  expect(presentationBuilds(objects.toArray()).map(build=>[build.index,build.delay])).toEqual([[0,0],[1,0],[2,PRESENTATION_ANIMATION_MS],[3,0]]);
  const bytes=exportPptxPresentation(Y.encodeStateAsUpdate(doc)),slide=readOfficeXml(readOfficePackage(bytes)['ppt/slides/slide1.xml']);
  expect(slide.match(/nodeType="afterEffect"/g)).toHaveLength(3);expect(slide).toContain('<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>');
  const parsed=encodePptxPresentation(bytes);try{expect(presentationBuilds(presentationSlides(parsed.doc)[0].objects.toArray()).map(build=>[build.index,build.trigger,build.step,build.delay])).toEqual([[0,'after',0,0],[1,'after',0,0],[2,'after',0,PRESENTATION_ANIMATION_MS],[3,'click',1,0]]);}finally{parsed.doc.destroy();}
 }finally{doc.destroy();}
});
