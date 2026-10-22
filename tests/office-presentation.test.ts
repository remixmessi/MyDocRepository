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
