import {it,expect} from 'vitest';
import {Y} from '@mydoc/crdt';
import {compareOfficeVersions,officeVersionContent} from '../packages/frontend/core/src/office-version-diff';
import {encodeBinary} from '../packages/frontend/core/src/backup';
import type {OfficeKind} from '../packages/common/nbstore/src/office';
const snapshot=(doc:Y.Doc,kind:OfficeKind)=>({resource:{kind,name:'比较'},state:{data:encodeBinary(Y.encodeStateAsUpdate(doc)),sha256:''},original:null});
it('版本比较忽略 Yjs 操作回执和事务身份，保留文字、格式与页面设置变化',()=>{
 const a=new Y.Doc(),b=new Y.Doc(),text=new Y.Text('中文😀');a.getArray('paragraphs').push([text]);Y.applyUpdate(b,Y.encodeStateAsUpdate(a));expect(compareOfficeVersions(snapshot(a,'word'),snapshot(b,'word')).equal).toBe(true);
 b.getArray<Y.Text>('paragraphs').get(0).format(0,2,{bold:true});b.getMap('pageSettings').set('layout',{width:16000});const diff=compareOfficeVersions(snapshot(a,'word'),snapshot(b,'word'));expect(diff.equal).toBe(false);expect(diff.after).toContain('bold');expect(diff.after).toContain('16000');a.destroy();b.destroy();
});
it('工作簿类型、PPT 多页备注、TXT 编码与 PDF 几何均进入比较',()=>{
 for(const [kind,root]of [['spreadsheet','sheetCellTypes'],['presentation','slides'],['text','textSettings']]as const){const a=new Y.Doc(),b=new Y.Doc();a.getMap(root).set('value','旧');b.getMap(root).set('value','新');expect(compareOfficeVersions(snapshot(a,kind),snapshot(b,kind)).equal).toBe(false);a.destroy();b.destroy();}
 const pdf=new Y.Doc(),annotations=pdf.getArray('annotations');annotations.push([{x:10}]);expect(officeVersionContent(snapshot(pdf,'pdf'))).toContain('10');pdf.destroy();
});
it('回执不会造成内容差异，原件和图片改变可见，长比较有显示配额',()=>{
 const a=new Y.Doc(),b=new Y.Doc();a.getMap('pdfOperations').set('receipt','旧');b.getMap('pdfOperations').set('receipt','新');expect(compareOfficeVersions(snapshot(a,'pdf'),snapshot(b,'pdf')).equal).toBe(true);const old={...snapshot(a,'pdf'),original:{sha256:'original-a'}},now={...snapshot(a,'pdf'),original:{sha256:'original-b'}};expect(compareOfficeVersions(old,now).after).toContain('original-b');a.destroy();b.destroy();
 const x=new Y.Doc(),y=new Y.Doc();x.getArray('sheet').push(Array.from({length:1500},()=>['before']));y.getArray('sheet').push(Array.from({length:1500},()=>['after']));const diff=compareOfficeVersions(snapshot(x,'spreadsheet'),snapshot(y,'spreadsheet'));expect(diff.truncated).toBe(true);expect(diff.before.split('\n').length).toBeLessThanOrEqual(1000);x.destroy();y.destroy();
});
