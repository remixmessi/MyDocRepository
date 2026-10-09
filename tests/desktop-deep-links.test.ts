import {describe,it,expect} from 'vitest';
import {parseDocumentLink,documentLink} from '../packages/common/nbstore/src/deep-link';
import {DocumentLinkQueue,commandLineLinks} from '../packages/frontend/apps/desktop/src/deep-links';
import {emptyWorkbench,openLocation,navigateHistory,restoreWorkbench} from '../packages/frontend/core/src/workbench';
describe('desktop document links',()=>{
 it('round trips page, block and board destinations and rejects malformed URLs',()=>{
  for(const target of [{workspace:'local',doc:'abc',mode:'page' as const},{workspace:'w1',doc:'d1',mode:'page' as const,block:'code-1'},{workspace:'w1',doc:'d1',mode:'board' as const,frame:'frame-1'}])expect(parseDocumentLink(documentLink(target))).toEqual(target);
  for(const url of ['https://document?workspace=w&doc=d','mydoc://user@document?workspace=w&doc=d','mydoc://document/path?workspace=w&doc=d','mydoc://document?workspace=w&doc=d&doc=e','mydoc://document?workspace=w&doc=d&frame=f&block=b','mydoc://document?workspace=w&doc=d&mode=bad','mydoc://document?workspace=w&doc=d&block=','mydoc://document?workspace=w&doc=d&path=/etc/passwd','mydoc://document?workspace=w&doc=%0A','mydoc://document?workspace=w&doc=d#x'])expect(parseDocumentLink(url)).toBeNull();
 });
 it('bounds and deduplicates queued requests, consumes only matching acknowledgements',()=>{
  const queue=new DocumentLinkQueue(),url=documentLink({workspace:'local',doc:'d',mode:'page'});queue.enqueue(url);queue.enqueue(url);expect(queue.pending).toBe(1);queue.acknowledge('invalid');expect(queue.pending).toBe(1);queue.acknowledge(queue.take()!.id);expect(queue.pending).toBe(0);queue.enqueue('mydoc://broken');expect(queue.take()?.error).toContain('无效');for(let i=0;i<40;i++)queue.enqueue(documentLink({workspace:'local',doc:'d'+i,mode:'page'}));expect(queue.pending).toBe(32);expect(commandLineLinks(['app','--flag',url,'a.md'])).toEqual([url]);
 });
 it('records distinct block destinations, restores and navigates their modes',()=>{
  let state=openLocation(emptyWorkbench(),{doc:'d',mode:'board',frame:'f'});state=openLocation(state,{doc:'d',mode:'page',block:'b'});state=openLocation(state,{doc:'d',mode:'page',block:'c'});expect(state.left.history?.entries).toHaveLength(3);state=restoreWorkbench(state,new Set(['d']));expect(state.left.tabs[0].blockId).toBe('c');state=navigateHistory(state,'left',1);expect(state.left.tabs[0].blockId).toBe('b');state=navigateHistory(state,'left',0);expect(state.left.tabs[0].mode).toBe('board');expect(state.left.tabs[0].frameId).toBe('f');
 });
});
