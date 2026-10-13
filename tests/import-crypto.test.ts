import {createHash} from 'node:crypto';
import {afterEach,expect,it,vi} from 'vitest';
import {zipSync,strToU8} from 'fflate';
import {prepareImport} from '../packages/frontend/core/src/document-import';

const hash=(bytes:Uint8Array|string)=>createHash('sha256').update(bytes).digest('hex');
afterEach(()=>vi.unstubAllGlobals());

it.each(['missing subtle','missing crypto','rejected digest'] as const)(
 'Markdown and Notion fingerprints remain SHA-256 with %s',async scenario=>{
  const digest=vi.fn().mockRejectedValue(new Error('Web Crypto unavailable'));
  vi.stubGlobal('crypto',scenario==='missing crypto'?undefined:scenario==='missing subtle'?{}:{subtle:{digest}});
  const markdown='# 中文文档\n\n![图](pic.png)',image=Uint8Array.from([137,80,78,71,13,10,26,10]);
  const files={'Page.md':strToU8(markdown),'pic.png':image};
  const expected=hash(hash(files['Page.md'])+'\npic.png:'+hash(image));
  for(const kind of ['markdown','notion'] as const){
   const plan=await prepareImport([new File([zipSync(files)],'import.zip')],kind);
   expect(plan.pages[0].error).toBeUndefined();
   expect(plan.pages[0].node?.textContent).toContain('中文文档');
   expect(plan.pages[0].fingerprint).toBe(expected);
   const changed=await prepareImport([new File([zipSync({...files,'pic.png':new Uint8Array([...image,1])})],'import.zip')],kind);
   expect(changed.pages[0].fingerprint).not.toBe(expected);
  }
  if(scenario==='rejected digest')expect(digest).toHaveBeenCalled();
 });
