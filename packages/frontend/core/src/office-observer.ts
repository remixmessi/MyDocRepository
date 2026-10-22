import type {Y} from '@mydoc/crdt';

/** Observe the whole document so nested edits and structural operations both refresh. */
export function observeOfficeDocument(doc:Y.Doc,notify:()=>void){
  doc.on('update',notify);
  return ()=>doc.off('update',notify);
}
