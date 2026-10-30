import type {ResultPages} from '@mydoc/agent-contracts';
/** Run-local bounded artifacts; durable/account-scoped storage is supplied by the host in A1. */
export class MemoryResultPages implements ResultPages {
 private readonly values=new Map<string,string>();
 private size=0;
 constructor(private readonly maxChars=1_000_000){}
 async put(id:string,text:string){
  if(this.values.has(id))throw new Error('Duplicate result artifact');
  if(this.size+text.length>this.maxChars)throw new Error('Result artifact budget exceeded');
  this.values.set(id,text);this.size+=text.length;
 }
 async read(id:string,offset:number,limit:number){
  if(!Number.isSafeInteger(offset)||offset<0||!Number.isSafeInteger(limit)||limit<1||limit>8000)throw new Error('Invalid result page');
  const value=this.values.get(id);if(value===undefined)throw new Error('Result artifact not found');
  return {text:value.slice(offset,offset+limit),nextOffset:offset+limit<value.length?offset+limit:null};
 }
}
