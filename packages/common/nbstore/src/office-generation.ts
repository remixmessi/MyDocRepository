/** Queue envelope stored in the existing binary payload column; snapshots remain ordinary Yjs updates. */
const magic=new TextEncoder().encode('MYDOC-OFFICE-UPDATE:1\n');
export function packOfficeUpdate(update:Uint8Array,generation?:string){
 if(!generation)return update;
 if(!/^[a-zA-Z0-9-]{1,100}$/.test(generation))throw new Error('办公资源代际无效');
 const prefix=new TextEncoder().encode(generation+'\n'),bytes=new Uint8Array(magic.length+prefix.length+update.length);bytes.set(magic);bytes.set(prefix,magic.length);bytes.set(update,magic.length+prefix.length);return bytes;
}
export function unpackOfficeUpdate(bytes:Uint8Array){
 if(!magic.every((n,i)=>bytes[i]===n))return {update:bytes,generation:undefined};
 const end=bytes.indexOf(10,magic.length);if(end<magic.length+1||end>magic.length+100)throw new Error('办公更新代际封包无效');
 const generation=new TextDecoder('utf-8',{fatal:true}).decode(bytes.subarray(magic.length,end));if(!/^[a-zA-Z0-9-]{1,100}$/.test(generation))throw new Error('办公更新代际封包无效');return {update:bytes.subarray(end+1),generation};
}
export function officeGenerationMatches(current:string,supplied:unknown){return typeof supplied==='string'?supplied===current:!current.startsWith('epoch-');}
