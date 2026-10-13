import {test} from 'node:test';
import assert from 'node:assert/strict';
import {namespaceExecutableUUID} from './mac-identity.mjs';

function thin(little=true){
 const bytes=Buffer.alloc(64,0),write=(value,offset)=>little?bytes.writeUInt32LE(value,offset):bytes.writeUInt32BE(value,offset);
 write(0xfeedfacf,0);write(1,16);write(24,20);write(0x1b,32);write(24,36);bytes.fill(0xab,40,56);bytes.fill(0xcd,56);
 return bytes;
}
test('UUID is distinct from Electron and other apps, with reproducible builds',()=>{
 for(const little of [true,false]){
  const source=thin(little),copy=Buffer.from(source),first=namespaceExecutableUUID(source,'com.mydoc.repository/main');
  assert.notDeepEqual(first.subarray(40,56),source.subarray(40,56));
  assert.deepEqual(first,namespaceExecutableUUID(source,'com.mydoc.repository/main'));
  assert.notDeepEqual(first,namespaceExecutableUUID(source,'another.app/main'));
  assert.notDeepEqual(first,namespaceExecutableUUID(source,'com.mydoc.repository/helper'));
  assert.deepEqual(first.subarray(0,40),source.subarray(0,40));assert.deepEqual(first.subarray(56),source.subarray(56));assert.deepEqual(source,copy);
 }
});
test('universal executables namespace each architecture without altering payloads',()=>{
 for(const wide of [false,true])for(const little of [false,true]){
  const source=Buffer.alloc(256),write=(n,offset)=>little?source.writeUInt32LE(n,offset):source.writeUInt32BE(n,offset);
  source.writeUInt32BE(wide?(little?0xbfbafeca:0xcafebabf):(little?0xbebafeca:0xcafebabe),0);write(2,4);
  for(let i=0;i<2;i++){
   const offset=8+i*(wide?32:20),start=128+i*64;
   if(wide){if(little){source.writeBigUInt64LE(BigInt(start),offset+8);source.writeBigUInt64LE(64n,offset+16);}else{source.writeBigUInt64BE(BigInt(start),offset+8);source.writeBigUInt64BE(64n,offset+16);}}
   else{write(start,offset+8);write(64,offset+12);}
   const slice=thin(i===0);slice.fill(i+1,40,56);slice.copy(source,start);
  }
  const result=namespaceExecutableUUID(source,'com.mydoc.repository/main');
  const restored=Buffer.from(result);
  for(const start of [128,192]){assert.notDeepEqual(result.subarray(start+40,start+56),source.subarray(start+40,start+56));source.copy(restored,start+40,start+40,start+56);}
  assert.deepEqual(restored,source);
 }
});
test('reject missing UUID and malformed commands instead of producing a damaged executable',()=>{
 assert.throws(()=>namespaceExecutableUUID(Buffer.alloc(3),'app'),/Truncated/);
 const missing=thin();missing.writeUInt32LE(0x19,32);assert.throws(()=>namespaceExecutableUUID(missing,'app'),/LC_UUID/);
 const broken=thin();broken.writeUInt32LE(0xffffffff,36);assert.throws(()=>namespaceExecutableUUID(broken,'app'),/size/);
 const short=thin().subarray(0,50);assert.throws(()=>namespaceExecutableUUID(short,'app'),/load commands/);
});
