import type {Doc} from 'yjs';import {Awareness,applyAwarenessUpdate} from 'y-protocols/awareness';import * as encoding from 'lib0/encoding';
const instances=new WeakMap<Doc,Awareness>();
export function documentAwareness(doc:Doc){return instances.get(doc);}
export function createDocumentAwareness(doc:Doc){const awareness=new Awareness(doc);instances.set(doc,awareness);return awareness;}
export function applyPresence(awareness:Awareness,client:number,clock:number,state:unknown){const encoder=encoding.createEncoder();encoding.writeVarUint(encoder,1);encoding.writeVarUint(encoder,client);encoding.writeVarUint(encoder,clock);encoding.writeVarString(encoder,JSON.stringify(state));applyAwarenessUpdate(awareness,encoding.toUint8Array(encoder),'remote');}
