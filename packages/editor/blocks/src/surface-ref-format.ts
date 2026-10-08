import {validRegion} from './surface-region';
import type {Bounds} from './whiteboard/model';
export interface SurfaceRefAttrs {docId:string;workspaceId:string;frameId:string;label:string;collapsed:boolean;region?:Bounds|null}
export function validSurfaceRef(value:unknown):value is SurfaceRefAttrs{if(!value||typeof value!=='object')return false;const s=value as SurfaceRefAttrs;return ['docId','workspaceId','frameId'].every(key=>{const v=s[key as 'docId'];return typeof v==='string'&&!!v&&v.length<=200;})&&typeof s.label==='string'&&s.label.length<=500&&typeof s.collapsed==='boolean'&&(s.region===undefined||s.region===null||validRegion(s.region));}
export function parseSurfaceRef(text:string):SurfaceRefAttrs|null{try{const s:unknown=JSON.parse(text);return validSurfaceRef(s)?s:null;}catch{return null;}}
