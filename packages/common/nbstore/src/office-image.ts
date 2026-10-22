import {Y} from '../../crdt/src/index';
import type {OfficeKind} from './office';
export const OFFICE_IMAGE_BYTES=5*1024*1024,OFFICE_MEDIA_BYTES=8*1024*1024,OFFICE_IMAGE_COUNT=1000;
const crcTable=Uint32Array.from({length:256},(_,value)=>{let crc=value;for(let bit=0;bit<8;bit++)crc=crc&1?0xedb88320^(crc>>>1):crc>>>1;return crc>>>0;});
function pngCrc(bytes:Uint8Array){let crc=0xffffffff;for(const value of bytes)crc=crcTable[(crc^value)&255]^(crc>>>8);return (crc^0xffffffff)>>>0;}
export function officeImageMime(bytes:Uint8Array):'image/png'|'image/jpeg'{
 if(!bytes.length||bytes.length>OFFICE_IMAGE_BYTES)throw new Error('办公图片须为 5 MB 以内的 PNG 或 JPEG');
 if(bytes.length>=24&&[137,80,78,71,13,10,26,10].every((n,i)=>bytes[i]===n)&&String.fromCharCode(...bytes.subarray(12,16))==='IHDR'){
  const view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength),w=view.getUint32(16),h=view.getUint32(20);if(!w||!h||w*h>40000000)throw new Error('图片像素超过 4000 万或尺寸无效');let offset=8,data=false,end=false;while(offset+12<=bytes.length){const size=view.getUint32(offset),type=String.fromCharCode(...bytes.subarray(offset+4,offset+8));if(size>bytes.length-offset-12||view.getUint32(offset+8+size)!==pngCrc(bytes.subarray(offset+4,offset+8+size))||offset===8&&(type!=='IHDR'||size!==13))throw new Error('PNG 图片损坏');offset+=size+12;if(type==='IDAT')data=true;if(type==='IEND'){if(size!==0||offset!==bytes.length)throw new Error('PNG 图片损坏');end=true;break;}}if(!data||!end)throw new Error('PNG 图片不完整');return 'image/png';
 }
 if(bytes[0]===255&&bytes[1]===216){let i=2;while(i+4<=bytes.length){if(bytes[i++]!==255)break;while(bytes[i]===255)i++;const marker=bytes[i++];if(marker===217||marker===218)break;const length=(bytes[i]<<8)|bytes[i+1];if(length<2||i+length>bytes.length)break;if([192,193,194,195,197,198,199,201,202,203,205,206,207].includes(marker)&&length>=8){const h=(bytes[i+3]<<8)|bytes[i+4],w=(bytes[i+5]<<8)|bytes[i+6];if(!w||!h||w*h>40000000)throw new Error('图片像素超过 4000 万或尺寸无效');return 'image/jpeg';}i+=length;}}
 throw new Error('办公图片仅支持有效 PNG 或 JPEG');
}
export function validOfficeAssetId(value:unknown):value is string{return typeof value==='string'&&/^[a-zA-Z0-9_-]{1,100}$/.test(value);}
export function isWordImage(value:unknown):value is Y.Map<unknown>{return value instanceof Y.Map&&value.get('kind')==='image';}
export function createWordImage(assetId:string,width=320,height=240,alt='图片') {const block=new Y.Map<unknown>();for(const [key,value]of Object.entries({kind:'image',assetId,width,height,alt}))block.set(key,value);return block;}
export function validateWordImage(block:Y.Map<unknown>){if(block._start||block.size!==5||[...block.keys()].some(k=>!['kind','assetId','width','height','alt'].includes(k))||!validOfficeAssetId(block.get('assetId'))||!['width','height'].every(k=>typeof block.get(k)==='number'&&Number.isFinite(block.get(k))&&Number(block.get(k))>0&&Number(block.get(k))<=10000)||typeof block.get('alt')!=='string'||String(block.get('alt')).length>1000)throw new Error('Word 图片字段无效');}
export function officeImageNodes(doc:Y.Doc,kind:OfficeKind){const nodes:Y.Map<unknown>[]=[];if(kind==='word')for(const block of doc.getArray('paragraphs')){if(isWordImage(block))nodes.push(block);}if(kind==='presentation'){
 const add=(objects:unknown)=>{if(objects instanceof Y.Array)for(const object of objects)if(object instanceof Y.Map&&object.get('type')==='image')nodes.push(object);};add(doc.getArray('objects'));for(const slide of doc.getMap('slides').values())if(slide instanceof Y.Map)add(slide.get('objects'));
 }if(nodes.length>OFFICE_IMAGE_COUNT)throw new Error('办公图片超过 1000 个');return nodes;}
