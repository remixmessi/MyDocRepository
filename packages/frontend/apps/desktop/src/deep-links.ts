import {randomUUID} from 'node:crypto';
import {parseDocumentLink,type DocumentLocation} from '../../../../common/nbstore/src/deep-link';
export interface PendingDocumentLink {id:string;target?:DocumentLocation;error?:string;}
export class DocumentLinkQueue {
 private requests:PendingDocumentLink[]=[];
 enqueue(url:string){if(this.requests.length>=32)return false;const target=parseDocumentLink(url);if(target&&this.requests.some(request=>JSON.stringify(request.target)===JSON.stringify(target)))return true;this.requests.push({id:randomUUID(),...(target?{target}:{error:'桌面文档链接无效，请重新复制链接。'})});return true;}
 take(){return this.requests[0];}
 acknowledge(id:unknown){if(this.requests[0]?.id===id)this.requests.shift();}
 get pending(){return this.requests.length;}
}
export function commandLineLinks(argv:string[]){return argv.filter(arg=>/^mydoc:/i.test(arg));}
