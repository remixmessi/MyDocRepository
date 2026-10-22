import {Y} from '@mydoc/crdt';
import {wordTextIdentity} from './office-word-flow';
import {captureWordFlowSelection,resolveWordFlowSelection,type WordFlowSelection} from './word-flow-selection';
import type {WordBlock} from '../../../common/nbstore/src/office-word-table';
export function wordDomPoint(element:HTMLElement,offset:number):[Node,number]{const walker=document.createTreeWalker(element,NodeFilter.SHOW_TEXT);let node:Node|null;while(node=walker.nextNode()){const length=node.textContent?.length??0;if(offset<=length)return[node,offset];offset-=length;}return [element,0];}
export function readWordFlowDom(root:HTMLElement,doc:Y.Doc){
 const selection=window.getSelection();if(!selection?.rangeCount)return;const range=selection.getRangeAt(0);
 const endpoint=(node:Node,offset:number)=>{const element=(node instanceof HTMLElement?node:node.parentElement)?.closest<HTMLElement>('.word-rich-input');if(!element||!root.contains(element))return;const text=doc.getArray<WordBlock>('paragraphs').toArray().find((text):text is Y.Text=>text instanceof Y.Text&&wordTextIdentity(text)===element.dataset.wordText);if(!text)return;const before=document.createRange();before.selectNodeContents(element);before.setEnd(node,offset);return {text,offset:before.toString().length};};
 const a=endpoint(range.startContainer,range.startOffset),b=endpoint(range.endContainer,range.endOffset);if(!a||!b||a.text===b.text)return;
 return captureWordFlowSelection(a.text,a.offset,b.text,b.offset);
}
export function restoreWordFlowDom(root:HTMLElement,doc:Y.Doc,selection:WordFlowSelection){
 const {ranges}=resolveWordFlowSelection(doc,selection),first=ranges[0],last=ranges[ranges.length-1],a=root.querySelector<HTMLElement>(`[data-word-text="${wordTextIdentity(first.text)}"]`),b=root.querySelector<HTMLElement>(`[data-word-text="${wordTextIdentity(last.text)}"]`);if(!a||!b)return;
 const range=document.createRange();range.setStart(...wordDomPoint(a,first.start));range.setEnd(...wordDomPoint(b,last.end));const dom=window.getSelection();dom?.removeAllRanges();dom?.addRange(range);
}
